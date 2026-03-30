VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmEInvoice 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·›« Ê—… «·÷—Ì»Ì…"
   ClientHeight    =   9525
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   16200
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   580
   Icon            =   "FrmEInvoice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9525
   ScaleWidth      =   16200
   Visible         =   0   'False
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   9525
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   16200
      _cx             =   28575
      _cy             =   16801
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
      BorderWidth     =   2
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
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   8340
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   16140
         _cx             =   28469
         _cy             =   14711
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
         Caption         =   "."
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7920
            Index           =   2
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   16050
            _cx             =   28310
            _cy             =   13970
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   765
               Index           =   5
               Left            =   0
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   0
               Width           =   15855
               _cx             =   27966
               _cy             =   1349
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   24
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
               Picture         =   "FrmEInvoice.frx":038A
               Caption         =   "«·›« Ê—… «·÷—Ì»Ì…"
               Align           =   0
               AutoSizeChildren=   0
               BorderWidth     =   0
               ChildSpacing    =   0
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
               PicturePos      =   0
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
               Begin VB.TextBox txtPassword 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  IMEMode         =   3  'DISABLE
                  Left            =   3600
                  PasswordChar    =   "."
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   90
                  Width           =   405
               End
               Begin ImpulseButton.ISButton XPBtnMove 
                  Height          =   375
                  Index           =   0
                  Left            =   1695
                  TabIndex        =   23
                  Top             =   90
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmEInvoice.frx":1064
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
                  Height          =   375
                  Index           =   2
                  Left            =   630
                  TabIndex        =   24
                  Top             =   90
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmEInvoice.frx":13FE
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
                  Height          =   375
                  Index           =   1
                  Left            =   2220
                  TabIndex        =   25
                  Top             =   90
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmEInvoice.frx":1798
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
                  Height          =   375
                  Index           =   3
                  Left            =   1155
                  TabIndex        =   26
                  Top             =   90
                  Width           =   495
                  _ExtentX        =   873
                  _ExtentY        =   661
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
                  ButtonImage     =   "FrmEInvoice.frx":1B32
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
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   7905
               Index           =   1
               Left            =   0
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   -120
               Width           =   17295
               _cx             =   30506
               _cy             =   13944
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
               Begin VB.Frame Frame2 
                  BackColor       =   &H00E2E9E9&
                  Height          =   3840
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   27
                  Top             =   840
                  Width           =   15990
                  Begin VB.TextBox txtBranch_Code 
                     Alignment       =   1  'Right Justify
                     Height          =   345
                     Left            =   6570
                     RightToLeft     =   -1  'True
                     TabIndex        =   80
                     Top             =   210
                     Width           =   645
                  End
                  Begin VB.PictureBox Picture1 
                     Height          =   2145
                     Left            =   0
                     ScaleHeight     =   2085
                     ScaleWidth      =   2025
                     TabIndex        =   77
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   2085
                  End
                  Begin VB.TextBox txtManualInvoiceNo 
                     Alignment       =   2  'Center
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   9810
                     RightToLeft     =   -1  'True
                     TabIndex        =   74
                     Top             =   180
                     Width           =   1440
                  End
                  Begin VB.OptionButton ComResid 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Œ«÷⁄"
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Index           =   1
                     Left            =   1740
                     RightToLeft     =   -1  'True
                     TabIndex        =   73
                     Top             =   270
                     Width           =   975
                  End
                  Begin VB.OptionButton ComResid 
                     Alignment       =   1  'Right Justify
                     Caption         =   "€Ì— Œ«÷⁄"
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Index           =   0
                     Left            =   2730
                     RightToLeft     =   -1  'True
                     TabIndex        =   72
                     Top             =   270
                     Width           =   1095
                  End
                  Begin VB.TextBox txtGroupUniqueFileMaster 
                     Alignment       =   2  'Center
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   150
                     RightToLeft     =   -1  'True
                     TabIndex        =   70
                     Top             =   2940
                     Width           =   3285
                  End
                  Begin VB.TextBox TXTNewNO 
                     Alignment       =   2  'Center
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   150
                     RightToLeft     =   -1  'True
                     TabIndex        =   68
                     Top             =   2460
                     Width           =   3285
                  End
                  Begin VB.TextBox txtIdentificationid 
                     Alignment       =   2  'Center
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   150
                     RightToLeft     =   -1  'True
                     TabIndex        =   66
                     Top             =   1980
                     Width           =   3285
                  End
                  Begin VB.TextBox txtErrorMessageS 
                     Alignment       =   2  'Center
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   150
                     RightToLeft     =   -1  'True
                     TabIndex        =   64
                     Top             =   1500
                     Width           =   3285
                  End
                  Begin VB.TextBox txtwarrningmessage 
                     Alignment       =   2  'Center
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   62
                     Top             =   990
                     Width           =   3285
                  End
                  Begin VB.TextBox txtzatcaStatus 
                     Alignment       =   2  'Center
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   60
                     Top             =   510
                     Width           =   3285
                  End
                  Begin VB.TextBox txtIqarName 
                     Alignment       =   2  'Center
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   6000
                     RightToLeft     =   -1  'True
                     TabIndex        =   58
                     Top             =   3060
                     Width           =   3285
                  End
                  Begin VB.TextBox Text9 
                     Alignment       =   2  'Center
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   6000
                     RightToLeft     =   -1  'True
                     TabIndex        =   56
                     Top             =   2580
                     Visible         =   0   'False
                     Width           =   3285
                  End
                  Begin VB.TextBox txtVatValue 
                     Alignment       =   2  'Center
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   6000
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   54
                     Top             =   2100
                     Width           =   3285
                  End
                  Begin VB.TextBox txtPayableAmount 
                     Alignment       =   2  'Center
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   6000
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   52
                     Top             =   1620
                     Width           =   3285
                  End
                  Begin VB.TextBox Text6 
                     Alignment       =   2  'Center
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   6000
                     RightToLeft     =   -1  'True
                     TabIndex        =   50
                     Top             =   1110
                     Visible         =   0   'False
                     Width           =   3285
                  End
                  Begin VB.TextBox txtCompanyID 
                     Alignment       =   2  'Center
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   6000
                     RightToLeft     =   -1  'True
                     TabIndex        =   48
                     Top             =   630
                     Width           =   3285
                  End
                  Begin VB.TextBox txtCitySubdivisionName 
                     Alignment       =   2  'Center
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   10920
                     RightToLeft     =   -1  'True
                     TabIndex        =   46
                     Top             =   3120
                     Width           =   3285
                  End
                  Begin VB.TextBox txtPostalZone 
                     Alignment       =   2  'Center
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   10920
                     RightToLeft     =   -1  'True
                     TabIndex        =   44
                     Top             =   2640
                     Width           =   3285
                  End
                  Begin VB.TextBox txtBuildingNumber 
                     Alignment       =   2  'Center
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   10920
                     RightToLeft     =   -1  'True
                     TabIndex        =   42
                     Top             =   2160
                     Width           =   3285
                  End
                  Begin VB.TextBox txtStreetName 
                     Alignment       =   2  'Center
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   10920
                     RightToLeft     =   -1  'True
                     TabIndex        =   40
                     Top             =   1680
                     Width           =   3285
                  End
                  Begin VB.TextBox txtCityName 
                     Alignment       =   2  'Center
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   10890
                     RightToLeft     =   -1  'True
                     TabIndex        =   38
                     Top             =   1170
                     Width           =   3285
                  End
                  Begin VB.TextBox txtRegistrationName 
                     Alignment       =   2  'Center
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   10890
                     RightToLeft     =   -1  'True
                     TabIndex        =   35
                     Top             =   690
                     Width           =   3285
                  End
                  Begin VB.TextBox txtInvoiceID 
                     Alignment       =   2  'Center
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   330
                     Left            =   12420
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   29
                     Top             =   210
                     Width           =   1740
                  End
                  Begin VB.TextBox txtRemarks 
                     Alignment       =   1  'Right Justify
                     BeginProperty Font 
                        Name            =   "Times New Roman"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   435
                     Left            =   150
                     MultiLine       =   -1  'True
                     RightToLeft     =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   28
                     Top             =   3300
                     Width           =   4290
                  End
                  Begin MSComCtl2.DTPicker txtIssueDate 
                     Height          =   315
                     Left            =   3900
                     TabIndex        =   30
                     Top             =   180
                     Width           =   1575
                     _ExtentX        =   2778
                     _ExtentY        =   556
                     _Version        =   393216
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Format          =   112066561
                     CurrentDate     =   38784
                  End
                  Begin MSDataListLib.DataCombo DcBranches 
                     Height          =   315
                     Left            =   7200
                     TabIndex        =   78
                     Top             =   210
                     Width           =   2115
                     _ExtentX        =   3731
                     _ExtentY        =   582
                     _Version        =   393216
                     Text            =   ""
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·›—⁄"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   23
                     Left            =   9090
                     RightToLeft     =   -1  'True
                     TabIndex        =   79
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·—ﬁ„ «·„—Ã⁄Ì"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   270
                     Index           =   2
                     Left            =   11280
                     RightToLeft     =   -1  'True
                     TabIndex        =   75
                     Top             =   240
                     Width           =   1065
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "GroupUniqueFileMaster"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   22
                     Left            =   3960
                     RightToLeft     =   -1  'True
                     TabIndex        =   71
                     Top             =   3030
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "—ﬁ„ «·⁄ﬁœ «·„ÊÕœ"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   21
                     Left            =   3690
                     RightToLeft     =   -1  'True
                     TabIndex        =   69
                     Top             =   2520
                     Width           =   1140
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·”Ã·"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   20
                     Left            =   3960
                     RightToLeft     =   -1  'True
                     TabIndex        =   67
                     Top             =   2070
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   19
                     Left            =   3960
                     RightToLeft     =   -1  'True
                     TabIndex        =   65
                     Top             =   1560
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ⁄·Ì„« "
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   18
                     Left            =   3930
                     RightToLeft     =   -1  'True
                     TabIndex        =   63
                     Top             =   1080
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ«·… «·›« Ê—…"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   17
                     Left            =   3930
                     RightToLeft     =   -1  'True
                     TabIndex        =   61
                     Top             =   570
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·⁄ﬁ«—"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   16
                     Left            =   9810
                     RightToLeft     =   -1  'True
                     TabIndex        =   59
                     Top             =   3150
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·⁄„Ì·"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   15
                     Left            =   9810
                     RightToLeft     =   -1  'True
                     TabIndex        =   57
                     Top             =   2640
                     Visible         =   0   'False
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·÷—Ì»…"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   14
                     Left            =   9810
                     RightToLeft     =   -1  'True
                     TabIndex        =   55
                     Top             =   2190
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·«Ã„«·Ì"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   13
                     Left            =   9810
                     RightToLeft     =   -1  'True
                     TabIndex        =   53
                     Top             =   1680
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„œÌ‰…"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   12
                     Left            =   9780
                     RightToLeft     =   -1  'True
                     TabIndex        =   51
                     Top             =   1200
                     Visible         =   0   'False
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·—ﬁ„ «·÷—Ì»Ì"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   11
                     Left            =   9780
                     RightToLeft     =   -1  'True
                     TabIndex        =   49
                     Top             =   690
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·⁄‰Ê«‰"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   10
                     Left            =   14730
                     RightToLeft     =   -1  'True
                     TabIndex        =   47
                     Top             =   3210
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "PostalZone"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   9
                     Left            =   14730
                     RightToLeft     =   -1  'True
                     TabIndex        =   45
                     Top             =   2700
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„»‰Ï"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   8
                     Left            =   14730
                     RightToLeft     =   -1  'True
                     TabIndex        =   43
                     Top             =   2250
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·‘«—⁄"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   6
                     Left            =   14730
                     RightToLeft     =   -1  'True
                     TabIndex        =   41
                     Top             =   1740
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„œÌ‰…"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   4
                     Left            =   14700
                     RightToLeft     =   -1  'True
                     TabIndex        =   39
                     Top             =   1260
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "—ﬁ„ «·›« Ê—…"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   270
                     Index           =   7
                     Left            =   14520
                     RightToLeft     =   -1  'True
                     TabIndex        =   34
                     Top             =   210
                     Width           =   1065
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " «—ÌŒ «·›« Ê—…"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Index           =   5
                     Left            =   5490
                     RightToLeft     =   -1  'True
                     TabIndex        =   33
                     Top             =   240
                     Width           =   1110
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·⁄„Ì·"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   0
                     Left            =   14700
                     RightToLeft     =   -1  'True
                     TabIndex        =   32
                     Top             =   750
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„·«ÕŸ« "
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   3
                     Left            =   4440
                     RightToLeft     =   -1  'True
                     TabIndex        =   31
                     Top             =   3480
                     Width           =   840
                  End
               End
               Begin VB.TextBox txtID 
                  Alignment       =   1  'Right Justify
                  Height          =   420
                  Index           =   0
                  Left            =   -4560
                  RightToLeft     =   -1  'True
                  TabIndex        =   8
                  Top             =   12420
                  Width           =   2610
               End
               Begin VB.TextBox TxtModFlg 
                  Alignment       =   1  'Right Justify
                  Height          =   420
                  Left            =   6675
                  RightToLeft     =   -1  'True
                  TabIndex        =   6
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   2415
               End
               Begin VSFlex8Ctl.VSFlexGrid grd 
                  Height          =   2670
                  Index           =   0
                  Left            =   -390
                  TabIndex        =   37
                  Top             =   4710
                  Width           =   16560
                  _cx             =   29210
                  _cy             =   4710
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
                  AllowUserResizing=   1
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   50
                  Cols            =   83
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   320
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmEInvoice.frx":1ECC
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
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   495
                  Left            =   15720
                  RightToLeft     =   -1  'True
                  TabIndex        =   7
                  Top             =   1275
                  Width           =   945
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„ÊŸ›"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   8400
               RightToLeft     =   -1  'True
               TabIndex        =   3
               Top             =   90
               Width           =   1125
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   1110
         Left            =   30
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   8385
         Width           =   16140
         _cx             =   28469
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
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   330
            Left            =   13080
            TabIndex        =   10
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   0
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
            ButtonImage     =   "FrmEInvoice.frx":2DD0
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   330
            Left            =   13320
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   120
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
            ButtonImage     =   "FrmEInvoice.frx":316A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   0
            Left            =   10680
            TabIndex        =   14
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   1
            Left            =   9360
            TabIndex        =   15
            Top             =   480
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   2
            Left            =   7920
            TabIndex        =   16
            Top             =   480
            Width           =   1245
            _ExtentX        =   2196
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            CausesValidation=   0   'False
            Height          =   375
            Index           =   3
            Left            =   6360
            TabIndex        =   17
            Top             =   480
            Width           =   1365
            _ExtentX        =   2408
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   435
            Index           =   4
            Left            =   4920
            TabIndex        =   18
            Top             =   480
            Width           =   1365
            _ExtentX        =   2408
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            CausesValidation=   0   'False
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   19
            Top             =   480
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   661
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   5
            Left            =   3480
            TabIndex        =   20
            Top             =   480
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   661
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ALLButtonS.ALLButton CmdRemove 
            Height          =   375
            Left            =   11160
            TabIndex        =   21
            Tag             =   "Delete Row"
            Top             =   0
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Õ–› ”ÿ—"
            ENAB            =   0   'False
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
            BCOL            =   12632319
            BCOLO           =   12632319
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmEInvoice.frx":3504
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   405
            Left            =   1800
            TabIndex        =   36
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   480
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
            ButtonImage     =   "FrmEInvoice.frx":3520
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label LabCountRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   225
            Width           =   1740
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   240
            Width           =   1515
         End
      End
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   345
      Left            =   3360
      TabIndex        =   4
      Top             =   6840
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
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
      ButtonImage     =   "FrmEInvoice.frx":9D82
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmEInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSearchDCombo As clsDCboSearch
Dim BKGrndPic As ClsBackGroundPic
Dim net_value As Double
Dim net_value1 As Double
Dim My_SQL  As String
Dim StrSQL  As String
Dim rs As ADODB.Recordset

Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long
Private Sub ChkDetails_Click()
    FillGridWithData
End Sub
Private Sub ALLButton1_Click()
    FrmShowCol1.show
End Sub

Function check_previous_dev(year As String, Month As String) As Boolean
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from notes where salary=" & year & Month
 
    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs2.RecordCount = 0 Then
        check_previous_dev = False
    Else
        check_previous_dev = True
    End If
 
End Function

Function check_previous_dev1(year As String, Month As String) As Boolean
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from salary_voucher where m_year='" & year & "' and m_month='" & Month & "'"
 
    rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs2.RecordCount = 0 Then
        check_previous_dev1 = False
    Else
        check_previous_dev1 = True
    End If
 
End Function

Function Create_dev()
 
  End Function

Function Create_dev1()
  
End Function

Function create_report_data()

End Function


Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long

    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
 
        
 
    End If

    '-------------------------------------------------------------------------------------------
    Dim IsFromUser As Integer
    Cn.BeginTrans
    BeginTrans = True
    
    If TxtModFlg.text = "N" Then
        rs.AddNew
        rs!IsFromUser = 1
        IsFromUser = 1
'        rs("id").value = 5
        rs!Transaction_ID = val(CStr(new_id("tblEInvoice", "Transaction_ID", "", True)))
        If txtInvoiceID.text = "" Then
            txtInvoiceID.text = Voucher_coding(1, txtIssueDate.value, 86, 5003, , , , , , , "")
        End If
        NewInvoice
        Dim mm As String
        mm = GenerateGUID
        rs!GroupUniqueCode = "{" & mm & "}"
        rs!GroupUniqueFileMaster = rs!GroupUniqueCode
    ElseIf Me.TxtModFlg.text = "E" Then
        IsFromUser = val(rs!IsFromUser & "")
        If IsFromUser = 1 Then
             Cn.Execute "delete tblEInvoice2 where invoiceid='" & Trim(Me.txtInvoiceID.text) & "'"
        End If
   
    End If
    
    'rs("id").value = txtID(0).text
    If txtInvoiceID.text = "" Then
        NewInvoice
    End If
    rs("InvoiceID").value = IIf(txtInvoiceID.text <> "", Trim(txtInvoiceID.text), Null)
    rs("RegistrationName").value = IIf(txtRegistrationName.text <> "", Trim(txtRegistrationName.text), Null)
    rs("CityName").value = IIf(txtCityName.text <> "", Trim(txtCityName.text), Null)
    rs("StreetName").value = IIf(txtstreetname.text <> "", Trim(txtstreetname.text), Null)
    rs("BuildingNumber").value = IIf(txtBuildingNumber.text <> "", Trim(txtBuildingNumber.text), Null)
    rs("PostalZone").value = IIf(txtPostalZone.text <> "", Trim(txtPostalZone.text), Null)
    rs("CitySubdivisionName").value = IIf(txtCitySubdivisionName.text <> "", Trim(txtCitySubdivisionName.text), Null)
    rs("CompanyID").value = IIf(txtCompanyID.text <> "", Trim(txtCompanyID.text), Null)
    rs("Identificationid").value = IIf(txtIdentificationid.text <> "", Trim(txtIdentificationid.text), Null)
    rs("ManualInvoiceNo").value = IIf(txtManualInvoiceNo.text <> "", Trim(txtManualInvoiceNo.text), Null)
    rs("IqarName").value = IIf(txtIqarName.text <> "", Trim(txtIqarName.text), Null)
    rs("branch_id").value = val(DcBranches.BoundText)


    rs("PayableAmount").value = IIf(txtPayableAmount.text <> "", Trim(txtPayableAmount.text), Null)
    
    rs("VatValue").value = IIf(TxtVATValue.text <> "", Trim(TxtVATValue.text), Null)
    

    
  '  rs("TxtVendorCode").value = IIf(TxtSearchCode.text <> "", Trim(TxtSearchCode.text), Null)
    rs("IssueDate").value = txtIssueDate.value
    
      
          If ComResid(1).value = True Then
            rs.Fields("ComResid").value = 1
        Else
            rs.Fields("ComResid").value = 0
        End If
        rs("NewNO").value = IIf(TXTNewNO.text = "", Null, TXTNewNO.text)
        
      
      
      

      
      


    rs.update
    CuurentLogdata
    Dim s As String
    Set RsDev = New ADODB.Recordset
    s = "Select * from tblEInvoice2 where invoiceid='" & Trim(Me.txtInvoiceID.text) & "'"
    RsDev.Open s, Cn, adOpenKeyset, adLockOptimistic
    'RsDev.Open "tblEInvoice2", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
    Dim i As Integer

    With Me.grd(0)

        For i = 1 To .rows - 1

            
                If IsFromUser = 1 Then
                    RsDev.AddNew
                    RsDev("InvoiceID").value = Trim(txtInvoiceID)
                End If
                
                RsDev!Transaction_ID = rs!Transaction_ID
                
                RsDev("Qty").value = val(.TextMatrix(i, .ColIndex("Qty")))
                RsDev("ItemName").value = Trim(.TextMatrix(i, .ColIndex("ItemName")))
                RsDev("Price").value = val(.TextMatrix(i, .ColIndex("Price")))
                RsDev("PayableAmount").value = val(.TextMatrix(i, .ColIndex("PayableAmount")))
                RsDev("VatValue").value = val(.TextMatrix(i, .ColIndex("VatValue")))
                
                
                RsDev.update
                    
            
            
            '
        Next i

        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg
            
    End With
 
    Cn.CommitTrans
    BeginTrans = False
 
    Select Case Me.TxtModFlg.text

        Case "N"
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

            '    Fg_Journal.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

        Case "E"
            MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '  Fg_Journal.Enabled = False
    End Select

    TxtModFlg.text = "R"
    'End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0
     
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            grd(0).rows = 1
            grd(0).rows = 2
            clear_all Me
      '  Me.TxtTblVendorContractD.text = CStr(new_id("tblEInvoice", "tblEInvoiceD", "", True))
       
      

        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"
            '         Grid.Rows = Grid.Rows + 1
            grd(0).Enabled = True
         
            CuurentLogdata

        Case 2
    
            SaveData
           
        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If
            FrmVendorContractSearch.mIndex = 1
            Load FrmVendorContractSearch
            FrmVendorContractSearch.show vbModal

        Case 6
            Unload Me

        Case 7

            '   ViewDataList
        Case 20
            

        Case 21
           
    End Select

    Exit Sub
ErrTrap:

End Sub

Private Sub Del_Trans()
    
    Dim Msg  As String

End Sub


Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
 
            Retrive
            Me.TxtModFlg.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub ComResid_Click(Index As Integer)
If Me.TxtModFlg <> "E" Then Exit Sub
    Dim i As Long
    For i = 1 To grd(0).rows - 1
        grd_AfterEdit 0, i, grd(0).ColIndex("Price")
    Next
End Sub

Private Sub DcBranches_Change()
Dim bid As Long
    bid = val(DcBranches.BoundText & "")
    If bid > 0 Then
        txtBranch_Code.text = GetBranchCode(bid)
    Else
        txtBranch_Code.text = ""
    End If
End Sub

Private Sub Form_Load()

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    ScreenNameArabic = " «·›« Ê—… «·÷—Ì»Ì…  "
    ScreenNameEnglish = "  Supplier Contracts "
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    'Set CmdHelp.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Help").Picture
    'Set Cmd(7).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("FillData").Picture
    Dim My_SQL As String

    Dim GrdBack As ClsBackGroundPic
    Set GrdBack = New ClsBackGroundPic

    With Me.grd(0)
        Set .WallPaper = GrdBack.Picture
     
    End With

    'My_SQL = " select id,Project_name from projects"
    'fill_combo dcproject, My_SQL
    '
    'My_SQL = " select  fullcode,des from projects_des"
    'fill_combo Dcterm, My_SQL

    'My_SQL = " select  fullcode,name from terms_operations"
    'fill_combo dcopr, My_SQL

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
     
    
      Dcombos.GetBranches Me.DcBranches
    Set BKGrndPic = New ClsBackGroundPic

 
    With Me.grd(0)
        .rows = 1
        .ExplorerBar = flexExSortShowAndMove
        .RowHeightMin = 300
        .ExtendLastCol = True
    End With
      
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set rs = New ADODB.Recordset
    StrSQL = "select * From tblEInvoice  "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

End Sub

Private Sub ChangeLang()

    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    ISButton2.Caption = "Print"
    Cmd(6).Caption = "Exit"
    'CmdHelp.Caption = "Help"

    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = "Supplier Contract"
    Ele(5).Caption = Me.Caption
    lbl(7).Caption = "ID"
    lbl(5).Caption = "Start Date"
    lbl(2).Caption = "End Date"
    lbl(0).Caption = "Supplier"
    lbl(3).Caption = "Remarks"

End Sub

Public Sub get_all_employee()
    
End Sub

Public Sub FillGridWithData()
    Exit Sub

   
End Sub
 
Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

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
     
                SaveData

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

Function CuurentLogdata(Optional Currentmode As String)
   
    
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish
End Sub



Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer

    With Me.grd(0)

        For i = .FixedRows To .rows - 1
    
            If .TextMatrix(i, .ColIndex("ItemId")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i
   
    End With

End Sub

' ÷⁄ «·œ«·… ›Ì «·›Ê—„ √Ê „ÊœÊ· ⁄«„
Private Function GetBranchCode(ByVal BranchID As Long) As String
    On Error GoTo eh
    Dim rs As New ADODB.Recordset
    Dim sql As String

    sql = "SELECT branch_Code FROM TblBranchesData WHERE branch_id = " & BranchID
    rs.Open sql, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText

    If Not rs.EOF Then
        If IsNull(rs!branch_Code) Then
            GetBranchCode = ""
        Else
            GetBranchCode = CStr(rs!branch_Code)
        End If
    Else
        GetBranchCode = ""
    End If

    rs.Close
    Set rs = Nothing
    Exit Function
eh:
    GetBranchCode = ""
End Function
Private Sub DcBranches_Click(Area As Integer)
   txtInvoiceID.text = ""
    
    Dim bid As Long
    bid = val(DcBranches.BoundText & "")
    If bid > 0 Then
        txtBranch_Code.text = GetBranchCode(bid)
    Else
        txtBranch_Code.text = ""
    End If
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap
    grd(0).Clear flexClearScrollable, flexClearEverything
    grd(0).rows = 1
          
    If rs.RecordCount < 1 Then
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

    End If
    txtid(0) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
    txtInvoiceID = IIf(IsNull(rs("InvoiceID").value), "", rs("InvoiceID").value)
    txtRegistrationName = IIf(IsNull(rs("RegistrationName").value), "", rs("RegistrationName").value)
    txtCityName = IIf(IsNull(rs("CityName").value), "", rs("CityName").value)
    txtManualInvoiceNo = IIf(IsNull(rs("ManualInvoiceNo").value), "", rs("ManualInvoiceNo").value)
    txtIqarName = IIf(IsNull(rs("IqarName").value), "", rs("IqarName").value)
    DcBranches.BoundText = val(rs("branch_id").value & "")
    txtstreetname = IIf(IsNull(rs("StreetName").value), "", rs("StreetName").value)
    txtBuildingNumber = IIf(IsNull(rs("BuildingNumber").value), "", rs("BuildingNumber").value)
    txtPostalZone = IIf(IsNull(rs("PostalZone").value), "", rs("PostalZone").value)
    txtCitySubdivisionName = IIf(IsNull(rs("CitySubdivisionName").value), "", rs("CitySubdivisionName").value)
    txtCompanyID = IIf(IsNull(rs("CompanyID").value), "", rs("CompanyID").value)
    txtPayableAmount = IIf(IsNull(rs("PayableAmount").value), "", rs("PayableAmount").value)
    TxtVATValue = IIf(IsNull(rs("VatValue").value), "", rs("VatValue").value)
    txtzatcaStatus = IIf(IsNull(rs("zatcaStatus").value), "", rs("zatcaStatus").value)
    txtwarrningmessage = IIf(IsNull(rs("warrningmessage").value), "", rs("warrningmessage").value)
    txtErrorMessageS = IIf(IsNull(rs("ErrorMessageS").value), "", rs("ErrorMessageS").value)
    txtwarrningmessage = IIf(IsNull(rs("warrningmessage").value), "", rs("warrningmessage").value)
    txtGroupUniqueFileMaster = IIf(IsNull(rs("GroupUniqueFileMaster").value), "", rs("GroupUniqueFileMaster").value)
    txtIdentificationid = IIf(IsNull(rs("Identificationid").value), "", rs("Identificationid").value)
        
    If Not IsNull(rs.Fields("ComResid").value) Then
   If rs.Fields("ComResid").value = 1 Then
   ComResid(1).value = True
   Else
   ComResid(0).value = True
   End If
   Else
   ComResid(1).value = True
   End If
   TXTNewNO.text = IIf(IsNull(rs("NewNO").value), "", rs("NewNO").value)
    

 
    txtIssueDate.value = IIf(IsNull(rs("IssueDate").value), Date, rs("IssueDate").value)
   ' TxtSearchCode.text = IIf(IsNull(rs("TxtVendorCode").value), "", rs("TxtVendorCode").value)
   
'    txtRemarks.text = IIf(IsNull(rs("Remarks").value), 0, rs("Remarks").value)
    

  

    Dim s
    s = ""
s = s & "SELECT " & vbCrLf
s = s & "    InvoiceID,GroupUniqueFileMaster,GroupUniqueCode," & vbCrLf
s = s & "    InvoiceID as Transaction_ID," & vbCrLf

s = s & "    TotalB =(IsNull(tblEInvoice2.Price,0) * IsNull(tblEInvoice2.Qty,0)),"



s = s & "    CompanyID,ItemName," & vbCrLf
s = s & "    CoCRCode," & vbCrLf
s = s & "    (PayableAmount) AS PayableAmount," & vbCrLf
s = s & "    (VatValue) AS VatValue," & vbCrLf
s = s & "    round((VatValue) /((Price*qty)   ) * 100,1)  AS TaxCategoryPercent ," & vbCrLf
'15' AS TaxCategoryPercent
s = s & "    Id700," & vbCrLf
s = s & "    QrCodeData,Price,Qty," & vbCrLf
s = s & "    QrCodeDataPath," & vbCrLf
s = s & "    zatcaStatus," & vbCrLf
s = s & "    InvoiceTypeCodeID," & vbCrLf
s = s & "    InvoiceTypeCodename," & vbCrLf
s = s & "    AdditionalDocumentReferencePIH," & vbCrLf
s = s & "    InvoiceDocumentReferenceID," & vbCrLf
s = s & "    AdditionalDocumentReferenceICVUUID," & vbCrLf
s = s & "    ActualDeliveryDate," & vbCrLf
s = s & "    LatestDeliveryDate," & vbCrLf
s = s & "    RecTime " & vbCrLf
s = s & "FROM tblEInvoice2 " & vbCrLf
s = s & " where tblEInvoice2.InvoiceID = '" & Trim(Me.txtInvoiceID.text) & "'"


        loadgrid s, grd(0), True, False
         grd(0).rows = grd(0).rows + 1

        If SystemOptions.UserInterface = ArabicInterface Then
            grd(0).TextMatrix(grd(0).rows - 1, grd(0).ColIndex("Ser")) = "«·√Ã„«·Ï"
        Else
            grd(0).TextMatrix(grd(0).rows - 1, grd(0).ColIndex("Ser")) = "Total"
        End If

        grd(0).IsSubtotal(grd(0).rows - 1) = True
        
        Dim SngTotal As Double
        SngTotal = grd(0).Aggregate(flexSTSum, grd(0).FixedRows, grd(0).ColIndex("TotalB"), grd(0).rows - 1, grd(0).ColIndex("TotalB"))
            grd(0).TextMatrix(grd(0).rows - 1, grd(0).ColIndex("TotalB")) = SngTotal
            
        SngTotal = grd(0).Aggregate(flexSTSum, grd(0).FixedRows, grd(0).ColIndex("PayableAmount"), grd(0).rows - 1, grd(0).ColIndex("PayableAmount"))
            grd(0).TextMatrix(grd(0).rows - 1, grd(0).ColIndex("PayableAmount")) = SngTotal
            
        grd(0).cell(flexcpBackColor, grd(0).rows - 1, 1, grd(0).rows - 1, grd(0).Cols - 1) = vbYellow
        grd(0).cell(flexcpFontBold, grd(0).rows - 1, 1, grd(0).rows - 1, grd(0).Cols - 1) = True
        grd(0).cell(flexcpFontSize, grd(0).rows - 1, 1, grd(0).rows - 1, grd(0).Cols - 1) = 10
        grd(0).cell(flexcpFontName, grd(0).rows - 1, 1, grd(0).rows - 1, grd(0).Cols - 1) = "Tahoma"
        grd(0).AutoSize 0, grd(0).Cols - 1, False
 
 
 
  
  grd(0).Subtotal flexSTClear
'  grd(0).Subtotal flexSTSum
  
  
        SngTotal = grd(0).Aggregate(flexSTSum, grd(0).FixedRows, grd(0).ColIndex("TotalB"), grd(0).rows - 1, grd(0).ColIndex("TotalB"))
            grd(0).TextMatrix(grd(0).rows - 1, grd(0).ColIndex("TotalB")) = SngTotal
            
      SngTotal = grd(0).Aggregate(flexSTSum, grd(0).FixedRows, grd(0).ColIndex("VATValue"), grd(0).rows - 1, grd(0).ColIndex("VATValue"))
            grd(0).TextMatrix(grd(0).rows - 1, grd(0).ColIndex("VATValue")) = SngTotal
              TxtVATValue = SngTotal
        
        SngTotal = grd(0).Aggregate(flexSTSum, grd(0).FixedRows, grd(0).ColIndex("PayableAmount"), grd(0).rows - 1, grd(0).ColIndex("PayableAmount"))
            grd(0).TextMatrix(grd(0).rows - 1, grd(0).ColIndex("PayableAmount")) = SngTotal
            txtPayableAmount = SngTotal
            
            
  '  ReLineGrid
    Exit Sub
ErrTrap:
End Sub
 
' [ ⁄œÌ·] ≈÷«›… »«—«„Ì — mBranchID ›Ì ‰Â«Ì… «·œ«·…

 Function print_report(Optional NoteSerial As String)
On Error GoTo ErrTrap
    
    '--- (·«  €ÌÌ— ›Ì «·„ €Ì—« ) ---
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    Dim SellerName As String
    Dim SellerNameAr As String
    Dim Vatregestriationnumber As String
    Dim RegNoCom As String
    Dim cOptions As New ClsCompanyInfo
    
    '--- (Ã·» »Ì«‰«  «·‘—ﬂ… ﬂ„« ÂÌ) ---
    SellerName = cOptions.EngCompanyName
    SellerNameAr = cOptions.ArabCompanyName
    RegNoCom = cOptions.ArabComment
    Vatregestriationnumber = IIf(cOptions.VATRegNo = "", "123456789", cOptions.VATRegNo)

    '--- [«· €ÌÌ— «·—∆Ì”Ì Â‰«] ---
    ' »œ·« „‰ »‰«¡ Ã„·… SQL ﬂ‰’° ”‰” Œœ„ ADODB.Command
    Dim oCmd As New ADODB.Command
    Set RsData = New ADODB.Recordset
    
    With oCmd
        .ActiveConnection = Cn  ' (Cn ÂÊ „ €Ì— «·« ’«· «·⁄«„ ⁄‰œﬂ)
        .CommandType = adCmdStoredProc
        .CommandText = "sp_GetEInvoiceDataForReport"
        
        ' 1. ≈÷«›… »«—«„Ì — «·›« Ê—… («·√Â„)
        .Parameters.Append .CreateParameter("@InvoiceID", adVarWChar, adParamInput, 100, Trim(txtInvoiceID.text))
        
        ' 2. ≈÷«›… »«—«„Ì —«  »Ì«‰«  «·‘—ﬂ…
        .Parameters.Append .CreateParameter("@SellerNameEn", adVarWChar, adParamInput, 255, SellerName)
        .Parameters.Append .CreateParameter("@SellerNameAr", adVarWChar, adParamInput, 255, SellerNameAr)
        .Parameters.Append .CreateParameter("@RegNoCom", adVarWChar, adParamInput, 100, RegNoCom)
        .Parameters.Append .CreateParameter("@Vatregestriationnumber", adVarWChar, adParamInput, 50, Vatregestriationnumber)
    End With
    
    '--- ( ÕœÌœ „”«— «· ﬁ—Ì— - ·«  €ÌÌ—) ---
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Einvoice.rpt"
    Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Einvoice.rpt"
    End If
    If Dir(StrFileName) = "" Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
    
    '--- [ €ÌÌ—] › Õ «·‹ Recordset »«” Œœ«„ «·‹ Command ---
    RsData.Open oCmd, , adOpenKeyset, adLockOptimistic
    
    If RsData.BOF Or RsData.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Set oCmd = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    '--- [ ⁄œÌ· „Â„] ---
    ' ﬁ„ » „—Ì— —ﬁ„ «·›—⁄ (branch_id) ≈·Ï œ«·… SaveQRCode
    ' Â–« Ì„‰⁄Â« „‰ ≈Ã—«¡ «” ⁄·«„ ≈÷«›Ì Ê€Ì— ¬„‰
    Dim BranchID As Long
    BranchID = val(RsData!branch_id & "")
    
    SaveQRCode6 "tblEInvoice", "ID", val(RsData!ID & ""), Trim(RsData!invoiceID & ""), RsData!IssueDate & "", _
        val(RsData!PayableAmount & ""), Picture1, 0, val(RsData!VATValue & ""), val(RsData!PayableAmount & ""), BranchID

    RsData.Close
    
    '--- [ €ÌÌ—] ≈⁄«œ… › Õ «·‹ Recordset ·· ﬁ—Ì— ---
    ' (Crystal Reports √ÕÌ«‰« Ì›÷· adOpenStatic)
    Set RsData = New ADODB.Recordset
    RsData.Open oCmd, , adOpenStatic, adLockReadOnly

    '--- (»«ﬁÌ ﬂÊœ «· ﬁ—Ì— - ·«  €ÌÌ— ﬂ»Ì—) ---
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
    '--- [ €ÌÌ—] ·« ‰„—— Ã„·… SQL «·ﬁœÌ„… ---
    ' CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , sql ' («·”ÿ— «·ﬁœÌ„)
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , "" ' («·”ÿ— «·ÃœÌœ)
    
    RsData.Close
    Set RsData = Nothing
    Set oCmd = Nothing ' ( ‰ŸÌ› «·‹ Command)
    Screen.MousePointer = vbDefault
    
ErrTrap:
    ' (Ì›÷· ≈÷«›… „⁄«·Ã… Œÿ√ Â‰«)
    Set RsData = Nothing
    Set oCmd = Nothing
    Screen.MousePointer = vbDefault
End Function
Private Sub grd_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)

If Index = 0 Then

If Me.TxtModFlg <> "E" And Me.TxtModFlg <> "N" Then Exit Sub
    
    Dim Percetage As Double
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    If Col = grd(0).ColIndex("Price") Or Col = grd(0).ColIndex("Qty") Then
        PercentgValueAddedAccount_Transec txtIssueDate.value, 21, 0, , Percetage
        If ComResid(0).value Then
            Percetage = 0
        End If
        grd(0).TextMatrix(Row, grd(0).ColIndex("TotalB")) = val(grd(0).TextMatrix(Row, grd(0).ColIndex("Price"))) * val(grd(0).TextMatrix(Row, grd(0).ColIndex("Qty")))
        grd(0).TextMatrix(Row, grd(0).ColIndex("VATValue")) = val(grd(0).TextMatrix(Row, grd(0).ColIndex("TotalB"))) * Percetage / 100
        
        grd(0).TextMatrix(Row, grd(0).ColIndex("PayableAmount")) = val(grd(0).TextMatrix(Row, grd(0).ColIndex("TotalB"))) + val(grd(0).TextMatrix(Row, grd(0).ColIndex("VATValue")))
    End If
    
End If
Dim i As Long
txtPayableAmount = 0
TxtVATValue = 0
  Dim SngTotal As Double
  grd(0).Subtotal flexSTClear
'  grd(0).Subtotal flexSTSum
  
  
        SngTotal = grd(0).Aggregate(flexSTSum, grd(0).FixedRows, grd(0).ColIndex("TotalB"), grd(0).rows - 1, grd(0).ColIndex("TotalB"))
            grd(0).TextMatrix(grd(0).rows - 1, grd(0).ColIndex("TotalB")) = SngTotal
            
      SngTotal = grd(0).Aggregate(flexSTSum, grd(0).FixedRows, grd(0).ColIndex("VATValue"), grd(0).rows - 1, grd(0).ColIndex("VATValue"))
            grd(0).TextMatrix(grd(0).rows - 1, grd(0).ColIndex("VATValue")) = SngTotal
              TxtVATValue = SngTotal
        
        SngTotal = grd(0).Aggregate(flexSTSum, grd(0).FixedRows, grd(0).ColIndex("PayableAmount"), grd(0).rows - 1, grd(0).ColIndex("PayableAmount"))
            grd(0).TextMatrix(grd(0).rows - 1, grd(0).ColIndex("PayableAmount")) = SngTotal
            txtPayableAmount = SngTotal
            
'For i = 1 To grd(0).rows - 1
'    txtPayableAmount = val(txtPayableAmount) + val(grd(0).TextMatrix(i, grd(0).ColIndex("PayableAmount")))
'    txtVatValue = val(txtPayableAmount) + val(grd(0).TextMatrix(i, grd(0).ColIndex("VATValue")))
'Next
End Sub

Private Sub ISButton2_Click()
On Error GoTo ErrTrap
   If val(Me.txtid(0).text) <> 0 Then
    
       print_report
   End If
ErrTrap:
End Sub
'
Function print_reportOld(Optional NoteSerial As String)
On Error GoTo ErrTrap
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    Dim SellerName As String
    Dim SellerNameAr As String
    Dim Vatregestriationnumber As String
    Dim RegNoCom As String
 Dim cOptions As New ClsCompanyInfo
    SellerName = cOptions.EngCompanyName
    SellerNameAr = cOptions.ArabCompanyName
    
    RegNoCom = cOptions.ArabComment
    Vatregestriationnumber = IIf(cOptions.VATRegNo = "", "123456789", cOptions.VATRegNo)
     
        

sql = " SELECT   SellerNameEn = N'" & SellerName & "', SellerNameAr = N'" & SellerNameAr & "', RegNoComCom2  = N'" & RegNoCom & "',VatregestriationnumberCom = N'" & Vatregestriationnumber & "',    dbo.tblEInvoice.ID, dbo.tblEInvoice.InvoiceID,tblEInvoice.IqarName,tblEInvoice.ManualInvoiceNo, dbo.tblEInvoice.DefaultInvoicetype, dbo.tblEInvoice.ErrorMessageS, dbo.tblEInvoice.IssueDate, dbo.tblEInvoice.IssueTim, dbo.tblEInvoice.DocumentCurrencyCode,tblEInvoice.NewNO,tblEInvoice.ComResid,"
sql = sql & "                          dbo.tblEInvoice.TaxCurrencyCode, dbo.tblEInvoice.StreetName, dbo.tblEInvoice.BuildingNumber, dbo.tblEInvoice.CityName, dbo.tblEInvoice.PostalZone, dbo.tblEInvoice.CitySubdivisionName, dbo.tblEInvoice.RegistrationName,dbo.tblEInvoice.InvoiceID as NoteSerial11,"
sql = sql & "                          dbo.tblEInvoice.CompanyID, dbo.tblEInvoice.CoCRCode, dbo.tblEInvoice.PayableAmount, dbo.tblEInvoice.VatValue, dbo.tblEInvoice.Id700, dbo.tblEInvoice.serial, dbo.tblEInvoice.ExcelRow, dbo.tblEInvoice.ExcelFile,"
sql = sql & "                          dbo.tblEInvoice.QrCodeData, dbo.tblEInvoice.QrCodeDataPath, dbo.tblEInvoice.QrCodeImage, dbo.tblEInvoice.zatcaStatus, dbo.tblEInvoice.InvoiceTypeCodeID, dbo.tblEInvoice.InvoiceTypeCodename,"
sql = sql & "                          dbo.tblEInvoice.AdditionalDocumentReferencePIH, dbo.tblEInvoice.InvoiceDocumentReferenceID, dbo.tblEInvoice.AdditionalDocumentReferenceICVUUID, dbo.tblEInvoice.ActualDeliveryDate, dbo.tblEInvoice.LatestDeliveryDate,"
sql = sql & "                          dbo.tblEInvoice.RecTime, dbo.tblEInvoice.Transaction_ID, dbo.tblEInvoice.warrningmessage, dbo.tblEInvoice.PaymentMeansCode, dbo.tblEInvoice.InstructionNote, dbo.tblEInvoice.Iban, dbo.tblEInvoice.paymentnote,"
sql = sql & "                          dbo.tblEInvoice.AdditionalStreetName, dbo.tblEInvoice.PlotIdentification, dbo.tblEInvoice.CountrySubentity, dbo.tblEInvoice.IdentificationCode, dbo.tblEInvoice.last_changed, dbo.tblEInvoice.Identificationid,"
sql = sql & "                          dbo.tblEInvoice.schemeID, dbo.tblEInvoice.TaxCategoryPercent, dbo.tblEInvoice.TaxCategoryID, dbo.tblEInvoice.Export, dbo.tblEInvoice.branch_id, dbo.tblEInvoice.branch_name, dbo.tblEInvoice.branchname,"
sql = sql & "                          dbo.tblEInvoice.GroupUniqueCode, dbo.tblEInvoice.GroupUniqueFileMaster, dbo.tblEInvoice2.ItemName , dbo.tblEInvoice2.Price , dbo.tblEInvoice2.Qty ,"
sql = sql & "                          dbo.tblEInvoice2.Price , dbo.tblEInvoice2.PayableAmount AS PayableAmount2, dbo.tblEInvoice2.VatValue  VatValue2"
sql = sql & "        ,SellerName = N'" & SellerName & "',RegNoCom = N'" & RegNoCom & "',Vatregestriationnumber = '" & Vatregestriationnumber & "'"
    sql = sql & " from tblEInvoice Left outer JOIN  tblEInvoice2 On tblEInvoice.InvoiceID =tblEInvoice2.InvoiceID"

    sql = sql & " Where (dbo.tblEInvoice.InvoiceID = '" & Trim(txtInvoiceID.text) & "')"
        
        
        
 
                    
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Einvoice.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Einvoice.rpt"
        End If
   If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
   Set RsData = New ADODB.Recordset
    RsData.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     If RsData.BOF Or RsData.EOF Then
       Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
      MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
      Exit Function
   End If
   SaveQRCode "tblEInvoice", "ID", val(RsData!ID & ""), Trim(RsData!invoiceID & ""), RsData!IssueDate & "", _
        val(RsData!PayableAmount & ""), Picture1, 0, val(RsData!VATValue & ""), val(RsData!PayableAmount & "")
      RsData.Close
      RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    
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
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , sql
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
ErrTrap:
  End Function

Private Sub TxtItemCode_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 20915
        FrmItemSearch.show vbModal
    End If

End Sub

Private Sub TxtModFlg_Change()

    If Me.TxtModFlg.text = "N" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(5).Enabled = False
        ISButton2.Enabled = False
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

    ElseIf Me.TxtModFlg.text = "E" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True
        ISButton2.Enabled = False
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        Cmd(5).Enabled = False

    Else
        Ele(1).Enabled = False
        ISButton2.Enabled = True
        CmdRemove.Enabled = False
        Cmd(2).Enabled = False
        Cmd(3).Enabled = False
        Cmd(0).Enabled = True
        Cmd(1).Enabled = True
        Cmd(4).Enabled = True

        Cmd(5).Enabled = True

    End If

End Sub

Private Sub txtPassword_Change()
    If Trim(txtPassword) = "Edit2025" Then
        Cmd(1).Visible = True
    Else
        Cmd(1).Visible = False
    End If
End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
    End If

    On Error GoTo ErrTrap

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
 Public Function FindRec(ByVal TblVendorContractD As Long)
    On Error GoTo ErrTrap
    rs.Find "ID=" & TblVendorContractD, , adSearchForward, 1
    If Not (rs.EOF) Then
        Retrive
        End If
    Exit Function
ErrTrap:
    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
       ' BtnUndo_Click
    End If
  End Function
' aladein add
''''''''''''''''''''''''''''''''
Private Sub NewInvoice()
    Dim pref As String
    Dim ID As String
    Dim d As Date

    ' «Œ Û— «· «—ÌŒ («·ÌÊ„) √Ê  «—ÌŒ „Œ’’
    d = txtIssueDate.value

    ' „À·« „‰ «·ﬂÊ„»Ê: DcBranches.BoundText = branch_id ? Â«  «·ﬂÊœ „‰ «·ÃœÊ·
    ' √Ê ·Ê ⁄‰œﬂ «·ﬂÊœ „»«‘—…:
    pref = txtBranch_Code.text  ' „À«·: "IR" √Ê "SCQ"

    ID = NextInvoiceID(pref, d, 5) ' 5 Œ«‰«  ··”Ì—Ì«·

    If ID <> "" Then
        txtInvoiceID.text = ID
        txtIssueDate.value = d   ' √Ê √Ì ﬂ‰ —Ê·  «—ÌŒ ⁄‰œﬂ
    Else
        MsgBox " ⁄–—  Ê·Ìœ —ﬁ„ «·›« Ê—….", vbExclamation
    End If
End Sub


'  Ê·Ìœ —ﬁ„ ›« Ê—… ÃœÌœ ⁄·Ï «·‰„ÿ Prefix/YYYY/Serial
' BranchPrefix: ﬂÊœ «·›—⁄ (IR, SCQ, SPI, ...)
' IssueDt:  «—ÌŒ «·≈’œ«— (·Ê 0 ‰” Œœ„ Date)
' PadLen: ⁄œœ Œ«‰«  «·”Ì—Ì«· »’›— »«œ∆ («› —«÷Ì 5 -> 00001)
' Ì ÿ·»: „—Ã⁄ Microsoft ActiveX Data Objects
Public Function NextInvoiceID(ByVal BranchPrefix As String, _
                              Optional ByVal IssueDt As Date = 0, _
                              Optional ByVal PadLen As Integer = 5, _
                              Optional ByRef newSerial As Long) As String
    On Error GoTo eh

    Dim Y As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim lastSer As Long
    Dim safePrefix As String
    Dim constPartLen As Long   ' <-- ﬂ«‰ ‰«ﬁ’  ⁄—Ì›Â

    If IssueDt = 0 Then IssueDt = Date
    Y = Format$(IssueDt, "yyyy")

    '  √„Ì‰ ⁄·«„«  «·«ﬁ »«” œ«Œ· «·»—Ì›ﬂ”
    safePrefix = Replace(BranchPrefix, "'", "''")

    ' ÿÊ· «·Ã“¡ «·À«»  ﬁ»· «·”Ì—Ì«·: Prefix + "/" + YYYY + "/"
    ' = Len(prefix) + 1 + 4 + 1 = Len(prefix) + 6
    constPartLen = Len(BranchPrefix) + 6

    ' «·ﬂÊÌ—Ì «·’ÕÌÕ (»œÊ‰ TRY_CONVERT)
    sql = "SELECT ISNULL(MAX(CAST(RIGHT(InvoiceID, LEN(InvoiceID) - (" & constPartLen & _
          ")) AS INT)),0) AS lastSer " & _
          "FROM tblEInvoice " & _
          "WHERE InvoiceID LIKE '" & safePrefix & "/" & Y & "/%' " & _
          "AND PATINDEX('%[^0-9]%', RIGHT(InvoiceID, LEN(InvoiceID) - (" & constPartLen & "))) = 0;"

    rs.Open sql, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    If Not rs.EOF Then lastSer = val(rs!lastSer & "")
    rs.Close: Set rs = Nothing

    newSerial = lastSer + 1

    ' »‰«¡ —ﬁ„ «·›« Ê—…: Prefix/YYYY/Serial „⁄ ’›— »«œ∆
    NextInvoiceID = BranchPrefix & "/" & Y & "/" & right$(String$(PadLen, "0") & CStr(newSerial), PadLen)
    Exit Function

eh:
    NextInvoiceID = ""
    newSerial = 0
End Function


