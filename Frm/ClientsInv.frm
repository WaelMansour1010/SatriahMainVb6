VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form ClientsInv 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "›Ê« Ì— «·⁄„·«¡"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15045
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8605.668
   ScaleMode       =   0  'User
   ScaleTop        =   60
   ScaleWidth      =   14550
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
      Height          =   9645
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15045
      _cx             =   26538
      _cy             =   17013
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
      Frame           =   0
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
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   " ÕœÌœ «·ﬂ·"
         Height          =   195
         Left            =   5940
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   4470
         Width           =   1200
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         Caption         =   " ÕœÌœ «·ﬂ·"
         Height          =   195
         Left            =   13620
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   4440
         Width           =   1200
      End
      Begin VB.CommandButton Command1 
         Caption         =   " ‰›Ì– «·—»ÿ"
         Height          =   300
         Left            =   12000
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   3960
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.TextBox RecId 
         Alignment       =   1  'Right Justify
         Height          =   270
         Left            =   630
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   1410
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.TextBox XPTxtVal 
         Alignment       =   1  'Right Justify
         Height          =   390
         Left            =   1545
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   810
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox TxtValueTemp 
         Alignment       =   1  'Right Justify
         Height          =   390
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   810
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Frame FraHeader 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   645
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   0
         Width           =   15045
         Begin VB.TextBox TxtVac_ID 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            Height          =   240
            Left            =   7920
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   240
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            Enabled         =   0   'False
            Height          =   285
            Left            =   3900
            RightToLeft     =   -1  'True
            TabIndex        =   40
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
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   120
            Visible         =   0   'False
            Width           =   3105
            Begin MSDataListLib.DataCombo DCUser 
               CausesValidation=   0   'False
               Height          =   315
               Left            =   5865
               TabIndex        =   38
               Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
               Top             =   375
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
               TabIndex        =   39
               Top             =   45
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
                  Picture         =   "ClientsInv.frx":0000
                  Key             =   "CompanyName"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ClientsInv.frx":039A
                  Key             =   "Ser"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ClientsInv.frx":0734
                  Key             =   "Vac_Name"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ClientsInv.frx":0ACE
                  Key             =   "ShareCount"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ClientsInv.frx":0E68
                  Key             =   "Dis_Count"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ClientsInv.frx":1202
                  Key             =   "Bouns"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ClientsInv.frx":159C
                  Key             =   "SharesValue"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "ClientsInv.frx":1B36
                  Key             =   "BuyValue"
               EndProperty
            EndProperty
         End
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   930
            TabIndex        =   42
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
            ButtonImage     =   "ClientsInv.frx":1ED0
            ColorButton     =   14871017
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   1395
            TabIndex        =   43
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
            ButtonImage     =   "ClientsInv.frx":226A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1995
            TabIndex        =   44
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
            ButtonImage     =   "ClientsInv.frx":2604
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   2460
            TabIndex        =   45
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
            ButtonImage     =   "ClientsInv.frx":299E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "›Ê« Ì— «·⁄„·«¡"
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
            Height          =   495
            Index           =   2
            Left            =   5880
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   210
            Width           =   8040
         End
      End
      Begin VB.TextBox txtcode 
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
         Left            =   12825
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   900
         Width           =   1605
      End
      Begin VB.CommandButton ShowInvData 
         Caption         =   "⁄—÷ «·”‰œ« "
         Height          =   300
         Left            =   12000
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   3600
         Width           =   2505
      End
      Begin MSComCtl2.DTPicker DtRecord 
         Height          =   330
         Left            =   9690
         TabIndex        =   3
         Top             =   900
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         CalendarBackColor=   12648447
         CustomFormat    =   "yyyy/M/d"
         Format          =   322371587
         CurrentDate     =   38718
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   450
         Index           =   1
         Left            =   30
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   9180
         Width           =   14970
         _cx             =   26405
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
            Height          =   330
            Index           =   0
            Left            =   13740
            TabIndex        =   5
            Top             =   60
            Width           =   1050
            _ExtentX        =   1852
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
            Height          =   330
            Index           =   1
            Left            =   12135
            TabIndex        =   6
            Top             =   60
            Width           =   1170
            _ExtentX        =   2064
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
            Height          =   330
            Index           =   2
            Left            =   10485
            TabIndex        =   7
            Top             =   60
            Width           =   1125
            _ExtentX        =   1984
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
            Height          =   330
            Index           =   3
            Left            =   8880
            TabIndex        =   8
            Top             =   60
            Width           =   1035
            _ExtentX        =   1826
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
            Height          =   330
            Index           =   4
            Left            =   7095
            TabIndex        =   9
            Top             =   60
            Width           =   1140
            _ExtentX        =   2011
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
            Height          =   330
            Index           =   5
            Left            =   5475
            TabIndex        =   10
            Top             =   60
            Width           =   1080
            _ExtentX        =   1905
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
            Height          =   330
            Index           =   6
            Left            =   555
            TabIndex        =   11
            Top             =   60
            Width           =   855
            _ExtentX        =   1508
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
            Height          =   330
            Index           =   7
            Left            =   4065
            TabIndex        =   12
            Top             =   60
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   582
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton CmdHelp 
            Height          =   330
            Left            =   2205
            TabIndex        =   13
            Top             =   60
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "„”«⁄œ…"
            BackColor       =   14871017
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
      End
      Begin XtremeSuiteControls.RadioButton Rd 
         Height          =   255
         Index           =   2
         Left            =   13200
         TabIndex        =   16
         Top             =   1680
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "·⁄„Ì· „Õœœ"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton Rd 
         Height          =   255
         Index           =   3
         Left            =   6480
         TabIndex        =   17
         Top             =   1680
         Visible         =   0   'False
         Width           =   1545
         _Version        =   786432
         _ExtentX        =   2725
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "·„Ã„Ê⁄… ⁄„·«¡"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin C1SizerLibCtl.C1Elastic PayTypeFram 
         Height          =   660
         Left            =   8520
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2760
         Visible         =   0   'False
         Width           =   6060
         _cx             =   10689
         _cy             =   1164
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
         CaptionPos      =   7
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
         Frame           =   0
         FrameStyle      =   5
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin XtremeSuiteControls.RadioButton Rd 
            Height          =   225
            Index           =   4
            Left            =   3570
            TabIndex        =   28
            Top             =   240
            Width           =   1800
            _Version        =   786432
            _ExtentX        =   3175
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "FIFO"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton Rd 
            Height          =   225
            Index           =   5
            Left            =   615
            TabIndex        =   29
            Top             =   240
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "›Ê« Ì—"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
      End
      Begin C1SizerLibCtl.C1Elastic CusTypeFrame 
         Height          =   585
         Left            =   3285
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   780
         Width           =   5355
         _cx             =   9446
         _cy             =   1032
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
         CaptionPos      =   7
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
         Frame           =   0
         FrameStyle      =   5
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin XtremeSuiteControls.RadioButton Rd 
            Height          =   270
            Index           =   0
            Left            =   2145
            TabIndex        =   21
            Top             =   225
            Width           =   2580
            _Version        =   786432
            _ExtentX        =   4551
            _ExtentY        =   476
            _StockProps     =   79
            Caption         =   "·⁄„Ì·"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton Rd 
            Height          =   270
            Index           =   1
            Left            =   390
            TabIndex        =   22
            Top             =   225
            Width           =   1680
            _Version        =   786432
            _ExtentX        =   2963
            _ExtentY        =   476
            _StockProps     =   79
            Caption         =   "·„Ê—œ"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
      End
      Begin C1SizerLibCtl.C1Elastic SingleCusFrame 
         Height          =   1080
         Left            =   8400
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1680
         Width           =   6060
         _cx             =   10689
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
         CaptionPos      =   7
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
         Frame           =   0
         FrameStyle      =   5
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.TextBox CusCodeText 
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
            Left            =   4005
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   465
            Width           =   1410
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   660
            TabIndex        =   26
            Top             =   465
            Width           =   3165
            _ExtentX        =   5583
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
      End
      Begin C1SizerLibCtl.C1Elastic GroupCusFram 
         Height          =   2550
         Left            =   0
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1860
         Visible         =   0   'False
         Width           =   8190
         _cx             =   14446
         _cy             =   4498
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
         CaptionPos      =   7
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
         Frame           =   0
         FrameStyle      =   5
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.ListBox SelectedCusList 
            Height          =   1815
            ItemData        =   "ClientsInv.frx":2D38
            Left            =   300
            List            =   "ClientsInv.frx":2D3F
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   345
            Width           =   3255
         End
         Begin VB.ListBox CusList 
            Height          =   1815
            ItemData        =   "ClientsInv.frx":2D52
            Left            =   4725
            List            =   "ClientsInv.frx":2D59
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   345
            Width           =   3255
         End
         Begin VB.Label SelectSingleCus 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
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
            Left            =   3660
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   480
            Width           =   990
         End
         Begin VB.Label SelectAllCus 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
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
            Height          =   225
            Left            =   3660
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   900
            Width           =   990
         End
         Begin VB.Label RemoveAllCus 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
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
            Height          =   315
            Left            =   3660
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   1740
            Width           =   990
         End
         Begin VB.Label RemoveSingleCus 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
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
            Height          =   405
            Left            =   3660
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   1230
            Width           =   990
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid GR 
         Height          =   3885
         Left            =   7530
         TabIndex        =   47
         Top             =   4680
         Width           =   7335
         _cx             =   12938
         _cy             =   6853
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
         Cols            =   18
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"ClientsInv.frx":2D66
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
         ExplorerBar     =   1
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   600
         Left            =   150
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   8475
         Width           =   14745
         _cx             =   26009
         _cy             =   1058
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
         CaptionPos      =   7
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
         Frame           =   0
         FrameStyle      =   5
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   10005
            TabIndex        =   52
            Top             =   210
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   270
            Left            =   5385
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   210
            Width           =   675
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ì „”œœ «·Õ—ﬂ…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   3210
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   210
            Width           =   1755
         End
         Begin VB.Label TotalPaidLab 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "00.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   270
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   210
            Width           =   1365
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—.”"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   270
            Left            =   630
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   210
            Width           =   450
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ê«”ÿ…"
            Height          =   255
            Index           =   4
            Left            =   12735
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   210
            Width           =   1380
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   270
            Index           =   0
            Left            =   8445
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   225
            Width           =   1290
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   270
            Index           =   2
            Left            =   6420
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   225
            Width           =   1260
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   300
            Left            =   7680
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   210
            Width           =   660
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid GridCashing 
         Height          =   3885
         Left            =   120
         TabIndex        =   61
         Top             =   4680
         Width           =   7185
         _cx             =   12674
         _cy             =   6853
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
         Cols            =   18
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"ClientsInv.frx":303A
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
         ExplorerBar     =   1
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
         Caption         =   "„"
         Height          =   210
         Index           =   3
         Left            =   14310
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   900
         Width           =   570
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·Õ—ﬂ…"
         Height          =   225
         Index           =   0
         Left            =   11055
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   990
         Width           =   1380
      End
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   1065
      Index           =   2
      Left            =   1320
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2355
      _cx             =   4154
      _cy             =   1879
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
      CaptionPos      =   7
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
      Frame           =   0
      FrameStyle      =   5
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   30
      Index           =   4
      Left            =   0
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Width           =   2355
      _cx             =   4154
      _cy             =   53
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
      CaptionPos      =   7
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
      Frame           =   0
      FrameStyle      =   5
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
End
Attribute VB_Name = "ClientsInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'########################## Genrale Var Declaration ######################
Dim StrSQL  As String
Dim rs As ADODB.Recordset
Private Sub BtnFirst_Click()
    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveFirst
       ' txtcode.Text = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
       ' RecId.Text = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
    End If
    Retrive
End Sub

Function saveBillBuy()
    Dim StrSQL As String
   ' Dim StrSQL  As String
    Dim i As Integer
    Dim j As Integer
    Dim Diff As Double
    Dim Note_Value1 As Double
    Dim mNoteId As Long
    Diff = 0
Dim RsDetails As ADODB.Recordset
        If Me.TxtModFlg.Text = "E" Then
            For i = 1 To GridCashing.Rows - 1
                If GridCashing.Cell(flexcpChecked, i, GridCashing.ColIndex("payed")) = flexChecked Then
                    mNoteId = val(GridCashing.TextMatrix(i, GridCashing.ColIndex("NoteID")))
                
                    StrSQL = "Delete From TblNotesBillBuyPayment2 Where NoteID1=" & mNoteId & " and TransType is null"
                    Cn.Execute StrSQL, , adExecuteNoRecords
                    StrSQL = "Delete From TblBillBuyPayment2 Where TypTrans IS NULL and  NoteID=" & mNoteId & " and TransType is null"
                    Cn.Execute StrSQL, , adExecuteNoRecords
                End If
            Next
        End If

    Set RsDetails = New ADODB.Recordset
   ' RsDetails.Open "TblEmpAdvanceDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
Relin
    StrSQL = "SELECT     * from dbo.TblNotesBillBuyPayment2 Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
   
      '  If GridCashing.Cell(flexcpChecked, j, GridCashing.ColIndex("payed")) = flexChecked Then
                mNoteId = val(GridCashing.TextMatrix(j, GridCashing.ColIndex("NoteID")))
                With GR
               ' TxtValueTemp.Text = TotalPaidLab.Caption
                'TxtValueTemp.Text = val(GridCashing.TextMatrix(j, GridCashing.ColIndex("value")))
                 TxtValueTemp.Text = val(TotalPaidLab.Caption)
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
                        RsDetails.AddNew
                        RsDetails("NoteID1").value = val(mNoteId)
                        RsDetails("NoteID").value = val(.TextMatrix(i, .ColIndex("NoteID")))
                        RsDetails("branch_no").value = val(.TextMatrix(i, .ColIndex("branch_no")))
                        RsDetails("NoteSerial1").value = val(.TextMatrix(i, .ColIndex("NoteSerial1")))
                        RsDetails("Note_Value").value = val(.TextMatrix(i, .ColIndex("Note_Value")))
                        Note_Value1 = val(.TextMatrix(i, .ColIndex("RemainingValue")))
                        Diff = 0
                        If val(TxtValueTemp.Text) > 0 Then
                      If val(TxtValueTemp.Text) <= Note_Value1 Then
                      Diff = val(TxtValueTemp.Text)
                      TxtValueTemp.Text = val(TotalPaidLab.Caption) - Note_Value1
                      Else
                      Diff = Note_Value1
                      TxtValueTemp.Text = val(TotalPaidLab.Caption) - Note_Value1
                      End If
                        End If
                       .TextMatrix(i, .ColIndex("RemainingValue")) = val(.TextMatrix(i, .ColIndex("RemainingValue"))) - Diff
                        .TextMatrix(i, .ColIndex("TransPayedValue")) = Diff
                        .TextMatrix(i, .ColIndex("PayedValue")) = Diff
                        RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("PayedValue")))
                        
                        RsDetails("too").value = (.TextMatrix(i, .ColIndex("too")))
                        RsDetails("NoteDate").value = IIf((.TextMatrix(i, .ColIndex("NoteDate"))) = "", Null, (.TextMatrix(i, .ColIndex("NoteDate"))))
                        If .TextMatrix(i, .ColIndex("DueDate")) <> "" And .TextMatrix(i, .ColIndex("DueDate")) <> " " Then
                        RsDetails("DueDate").value = IIf((.TextMatrix(i, .ColIndex("DueDate"))) = "", Null, (.TextMatrix(i, .ColIndex("DueDate"))))
                        Else
                        RsDetails("DueDate").value = Null
                        End If
                        RsDetails("TransPayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
                       .TextMatrix(i, .ColIndex("NetValue")) = val(.TextMatrix(i, .ColIndex("RemainingValue"))) - val(.TextMatrix(i, .ColIndex("TransPayedValue")))
                        RsDetails("NetValue").value = val(.TextMatrix(i, .ColIndex("NetValue")))
                        RsDetails("RemainingValue").value = val(.TextMatrix(i, .ColIndex("RemainingValue")))
                        RsDetails.update
                            
                        If val(val(.TextMatrix(i, .ColIndex("NetValue")))) = 0 Then
                        StrSQL = "Update Transactions Set  TotalPayed=1 Where Transaction_ID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                            Cn.Execute StrSQL, , adExecuteNoRecords
                         Else
                             StrSQL = "Update Transactions Set  TotalPayed=0 Where Transaction_ID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                            Cn.Execute StrSQL, , adExecuteNoRecords
                        End If
                 End If
                Next i
            End With
        
    
    Set RsDetails = New ADODB.Recordset
   ' RsDetails.Open "TblEmpAdvanceDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    StrSQL = "SELECT     * from dbo.TblBillBuyPayment2 Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    For j = 1 To GridCashing.Rows - 1
        If GridCashing.Cell(flexcpChecked, j, GridCashing.ColIndex("payed")) = flexChecked Then
            mNoteId = val(GridCashing.TextMatrix(j, GridCashing.ColIndex("NoteID")))
            With GR
                For i = .FixedRows To .Rows - 1
                    If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
                        RsDetails.AddNew
                        RsDetails("NoteID").value = val(mNoteId)
                        RsDetails("RecDate").value = Trim(GridCashing.TextMatrix(j, GridCashing.ColIndex("RecordDate")))
                        
                        RsDetails("Serial").value = Trim(GridCashing.TextMatrix(j, GridCashing.ColIndex("NoteSerial1")))
                        
                        RsDetails("Transaction_ID").value = val(.TextMatrix(i, .ColIndex("NoteID")))
                        RsDetails("Note_Value").value = val(.TextMatrix(i, .ColIndex("Note_Value")))
                        RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
                        RsDetails.update
                    End If
                Next i
            End With
        End If
    Next j
        

End Function


Private Sub BtnLast_Click()
    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveLast
       ' txtcode.Text = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
      '  RecId.Text = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
    End If
    Retrive
End Sub

Private Sub BtnNext_Click()
    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveNext
        If rs.EOF Then rs.MoveLast
        txtcode.Text = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
        RecId.Text = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
    End If
    Retrive
End Sub

Private Sub BtnPrevious_Click()
    If Not (rs.EOF Or rs.BOF) Then
        rs.MovePrevious
        If rs.BOF Then rs.MoveFirst
        txtcode.Text = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
        RecId.Text = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
    End If
    Retrive
End Sub

Private Sub Check1_Click()
    Dim i As Integer

    If Check1.value = vbChecked Then

        With Me.GR
 
            For i = .FixedRows To .Rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = True
            Next i

        End With

    Else

        With Me.GR

            For i = .FixedRows To .Rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = False
            Next i

        End With

    End If
 '   RelineBuy
End Sub

Private Sub Check2_Click()

    Dim i As Integer

    If Check2.value = vbChecked Then

        With Me.GridCashing
 
            For i = .FixedRows To .Rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = True
            Next i

        End With

    Else

        With Me.GridCashing

            For i = .FixedRows To .Rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = False
            Next i

        End With

    End If
 '   RelineBuy
 Relin

End Sub

Private Sub DBCboClientName_Change()
DBCboClientName_Click (0)
End Sub

'#################################################################
Private Sub Form_Load() ' %%%%%% windows start %%%%%%%
    '############################### Set Icons for bottom Bar #############################
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    'Set CmdConvert.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Excute").Picture
    'Set CmdTemplate.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Excute").Picture
    'Set Accredit.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Required").Picture
    '######################################################################################
    
    '################################## Change The lung ###################################
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    '######################################################################################
    '############################### Defult window state ##################################
    'Rd(0).value = True
    'Rd(3).value = False
    '######################################################################################
    
    '########################## change windows state to read ##############################
    TxtModFlg.Text = "R"
    '######################################################################################
    
    '########################## Get data for all list and combos ##########################
    FillWindowsControlesData
    '######################################################################################
    
    '################################# Get the last recored ###############################
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblEndDebtAgingInv"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveLast
        txtcode.Text = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
        RecId.Text = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
    End If
    Retrive
    '######################################################################################
End Sub
Private Sub ChangeLang() '%%%%% Convert window object lung to english %%%%%%
Label1(2).Caption = "Client Invoices"
Label1(3).Caption = "Ser"
Label1(0).Caption = "Process date"
Rd(0).Caption = "Client"
Rd(1).Caption = "Supplier"
Rd(3).Caption = "Multiple"
Rd(2).Caption = "Single"
Rd(5).Caption = "Invoices"
ShowInvData.Caption = "Show Invoices"
With GR
    .TextMatrix(0, .ColIndex("Ser")) = "No."
    .TextMatrix(0, .ColIndex("payed")) = "Payed"
    .TextMatrix(0, .ColIndex("NoteSerial1")) = "Invoice No."
    .TextMatrix(0, .ColIndex("too")) = "For Client"
    .TextMatrix(0, .ColIndex("NoteDate")) = "Invoice data"
    .TextMatrix(0, .ColIndex("DueDate")) = "Due Date"
    .TextMatrix(0, .ColIndex("branch_name")) = "Branch"
    .TextMatrix(0, .ColIndex("Note_Value")) = "Invoice Value"
    .TextMatrix(0, .ColIndex("PayedValue")) = "Payed Value"
    .TextMatrix(0, .ColIndex("RemainingValue")) = "Remaining"
    .TextMatrix(0, .ColIndex("TransPayedValue")) = "Payed in Process"
    .TextMatrix(0, .ColIndex("NetValue")) = "Due Net"
End With
Label1(4).Caption = "By"
Label1(1).Caption = "Total payed int the process"
Label2.Caption = "S.R"
Cmd(0).Caption = "New"
Cmd(1).Caption = "Edit"
Cmd(2).Caption = "Save"
Cmd(3).Caption = "Cancel"
Cmd(4).Caption = "Delete"
Cmd(5).Caption = "Search"
Cmd(7).Caption = "Print"
Cmd(6).Caption = "Exit"
CmdHelp.Caption = "Help"
End Sub
Private Sub Cmd_Click(Index As Integer) '%%%%%%%%% Command Bar %%%%%%%%%
    '@@@@@@@@ declear Var @@@@@@@@@
    Dim intDef As Integer
    Dim StrSQL As String
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    
    'Handling an exception
    'On Error GoTo ErrTrap

    Select Case Index
        Case 0
        '######################### New Bottom ###########################
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
            
            TxtModFlg.Text = "N"
            
            Me.DCboUserName.BoundText = user_id
            
            txtcode.Text = new_id("TblEndDebtAgingInv", "ID", "", True)
            RecId.Text = new_id("TblEndDebtAgingInv", "ID", "", True)
            TotalPaidLab.Caption = 0
            'default is client
            Rd(0).value = True
            'default note group
            Rd(2).value = True
            GR.Rows = 1
            SelectedCusList.Clear
            DtRecord.value = Date
        '################################################################
        Case 1
        '######################## Edit Bottom ###########################
        'check if user have permission to EDIT recored
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
          If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ê› Ì „ Õ–› «·„œ›Ê⁄«  ›Ì Â–Â «·Õ—ﬂ…  .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
            Else
            Msg = "Payments will be deleted in this movement"
            End If
            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
            DeleteBillBuy

         StrSQL = "Delete From TblBillBuyPayment2 Where TypTrans=1 and  NoteID=" & val(Me.txtcode.Text)
         Cn.Execute StrSQL, , adExecuteNoRecords
             StrSQL = "Delete From TblBillBuyPayment2    Where TypTrans=1 and NoteID=" & val(Me.txtcode.Text)
         Cn.Execute StrSQL, , adExecuteNoRecords
            TxtModFlg.Text = "E"
            
            StrSQL = "Delete From TblEndDebtAgingInvDet Where EndDebAgInvID = " & val(txtcode.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            ShowInvData_Click
             Me.DCboUserName.BoundText = user_id
           Else
           Exit Sub
           End If
        Case 2
'        If val(TotalPaidLab.Caption) = 0 Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'        MsgBox "·«ÌÊÃœ «Ì „»·€ „”œœ"
'        Else
'        MsgBox "There is no amount paid"
'        End If
'        Exit Sub
'        End If
        '######################## Save Bottom ###########################
            'call save function
            SaveData
        '################################################################
        Case 3
        '######################## Undo Bottom ###########################
            'call undo function
            Undo
        '################################################################
        Case 4
        '######################## Delete Bottom ###########################
            ' check if user have permission to DELETE recored
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If
            
            If SystemOptions.UserInterface = EnglishInterface Then
            Else
            End If
           If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ê› Ì „ Õ–› «·⁄„·Ì… .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
            Else
            Msg = "Confirm Delete"
            End If
            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                'call Delete function
                DelRecored
                Me.TxtModFlg.Text = "R"
                BtnLast_Click
            End If
            

        '################################################################
        Case 5
        '######################## Search Bottom #########################
        '################################################################
        Case 7
        '######################## Print Bottom ##########################
            'call print report function
            PrintReport
        '################################################################
        Case 6
        '######################## Exit Bottom ##########################
            'clear all function and get the last recored
            Unload Me
        '################################################################
    End Select
ErrTrap:
'******************************** show Error Message *******************************
End Sub
Private Sub Undo()    '%%%%%%%% Undo Enteries and clear all fields also set text mode to R %%%%%%%%
    '@@@@@@@@ declear Var @@@@@@@@@
    Dim Msg As String
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    
    'Handling an exception
    On Error GoTo ErrTrap
    
    Select Case TxtModFlg.Text
        Case "N"
        
              If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "This process will be undone."
                Msg = Msg & CHR(13) & "do you want to continue"
            Else
                Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–Â «·⁄„·Ì… .."
                Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
            End If
          
            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                clear_all Me
                Me.TxtModFlg.Text = "R"
                BtnLast_Click
            End If
        Case "E"
        
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "This process will be undone."
                Msg = Msg & CHR(13) & "do you want to continue"
            Else
                Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–Â «·⁄„·Ì… .."
                Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
            End If
            
            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                'rs.Find "Transaction_ID='" & val(XPTxtBillID.Text) & "'", , adSearchForward, adBookmarkFirst
                'If Not rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.Text = "R"
                    BtnLast_Click
                    'Retrive
                'End If
            End If
    End Select
    
    'get data again
    Retrive
    
ErrTrap:
End Sub
Public Sub Retrive(Optional Lngid As Long = 0) '%%%%%%%%%% Get the last Recored %%%%%%%%%%
    '@@@@@@@@ declear Var @@@@@@@@@
    Dim sql As String
    Dim i As Integer
    Dim StrSQL As String
    
    'Grid Part
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    
    'Header part
    Dim RsHeader As ADODB.Recordset
    Set RsHeader = New ADODB.Recordset
    
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    
    'Handling an exception
    On Error GoTo ErrTrap
    
    '########################################################### Header Part ##################################################
    'StrSQL = "select * from TblEndDebtAgingInv where ID = " & val(RecId.Text)
    'RsHeader.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    txtcode.Text = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
    Me.DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    DtRecord.value = IIf(IsNull(rs("RecDate").value), Date, rs("RecDate").value)
    If (IIf(IsNull(rs("IsClient").value), True, rs("IsClient").value)) = True Then
        Rd(0).value = True
    Else
        Rd(1).value = True
    End If
    If (IIf(IsNull(rs("IsSingleCus").value), True, rs("IsSingleCus").value)) = True Then
        Rd(2).value = True
    Else
        Rd(3).value = True
    End If
    If (IIf(IsNull(rs("IsFIFO").value), True, rs("IsFIFO").value)) = True Then
        Rd(4).value = True
    Else
        Rd(5).value = True
    End If
    TotalPaidLab.Caption = IIf(IsNull(rs("TotalPaid").value), True, rs("TotalPaid").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), True, rs("UserID").value)
    
    
    StrSQL = "SELECT TblCustemers.CusID, TblCustemers.CusName, TblCustemers.CusNamee "
    StrSQL = StrSQL & "FROM TblEndDebtAgingInvDet INNER JOIN TblCustemers ON TblEndDebtAgingInvDet.CusID = TblCustemers.CusID "
    StrSQL = StrSQL & "Where (TblEndDebtAgingInvDet.EndDebAgInvID = " & val(txtcode.Text) & ") And (TblEndDebtAgingInvDet.IsHeaderRec = 1)"
    RsHeader.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    If Rd(3).value = True Then
    SelectedCusList.Clear
        For i = 1 To RsHeader.RecordCount
            If SystemOptions.UserInterface = ArabicInterface Then
                SelectedCusList.AddItem IIf(IsNull(RsHeader("CusName").value), "", RsHeader("CusName").value)
            Else
                SelectedCusList.AddItem IIf(IsNull(RsHeader("CusNamee").value), "", RsHeader("CusNamee").value)
            End If
            SelectedCusList.ItemData(SelectedCusList.NewIndex) = IIf(IsNull(RsHeader("CusID").value), 0, RsHeader("CusID").value)
            
            RsHeader.MoveNext
        Next i
    End If
    
    RsHeader.Close
    Set RsHeader = Nothing
    
   ' FillWindowsControlesData
    
    '##########################################################################################################################
    
    '################################################## Det Part (Grid Part) ##################################################
    StrSQL = "SELECT TblBranchesData.branch_name, TblBranchesData.branch_namee, TblEndDebtAgingInvDet.* "
    StrSQL = StrSQL & " FROM TblEndDebtAgingInvDet LEFT OUTER JOIN "
    StrSQL = StrSQL & " TblBranchesData ON TblEndDebtAgingInvDet.branch_no = TblBranchesData.branch_id "
    StrSQL = StrSQL & " Where (TblEndDebtAgingInvDet.EndDebAgInvID = " & val(txtcode.Text) & ")  and (TblEndDebtAgingInvDet.IsHeaderRec = 0 or TblEndDebtAgingInvDet.IsHeaderRec IS NULL)"
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    With GR
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows
        If Not (RsDetails.BOF Or RsDetails.EOF) Then
            RsDetails.MoveFirst
            .Rows = .FixedRows + RsDetails.RecordCount
            For i = .FixedRows To RsDetails.RecordCount
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("branch_no")) = IIf(IsNull(RsDetails("branch_no").value), 0, RsDetails("branch_no").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDetails("branch_name").value), "", RsDetails("branch_name").value)
                Else
                    .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDetails("branch_namee").value), 0, RsDetails("branch_namee").value)
                End If
                .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(RsDetails("NoteID").value), 0, RsDetails("NoteID").value)
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(RsDetails("NoteSerial1").value), 0, RsDetails("NoteSerial1").value)
                .TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(RsDetails("Note_Value").value), 0, RsDetails("Note_Value").value)
                .TextMatrix(i, .ColIndex("PayedValue")) = IIf(IsNull(RsDetails("PayedValue").value), 0, RsDetails("PayedValue").value)
                .TextMatrix(i, .ColIndex("TransPayedValue")) = IIf(IsNull(RsDetails("TransPayedValue").value), 0, RsDetails("TransPayedValue").value)
                .TextMatrix(i, .ColIndex("too")) = IIf(IsNull(RsDetails("too").value), "", RsDetails("too").value)
                .TextMatrix(i, .ColIndex("NetValue")) = IIf(IsNull(RsDetails("NetValue").value), 0, RsDetails("NetValue").value)
                .TextMatrix(i, .ColIndex("RemainingValue")) = IIf(IsNull(RsDetails("RemainingValue").value), 0, RsDetails("RemainingValue").value)
                .TextMatrix(i, .ColIndex("DueDate")) = IIf(IsNull(((RsDetails("DueDate").value))), " ", ((RsDetails("DueDate").value)))
                .TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(((RsDetails("NoteDate").value))), "", ((RsDetails("NoteDate").value)))
                .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked
                RsDetails.MoveNext
            Next i
        End If
    End With
    
         Dim s As String
     s = "Select * from TblEndDebtAgingInvDet2 Where EndDebAgInvID =" & val(txtcode)
     LoadGrid s, GridCashing, True
       
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Relin
    RsDetails.Close
    Set RsDetails = Nothing
    '#################################################################################################################
ErrTrap:
'******************************** show Error Message *******************************
End Sub
Private Sub DelRecored() '%%%%%%%% Delete current recored %%%%%%%%
    '@@@@@@@@ declear Var @@@@@@@@@
    Dim Msg As String
    Dim StrSQL As String
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    'Handling an exception
    On Error GoTo ErrTrap
    If rs.RecordCount > 0 Then
    DeleteBillBuy
    rs.delete
    StrSQL = "Delete From TblEndDebtAgingInv Where Id = " & val(txtcode.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    StrSQL = "Delete From TblEndDebtAgingInvDet Where EndDebAgInvID = " & val(txtcode.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords
  
   StrSQL = "Delete From TblEndDebtAgingInvDet2 Where EndDebAgInvID = " & val(txtcode.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords
   
  
    StrSQL = "Delete From TblBillBuyPayment2 Where TypTrans=1 and NoteID=" & val(Me.txtcode.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords
   StrSQL = "Delete From TblBillBuyPayment Where TypTrans=1 and NoteID=" & val(Me.txtcode.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    rs.MovePrevious
    End If
    If rs.RecordCount < 1 Then
        GR.Rows = 1
            SelectedCusList.Clear
        clear_all Me
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        TotalPaidLab.Caption = 0
    Else
        GR.Rows = 1
            SelectedCusList.Clear
            TotalPaidLab.Caption = 0
        clear_all Me
        'txtcode.Text = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
      '  RecId.Text = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
        Retrive
    End If
                    
ErrTrap:
'******************************** show Error Message *******************************

End Sub
Private Sub SaveData()
    '@@@@@@@@ declear Var @@@@@@@@@
    Dim StrSQL As String
    Dim i As Integer
    Dim Diff As Double
    Dim Note_Value1 As Double
    Dim RsHeader As ADODB.Recordset
    Set RsHeader = New ADODB.Recordset
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    
    'Handling an exception
    'On Error GoTo ErrTrap
    
    Diff = 0

    'Check if in edit mode or new recored mode
    If Me.TxtModFlg.Text = "E" Then
    DeleteBillBuy
        StrSQL = "Delete From TblEndDebtAgingInvDet Where EndDebAgInvID = " & val(txtcode.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
  StrSQL = "Delete From TblBillBuyPayment Where typTrans=1 and  NoteID=" & val(Me.txtcode.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblBillBuyPayment2 Where TypTrans=1 and NoteID=" & val(Me.txtcode.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords

        StrSQL = "Delete From TblEndDebtAgingInvDet2 Where EndDebAgInvID = " & val(txtcode.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords

         sql = "Delete From TblBillBuyPayment2 Where   Transaction_ID In ("
         
    
   sql = sql & " SELECT      dbo.Transactions.Transaction_ID"
sql = sql & " FROM         dbo.Transactions "
sql = sql & "  WHERE     (dbo.Transactions.PaymentType = 1) AND (dbo.Transactions.Transaction_Type = 21 OR"
sql = sql & "                       dbo.Transactions.Transaction_Type = 2 or dbo.Transactions.Transaction_Type = 71 ) "

   sql = sql & " AND (Transactions.CusID in (" & val(DBCboClientName.BoundText) & ")) )"
 Cn.Execute sql, , adExecuteNoRecords
    End If

    '################################################################## Header Part ##################################################################
    If Me.TxtModFlg.Text = "N" Then
        'get the last id and add one
        Dim str As String
        str = new_id("TblEndDebtAgingInv", "ID", "", True)
        rs.AddNew
        rs("ID").value = str
        txtcode.Text = str
        
    End If
    rs("RecDate").value = DtRecord.value
    If Rd(4).value = True Then
        rs("IsFIFO").value = True
    Else
        rs("IsFIFO").value = False
    End If
    If Rd(0).value = True Then
        rs("IsClient").value = True
    Else
        rs("IsClient").value = False
    End If
    If Rd(2).value = True Then
        rs("IsSingleCus").value = True
    Else
        rs("IsSingleCus").value = False
    End If
    rs("TotalPaid").value = val(TotalPaidLab.Caption)
    rs("CusID").value = val(DBCboClientName.BoundText)
    rs("UserID").value = val(Me.DCboUserName.BoundText)
    
    rs.update
    
    StrSQL = "SELECT * from TblEndDebtAgingInvDet Where (1 = -1)"
    RsHeader.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    'check if single or multi cus
    If Rd(3).value = True Then
        If Me.SelectedCusList.ListCount > 0 Then
            For i = 0 To Me.SelectedCusList.ListCount - 1
                RsHeader.AddNew
                RsHeader("EndDebAgInvID").value = val(txtcode.Text)
                RsHeader("IsHeaderRec").value = True
                RsHeader("CusID").value = val(SelectedCusList.ItemData(i))
                RsHeader.update
            Next i
        End If
    End If
    
    RsHeader.Close
    Set RsHeader = Nothing
                
    'FillWindowsControlesData
     saveBillBuy
    '#################################################################################################################################################

    '############################################################## Det Part (Grid part) #############################################################
    StrSQL = "SELECT * from TblEndDebtAgingInvDet Where (1 = -1)"
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    With GR
        TxtValueTemp.Text = val(TotalPaidLab.Caption)
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
                RsDetails.AddNew
                RsDetails("EndDebAgInvID").value = val(txtcode.Text)
                RsDetails("NoteID1").value = val(txtcode.Text)
                RsDetails("NoteID").value = val(.TextMatrix(i, .ColIndex("NoteID")))
                RsDetails("branch_no").value = val(.TextMatrix(i, .ColIndex("branch_no")))
                RsDetails("NoteSerial1").value = val(.TextMatrix(i, .ColIndex("NoteSerial1")))
                RsDetails("Note_Value").value = val(.TextMatrix(i, .ColIndex("Note_Value")))
                Note_Value1 = val(.TextMatrix(i, .ColIndex("RemainingValue")))
                Diff = 0
                
                If val(TxtValueTemp.Text) > 0 Then
                    If val(TxtValueTemp.Text) <= Note_Value1 Then
                        Diff = val(TxtValueTemp.Text)
                        TxtValueTemp.Text = val(TxtValueTemp.Text) - Note_Value1
                     Else
                        Diff = Note_Value1
                        TxtValueTemp.Text = val(TxtValueTemp.Text) - Note_Value1
                     End If
                End If
               ' .TextMatrix(i, .ColIndex("RemainingValue")) = val(.TextMatrix(i, .ColIndex("RemainingValue"))) - val(.TextMatrix(i, .ColIndex("RemainingValue")))
                .TextMatrix(i, .ColIndex("TransPayedValue")) = Diff
                RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("PayedValue")))
                RsDetails("too").value = (.TextMatrix(i, .ColIndex("too")))
                RsDetails("NoteDate").value = IIf((.TextMatrix(i, .ColIndex("NoteDate"))) = "", Null, (.TextMatrix(i, .ColIndex("NoteDate"))))
                RsDetails("DueDate").value = IIf((.TextMatrix(i, .ColIndex("DueDate"))) = "", Date, (.TextMatrix(i, .ColIndex("DueDate"))))
                RsDetails("TransPayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
                .TextMatrix(i, .ColIndex("NetValue")) = val(.TextMatrix(i, .ColIndex("RemainingValue"))) - val(.TextMatrix(i, .ColIndex("TransPayedValue")))
                RsDetails("NetValue").value = val(.TextMatrix(i, .ColIndex("NetValue")))
                RsDetails("RemainingValue").value = val(.TextMatrix(i, .ColIndex("RemainingValue")))
                
                RsDetails.update
                                    
                If val(val(.TextMatrix(i, .ColIndex("NetValue")))) = 0 Then
                    StrSQL = "Update Transactions Set  TotalPayed=1 Where Transaction_ID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                    Cn.Execute StrSQL, , adExecuteNoRecords
                Else
                    StrSQL = "Update Transactions Set TotalPayed=0 Where Transaction_ID =" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                    Cn.Execute StrSQL, , adExecuteNoRecords
                End If
            End If
        Next i
    End With
   
    
    RsDetails.Close
    Set RsDetails = Nothing
    Dim s As String
     s = "Select * from TblEndDebtAgingInvDet2 Where EndDebAgInvID =" & val(txtcode)
     saveGrid s, GridCashing, "NoteSerial1", "", "EndDebAgInvID", val(Me.txtcode.Text)
    
'''//////////////
'Set RsDetails = New ADODB.Recordset
'If Rd(0).value = True Then
'    StrSQL = "SELECT     * from dbo.TblBillBuyPayment2 Where (1 = -1)"
' Else
' StrSQL = "SELECT     * from dbo.TblBillBuyPayment Where (1 = -1)"
' End If
'
'   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
'
'    With GR
'    For i = .FixedRows To .Rows - 1
'        If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
'            RsDetails.AddNew
'            RsDetails("TypTrans").value = 1
'            RsDetails("NoteID").value = val(Me.txtcode.Text)
'            RsDetails("Transaction_ID").value = val(.TextMatrix(i, .ColIndex("NoteID")))
'            RsDetails("Note_Value").value = val(.TextMatrix(i, .ColIndex("Note_Value")))
'            RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
'            RsDetails.update
'        End If
'    Next i
'End With


saveBillBuy
    '#############################################################################################################################################
    If TxtModFlg.Text = "N" Then
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Data Saved Successfully" & CHR(13)
        Else
            Msg = " „ Õ›Ÿ «·»Ì«‰« " & CHR(13)
        End If
        
    ElseIf TxtModFlg.Text = "E" Then
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Data Edited Successfully" & CHR(13)
        Else
            Msg = " „  ⁄œÌ· «·»Ì«‰« " & CHR(13)
        End If
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Me.TxtModFlg.Text = "R"
    XPTxtCurrent.Caption = rs.RecordCount
    XPTxtCount.Caption = rs.RecordCount
ErrTrap:
'******************************** show Error Message *******************************
Retrive
End Sub
Private Sub PrintReport()
    On Error GoTo ErrTrap

    If XPTxtBillID.Text <> "" Then
        Set SaleReport = New ClsSaleReport
        SaleReport.ShowPrice XPTxtBillID.Text, 94, DcboEmp.Text
    End If
ErrTrap:
'******************************** show Error Message *******************************
End Sub

Private Sub GR_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With GR

Select Case .ColKey(Col)
Case "TransPayedValue"
If val(.TextMatrix(Row, .ColIndex("TransPayedValue"))) > val(.TextMatrix(Row, .ColIndex("RemainingValue"))) Then
            If SystemOptions.UserInterface = ArabicInterface Then
              MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ ﬁÌ„… «·œ›⁄… «ﬂ»— „‰ «·„ »ﬁÌ"
              Else
              MsgBox "Can Not PaymentValue Larger Than Total Value "
              End If
              .TextMatrix(Row, .ColIndex("TransPayedValue")) = 0
              Exit Sub
              End If
End Select
End With
Relin
End Sub
Sub DeleteBillBuy()
Dim i As Integer
Dim StrSQL As String
With Me.GR
 For i = .FixedRows To .Rows - 1
 If val(.TextMatrix(i, .ColIndex("NoteID"))) <> 0 Then
      StrSQL = "Update Transactions Set  TotalPayed=0 Where Transaction_ID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
     End If
     Next i
 End With
End Sub
Sub Relin()
Dim i As Integer
Dim SmValu As Double
SmValu = 0
With GridCashing
For i = 1 To .Rows - 1
    If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
        If val(.TextMatrix(i, .ColIndex("Value"))) <> 0 Then
            SmValu = SmValu + val(.TextMatrix(i, .ColIndex("Value")))
        End If
    End If
Next i
End With
TotalPaidLab.Caption = SmValu
End Sub

Private Sub GR_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Rd(4).value = True Then
Cancel = True
End If
With GR
Select Case .ColKey(Col)
Case "NoteSerial1"
Cancel = True
Case "too"
Cancel = True
Case "NoteDate"
Cancel = True
Case "DueDate"
Cancel = True
Case "branch_name"
Cancel = True
Case "Note_Value"
Cancel = True
Case "PayedValue"
Cancel = True
Case "NetValue"
Cancel = True
Case "RemainingValue"
Cancel = True
Case "TransPayedValue"
If .Cell(flexcpChecked, Row, .ColIndex("payed")) = flexChecked Then
.ComboList = ""
Else
Cancel = True
End If
End Select
End With
End Sub

Private Sub GridCashing_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Relin
End Sub

Private Sub GridCashing_Click()
Relin
End Sub

Private Sub TxtModFlg_Change() ' %%%%%%%% Set Windows Stutes %%%%%%%%%
Rd(0).Enabled = False
Rd(1).Enabled = False
    If Me.TxtModFlg.Text = "N" Then
    '################### case new recored ########################
    Rd(0).Enabled = True
    Rd(1).Enabled = True
    txtcode.Enabled = True
    CusTypeFrame.Enabled = True
    Rd(3).Enabled = True
    Rd(2).Enabled = True
    GroupCusFram.Enabled = True
    SingleCusFrame.Enabled = True
    PayTypeFram.Enabled = True
    ShowInvData.Enabled = True
    GR.Enabled = True
    Cmd(0).Enabled = False
    Cmd(1).Enabled = False
    Cmd(2).Enabled = True
    Cmd(3).Enabled = True
    Cmd(4).Enabled = False
    Cmd(5).Enabled = False
    Cmd(6).Enabled = True
    Cmd(7).Enabled = True
    btnFirst.Enabled = False
    btnPrevious.Enabled = False
    btnNext.Enabled = False
    btnLast.Enabled = False
    '#############################################################
    ElseIf Me.TxtModFlg.Text = "E" Then
    '################### case edit recored #######################
    txtcode.Enabled = True
    CusTypeFrame.Enabled = True
    Rd(3).Enabled = True
    Rd(2).Enabled = True
    GroupCusFram.Enabled = True
    SingleCusFrame.Enabled = True
    PayTypeFram.Enabled = True
    ShowInvData.Enabled = True
    GR.Enabled = True
    Cmd(0).Enabled = False
    Cmd(1).Enabled = False
    Cmd(2).Enabled = True
    Cmd(3).Enabled = True
    Cmd(4).Enabled = False
    Cmd(5).Enabled = False
    btnFirst.Enabled = False
    btnPrevious.Enabled = False
    btnNext.Enabled = False
    btnLast.Enabled = False
    '#############################################################
    ElseIf Me.TxtModFlg.Text = "R" Then
    '################### case read recored #######################
    ' lock all fields show only
    txtcode.Enabled = False
    CusTypeFrame.Enabled = False
    Rd(3).Enabled = False
    Rd(2).Enabled = False
    GroupCusFram.Enabled = False
    SingleCusFrame.Enabled = False
    PayTypeFram.Enabled = False
    ShowInvData.Enabled = False
    GR.Enabled = False
    Cmd(0).Enabled = True
    Cmd(1).Enabled = True
    Cmd(2).Enabled = False
    Cmd(3).Enabled = False
    Cmd(4).Enabled = True
    Cmd(5).Enabled = True
    Cmd(6).Enabled = True
    Cmd(7).Enabled = True
    btnFirst.Enabled = True
    btnPrevious.Enabled = True
    btnNext.Enabled = True
    btnLast.Enabled = True
    '#############################################################
    End If
End Sub
Private Sub Rd_Click(Index As Integer) '%%%%%%%% Change the UI and Drop list with User choesing  %%%%%%%%
    '@@@@@@@@ declear Var @@@@@@@@@
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    Select Case Index
        Case 0
        GR.Clear flexClearScrollable, flexClearEverything
        GR.Rows = GR.FixedRows
        CusCodeText.Text = ""
            '&&&&&&&&&&&&&&&&& UI part &&&&&&&&&&&&&&&&
            Rd(3).Caption = "·„Ã„Ê⁄… ⁄„·«¡"
            Rd(2).Caption = "·⁄„Ì· „Õœœ"
            With GR
                .TextMatrix(0, .ColIndex("too")) = "›« Ê—… «·⁄„Ì·"
            End With
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
            
            '&&&&&&&&&&&&& update UI &&&&&&&&&&&&&&&&&&
            FillWindowsControlesData
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        Case 1
         GR.Clear flexClearScrollable, flexClearEverything
        GR.Rows = GR.FixedRows
        CusCodeText.Text = ""
            '&&&&&&&&&&&&&&&&& UI part &&&&&&&&&&&&&&&&
            Rd(3).Caption = "·„Ã„Ê⁄… „Ê—œÌ‰"
            Rd(2).Caption = "·„Ê—œ „Õœœ"
            With GR
                .TextMatrix(0, .ColIndex("too")) = "›« Ê—… «·„Ê—œ"
            End With
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
            '&&&&&&&&&&&&& update UI &&&&&&&&&&&&&&&&&&
            FillWindowsControlesData
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        
            '&&&&&&&&&&&& Drop List Part &&&&&&&&&&&&&&
            Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    End Select
    
    Select Case Index
        Case 3
            DBCboClientName.BoundText = 0
            CusCodeText.Text = ""
        
            GroupCusFram.Enabled = True
            SingleCusFrame.Enabled = False
            '&&&&&&&&&&&&& update UI &&&&&&&&&&&&&&&&&&
            FillWindowsControlesData
            
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        Case 2
            CusList.Clear
            SelectedCusList.Clear
            
        
            GroupCusFram.Enabled = False
            SingleCusFrame.Enabled = True
            '&&&&&&&&&&&&& update UI &&&&&&&&&&&&&&&&&&
            FillWindowsControlesData
            '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
        Case 4, 5
     '   ShowInvData_Click
    End Select
End Sub
Private Sub DBCboClientName_Click(Area As Integer) '%%%%% bound cus code with code box %%%%
    '@@@@@@@@ declear Var @@@@@@@@@
    Dim CusCode  As String
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    ' in case no data
    If val(DBCboClientName.BoundText) = 0 Then Exit Sub
    
    'get selected cus code and show it code text box
    GetTblCustemersCode , , DBCboClientName.BoundText, CusCode
    Me.CusCodeText.Text = CusCode
End Sub

Private Sub CusCodeText_KeyPress(KeyAscii As Integer) '%%%%%% get cus associated with enter code %%%%%%
    '@@@@@@@@ declear Var @@@@@@@@@
    Dim CusCode As Integer
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    
    'get cus only when press enter
    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode CusCodeText.Text, CusCode
        DBCboClientName.BoundText = CusCode
    End If
End Sub

Private Sub SelectSingleCus_Click() '%%%%%% Add one Employee at a time %%%%%%
    '@@@@@@@@ declear Var @@@@@@@@@
    Dim Rs1  As ADODB.Recordset
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    If Me.CusList.ListIndex > -1 Then
        Me.SelectedCusList.AddItem CusList.List(CusList.ListIndex)
        SelectedCusList.ItemData(SelectedCusList.NewIndex) = CusList.ItemData(CusList.ListIndex)
    End If
End Sub
Private Sub SelectAllCus_Click()   '%%%%%% Add All Cus %%%%%%
    '@@@@@@@@ declear Var @@@@@@@@@
    Dim i As Integer
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    Me.SelectedCusList.Clear
    For i = 0 To Me.CusList.ListCount - 1
        Me.SelectedCusList.AddItem CusList.List(i)
        SelectedCusList.ItemData(i) = CusList.ItemData(i)
    Next i
End Sub
Private Sub RemoveSingleCus_Click()
    If SelectedCusList.ListIndex > -1 Then
        SelectedCusList.RemoveItem (SelectedCusList.ListIndex)
    End If
End Sub
Private Sub RemoveAllCus_Click()
    SelectedCusList.Clear
End Sub
Private Sub ShowInvData_Click()
    '@@@@@@@@ declear Var @@@@@@@@@
    Dim sql As String
    Dim Rs8 As ADODB.Recordset
    Dim i As Integer
    Set Rs8 = New ADODB.Recordset
    Dim CuID As Integer
    Dim j As Integer
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
If Me.TxtModFlg.Text <> "R" Then
    CUIDs = "0"
    If Rd(3).value = True Then
        If Me.SelectedCusList.ListCount > 0 Then
            For i = 0 To Me.SelectedCusList.ListCount - 1
                CUIDs = CUIDs & "," & SelectedCusList.ItemData(i)
            Next
        End If
        If CUIDs = "0" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "«·—Ã«¡  ÕœÌœ „” Œœ„ Ê«Õœ ⁄·Ï «·√ﬁ·"
            Else
                Msg = "Please select at least one user "
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
    ElseIf Rd(2).value = True Then
        CUIDs = DBCboClientName.BoundText
    End If
    
    
    With GR
        .Clear flexClearScrollable, flexClearEverything
        .Rows = 1
    End With

    sql = "Delete From TblBillBuyPayment2 Where   Transaction_ID In ("
         
    
   sql = sql & " SELECT      dbo.Transactions.Transaction_ID"
sql = sql & " FROM         dbo.Transactions "
sql = sql & "  WHERE     (dbo.Transactions.PaymentType = 1) AND (dbo.Transactions.Transaction_Type = 21 OR"
sql = sql & "                       dbo.Transactions.Transaction_Type = 2 or dbo.Transactions.Transaction_Type = 71 ) "

   sql = sql & " AND (Transactions.CusID in (" & CUIDs & ")) )"
    
         Cn.Execute sql, , adExecuteNoRecords
             sql = "Delete From TblNotesBillBuyPayment2    Where  NoteID In ("
              
         
     sql = sql & "  SELECT     dbo.Notes.NoteID "
    sql = sql & "  FROM         dbo.Notes INNER JOIN"
    sql = sql & "                        dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
    sql = sql & "                       dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType"
    sql = sql & "  WHERE     (dbo.Notes.NoteType = 4 OR dbo.Notes.NoteType = 200 or dbo.Notes.NoteType = 57 Or "
    sql = sql & "                       dbo.Notes.NoteType = 9 OR"
    sql = sql & "                       dbo.Notes.NoteType = 220) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = '" & AccountCode & "'))"
    
         Cn.Execute sql, , adExecuteNoRecords
         

    sql = "SELECT TOP 100 PERCENT Transactions.Transaction_ID, Transactions.Transaction_Date, Transactions.Transaction_Type, Transactions.NoteSerial1,"
    sql = sql & "Transactions.ManualNO, Transactions.BranchId, TblBranchesData.branch_name, TblBranchesData.branch_namee, Transactions.CusID, "
    sql = sql & "TblCustemers.CusName, TblCustemers.CusNamee, TblCustemers.Fullcode, Transactions.TotalPayed, Transactions.OldContID, "
    sql = sql & "Transactions.OldValue , Transactions.DueDate, Transactions.Transaction_NetValue "
    sql = sql & "FROM Transactions LEFT OUTER JOIN "
    sql = sql & "TblCustemers ON Transactions.CusID = TblCustemers.CusID LEFT OUTER JOIN TblBranchesData ON Transactions.BranchId = TblBranchesData.branch_id "
    sql = sql & "WHERE (1 = 1) "
    If Rd(0).value = True Then
    sql = sql & " and (Transactions.PaymentType = 1) "
    sql = sql & " AND (Transactions.Transaction_Type = 21 OR Transactions.Transaction_Type = 2 OR  Transactions.Transaction_Type = 71)"
    Else
    sql = sql & " AND (Transactions.Transaction_Type = 1 OR  Transactions.Transaction_Type = 22 OR  Transactions.Transaction_Type = 73 )"
    End If
   ' sql = sql & " AND (Transactions.TotalPayed IS NULL OR "
   ' sql = sql & "Transactions.TotalPayed = 0) "
    
    sql = sql & " AND (Transactions.CusID in (" & CUIDs & ")) "
    sql = sql & "ORDER BY Transactions.DueDate, dbo.Transactions.NoteSerial1"


sql = " SELECT     TOP 100 PERCENT dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.Transactions.NoteSerial1, "
sql = sql & "                      dbo.Transactions.ManualNO, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.Transactions.CusID,"
sql = sql & "                      dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.Transactions.TotalPayed, dbo.Transactions.OldContID,"
sql = sql & "                      dbo.transactions.OldValue , dbo.transactions.dueDate, dbo.transactions.Vat, dbo.transactions.Transaction_NetValue"
sql = sql & " FROM         dbo.Transactions LEFT OUTER JOIN"
sql = sql & "                      dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
sql = sql & "  WHERE     (dbo.Transactions.PaymentType = 1) AND (dbo.Transactions.Transaction_Type = 21 OR"
sql = sql & "                       dbo.Transactions.Transaction_Type = 2 or dbo.Transactions.Transaction_Type = 71 ) AND (dbo.Transactions.TotalPayed IS NULL OR"
sql = sql & "                       dbo.Transactions.TotalPayed = 0) "
   sql = sql & " AND (Transactions.CusID in (" & CUIDs & ")) "
sql = sql & "  ORDER BY dbo.Transactions.DueDate ,dbo.Transactions.NoteSerial1"


    Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs8.RecordCount > 0 Then
        GR.Enabled = True
        GR.Enabled = True
        With GR
            .Clear flexClearScrollable, flexClearEverything
            .Rows = 1
            .Rows = .Rows + Rs8.RecordCount
            .Rows = .FixedRows + Rs8.RecordCount
            Rs8.MoveFirst
            For i = .FixedRows To Rs8.RecordCount
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("branch_no")) = IIf(IsNull(Rs8("BranchId").value), 0, Rs8("BranchId").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), 0, Rs8("branch_name").value)
                Else
                    .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), 0, Rs8("branch_namee").value)
                End If
                .TextMatrix(i, .ColIndex("DueDate")) = IIf(IsNull(Rs8("DueDate").value), "", Rs8("DueDate").value)
                .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(Rs8("Transaction_ID").value), 0, Rs8("Transaction_ID").value)
                .TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(Rs8("Transaction_Date").value), "", Rs8("Transaction_Date").value)
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(Rs8("NoteSerial1").value), "", Rs8("NoteSerial1").value)
                .TextMatrix(i, .ColIndex("too")) = IIf(IsNull(Rs8("ManualNO").value), "", Rs8("ManualNO").value)
                .TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(Rs8("Transaction_NetValue").value), IIf(IsNull(Rs8("OldValue").value), 0, Rs8("OldValue").value), Rs8("Transaction_NetValue").value)
  'salimhere
  
                If val(.TextMatrix(i, .ColIndex("NoteID"))) <> 0 Then
                    .TextMatrix(i, .ColIndex("PayedValue")) = GeteBillBuy(val(.TextMatrix(i, .ColIndex("NoteID"))))
                Else
                    .TextMatrix(i, .ColIndex("PayedValue")) = 0
                End If
                
                .TextMatrix(i, .ColIndex("RemainingValue")) = val(.TextMatrix(i, .ColIndex("Note_Value"))) - val(.TextMatrix(i, .ColIndex("PayedValue")))
                If Rd(4).value = True Then
                .TextMatrix(i, .ColIndex("TransPayedValue")) = val(.TextMatrix(i, .ColIndex("Note_Value"))) - val(.TextMatrix(i, .ColIndex("PayedValue")))
                .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked
                Else
                .Cell(flexcpChecked, i, .ColIndex("payed")) = flexUnchecked
                .TextMatrix(i, .ColIndex("TransPayedValue")) = 0
                End If
                Rs8.MoveNext
            Next i
        End With
    End If
    Rs8.Close
    Set Rs8 = Nothing
  End If
  Relin



'part2 Show Cashing
'Dim AccountCode As String

 AccountCode = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code")
     
         

Set Rs8 = New ADODB.Recordset

    With GridCashing
        .Clear flexClearScrollable, flexClearEverything
        .Rows = 1
    End With

     If Rd(0).value = True Then
  
 sql = "SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value], dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate, dbo.Notes.NoteID, "
    sql = sql & "                       dbo.TblNotesTypes.NotesTypeName , dbo.TblNotesTypes.NotesTypeNameE, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1, Notes.Remark"
    sql = sql & "  FROM         dbo.Notes INNER JOIN"
    sql = sql & "                        dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
    sql = sql & "                       dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType"
    sql = sql & "  WHERE     (dbo.Notes.NoteType = 4 OR dbo.Notes.NoteType = 200 or dbo.Notes.NoteType = 57 Or "
    sql = sql & "                       dbo.Notes.NoteType = 9 OR"
    sql = sql & "                       dbo.Notes.NoteType = 220) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = '" & AccountCode & "')"
    sql = sql & "  order by dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate, dbo.Notes.NoteSerial"
 
 
    Else
  
    End If
   ' sql = sql & " AND (Transactions.TotalPayed IS NULL OR "
   ' sql = sql & "Transactions.TotalPayed = 0) "
    
'    sql = sql & " AND (Transactions.CusID in (" & CUIDs & ")) "
'    sql = sql & "ORDER BY Transactions.DueDate, dbo.Transactions.NoteSerial1"


    Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs8.RecordCount > 0 Then
        GridCashing.Enabled = True
        GridCashing.Enabled = True
        With GridCashing
            .Clear flexClearScrollable, flexClearEverything
            .Rows = 1
            .Rows = .Rows + Rs8.RecordCount
            .Rows = .FixedRows + Rs8.RecordCount
            Rs8.MoveFirst
            For i = .FixedRows To Rs8.RecordCount
                .TextMatrix(i, .ColIndex("Ser")) = i
         '       .TextMatrix(i, .ColIndex("branch_no")) = IIf(IsNull(Rs8("BranchId").value), 0, Rs8("BranchId").value)
         '       If SystemOptions.UserInterface = ArabicInterface Then
         '           .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), 0, Rs8("branch_name").value)
         '       Else
         '           .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), 0, Rs8("branch_namee").value)
         '       End If
         '       .TextMatrix(i, .ColIndex("DueDate")) = IIf(IsNull(Rs8("DueDate").value), "", Rs8("DueDate").value)
                
                
                  
                .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(Rs8("Value").value), 0, Rs8("Value").value)
                .TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(Rs8("RecordDate").value), "", Rs8("RecordDate").value)
                .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(Rs8("NoteID").value), 0, Rs8("NoteID").value)
                .TextMatrix(i, .ColIndex("NotesTypeName")) = IIf(IsNull(Rs8("NotesTypeName").value), 0, Rs8("NotesTypeName").value)
                .TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(Rs8("NoteSerial").value), "", Rs8("NoteSerial").value)
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(Rs8("NoteSerial1").value), "", Rs8("NoteSerial1").value)
                .TextMatrix(i, .ColIndex("remark")) = IIf(IsNull(Rs8("remark").value), "", Rs8("remark").value)
       
                Rs8.MoveNext
            Next i
        End With
        
    
    
    End If
    Rs8.Close
    Set Rs8 = Nothing
   
End Sub
Function GeteBillBuy2(Optional Transaction_ID As Double = 0) As Double
    '@@@@@@@@ declear Var @@@@@@@@@
    Dim sql As String
    Dim Rs8 As ADODB.Recordset
    Set Rs8 = New ADODB.Recordset
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    
    If Rd(0).value = True Then
   sql = " SELECT   SUM(TransPayedValue) AS Smatiobn"
    sql = sql & " From dbo.TblNotesBillBuyPayment2"
     sql = sql & " Where (noteid = " & Transaction_ID & ")"
    sql = sql & " GROUP BY notedate"
    
    Else
    sql = " SELECT   SUM(PayedValue) AS Smatiobn"
    sql = sql & " From dbo.TblBillBuyPayment"
     sql = sql & " Where (Transaction_ID = " & Transaction_ID & ")"
    sql = sql & " GROUP BY Transaction_ID"
    End If
   
    Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs8.RecordCount > 0 Then
        GeteBillBuy = IIf(IsNull(Rs8("Smatiobn").value), 0, Rs8("Smatiobn").value)
    Else
        GeteBillBuy = 0
    End If
    Rs8.Close
    Set Rs8 = Nothing
End Function

Function GeteBillBuy(Optional Transaction_ID As Double = 0) As Double
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
sql = " SELECT   SUM(PayedValue) AS Smatiobn"
sql = sql & " From dbo.TblBillBuyPayment2"
sql = sql & " Where (Transaction_ID = " & Transaction_ID & ")"
sql = sql & " GROUP BY Transaction_ID"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
GeteBillBuy = IIf(IsNull(Rs8("Smatiobn").value), 0, Rs8("Smatiobn").value)
Else
GeteBillBuy = 0
End If
End Function

Sub FillWindowsControlesData()
'############################################### Cus List Part ##################################################
    '@@@@@@@@ declear Var @@@@@@@@@
    Dim rs2 As ADODB.Recordset
    Dim i As Integer
    Set rs2 = New ADODB.Recordset
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    
    'only get data only when in edit or new mode as will as not single client mode
    If Rd(3).value = True Then
        'check client or supplier
        If Rd(0).value = True Then
            sql = "SELECT * from  TblCustemers where Type = 1 and CusID <> 2"
        ElseIf Rd(1).value = True Then
            sql = " SELECT * from  TblCustemers where Type = 2 and CusID <> 1"
        End If
        
        rs2.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
        Me.CusList.Clear
        'Me.SelectedCusList.Clear
        If rs2.RecordCount > 0 Then
            For i = 1 To rs2.RecordCount
                If SystemOptions.UserInterface = ArabicInterface Then
                    CusList.AddItem IIf(IsNull(rs2("CusName").value), "", rs2("CusName").value)
                Else
                    CusList.AddItem IIf(IsNull(rs2("CusNamee").value), "", rs2("CusNamee").value)
                End If
                CusList.ItemData(CusList.NewIndex) = IIf(IsNull(rs2("CusID").value), 0, rs2("CusID").value)
            rs2.MoveNext
            Next i
        End If
        rs2.Close
        Set rs2 = Nothing
        
    End If
    '#############################################################################################################
    
    '################################## Users Part ###############################
    Dcombos.GetUsers Me.DCboUserName
    '#############################################################################
    
    '################################## Cus Part #################################
    If Rd(0) = True Then
        Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName
    ElseIf Rd(1) = True Then
        Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName
    End If
    '#############################################################################
End Sub
