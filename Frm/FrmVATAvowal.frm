VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmVATAvowal 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‰„Ê–Ã «·«ﬁ—«— «·÷—»Ì "
   ClientHeight    =   10350
   ClientLeft      =   6705
   ClientTop       =   1620
   ClientWidth     =   15015
   Icon            =   "FrmVATAvowal.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10350
   ScaleWidth      =   15015
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   10350
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   15015
      _cx             =   26485
      _cy             =   18256
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
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   285
         Left            =   0
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   15120
         _cx             =   26670
         _cy             =   503
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   22.5
            Charset         =   178
            Weight          =   700
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
         CaptionStyle    =   1
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
            Height          =   345
            Left            =   2250
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1155
            TabIndex        =   5
            Top             =   0
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
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
            ButtonImage     =   "FrmVATAvowal.frx":038A
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
            Height          =   345
            Index           =   2
            Left            =   90
            TabIndex        =   6
            Top             =   0
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
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
            ButtonImage     =   "FrmVATAvowal.frx":0724
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
            Height          =   345
            Index           =   1
            Left            =   1680
            TabIndex        =   7
            Top             =   0
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
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
            ButtonImage     =   "FrmVATAvowal.frx":0ABE
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
            Height          =   345
            Index           =   3
            Left            =   615
            TabIndex        =   8
            Top             =   0
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
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
            ButtonImage     =   "FrmVATAvowal.frx":0E58
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
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "‰„Ê–Ã «·«ﬁ—«— «·÷—»Ì "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   9600
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   0
            Width           =   5520
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic8 
         Height          =   570
         Left            =   0
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   120
         Width           =   15120
         _cx             =   26670
         _cy             =   1005
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
         Begin VB.TextBox ID 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   13350
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   165
            Width           =   870
         End
         Begin MSComCtl2.DTPicker RecoredDate 
            Height          =   315
            Left            =   10125
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   165
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   151781379
            CurrentDate     =   37140
         End
         Begin MSDataListLib.DataCombo DcBranch 
            Height          =   315
            Left            =   180
            TabIndex        =   13
            Top             =   165
            Width           =   8025
            _ExtentX        =   14155
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·—ﬁ„"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   14385
            TabIndex        =   280
            Top             =   225
            Width           =   690
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   330
            Index           =   40
            Left            =   8580
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   195
            Width           =   1275
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«· «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   12615
            TabIndex        =   15
            Top             =   165
            Width           =   690
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”·”·"
            Height          =   330
            Index           =   8
            Left            =   17940
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   225
            Width           =   1680
         End
      End
      Begin C1SizerLibCtl.C1Tab C1Tab2 
         Height          =   8325
         Left            =   0
         TabIndex        =   17
         Top             =   1020
         Width           =   15105
         _cx             =   26644
         _cy             =   14684
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
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483633
         BackTabColor    =   14871017
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "«·«ﬁ—«—|»Ì«‰«   Õ·Ì·ÌÂ ··‰”»… «·«”«”Ì… |»Ì«‰«   Õ·Ì·ÌÂ ··‰”»… «·’›—Ì…|»Ì«‰«   Õ··Ì·Ì ··⁄„·Ì«  «·„⁄›«Â|„·«ÕŸ«  Â«„…"
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   7950
            Left            =   16050
            TabIndex        =   128
            TabStop         =   0   'False
            Top             =   45
            Width           =   15015
            _cx             =   26485
            _cy             =   14023
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
            Begin VB.TextBox TxtProjCusValuezero 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   5805
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   260
               Top             =   2295
               Width           =   2040
            End
            Begin VB.TextBox TxtProjCusValueRetzero 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   270
               Left            =   3615
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   259
               Top             =   2295
               Width           =   2040
            End
            Begin VB.TextBox TxtMinisterReValuez 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   3615
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   247
               Top             =   1710
               Width           =   2040
            End
            Begin VB.TextBox SalesRetZero 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   270
               Left            =   3615
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   246
               Top             =   1200
               Width           =   2040
            End
            Begin VB.TextBox manulaSAlesZeroRet 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   3615
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   245
               Top             =   645
               Width           =   2040
            End
            Begin VB.TextBox txtmanulPurcahsezeroRetur 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   3615
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   244
               Top             =   4695
               Width           =   2040
            End
            Begin VB.TextBox TxtPurchaseZeroRet 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   3615
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   243
               Top             =   5175
               Width           =   2040
            End
            Begin VB.TextBox TotalRetSalesZero 
               Alignment       =   2  'Center
               BackColor       =   &H0080FF80&
               Enabled         =   0   'False
               Height          =   255
               Left            =   3615
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   242
               Top             =   2835
               Width           =   2040
            End
            Begin VB.TextBox TotalReturnPurchaseZero 
               Alignment       =   2  'Center
               BackColor       =   &H0080C0FF&
               Enabled         =   0   'False
               Height          =   300
               Left            =   3615
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   241
               Top             =   6630
               Width           =   2040
            End
            Begin VB.TextBox TxtprojectsuppRet 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   3615
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   240
               Top             =   5565
               Width           =   2040
            End
            Begin VB.TextBox Txtprojectsupp 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   5805
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   165
               Top             =   5565
               Width           =   2040
            End
            Begin VB.TextBox TotalPurchaseZero 
               Alignment       =   2  'Center
               BackColor       =   &H0080C0FF&
               Enabled         =   0   'False
               Height          =   300
               Left            =   5805
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   163
               Top             =   6630
               Width           =   2040
            End
            Begin VB.TextBox TotalSalesZero 
               Alignment       =   2  'Center
               BackColor       =   &H0080FF80&
               Enabled         =   0   'False
               Height          =   255
               Left            =   5805
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   161
               Top             =   2835
               Width           =   2040
            End
            Begin VB.TextBox TxtPurchaseZero 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   5805
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   157
               Top             =   5205
               Width           =   2040
            End
            Begin VB.TextBox txtmanulPurcahsezero 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   5805
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   156
               Top             =   4695
               Width           =   2040
            End
            Begin VB.TextBox manulaSAlesZero 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   5805
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   154
               Top             =   645
               Width           =   2040
            End
            Begin VB.TextBox SalesZero 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   270
               Left            =   5805
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   140
               Top             =   1200
               Width           =   2040
            End
            Begin VB.TextBox TxtMinisterValuez 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   5805
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   129
               Top             =   1680
               Width           =   2040
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ⁄œÌ·« "
               Height          =   240
               Index           =   8
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   313
               Top             =   4395
               Width           =   1785
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„»·€"
               Height          =   240
               Index           =   7
               Left            =   6060
               RightToLeft     =   -1  'True
               TabIndex        =   312
               Top             =   4320
               Width           =   1785
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ⁄œÌ·« "
               Height          =   240
               Index           =   6
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   311
               Top             =   315
               Width           =   1785
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„»·€"
               Height          =   240
               Index           =   5
               Left            =   5820
               RightToLeft     =   -1  'True
               TabIndex        =   310
               Top             =   240
               Width           =   1785
            End
            Begin VB.Label Label85 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„” Œ·’«  «·„‘«—Ì⁄ ··⁄„·«¡"
               Height          =   210
               Left            =   9360
               RightToLeft     =   -1  'True
               TabIndex        =   261
               Top             =   2340
               Width           =   2565
            End
            Begin VB.Line Line8 
               X1              =   19020
               X2              =   0
               Y1              =   3990
               Y2              =   3990
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«Ã„«·Ì"
               Height          =   270
               Index           =   13
               Left            =   9360
               RightToLeft     =   -1  'True
               TabIndex        =   164
               Top             =   6630
               Width           =   2565
            End
            Begin VB.Label Label64 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«Ã„«·Ì"
               Height          =   240
               Left            =   9360
               RightToLeft     =   -1  'True
               TabIndex        =   162
               Top             =   2835
               Width           =   2565
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„” Œ·’«  «·„‘«—Ì⁄ ··„ﬁ«Ê·"
               Height          =   195
               Index           =   12
               Left            =   9585
               RightToLeft     =   -1  'True
               TabIndex        =   160
               Top             =   5610
               Width           =   2115
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "›Ê« Ì— «·„‘ —Ì« "
               Height          =   300
               Index           =   11
               Left            =   9585
               RightToLeft     =   -1  'True
               TabIndex        =   159
               Top             =   5190
               Width           =   2115
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«œŒ«·«  ÌœÊÌ…"
               Height          =   180
               Index           =   10
               Left            =   9285
               RightToLeft     =   -1  'True
               TabIndex        =   158
               Top             =   4770
               Width           =   2715
            End
            Begin VB.Label Label55 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«œŒ«·«  ÌœÊÌ…"
               Height          =   210
               Left            =   9285
               RightToLeft     =   -1  'True
               TabIndex        =   155
               Top             =   675
               Width           =   2640
            End
            Begin VB.Label Label47 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "›Ê« Ì— «·„»Ì⁄«  «·„Õ·Ì… «·Œ«÷⁄… ··‰”»… «·’›—Ì… "
               Height          =   285
               Left            =   8985
               RightToLeft     =   -1  'True
               TabIndex        =   141
               Top             =   1185
               Width           =   3240
            End
            Begin VB.Label Label35 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   " Õ·Ì· «·÷—Ì»… ⁄·Ï  «·„‘ —Ì« 0%"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000080FF&
               Height          =   420
               Left            =   8445
               RightToLeft     =   -1  'True
               TabIndex        =   133
               Top             =   4155
               Width           =   4455
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«” Õﬁ«ﬁ «·Ê“«—Ì 0%"
               Height          =   195
               Left            =   9360
               RightToLeft     =   -1  'True
               TabIndex        =   131
               Top             =   1770
               Width           =   2565
            End
            Begin VB.Label Label34 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   " Õ·Ì· «·÷—Ì»… ⁄·Ï  «·„»Ì⁄«  0%"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   390
               Left            =   8670
               RightToLeft     =   -1  'True
               TabIndex        =   130
               Top             =   120
               Width           =   4005
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   7950
            Left            =   45
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   45
            Width           =   15015
            _cx             =   26485
            _cy             =   14023
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
            Begin VB.CheckBox ChkIsFree 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   960
               RightToLeft     =   -1  'True
               TabIndex        =   321
               Top             =   2250
               Width           =   1095
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   "«·—’Ìœ Õ Ì "
               Height          =   255
               Left            =   6345
               RightToLeft     =   -1  'True
               TabIndex        =   297
               Top             =   7200
               Width           =   1200
            End
            Begin VB.TextBox TxtNetVat 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0C0&
               Enabled         =   0   'False
               ForeColor       =   &H80000005&
               Height          =   300
               Left            =   825
               MultiLine       =   -1  'True
               TabIndex        =   206
               Top             =   7515
               Width           =   3705
            End
            Begin VB.TextBox TxtOldVat 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0C0&
               Enabled         =   0   'False
               ForeColor       =   &H80000005&
               Height          =   300
               Left            =   825
               MultiLine       =   -1  'True
               TabIndex        =   205
               Top             =   7185
               Width           =   3705
            End
            Begin VB.TextBox TxtCorrect1 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H80000001&
               Height          =   315
               Left            =   825
               MultiLine       =   -1  'True
               TabIndex        =   204
               Top             =   6735
               Width           =   3705
            End
            Begin VB.TextBox NotesTxt 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   7245
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   8835
               Width           =   6030
            End
            Begin VB.TextBox TotalNetTxt 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0C0&
               Enabled         =   0   'False
               ForeColor       =   &H80000005&
               Height          =   285
               Left            =   825
               MultiLine       =   -1  'True
               TabIndex        =   54
               Top             =   6435
               Width           =   3705
            End
            Begin VB.TextBox VATPurchasesTotal 
               Alignment       =   2  'Center
               BackColor       =   &H000080FF&
               Enabled         =   0   'False
               ForeColor       =   &H80000005&
               Height          =   300
               Left            =   825
               MultiLine       =   -1  'True
               TabIndex        =   53
               Top             =   6075
               Width           =   3705
            End
            Begin VB.TextBox Sales1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   270
               Left            =   7770
               MultiLine       =   -1  'True
               TabIndex        =   52
               Top             =   1110
               Width           =   2565
            End
            Begin VB.TextBox RSales1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   270
               Left            =   4830
               MultiLine       =   -1  'True
               TabIndex        =   51
               Top             =   1080
               Width           =   2715
            End
            Begin VB.TextBox SalesTotal 
               Alignment       =   2  'Center
               BackColor       =   &H0080FF80&
               Enabled         =   0   'False
               Height          =   300
               Left            =   7770
               MultiLine       =   -1  'True
               TabIndex        =   50
               Top             =   3045
               Width           =   2565
            End
            Begin VB.TextBox Sales5 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   300
               Left            =   7770
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   49
               Top             =   2610
               Width           =   2565
            End
            Begin VB.TextBox Sales4 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   360
               Left            =   7770
               MultiLine       =   -1  'True
               TabIndex        =   48
               Top             =   2235
               Width           =   2565
            End
            Begin VB.TextBox Sales3 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   7770
               MultiLine       =   -1  'True
               TabIndex        =   47
               Top             =   1830
               Width           =   2565
            End
            Begin VB.TextBox Sales2 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   315
               Left            =   7770
               MultiLine       =   -1  'True
               TabIndex        =   46
               Top             =   1455
               Width           =   2565
            End
            Begin VB.TextBox RSalesTotal 
               Alignment       =   2  'Center
               BackColor       =   &H0080FF80&
               Enabled         =   0   'False
               Height          =   300
               Left            =   4830
               MultiLine       =   -1  'True
               TabIndex        =   45
               Top             =   3045
               Width           =   2715
            End
            Begin VB.TextBox RSales5 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFF80&
               ForeColor       =   &H00000000&
               Height          =   300
               Left            =   4830
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   44
               Top             =   2655
               Width           =   2715
            End
            Begin VB.TextBox RSales4 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   360
               Left            =   4830
               MultiLine       =   -1  'True
               TabIndex        =   43
               Top             =   2235
               Width           =   2715
            End
            Begin VB.TextBox RSales3 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   300
               Left            =   4830
               MultiLine       =   -1  'True
               TabIndex        =   42
               Top             =   1785
               Width           =   2715
            End
            Begin VB.TextBox RSales2 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   315
               Left            =   4830
               MultiLine       =   -1  'True
               TabIndex        =   41
               Top             =   1455
               Width           =   2715
            End
            Begin VB.TextBox VATSales1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   270
               Left            =   900
               MultiLine       =   -1  'True
               TabIndex        =   40
               Top             =   1110
               Width           =   3555
            End
            Begin VB.TextBox VATSalesTotal 
               Alignment       =   2  'Center
               BackColor       =   &H0000C000&
               Enabled         =   0   'False
               Height          =   300
               Left            =   900
               MultiLine       =   -1  'True
               TabIndex        =   39
               Top             =   3045
               Width           =   3555
            End
            Begin VB.TextBox TxtRetBuyVAT 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   300
               Left            =   600
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   -285
               Visible         =   0   'False
               Width           =   840
            End
            Begin VB.TextBox TxtBuyVAT 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   300
               Left            =   1890
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   -285
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.TextBox TxtRetSalesVAT 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   300
               Left            =   2865
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   36
               Top             =   -285
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.TextBox TxtSalesVAT 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   165
               Left            =   4230
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   240
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.TextBox RPurchasesTotal 
               Alignment       =   2  'Center
               BackColor       =   &H0080C0FF&
               Enabled         =   0   'False
               Height          =   300
               Left            =   4830
               MultiLine       =   -1  'True
               TabIndex        =   34
               Top             =   6075
               Width           =   2715
            End
            Begin VB.TextBox PurchasesTotal 
               Alignment       =   2  'Center
               BackColor       =   &H0080C0FF&
               Enabled         =   0   'False
               Height          =   300
               Left            =   7770
               MultiLine       =   -1  'True
               TabIndex        =   33
               Top             =   6075
               Width           =   2565
            End
            Begin VB.TextBox VATPurchases1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   825
               MultiLine       =   -1  'True
               TabIndex        =   32
               Top             =   4080
               Width           =   3705
            End
            Begin VB.TextBox VATPurchases2 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   270
               Left            =   825
               MultiLine       =   -1  'True
               TabIndex        =   31
               Top             =   4530
               Width           =   3705
            End
            Begin VB.TextBox VATPurchases3 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   300
               Left            =   825
               MultiLine       =   -1  'True
               TabIndex        =   30
               Top             =   4995
               Width           =   3705
            End
            Begin VB.TextBox RPurchases2 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   270
               Left            =   4830
               MultiLine       =   -1  'True
               TabIndex        =   29
               Top             =   4530
               Width           =   2715
            End
            Begin VB.TextBox RPurchases3 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   300
               Left            =   4830
               MultiLine       =   -1  'True
               TabIndex        =   28
               Top             =   4995
               Width           =   2715
            End
            Begin VB.TextBox RPurchases4 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   360
               Left            =   4830
               MultiLine       =   -1  'True
               TabIndex        =   27
               Top             =   5370
               Width           =   2715
            End
            Begin VB.TextBox RPurchases5 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   270
               Left            =   4830
               MultiLine       =   -1  'True
               TabIndex        =   26
               Top             =   5760
               Width           =   2715
            End
            Begin VB.TextBox RPurchases1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   4830
               MultiLine       =   -1  'True
               TabIndex        =   25
               Top             =   4110
               Width           =   2715
            End
            Begin VB.TextBox Purchases2 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   270
               Left            =   7770
               MultiLine       =   -1  'True
               TabIndex        =   24
               Top             =   4530
               Width           =   2565
            End
            Begin VB.TextBox Purchases3 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   300
               Left            =   7770
               MultiLine       =   -1  'True
               TabIndex        =   23
               Top             =   4995
               Width           =   2565
            End
            Begin VB.TextBox Purchases4 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   360
               Left            =   7770
               MultiLine       =   -1  'True
               TabIndex        =   22
               Top             =   5370
               Width           =   2565
            End
            Begin VB.TextBox Purchases5 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   270
               Left            =   7770
               MultiLine       =   -1  'True
               TabIndex        =   21
               Top             =   5760
               Width           =   2565
            End
            Begin VB.TextBox Purchases1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   7770
               MultiLine       =   -1  'True
               TabIndex        =   20
               Top             =   4080
               Width           =   2565
            End
            Begin VB.CheckBox paidchk 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " „ «·”œ«œ"
               ForeColor       =   &H00000000&
               Height          =   240
               Index           =   0
               Left            =   975
               RightToLeft     =   -1  'True
               TabIndex        =   19
               Top             =   120
               Width           =   1290
            End
            Begin ImpulseButton.ISButton ShowBtn 
               Height          =   465
               Left            =   2790
               TabIndex        =   56
               Top             =   75
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   820
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
               ButtonImage     =   "FrmVATAvowal.frx":11F2
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               LowerToggledContent=   0   'False
            End
            Begin MSComCtl2.DTPicker DateFrom 
               Height          =   330
               Left            =   7920
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   75
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   582
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
               Format          =   151650307
               CurrentDate     =   37140
            End
            Begin MSComCtl2.DTPicker DateTo 
               Height          =   330
               Left            =   4980
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   75
               Width           =   2040
               _ExtentX        =   3598
               _ExtentY        =   582
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
               Format          =   151584771
               CurrentDate     =   37140
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   330
               Left            =   4680
               TabIndex        =   300
               TabStop         =   0   'False
               Top             =   7200
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   582
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
               Format          =   151650307
               CurrentDate     =   37140
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… ÷—Ì»… «·ﬁÌ„… «·„÷«›…"
               Height          =   225
               Index           =   2
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   299
               Top             =   720
               Width           =   2655
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„»·€"
               Height          =   225
               Index           =   1
               Left            =   7770
               RightToLeft     =   -1  'True
               TabIndex        =   298
               Top             =   720
               Width           =   2640
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "9-«·«” Ì—«œ«  «·Œ«÷⁄… ·÷—Ì»… «·ﬁÌ„… «·„÷«›… «· Ì  ÿ»ﬁ ⁄·ÌÂ« «·Ì… «·«Õ ”«» «·⁄ﬂ”Ì "
               Height          =   360
               Left            =   10410
               RightToLeft     =   -1  'True
               TabIndex        =   225
               Top             =   4980
               Width           =   4380
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "8-«·«” Ì—«œ«  «·Œ«÷⁄… ·÷—Ì»… «·ﬁÌ„… «·„÷«›… «· Ì  œ›⁄ ›Ì «·Ã„«—ﬂ"
               Height          =   435
               Left            =   10410
               RightToLeft     =   -1  'True
               TabIndex        =   224
               Top             =   4425
               Width           =   4380
            End
            Begin VB.Label Label13 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "7-«·„‘ —Ì«  «·Œ«÷⁄… ··‰”»… «·«”«”Ì…"
               Height          =   180
               Left            =   10410
               RightToLeft     =   -1  'True
               TabIndex        =   223
               Top             =   4050
               Width           =   4380
            End
            Begin VB.Label Label21 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "12-«Ã„«·Ì «·„‘ —Ì« "
               Height          =   240
               Left            =   10410
               RightToLeft     =   -1  'True
               TabIndex        =   222
               Top             =   6105
               Width           =   4380
            End
            Begin VB.Label Label22 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "11-«·„‘ —Ì«  «·„⁄›«…"
               Height          =   225
               Left            =   10410
               RightToLeft     =   -1  'True
               TabIndex        =   221
               Top             =   5775
               Width           =   4380
            End
            Begin VB.Label Label23 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "10- «·„‘ —Ì«  «·Œ«÷⁄… ··‰”»… «·’›—Ì… "
               Height          =   255
               Left            =   10410
               RightToLeft     =   -1  'True
               TabIndex        =   220
               Top             =   5430
               Width           =   4380
            End
            Begin VB.Label Label20 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—Ì»… ⁄·Ï «·„‘ —Ì« "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000080FF&
               Height          =   495
               Left            =   10710
               RightToLeft     =   -1  'True
               TabIndex        =   219
               Top             =   3450
               Width           =   3855
            End
            Begin VB.Label Label27 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "13-«Ã„«·Ì ÷—Ì»… «·ﬁÌ„… «·„÷«›… «·„” Õﬁ… ⁄‰ «·› —… «·Õ«·Ì…"
               Height          =   270
               Left            =   10410
               RightToLeft     =   -1  'True
               TabIndex        =   218
               Top             =   6435
               Width           =   4380
            End
            Begin VB.Label Label81 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "14-  ’ÕÌÕ«  „‰ › —«  ”«»ﬁ… »Ì‰ (5000) —Ì«·"
               Height          =   210
               Left            =   10410
               RightToLeft     =   -1  'True
               TabIndex        =   217
               Top             =   6795
               Width           =   4380
            End
            Begin VB.Label Label82 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "15- ÷—Ì»… «·ﬁÌ„… «·„÷«›… «· Ì  „  —ÕÌ·Â« „‰ «·› —…/ «·› —«  «·”«»ﬁ…"
               Height          =   450
               Left            =   10410
               RightToLeft     =   -1  'True
               TabIndex        =   216
               Top             =   7095
               Width           =   4380
            End
            Begin VB.Label Label24 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "16- ’«›Ì «·÷—Ì»… «·„” Õﬁ… «Ê «·„” —œ…"
               Height          =   270
               Left            =   10410
               RightToLeft     =   -1  'True
               TabIndex        =   215
               Top             =   7530
               Width           =   4380
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "4-«·’«œ—« "
               Height          =   300
               Index           =   3
               Left            =   10410
               RightToLeft     =   -1  'True
               TabIndex        =   207
               Top             =   2265
               Width           =   4380
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ« "
               Height          =   225
               Index           =   24
               Left            =   13575
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   9285
               Width           =   1215
            End
            Begin VB.Line Line4 
               X1              =   14865
               X2              =   0
               Y1              =   9225
               Y2              =   9225
            End
            Begin VB.Label Label17 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—Ì»… ⁄·Ï «·„»Ì⁄« "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   420
               Left            =   10335
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   690
               Width           =   4530
            End
            Begin VB.Line Line3 
               X1              =   14865
               X2              =   -225
               Y1              =   630
               Y2              =   630
            End
            Begin VB.Line Line1 
               X1              =   14865
               X2              =   75
               Y1              =   3405
               Y2              =   3405
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "2-„»Ì⁄«  ··⁄„·«¡ ›Ì œÊ· „Ã·” «· ⁄«Ê‰ «·Œ·ÌÃÌ «· Ì  ÿ»ﬁ ÷—Ì»Â «·ﬁÌ„… «·„÷«›… "
               Height          =   390
               Left            =   10410
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   1395
               Width           =   4380
            End
            Begin VB.Label Label14 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ï"
               Height          =   210
               Left            =   7020
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   120
               Width           =   750
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "1-«·„»Ì⁄«  «·Œ«÷⁄… ··‰”»… «·«”«”Ì… "
               Height          =   225
               Left            =   10410
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   1140
               Width           =   4380
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·› —…"
               Height          =   240
               Left            =   12225
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   120
               Width           =   1665
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "3-«·„»Ì⁄«  «·„Õ·Ì… «·Œ«÷⁄… ··‰”»… «·’›—Ì… "
               Height          =   240
               Left            =   10410
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   1845
               Width           =   4380
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰ "
               Height          =   240
               Left            =   10035
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   120
               Width           =   750
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ⁄œÌ·« "
               Height          =   225
               Index           =   0
               Left            =   4905
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   750
               Width           =   2640
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "5-«·„»Ì⁄«  «·„⁄›«…"
               Height          =   240
               Left            =   10410
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   2700
               Width           =   4380
            End
            Begin VB.Label Label12 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "6-«Ã„«·Ì «·„»Ì⁄« "
               Height          =   255
               Left            =   10410
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   3060
               Width           =   4380
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic5 
            Height          =   7950
            Left            =   15750
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   45
            Width           =   15015
            _cx             =   26485
            _cy             =   14023
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
            Begin VB.TextBox tztAdvBill 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   3960
               TabIndex        =   319
               Top             =   600
               Width           =   1125
            End
            Begin VB.CheckBox paidchk 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Œ›«¡«·⁄ﬁÊœ «·„’›«Â"
               ForeColor       =   &H000000FF&
               Height          =   390
               Index           =   1
               Left            =   11325
               RightToLeft     =   -1  'True
               TabIndex        =   304
               Top             =   870
               Width           =   1275
            End
            Begin VB.TextBox txtFaBuy3 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   255
               Left            =   150
               TabIndex        =   295
               Top             =   3810
               Width           =   3465
            End
            Begin VB.TextBox txtFaBuy2 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   255
               Left            =   3930
               TabIndex        =   294
               Top             =   3810
               Width           =   3315
            End
            Begin VB.TextBox txtFaBuy 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   240
               Left            =   7770
               TabIndex        =   293
               Top             =   3870
               Width           =   3465
            End
            Begin VB.TextBox TxtDept 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Index           =   7
               Left            =   7770
               TabIndex        =   273
               Top             =   4710
               Width           =   3465
            End
            Begin VB.TextBox TxtDept 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Index           =   6
               Left            =   225
               TabIndex        =   272
               Top             =   4710
               Width           =   3465
            End
            Begin VB.TextBox TxtDept 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Index           =   5
               Left            =   7770
               TabIndex        =   270
               Top             =   -30
               Width           =   3465
            End
            Begin VB.TextBox TxtDept 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   315
               Index           =   4
               Left            =   225
               TabIndex        =   269
               Top             =   -60
               Width           =   3465
            End
            Begin VB.TextBox TxtDept 
               Alignment       =   2  'Center
               BackColor       =   &H0000FFFF&
               Enabled         =   0   'False
               Height          =   255
               Index           =   3
               Left            =   225
               TabIndex        =   268
               Top             =   7230
               Width           =   3465
            End
            Begin VB.TextBox TxtDept 
               Alignment       =   2  'Center
               BackColor       =   &H0000FFFF&
               Enabled         =   0   'False
               Height          =   315
               Index           =   2
               Left            =   4005
               TabIndex        =   267
               Top             =   7200
               Width           =   3315
            End
            Begin VB.TextBox TxtDept 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0FF&
               Enabled         =   0   'False
               Height          =   285
               Index           =   1
               Left            =   225
               TabIndex        =   266
               Top             =   4080
               Width           =   3465
            End
            Begin VB.TextBox TxtDept 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0FF&
               Enabled         =   0   'False
               Height          =   285
               Index           =   0
               Left            =   4005
               TabIndex        =   263
               Top             =   4080
               Width           =   3315
            End
            Begin MSComctlLib.TabStrip TabStrip1 
               Height          =   360
               Left            =   6345
               TabIndex        =   262
               Top             =   -570
               Visible         =   0   'False
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   635
               _Version        =   393216
               BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
                  NumTabs         =   1
                  BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                     ImageVarType    =   2
                  EndProperty
               EndProperty
            End
            Begin VB.TextBox TxtPReVatTotal5V 
               Alignment       =   2  'Center
               BackColor       =   &H0000FFFF&
               Enabled         =   0   'False
               Height          =   270
               Left            =   7770
               TabIndex        =   256
               Top             =   6885
               Width           =   3465
            End
            Begin VB.TextBox TxtPReVatVAT5V 
               Alignment       =   2  'Center
               BackColor       =   &H0000FFFF&
               Enabled         =   0   'False
               Height          =   270
               Left            =   225
               TabIndex        =   255
               Top             =   6885
               Width           =   3465
            End
            Begin VB.TextBox TxtPReVatTotal5 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   240
               Left            =   7770
               TabIndex        =   254
               Top             =   3570
               Width           =   3465
            End
            Begin VB.TextBox TxtPReVatVAT5 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   240
               Left            =   225
               TabIndex        =   252
               Top             =   3570
               Width           =   3465
            End
            Begin VB.TextBox TxtContractVaue 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   7770
               TabIndex        =   239
               Top             =   915
               Width           =   3465
            End
            Begin VB.TextBox TxtContractReVaue 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   4005
               MultiLine       =   -1  'True
               TabIndex        =   238
               Top             =   945
               Width           =   3315
            End
            Begin VB.TextBox TxtContractVAT 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   225
               TabIndex        =   237
               Top             =   945
               Width           =   3465
            End
            Begin VB.TextBox transport5 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   300
               Left            =   7770
               TabIndex        =   213
               Top             =   2940
               Width           =   3465
            End
            Begin VB.TextBox transport5re 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   300
               Left            =   4005
               TabIndex        =   212
               Top             =   2955
               Width           =   3315
            End
            Begin VB.TextBox transport5vat 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   300
               Left            =   225
               TabIndex        =   211
               Top             =   2940
               Width           =   3465
            End
            Begin VB.TextBox TxtServiceInvoice5 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   240
               Left            =   7770
               TabIndex        =   203
               Top             =   3270
               Width           =   3465
            End
            Begin VB.TextBox TxtServiceInvoice5REt 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   255
               Left            =   4005
               TabIndex        =   202
               Top             =   3300
               Width           =   3315
            End
            Begin VB.TextBox TxtServiceInvoice5Vat 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   255
               Left            =   225
               TabIndex        =   201
               Top             =   3300
               Width           =   3465
            End
            Begin VB.TextBox SalesTVAT 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   315
               Left            =   225
               TabIndex        =   153
               Top             =   585
               Width           =   3465
            End
            Begin VB.TextBox Expensesvat 
               Alignment       =   2  'Center
               BackColor       =   &H0000FFFF&
               Enabled         =   0   'False
               Height          =   270
               Left            =   225
               TabIndex        =   152
               Top             =   6570
               Width           =   3465
            End
            Begin VB.TextBox Expenses 
               Alignment       =   2  'Center
               BackColor       =   &H0000FFFF&
               Enabled         =   0   'False
               Height          =   255
               Left            =   7770
               TabIndex        =   151
               Top             =   6570
               Width           =   3465
            End
            Begin VB.TextBox txtManulaEntryP5Vat 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   270
               Left            =   225
               TabIndex        =   148
               Top             =   4965
               Width           =   3465
            End
            Begin VB.TextBox txtManulaEntryP5 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   270
               Left            =   7770
               TabIndex        =   147
               Top             =   4965
               Width           =   3465
            End
            Begin VB.TextBox txtManulaEntryP5Ret 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   270
               Left            =   4005
               TabIndex        =   146
               Top             =   4965
               Width           =   3315
            End
            Begin VB.TextBox manulaEntey5Vat 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   315
               Left            =   225
               TabIndex        =   144
               Top             =   255
               Width           =   3465
            End
            Begin VB.TextBox manulaEntey5 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   270
               Left            =   7770
               TabIndex        =   143
               Top             =   270
               Width           =   3465
            End
            Begin VB.TextBox manulaEnteyRet5 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   270
               Left            =   4005
               TabIndex        =   142
               Top             =   270
               Width           =   3315
            End
            Begin VB.TextBox Purchasest5vat 
               Alignment       =   2  'Center
               Height          =   270
               Left            =   225
               Locked          =   -1  'True
               TabIndex        =   138
               Top             =   5280
               Width           =   3465
            End
            Begin VB.TextBox PurchasesRett5 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   270
               Left            =   4005
               Locked          =   -1  'True
               TabIndex        =   137
               Top             =   5295
               Width           =   3315
            End
            Begin VB.TextBox PurchasesT5 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   255
               Left            =   7770
               Locked          =   -1  'True
               TabIndex        =   136
               Top             =   5280
               Width           =   3465
            End
            Begin VB.TextBox SalesT5 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   7770
               Locked          =   -1  'True
               TabIndex        =   135
               Top             =   585
               Width           =   3465
            End
            Begin VB.TextBox SalesRet5 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   5685
               Locked          =   -1  'True
               TabIndex        =   134
               Top             =   585
               Width           =   1635
            End
            Begin VB.TextBox TxtAssestReValue 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   270
               Left            =   4005
               TabIndex        =   100
               Top             =   6240
               Width           =   3315
            End
            Begin VB.TextBox TxtAssestVAT 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   270
               Left            =   225
               TabIndex        =   99
               Top             =   6240
               Width           =   3465
            End
            Begin VB.TextBox TxtMaintCarVAT 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   315
               Left            =   225
               TabIndex        =   98
               Top             =   2580
               Width           =   3465
            End
            Begin VB.TextBox TxtMaintCarReValue 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   315
               Left            =   4005
               TabIndex        =   97
               Top             =   2580
               Width           =   3315
            End
            Begin VB.TextBox TxtMaintCarValue 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   315
               Left            =   7770
               TabIndex        =   96
               Top             =   2580
               Width           =   3465
            End
            Begin VB.TextBox TxtProjConValue 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   7770
               TabIndex        =   95
               Top             =   5595
               Width           =   3465
            End
            Begin VB.TextBox TxtAssestValue 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   270
               Left            =   7770
               TabIndex        =   94
               Top             =   6240
               Width           =   3465
            End
            Begin VB.TextBox TxtReqConValue 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   240
               Left            =   7770
               TabIndex        =   93
               Top             =   5925
               Width           =   3465
            End
            Begin VB.TextBox TxtProjConReValue 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   300
               Left            =   4005
               TabIndex        =   92
               Top             =   5595
               Width           =   2340
            End
            Begin VB.TextBox TxtReqConReValue 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   240
               Left            =   4005
               TabIndex        =   91
               Top             =   5925
               Width           =   3315
            End
            Begin VB.TextBox TxtReqConVAT 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   240
               Left            =   225
               TabIndex        =   90
               Top             =   5925
               Width           =   3465
            End
            Begin VB.TextBox TxtProjConVAT 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   300
               Left            =   225
               TabIndex        =   89
               Top             =   5595
               Width           =   3465
            End
            Begin VB.TextBox TxtTotalReceValue 
               Alignment       =   2  'Center
               BackColor       =   &H0080C0FF&
               Enabled         =   0   'False
               Height          =   315
               Left            =   7770
               TabIndex        =   88
               Top             =   7530
               Width           =   3465
            End
            Begin VB.TextBox TxtTotalReceReValue 
               Alignment       =   2  'Center
               BackColor       =   &H0080C0FF&
               Enabled         =   0   'False
               Height          =   315
               Left            =   4005
               TabIndex        =   87
               Top             =   7530
               Width           =   3315
            End
            Begin VB.TextBox TxtTotalPayVAT 
               Alignment       =   2  'Center
               BackColor       =   &H0000C000&
               Enabled         =   0   'False
               Height          =   210
               Left            =   225
               TabIndex        =   86
               Top             =   4410
               Width           =   3465
            End
            Begin VB.TextBox TxtMinisterVAT 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   255
               Left            =   225
               TabIndex        =   85
               Top             =   2265
               Width           =   3465
            End
            Begin VB.TextBox TxtHajjVAT 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   225
               TabIndex        =   84
               Top             =   1935
               Width           =   3465
            End
            Begin VB.TextBox TxtProjCusReValue 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   270
               Left            =   4005
               TabIndex        =   83
               Top             =   1260
               Width           =   2565
            End
            Begin VB.TextBox TxtOmraReValue 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   270
               Left            =   4005
               TabIndex        =   82
               Top             =   1605
               Width           =   3315
            End
            Begin VB.TextBox TxtHajjReValue 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   4005
               TabIndex        =   81
               Top             =   1935
               Width           =   3315
            End
            Begin VB.TextBox TxtMinisterReValue 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   255
               Left            =   4005
               TabIndex        =   80
               Top             =   2265
               Width           =   3315
            End
            Begin VB.TextBox TxtTotalRePayValue 
               Alignment       =   2  'Center
               BackColor       =   &H0080FF80&
               Enabled         =   0   'False
               Height          =   210
               Left            =   4005
               TabIndex        =   79
               Top             =   4410
               Width           =   3315
            End
            Begin VB.TextBox TxtProjCusVAT 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   255
               Left            =   225
               TabIndex        =   78
               Top             =   1260
               Width           =   3465
            End
            Begin VB.TextBox TxtOmraVAT 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   225
               TabIndex        =   77
               Top             =   1620
               Width           =   3465
            End
            Begin VB.TextBox TxtProjCusValue 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   255
               Left            =   7770
               TabIndex        =   76
               Top             =   1350
               Width           =   3465
            End
            Begin VB.TextBox TxtOmraValue 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   7770
               TabIndex        =   75
               Top             =   1620
               Width           =   3465
            End
            Begin VB.TextBox TxtHajjValue 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Left            =   7770
               TabIndex        =   74
               Top             =   1935
               Width           =   3465
            End
            Begin VB.TextBox TxtMinisterValue 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   255
               Left            =   7770
               TabIndex        =   73
               Top             =   2265
               Width           =   3465
            End
            Begin VB.TextBox TxtTotalPayValue 
               Alignment       =   2  'Center
               BackColor       =   &H0080FF80&
               Enabled         =   0   'False
               Height          =   285
               Left            =   7770
               TabIndex        =   72
               Top             =   4320
               Width           =   3465
            End
            Begin VB.TextBox TxtTotalReceVAT 
               Alignment       =   2  'Center
               BackColor       =   &H000080FF&
               Enabled         =   0   'False
               ForeColor       =   &H80000005&
               Height          =   315
               Left            =   225
               TabIndex        =   71
               Top             =   7530
               Width           =   3465
            End
            Begin VB.Label Label84 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "œ „ „ﬁ›·Â ﬁ.„"
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   1
               Left            =   5040
               RightToLeft     =   -1  'True
               TabIndex        =   320
               Top             =   600
               Width           =   750
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· Œ·’ „‰ «·«’·"
               Height          =   180
               Index           =   14
               Left            =   12075
               RightToLeft     =   -1  'True
               TabIndex        =   296
               Top             =   3870
               Width           =   2715
            End
            Begin VB.Label Label86 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "œ „ „ﬁ›·Â ﬁ.„"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   6495
               RightToLeft     =   -1  'True
               TabIndex        =   283
               Top             =   5640
               Width           =   750
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   " ”ÊÌ«  ÌœÊÌ…"
               Height          =   240
               Index           =   9
               Left            =   11475
               RightToLeft     =   -1  'True
               TabIndex        =   274
               Top             =   4710
               Width           =   3240
            End
            Begin VB.Label Label29 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   " Õ·Ì· «·÷—Ì»… ⁄·Ï  «·„‘ —Ì«   %"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000080FF&
               Height          =   2505
               Left            =   14115
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   5100
               Width           =   825
            End
            Begin VB.Label Label30 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   " Õ·Ì· «·÷—Ì»… ⁄·Ï  «·„»Ì⁄«   %"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   1845
               Left            =   14040
               RightToLeft     =   -1  'True
               TabIndex        =   236
               Top             =   1125
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   " ”ÊÌ«  ÌœÊÌ… "
               Height          =   210
               Index           =   7
               Left            =   11550
               RightToLeft     =   -1  'True
               TabIndex        =   271
               Top             =   30
               Width           =   3165
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«‘⁄«— „œÌ‰"
               Height          =   195
               Index           =   6
               Left            =   11700
               RightToLeft     =   -1  'True
               TabIndex        =   265
               Top             =   7230
               Width           =   3240
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«‘⁄«— œ«∆‰"
               Height          =   225
               Index           =   5
               Left            =   11475
               RightToLeft     =   -1  'True
               TabIndex        =   264
               Top             =   4080
               Width           =   3240
            End
            Begin VB.Label Label84 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "œ „ „ﬁ›·Â ﬁ.„"
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   0
               Left            =   6570
               RightToLeft     =   -1  'True
               TabIndex        =   258
               Top             =   1260
               Width           =   750
            End
            Begin VB.Label Label83 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·œ›⁄«  «·„ﬁœ„… ··„Ê—Ì‰"
               Height          =   270
               Left            =   11475
               RightToLeft     =   -1  'True
               TabIndex        =   257
               Top             =   6990
               Width           =   3465
            End
            Begin VB.Label Label60 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·œ›⁄«  «·„ﬁœ„… „‰ «·⁄„·«¡"
               Height          =   210
               Left            =   11475
               RightToLeft     =   -1  'True
               TabIndex        =   253
               Top             =   3585
               Width           =   3240
            End
            Begin VB.Label Label41 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„” Œ·’«  «·„‘«—Ì⁄ ··⁄„Ì· "
               Height          =   210
               Left            =   11550
               RightToLeft     =   -1  'True
               TabIndex        =   235
               Top             =   1290
               Width           =   3240
            End
            Begin VB.Label Label43 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄ﬁœ «·«ÌÃ«—  Ã«—Ì"
               Height          =   240
               Left            =   11550
               RightToLeft     =   -1  'True
               TabIndex        =   234
               Top             =   960
               Width           =   3240
            End
            Begin VB.Label Label46 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«„— ‘€· ⁄„—Â"
               Height          =   195
               Left            =   11550
               RightToLeft     =   -1  'True
               TabIndex        =   233
               Top             =   1665
               Width           =   3240
            End
            Begin VB.Label Label49 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«” Õﬁ«ﬁ «·Ê“«—Ì 5%"
               Height          =   210
               Left            =   11550
               RightToLeft     =   -1  'True
               TabIndex        =   232
               Top             =   2295
               Width           =   3240
            End
            Begin VB.Label Label50 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "›« Ê—… ’Ì«‰… ”Ì«—« "
               Height          =   210
               Left            =   11550
               RightToLeft     =   -1  'True
               TabIndex        =   231
               Top             =   2655
               Width           =   3240
            End
            Begin VB.Label Label51 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«„— ‘€· ÕÃ"
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   11550
               TabIndex        =   230
               Top             =   1965
               Width           =   3240
            End
            Begin VB.Label Label28 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«Ã„«·Ì"
               Height          =   180
               Left            =   11550
               RightToLeft     =   -1  'True
               TabIndex        =   229
               Top             =   4425
               Width           =   3240
            End
            Begin VB.Label Label42 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "›Ê« Ì— «·„»Ì⁄« "
               Height          =   240
               Left            =   11550
               RightToLeft     =   -1  'True
               TabIndex        =   228
               Top             =   630
               Width           =   3240
            End
            Begin VB.Label Label80 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›Ê« Ì— «·Œœ„Ì…"
               Height          =   195
               Left            =   11700
               RightToLeft     =   -1  'True
               TabIndex        =   227
               Top             =   3300
               Width           =   3240
            End
            Begin VB.Label Label25 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "›« Ê—… Œœ„«  «·‰ﬁ·"
               Height          =   240
               Left            =   11550
               RightToLeft     =   -1  'True
               TabIndex        =   226
               Top             =   2970
               Width           =   3240
            End
            Begin VB.Label Label54 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„’—Ê›« "
               Height          =   285
               Left            =   11475
               RightToLeft     =   -1  'True
               TabIndex        =   150
               Top             =   6660
               Width           =   3465
            End
            Begin VB.Label Label53 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«œŒ«·«  ÌœÊÌ…"
               Height          =   285
               Left            =   11475
               RightToLeft     =   -1  'True
               TabIndex        =   149
               Top             =   5070
               Width           =   3465
            End
            Begin VB.Label Label52 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«œŒ«·«  ÌœÊÌ…"
               Height          =   180
               Left            =   11550
               RightToLeft     =   -1  'True
               TabIndex        =   145
               Top             =   330
               Width           =   3240
            End
            Begin VB.Label Label44 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "›Ê« Ì— «·„‘ —Ì« "
               Height          =   240
               Left            =   11475
               RightToLeft     =   -1  'True
               TabIndex        =   139
               Top             =   5400
               Width           =   3465
            End
            Begin VB.Line Line7 
               X1              =   18330
               X2              =   -750
               Y1              =   4665
               Y2              =   4665
            End
            Begin VB.Label Label40 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘—«¡ «·«’Ê·"
               Height          =   240
               Left            =   11475
               RightToLeft     =   -1  'True
               TabIndex        =   108
               Top             =   6375
               Width           =   3465
            End
            Begin VB.Label Label39 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   " ÿ·»«  ’—› „ ⁄ÂœÌ‰"
               Height          =   225
               Left            =   11475
               RightToLeft     =   -1  'True
               TabIndex        =   107
               Top             =   6045
               Width           =   3465
            End
            Begin VB.Label Label38 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„” Œ·’«  «·„‘«—Ì⁄ ··„ﬁ«Ê·"
               Height          =   240
               Left            =   11475
               RightToLeft     =   -1  'True
               TabIndex        =   106
               Top             =   5730
               Width           =   3465
            End
            Begin VB.Line Line6 
               X1              =   19095
               X2              =   -225
               Y1              =   -45
               Y2              =   -45
            End
            Begin VB.Label Label36 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«Ã„«·Ì"
               Height          =   270
               Left            =   11475
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   7545
               Width           =   3465
            End
            Begin VB.Label Label33 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„»·€"
               Height          =   165
               Left            =   11850
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   4635
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.Label Label32 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ⁄œÌ·« "
               Height          =   195
               Left            =   4680
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   4515
               Visible         =   0   'False
               Width           =   1965
            End
            Begin VB.Label Label31 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… ÷—Ì»… «·ﬁÌ„… «·„÷«›…"
               Height          =   165
               Left            =   450
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   4425
               Width           =   3015
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   7950
            Left            =   16350
            TabIndex        =   132
            TabStop         =   0   'False
            Top             =   45
            Width           =   15015
            _cx             =   26485
            _cy             =   14023
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
            Begin VB.CheckBox paidchk 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Œ›«¡«·⁄ﬁÊœ «·„’›«Â"
               ForeColor       =   &H000000FF&
               Height          =   390
               Index           =   2
               Left            =   7845
               RightToLeft     =   -1  'True
               TabIndex        =   305
               Top             =   1680
               Width           =   1290
            End
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Index           =   1
               Left            =   5730
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   282
               Top             =   1800
               Width           =   1965
            End
            Begin VB.TextBox Text1 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Index           =   0
               Left            =   5730
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   251
               Top             =   1350
               Width           =   1965
            End
            Begin VB.TextBox Text5 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   270
               Left            =   5730
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   168
               Top             =   840
               Width           =   1965
            End
            Begin VB.TextBox Text4 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   270
               Left            =   3540
               MultiLine       =   -1  'True
               TabIndex        =   167
               Top             =   870
               Width           =   1965
            End
            Begin VB.TextBox Text2 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   315
               Left            =   3540
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   166
               Top             =   1380
               Width           =   1965
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ⁄œÌ·« "
               Height          =   240
               Index           =   4
               Left            =   3465
               RightToLeft     =   -1  'True
               TabIndex        =   309
               Top             =   480
               Width           =   1815
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„»·€"
               Height          =   240
               Index           =   3
               Left            =   5955
               RightToLeft     =   -1  'True
               TabIndex        =   308
               Top             =   405
               Width           =   1740
            End
            Begin VB.Label Label66 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄ﬁÊœ «·”ﬂ‰Ì…"
               Height          =   210
               Index           =   1
               Left            =   8370
               RightToLeft     =   -1  'True
               TabIndex        =   281
               Top             =   1800
               Width           =   2640
            End
            Begin VB.Label Label59 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—Ì»… ⁄·Ï «·„»Ì⁄« "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   405
               Left            =   8595
               RightToLeft     =   -1  'True
               TabIndex        =   250
               Top             =   210
               Width           =   2265
            End
            Begin VB.Label Label66 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "›Ê« Ì— «·„»Ì⁄«   «·„⁄›«Â"
               Height          =   210
               Index           =   0
               Left            =   8445
               RightToLeft     =   -1  'True
               TabIndex        =   249
               Top             =   1440
               Width           =   2640
            End
            Begin VB.Label Label69 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«œŒ«·«  ÌœÊÌ…"
               Height          =   195
               Left            =   8445
               RightToLeft     =   -1  'True
               TabIndex        =   248
               Top             =   915
               Width           =   2640
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic9 
            Height          =   7950
            Left            =   16650
            TabIndex        =   169
            TabStop         =   0   'False
            Top             =   45
            Width           =   15015
            _cx             =   26485
            _cy             =   14023
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
            Begin VB.TextBox TxtBillVstReverse 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   270
               Left            =   5760
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   317
               Top             =   1200
               Width           =   1890
            End
            Begin VB.TextBox TxtBillVstReverseREt 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   270
               Left            =   3360
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   316
               Top             =   1200
               Width           =   1815
            End
            Begin VB.TextBox txtExpensesVat 
               Alignment       =   2  'Center
               Height          =   300
               Index           =   6
               Left            =   3390
               Locked          =   -1  'True
               TabIndex        =   303
               Top             =   4680
               Width           =   1815
            End
            Begin VB.TextBox txtExpenses 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   300
               Index           =   6
               Left            =   5805
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   301
               Top             =   4695
               Width           =   1890
            End
            Begin VB.TextBox TxtMaintCarReValue2 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   255
               Left            =   3390
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   292
               Top             =   2730
               Width           =   1815
            End
            Begin VB.TextBox TxtMaintCarValue2 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   255
               Left            =   5955
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   291
               Top             =   2730
               Width           =   1815
            End
            Begin VB.TextBox TxtMaintCarValue1 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   285
               Left            =   5955
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   286
               Top             =   2415
               Width           =   1815
            End
            Begin VB.TextBox TxtMaintCarReValue1 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   255
               Left            =   3390
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   285
               Top             =   2445
               Width           =   1815
            End
            Begin VB.TextBox TxtVatADD 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFC0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   5805
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   277
               Top             =   7530
               Width           =   1815
            End
            Begin VB.TextBox TxtVatDis 
               Alignment       =   2  'Center
               BackColor       =   &H00C0C0FF&
               Enabled         =   0   'False
               Height          =   270
               Left            =   5805
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   276
               Top             =   7050
               Width           =   1890
            End
            Begin VB.TextBox txtExpenses 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   255
               Index           =   5
               Left            =   5805
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   209
               Top             =   3900
               Width           =   1890
            End
            Begin VB.TextBox txtExpensesVat 
               Alignment       =   2  'Center
               Height          =   255
               Index           =   5
               Left            =   3390
               Locked          =   -1  'True
               TabIndex        =   208
               Top             =   3900
               Width           =   1815
            End
            Begin VB.TextBox txtTotalExpenses 
               Alignment       =   2  'Center
               BackColor       =   &H0080C0FF&
               Enabled         =   0   'False
               Height          =   270
               Left            =   5805
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   200
               Top             =   6510
               Width           =   1890
            End
            Begin VB.TextBox TxtTotalVatExpenses 
               Alignment       =   2  'Center
               BackColor       =   &H0080C0FF&
               Enabled         =   0   'False
               Height          =   270
               Left            =   3390
               MultiLine       =   -1  'True
               TabIndex        =   199
               Top             =   6510
               Width           =   1815
            End
            Begin VB.TextBox txtExpensesVat 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   4
               Left            =   3390
               Locked          =   -1  'True
               TabIndex        =   198
               Top             =   6060
               Width           =   1815
            End
            Begin VB.TextBox txtExpensesVat 
               Alignment       =   2  'Center
               Height          =   285
               Index           =   3
               Left            =   3390
               Locked          =   -1  'True
               TabIndex        =   197
               Top             =   5565
               Width           =   1815
            End
            Begin VB.TextBox txtExpensesVat 
               Alignment       =   2  'Center
               Height          =   300
               Index           =   2
               Left            =   3390
               Locked          =   -1  'True
               TabIndex        =   196
               Top             =   5070
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.TextBox txtExpensesVat 
               Alignment       =   2  'Center
               Height          =   300
               Index           =   1
               Left            =   3390
               Locked          =   -1  'True
               TabIndex        =   195
               Top             =   4275
               Width           =   1815
            End
            Begin VB.TextBox txtExpensesVat 
               Alignment       =   2  'Center
               Height          =   270
               Index           =   0
               Left            =   3390
               Locked          =   -1  'True
               TabIndex        =   194
               Top             =   3480
               Width           =   1815
            End
            Begin VB.TextBox txtExpenses 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Index           =   4
               Left            =   5805
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   192
               Top             =   6060
               Width           =   1890
            End
            Begin VB.TextBox txtExpenses 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   285
               Index           =   3
               Left            =   5805
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   189
               Top             =   5565
               Width           =   1890
            End
            Begin VB.TextBox txtExpenses 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   300
               Index           =   2
               Left            =   5805
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   188
               Top             =   5070
               Visible         =   0   'False
               Width           =   1890
            End
            Begin VB.TextBox txtExpenses 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   300
               Index           =   1
               Left            =   5805
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   185
               Top             =   4275
               Width           =   1890
            End
            Begin VB.TextBox txtExpenses 
               Alignment       =   2  'Center
               Enabled         =   0   'False
               Height          =   270
               Index           =   0
               Left            =   5805
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   183
               Top             =   3450
               Width           =   1890
            End
            Begin VB.TextBox TxtRemarks1All2 
               Alignment       =   2  'Center
               BackColor       =   &H0080C0FF&
               Enabled         =   0   'False
               Height          =   270
               Left            =   3390
               MultiLine       =   -1  'True
               TabIndex        =   180
               Top             =   1575
               Width           =   1890
            End
            Begin VB.TextBox TxtRemarks1All 
               Alignment       =   2  'Center
               BackColor       =   &H0080C0FF&
               Enabled         =   0   'False
               Height          =   270
               Left            =   5805
               MultiLine       =   -1  'True
               TabIndex        =   179
               Top             =   1575
               Width           =   1890
            End
            Begin VB.TextBox PurcahseRemarks1Ret 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   255
               Left            =   3390
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   178
               Top             =   405
               Width           =   1815
            End
            Begin VB.TextBox PurcahseRemarks2Ret 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   270
               Left            =   3390
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   177
               Top             =   750
               Width           =   1815
            End
            Begin VB.TextBox PurcahseRemarks2 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   270
               Left            =   5805
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   174
               Top             =   750
               Width           =   1890
            End
            Begin VB.TextBox PurcahseRemarks1 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   255
               Left            =   5805
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   172
               Top             =   405
               Width           =   1890
            End
            Begin VB.Label Label78 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«Õ ”«» «·⁄ﬂ”Ì"
               Height          =   195
               Index           =   1
               Left            =   9180
               RightToLeft     =   -1  'True
               TabIndex        =   318
               Top             =   1245
               Width           =   2640
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ⁄œÌ·« "
               Height          =   240
               Index           =   10
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   315
               Top             =   3195
               Width           =   1785
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„»·€"
               Height          =   240
               Index           =   9
               Left            =   5820
               RightToLeft     =   -1  'True
               TabIndex        =   314
               Top             =   3120
               Width           =   1785
            End
            Begin VB.Label Label75 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·÷—Ì»… ⁄„Ê·«  «·‘»ﬂ…"
               ForeColor       =   &H000000FF&
               Height          =   270
               Index           =   18
               Left            =   9210
               RightToLeft     =   -1  'True
               TabIndex        =   302
               Top             =   4680
               Width           =   2025
            End
            Begin VB.Label Label75 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«ÃÊ— «·Ìœ"
               Height          =   270
               Index           =   16
               Left            =   9885
               RightToLeft     =   -1  'True
               TabIndex        =   290
               Top             =   2820
               Width           =   1515
            End
            Begin VB.Label Label75 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "›Ê« Ì— «’·«Õ ”Ì«—« "
               Height          =   270
               Index           =   15
               Left            =   9960
               RightToLeft     =   -1  'True
               TabIndex        =   289
               Top             =   2400
               Width           =   1515
            End
            Begin VB.Label Label75 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„»·€"
               Height          =   285
               Index           =   14
               Left            =   6420
               RightToLeft     =   -1  'True
               TabIndex        =   288
               Top             =   1965
               Width           =   975
            End
            Begin VB.Label Label75 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·÷—Ì»…"
               Height          =   195
               Index           =   13
               Left            =   3930
               RightToLeft     =   -1  'True
               TabIndex        =   287
               Top             =   1995
               Width           =   1200
            End
            Begin VB.Label Label75 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   " Õ·Ì· ’Ì«‰… «·„⁄œ« /«·”Ì«—« "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000080FF&
               Height          =   405
               Index           =   10
               Left            =   9210
               RightToLeft     =   -1  'True
               TabIndex        =   284
               Top             =   1950
               Width           =   2940
            End
            Begin VB.Label Label75 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈÷«›… ﬁÌ„… „÷«›…"
               Height          =   210
               Index           =   8
               Left            =   9510
               RightToLeft     =   -1  'True
               TabIndex        =   279
               Top             =   7530
               Width           =   2040
            End
            Begin VB.Label Label75 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "Œ’„ ﬁÌ„… „÷«›…"
               Height          =   210
               Index           =   7
               Left            =   9510
               RightToLeft     =   -1  'True
               TabIndex        =   278
               Top             =   7050
               Width           =   2040
            End
            Begin VB.Label Label75 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   " ’›Ì… «·⁄Âœ…"
               Height          =   195
               Index           =   2
               Left            =   9960
               RightToLeft     =   -1  'True
               TabIndex        =   210
               Top             =   3930
               Width           =   1200
            End
            Begin VB.Label Label75 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«Ã„«·Ì« "
               Height          =   210
               Index           =   6
               Left            =   9510
               RightToLeft     =   -1  'True
               TabIndex        =   193
               Top             =   6540
               Width           =   2040
            End
            Begin VB.Label Label75 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„Â «·„÷«›… ··ÕÊ«·« "
               Height          =   285
               Index           =   5
               Left            =   9510
               RightToLeft     =   -1  'True
               TabIndex        =   191
               Top             =   6060
               Width           =   2040
            End
            Begin VB.Label Label75 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„Â «·„÷«›… ··„œ›Ê⁄«  «·„ﬁœ„…"
               Height          =   375
               Index           =   4
               Left            =   9285
               RightToLeft     =   -1  'True
               TabIndex        =   190
               Top             =   5520
               Width           =   2490
            End
            Begin VB.Label Label75 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„Â «·„÷«›… ··Ã„«—ﬂ"
               Height          =   300
               Index           =   3
               Left            =   9735
               RightToLeft     =   -1  'True
               TabIndex        =   187
               Top             =   5070
               Visible         =   0   'False
               Width           =   1590
            End
            Begin VB.Label Label75 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›Ê« Ì— «·„«·Ì…"
               Height          =   270
               Index           =   1
               Left            =   9960
               RightToLeft     =   -1  'True
               TabIndex        =   186
               Top             =   4290
               Width           =   1200
            End
            Begin VB.Label Label75 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   " Õ·Ì·Ì «·„’—Ê›« "
               Height          =   270
               Index           =   0
               Left            =   9810
               RightToLeft     =   -1  'True
               TabIndex        =   184
               Top             =   3480
               Width           =   1515
            End
            Begin VB.Label Label75 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   " Õ·Ì· «·„’—Ê›« "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000080FF&
               Height          =   405
               Index           =   9
               Left            =   9060
               RightToLeft     =   -1  'True
               TabIndex        =   182
               Top             =   3060
               Width           =   2940
            End
            Begin VB.Label Label72 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ã„«·Ì «·„·«ÕŸ« "
               Height          =   255
               Left            =   9585
               RightToLeft     =   -1  'True
               TabIndex        =   181
               Top             =   1590
               Width           =   1890
            End
            Begin VB.Label Label75 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ⁄œÌ·« "
               Height          =   165
               Index           =   12
               Left            =   3780
               RightToLeft     =   -1  'True
               TabIndex        =   176
               Top             =   105
               Width           =   1200
            End
            Begin VB.Label Label75 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„»·€"
               Height          =   285
               Index           =   11
               Left            =   6270
               RightToLeft     =   -1  'True
               TabIndex        =   175
               Top             =   45
               Width           =   975
            End
            Begin VB.Label Label78 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‘ —Ì«  Œ«÷⁄Â Ê·ﬂ‰ «·„Ê—œ „⁄›Ì"
               Height          =   195
               Index           =   0
               Left            =   9210
               RightToLeft     =   -1  'True
               TabIndex        =   173
               Top             =   795
               Width           =   2640
            End
            Begin VB.Label Label77 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‘ —Ì«  Œ«÷⁄Â Ê·„ Ìﬁ„ «·„Ê—œ »«÷«› Â«"
               Height          =   195
               Left            =   8745
               RightToLeft     =   -1  'True
               TabIndex        =   171
               Top             =   435
               Width           =   3105
            End
            Begin VB.Label Label75 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·«ÕŸ«  ⁄·Ï «·„‘ —Ì« "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000080FF&
               Height          =   405
               Index           =   17
               Left            =   9060
               RightToLeft     =   -1  'True
               TabIndex        =   170
               Top             =   -30
               Width           =   2940
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   510
         Left            =   600
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   9390
         Width           =   14955
         _cx             =   26379
         _cy             =   900
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
         Begin VB.CommandButton Command9 
            Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
            Height          =   405
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   112
            Top             =   15
            Width           =   2295
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   111
            Top             =   165
            Width           =   3705
         End
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   22740
            RightToLeft     =   -1  'True
            TabIndex        =   110
            Top             =   345
            Visible         =   0   'False
            Width           =   2970
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   31635
            TabIndex        =   113
            Top             =   345
            Width           =   8595
            _ExtentX        =   15161
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   16777152
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   210
            Index           =   1
            Left            =   6570
            RightToLeft     =   -1  'True
            TabIndex        =   214
            Top             =   225
            Width           =   1860
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   210
            Index           =   4
            Left            =   9390
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   225
            Width           =   1380
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   225
            Index           =   2
            Left            =   11685
            RightToLeft     =   -1  'True
            TabIndex        =   118
            Top             =   240
            Width           =   1545
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   210
            Left            =   8610
            RightToLeft     =   -1  'True
            TabIndex        =   117
            Top             =   225
            Width           =   720
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   210
            Left            =   10890
            RightToLeft     =   -1  'True
            TabIndex        =   116
            Top             =   225
            Width           =   705
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   315
            Index           =   0
            Left            =   41040
            TabIndex        =   115
            Top             =   195
            Width           =   2985
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   255
            Index           =   35
            Left            =   27645
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   345
            Width           =   3420
         End
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   330
         Index           =   0
         Left            =   10860
         TabIndex        =   120
         Top             =   9945
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
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
         ButtonImage     =   "FrmVATAvowal.frx":7A54
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
         Height          =   330
         Index           =   1
         Left            =   9705
         TabIndex        =   121
         Top             =   9945
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
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
         ButtonImage     =   "FrmVATAvowal.frx":E2B6
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
         Height          =   330
         Index           =   2
         Left            =   8460
         TabIndex        =   122
         Top             =   9945
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
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
         ButtonImage     =   "FrmVATAvowal.frx":14B18
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
         Height          =   330
         Index           =   3
         Left            =   7395
         TabIndex        =   123
         Top             =   9945
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
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
         ButtonImage     =   "FrmVATAvowal.frx":1B37A
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
         Height          =   330
         Index           =   4
         Left            =   6165
         TabIndex        =   124
         Top             =   9945
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   582
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
         ButtonImage     =   "FrmVATAvowal.frx":21BDC
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
         Height          =   330
         Index           =   6
         Left            =   1980
         TabIndex        =   125
         Top             =   9945
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
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
         ButtonImage     =   "FrmVATAvowal.frx":2843E
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
         Height          =   330
         Index           =   7
         Left            =   4890
         TabIndex        =   126
         Top             =   9945
         Width           =   1200
         _ExtentX        =   2117
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
         ButtonImage     =   "FrmVATAvowal.frx":52060
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
         Height          =   330
         Index           =   9
         Left            =   3525
         TabIndex        =   127
         Top             =   9945
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
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
         ButtonImage     =   "FrmVATAvowal.frx":588C2
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… ÷—Ì»… «·ﬁÌ„… «·„÷«›…"
         Height          =   240
         Index           =   2
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   307
         Top             =   720
         Width           =   1785
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· ⁄œÌ·« "
         Height          =   240
         Index           =   1
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   306
         Top             =   720
         Width           =   1785
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„»·€"
         Height          =   240
         Index           =   0
         Left            =   8580
         RightToLeft     =   -1  'True
         TabIndex        =   275
         Top             =   765
         Width           =   1785
      End
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic7 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   10035
      Width           =   12765
      _cx             =   22516
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
      Begin ImpulseButton.ISButton Accredit 
         Height          =   420
         Left            =   1710
         TabIndex        =   1
         Top             =   90
         Visible         =   0   'False
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   741
         ButtonPositionImage=   1
         Caption         =   "«—”«· ··«⁄ „«œ"
         BackColor       =   -2147483635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColorButton     =   -2147483635
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
   End
   Begin VB.Line Line2 
      X1              =   17775
      X2              =   0
      Y1              =   0
      Y2              =   15
   End
End
Attribute VB_Name = "FrmVATAvowal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim TTP As clstooltip
Dim intervalVat As Double

Function checkfirst() As Boolean
Dim AccountVATCreit As String
Dim manyDes As String
checkfirst = True

 


 GetValueAddedAccount DateFrom.value, , AccountVATCreit, 1, 21
 If SystemOptions.UserInterface = ArabicInterface Then
        manyDes = "      «⁄œ«œ   «·„»Ì⁄«  ··ﬁÌ„Â «·„÷«›…"
 Else
        manyDes = "      VAT On Sales "
 End If
 If Trim(AccountVATCreit) = "" Then
    MsgBox manyDes
    GoTo ErrTrap
 End If


'                GetValueAddedAccount DateFrom.value, AccountVATCreit, , 1, 9
'                If Trim(AccountVATCreit) = "" Then
'                    MsgBox " «⁄œ«œ  „—œÊœ«  «·„»Ì⁄«  ··ﬁÌ„Â «·„÷«›…"
'                    GoTo ErrTrap
'                End If

 


        


       GetValueAddedAccount DateFrom.value, AccountVATCreit, , 1, 22
                 If SystemOptions.UserInterface = ArabicInterface Then
                       manyDes = "       «⁄œ«œ   «·„‘ —Ì«  ··ﬁÌ„Â «·„÷«›… "
                Else
                       manyDes = " VAT On Purchase "
                End If
                
             If Trim(AccountVATCreit) = "" Then
    MsgBox manyDes
    GoTo ErrTrap
 End If
   
                         


'       GetValueAddedAccount DateFrom.value, , AccountVATCreit, 1, 5
'
'          If SystemOptions.UserInterface = ArabicInterface Then
'                    manyDes = "       «⁄œ«œ   „—œÊœ«  «·„‘ —Ì«  ··ﬁÌ„Â «·„÷«›… "
'                Else
'                       manyDes = " VAT On  Return Purchase "
'                End If
'
'
'             If Trim(AccountVATCreit) = "" Then
'    MsgBox manyDes
'    GoTo ErrTrap
' End If
   
                         
checkfirst = True
Exit Function
ErrTrap:
checkfirst = False

 
                  
End Function

Private Sub ChkIsFree_Click()
    If ChkIsFree.value = vbChecked Then
        Sales5.Visible = False
        RSales5.Visible = False
        Label11.Visible = False
        Text1(0).Visible = False
        Text2.Visible = False
        Label66(0).Visible = False
    Else
         Sales5.Visible = True
        RSales5.Visible = True
        Label11.Visible = True
        Text1(0).Visible = True
        Text2.Visible = True
        Label66(0).Visible = True
    End If
End Sub

Private Sub Cmd_Click(Index As Integer)

    'On Error GoTo ErrTrap
    Dim AccountVATCreit As String
    Select Case Index
        Case 0
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
            TxtModFlg.Text = "N"
            clear_all Me
            
            ID.Text = CStr(new_id("TblVATAvowal", "ID", "", True))
            Me.DCboUserName.BoundText = user_id
'            Me.Dcbranch.BoundText = Current_branch
        Case 1
        
        
                If ChekClodePeriod(IIf(IsNull(DateFrom.value), Date, DateFrom.value)) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
              If ChekClodePeriod(IIf(IsNull(DateTo.value), Date, DateTo.value)) = True Then
               'If ChekClodePeriod(DateTo.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
            If CheckPayed() = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·«Ì„ﬂ‰ «· ⁄œÌ· ÌÊÃœ „œ›Ê⁄«  ⁄·Ï Â–Â «·Õ—ﬂ…"
            Else
            MsgBox "Can not edit .This process has payments"
            End If
            Exit Sub
            End If
            TxtModFlg.Text = "E"
        Case 2
        
        If checkfirst = False Then Exit Sub
                        If ChekClodePeriod(DateFrom.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
              
               If ChekClodePeriod(DateTo.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
    Dim Account_Code_dynamic As String
    Account_Code_dynamic = get_account_code_branch(145, my_branch)
    If Account_Code_dynamic = "NO branch" Then
    If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        Else
                MsgBox "Please Create Branch"
        End If
                Exit Sub
            Else

                If Account_Code_dynamic = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ÂÌ∆… «·“ﬂ«…", vbCritical
                 Else
                 MsgBox "Please Select Account"
                 End If
                   Exit Sub
                End If
            End If
       GetValueAddedAccount DateFrom.value, , AccountVATCreit, 1, 21
        If AccountVATCreit = "" And val(Me.TxtSalesVAT.Text) > 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «œŒ«· «⁄œ«œ  «·›«  ··„»Ì⁄« "
        Else
        MsgBox "Please VAT settings for sales"
        End If
        Exit Sub
        End If
        AccountVATCreit = ""
       GetValueAddedAccount DateFrom.value, AccountVATCreit, , 1, 9
        If AccountVATCreit = "" And val(Me.TxtRetSalesVAT.Text) > 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «œŒ«· «⁄œ«œ  «·›«  ·„—œÊœ«  «·„»Ì⁄« "
        Else
        MsgBox "Please VAT settings for sales return"
        End If
        Exit Sub
        End If
        AccountVATCreit = ""
       GetValueAddedAccount DateFrom.value, AccountVATCreit, , 1, 22
        If AccountVATCreit = "" And val(Me.TxtBuyVAT.Text) > 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «œŒ«· «⁄œ«œ  «·›«   ··„‘ —Ì« "
        Else
        MsgBox "Please VAT settings for purchases "
        End If
        Exit Sub
        End If
         AccountVATCreit = ""
       GetValueAddedAccount DateFrom.value, , AccountVATCreit, 1, 5
        If AccountVATCreit = "" And val(Me.TxtRetBuyVAT.Text) > 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «œŒ«· «⁄œ«œ  «·›«   ·„—œÊœ«  «·„‘ —Ì« "
        Else
        MsgBox "Please VAT settings for purchases return"
        End If
        Exit Sub
        End If
        AccountVATCreit = ""
        GetValueAddedAccount DateFrom.value, , AccountVATCreit, 1, 8
        If AccountVATCreit = "" And val(Me.TxtContractVAT.Text) > 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «œŒ«· «⁄œ«œ  «·›«  ·⁄ﬁœ «·«ÌÃ«—  Ã«—Ì"
        Else
        MsgBox "Please VAT settings for lease contract"
        End If
        Exit Sub
        End If
         AccountVATCreit = ""
         GetValueAddedAccount DateFrom.value, , AccountVATCreit, 1, 6
        If AccountVATCreit = "" And val(Me.TxtProjCusVAT.Text) > 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «œŒ«· «⁄œ«œ  «·›«    ·„” Œ·’«  «·„‘«—Ì⁄ ··⁄„Ì·"
        Else
        MsgBox "Please VAT settings for Projects for the client"
        End If
        Exit Sub
        End If
         AccountVATCreit = ""
         GetValueAddedAccount DateFrom.value, , AccountVATCreit, 1, 3
        If AccountVATCreit = "" And val(TxtOmraVAT.Text) > 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «œŒ«· «⁄œ«œ  «·›«  ·«„— ‘€· «·⁄„—…"
        Else
        MsgBox "Please VAT settings for Omrah"
        End If
        Exit Sub
        End If
         AccountVATCreit = ""
         GetValueAddedAccount DateFrom.value, , AccountVATCreit, 1, 4
        If AccountVATCreit = "" And val(TxtHajjVAT.Text) > 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «œŒ«· «⁄œ«œ  «·›«  ·«„— ‘€· «·ÕÃ"
        Else
        MsgBox "Please VAT settings for Hijj"
        End If
        Exit Sub
        End If
           AccountVATCreit = ""
         GetValueAddedAccount DateFrom.value, , AccountVATCreit, 1, 1
        If AccountVATCreit = "" And val(TxtMinisterVAT.Text) > 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «œŒ«· «⁄œ«œ  «·›«  ·«” Õﬁ«ﬁ «·Ê“«—…"
        Else
        MsgBox "Please VAT settings for Entitlement of the Ministry"
        End If
        Exit Sub
        End If
             AccountVATCreit = ""
         GetValueAddedAccount DateFrom.value, AccountVATCreit, , 1, 7
        If AccountVATCreit = "" And val(TxtProjConVAT.Text) > 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «œŒ«· «⁄œ«œ  «·›«   ·„” Œ·’«  «·„‘«—Ì⁄ ··„ﬁ«Ê·Ì‰"
        Else
        MsgBox "Please VAT settings for Contractors Projects"
        End If
        End If
                     AccountVATCreit = ""
        GetValueAddedAccount DateFrom.value, AccountVATCreit, , 1, 47
        If AccountVATCreit = "" And val(TxtReqConVAT.Text) > 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Ì—ÃÏ «œŒ«· «⁄œ«œ  «·›«     ·«” Õﬁ«ﬁ «·„ ⁄ÂœÌ‰"
        Else
        MsgBox "Please VAT settings for Entitlement of contractors"
        End If
        Exit Sub
        End If
            SaveData
        Case 3
            Undo
        Case 4
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If
                      If CheckPayed() = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·«Ì„ﬂ‰ «·Õ–› ÌÊÃœ „œ›Ê⁄«  ⁄·Ï Â–Â «·Õ—ﬂ…"
            Else
            MsgBox "Can not delete .This process has payments"
            End If
            Exit Sub
            End If
            Del_Action
        Case 5

        Case 6
            Unload Me
        Case 7
            print_report
        Case 9

    End Select
    Exit Sub
ErrTrap:
End Sub

Private Sub CmdAttach_Click()

End Sub

Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.Text, , 200
End Sub

Private Sub DateFrom_Change()
DTPicker1.value = DateFrom.value
End Sub

Private Sub Form_Load()

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    
    'On Error GoTo ErrTrap
    
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    C1Tab2.CurrTab = 0
    Dcombos.GetUsers DCboUserName
    Dcombos.GetBranches Me.DcBranch
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    
    Resize_Form Me
    
    Set rs = New ADODB.Recordset
    
    Dim StrSQL As String
    StrSQL = ""
    If SystemOptions.usertype <> UserAdminAll Then
        StrSQL = "SELECT  *  From TblVATAvowal"
    Else
        StrSQL = "SELECT  *  From TblVATAvowal"
    End If
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    RecoredDate.value = Date
    DateFrom.value = Date
    DateFrom.value = ""
    DateTo.value = Date
    DateTo.value = ""
    
    Me.TxtModFlg.Text = "R"
    XPBtnMove_Click 2

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
    
    Exit Sub
ErrTrap:
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
                SaveData
            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub
Private Sub ChangeLang()
    
    Dim XPic As IPictureDisp

    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    
    lbl(40).Caption = "Branch"
    lbl(8).Caption = "Ser."
    Label3.Caption = "Date"
    lbl(24).Caption = "Notes"
    Label19.Caption = "Period"
    Label7.Caption = "From"
    Label14.Caption = "To"
    paidchk(0).Caption = "Paid"
    ShowBtn.Caption = "Get Data"
    C1Tab2.TabCaption(0) = "All"
    C1Tab2.TabCaption(1) = "5% Details"
    C1Tab2.TabCaption(2) = "0% Details"
    C1Tab2.TabCaption(3) = "Non Vat"
    C1Tab2.TabCaption(4) = "Remarks"
    Label81.Caption = "Edited within 5000 SR"
    Label82.Caption = "VAT From Old Period"

    Label17.Caption = "VAT on Sales"
'    Label10(1).Caption = "Amount"
    Label33.Caption = "Amount"
    Label10(0).Caption = "Adjustment"
    Label10(1).Caption = "Adjustment"
    Label10(2).Caption = "VAT Amount"
    'Label48.Caption = "Adjustment"
    Label2.Caption = "Standard Rated Sales"
    Label6.Caption = "Sales to Customers in VAT Implementing GCC Countries"
    Label5.Caption = "Zero Rated Domestic Sales"
    lbl(3).Caption = "Exports"
    Label11.Caption = "Exempt Sales"
    Label12.Caption = "Total Sales"
    
    Label20.Caption = "VAT on Purchases"
    Label24.Caption = "Amount"
  
    Label31.Caption = "VAT Amount"
    
    Label13.Caption = "Standard Rated Purchases"
    Label9.Caption = "Imports Subject to VAT Paid at Customs"
    Label8.Caption = "Imports Subject to VAT Accounted Through Reverse Charge Mechanism"
    Label23.Caption = "Zero Rated Domestic Purchases"
    Label22.Caption = "Exempt Purchases"
    Label21.Caption = "Total Purchases"
    
    Label30.Caption = "Sales tax analysis for 5%"
    Label52.Caption = "Manual Entries"
    Label42.Caption = "Seals Invoices"
    Label43.Caption = "Commercial Lease Contract"
    Label41.Caption = "Projects Abstracts for Client"
    Label46.Caption = "Omra Order"
    Label51.Caption = "Hijj Order"
    Label49.Caption = "Ministerial Merit 5%"
    Label50.Caption = "Vehicles Maintenance Invoice"
    Label25.Caption = "Transportation Services Invoice"
    Label80.Caption = "Service Invoice"
    Label28.Caption = "Total"
    Label29.Caption = "Purchase tax analysis for 5%"
    Label53.Caption = "Manual Entries"
    Label44.Caption = "Purchase Invoices"
    Label38.Caption = "Projects Abstracts for Contractor"
    Label39.Caption = "Contractors Entitlements"
    Label40.Caption = "Assets Purchasing"
    Label54.Caption = "Expenses"
    Label36.Caption = "Total"
    
    Label34.Caption = "Sales tax analysis for 5%"
    Label55.Caption = "Manual Entries"
    Label47.Caption = "Domestic Sales Invoice With 0 Percentage"
    Label18.Caption = "Ministerial Merit 0%"
    Label64.Caption = "Total"
    Label35.Caption = "Purchase tax analysis for 0%"
    lbl(10).Caption = "Manual Entries"
    lbl(11).Caption = "Purchase Invoices"
    lbl(12).Caption = "Projects Abstracts for Contractor"
    lbl(13).Caption = "Total"
    
    Label59.Caption = "Taxes On"
    Label69.Caption = "Manual Entries"
     Label66(0).Caption = "Seals Invoce Wothout tax"
    
    Label75(17).Caption = "Remarks on Purchases"
    Label77.Caption = "Purchases has taxes but supplier didn't pay"
    
    Label72.Caption = "Total Remarks"
    
    Label75(9).Caption = "Analytical Expenses"
    Label75(0).Caption = "Analytical Expenses"
    Label75(2).Caption = "Liquidation Custody"
    Label75(1).Caption = "Financial Invoices"
    Label75(3).Caption = "VAT for Customs"
    Label75(4).Caption = "VAT for Advance Payments"
    Label75(5).Caption = "VAT for Bank Transfers"
    Label75(6).Caption = "Total"
    
   
    
  
    Label75(11).Caption = "Amount"
    Label75(12).Caption = "Adjustments"
 
  
    
 
    
    Label27.Caption = "Total Net"
    Label1(35).Caption = "Voucher No."
    lbl(1).Caption = "Voucher No."
    Command9.Caption = "Print Voucher"
    
    lbl(0).Caption = "By"
    
    lbl(2).Caption = "Current Record"
    lbl(4).Caption = "NO. Recordes"
    
    
    paidchk(0).RightToLeft = False
    
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(9).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
'    CmdAttach.Caption = "Attachment"
    
    Me.Caption = "VAT Form"
    Label1(2).Caption = Me.Caption
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  »Ì«‰«   ﬁÌÌ„ «·„ÊŸ›Ì‰   "
    LogTexte = " Exit Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

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



Private Sub Label73_Click()

End Sub

Private Sub Label61_Click()

End Sub

Private Sub PurcahseRemarks1_Click()
Dim X As Variant
X = getTotalVATItemsValue(22, False, , 0, True, Label77.Caption)
 
End Sub

Private Sub PurcahseRemarks1Ret_Click()
Dim X As Variant
PurcahseRemarks1Ret.Text = getTotalVATItemsValue(5, False, , 0, True, " ⁄œÌ·«  „‘ —Ì«  Œ«÷⁄Â Ê·„ Ìﬁ„ «·„Ê—œ »«÷«› Â«")

End Sub

Private Sub PurcahseRemarks2_Click()

Dim X As String
X = getTotalVATItemsValue(22, False, , 1, True, " „‘ —Ì«  Œ«÷⁄Â Ê·ﬂ‰ «·„Ê—œ „⁄›Ì")


End Sub

Private Sub PurcahseRemarks2Ret_Click()
Dim X As String
X = getTotalVATItemsValue(5, False, , 1, True, " ⁄œÌ·«   „‘ —Ì«  Œ«÷⁄Â Ê·ﬂ‰ «·„Ê—œ „⁄›Ì")

End Sub

 

 

Private Sub PurchasesRett5_DblClick()
Dim X As String
X = getTotalVATItemsValue(5, True, 5, , True, " ⁄œÌ·«  «·„‘ —Ì«  «·Œ«÷⁄Â 5%")

End Sub

Private Sub PurchasesT5_DblClick()
Dim X As String
X = getTotalVATItemsValue(22, True, 5, , True, "«·„‘ —Ì«  «·Œ«÷⁄Â 5%")
End Sub

Private Sub RecoredDate_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
  TxtNoteSerial.Text = ""
End Sub

Function CalnNet()
 


TxtNetVat = val(TotalNetTxt) + val(TxtOldVat) + val(TxtCorrect1)

End Function

Private Sub RSales5_DblClick()

 

    Dim X As String
    
    X = getTotalVATItemsValue(9, False, , , True, "«·„—œÊœ«  «·„⁄›«Â")
    
End Sub

 

Private Sub Sales5_DblClick()
Dim X As String
 

      X = getTotalVATItemsValue(21, False, , , True, " «·„»Ì⁄«  «·„⁄›«Â")
    
   ' X = getTotalVATItemsValue(9, False, , , Trueÿ, "«·„—œÊœ«  «·„⁄›«Â")
    

End Sub

Private Sub SalesRet5_Click()
Dim X As String
X = getTotalVATItemsValue(9, True, 5, , True, " ⁄œÌ·«  «·„»Ì⁄«  «·Œ«÷⁄Â 5%")
End Sub

Private Sub SalesRetZero_Click()
Dim X As String
  X = getTotalVATItemsValue(9, False, 0, , True, " ⁄œÌ·«  «·„»Ì⁄«  «·’›—Ì…")

End Sub

Private Sub SalesT5_Click()
Dim X As String
X = getTotalVATItemsValue(21, True, 5, , True, "«·„»Ì⁄«  «·Œ«÷⁄Â 5% ")
      
End Sub

Private Sub SalesZero_Click()
Dim X As String
X = getTotalVATItemsValue(21, True, 0, , True, "›Ê« Ì— «·„»Ì⁄«  «·„Õ·Ì… «·Œ«÷⁄… ··‰”»… «·’›—Ì… ")
End Sub

 

Private Sub TxtCorrect1_Change()
If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
CalnNet
End If
End Sub

Private Sub TxtModFlg_Change()
    
    On Error GoTo ErrTrap
    
    Select Case Me.TxtModFlg.Text
    
        Case "R"
            Cmd(0).Enabled = True
            Cmd(1).Enabled = True
            Cmd(2).Enabled = False
            Cmd(3).Enabled = False
            Cmd(4).Enabled = True
            Cmd(7).Enabled = True
            
            NotesTxt.Enabled = False
            DateFrom.Enabled = False
            DateTo.Enabled = False
            ShowBtn.Enabled = False
            paidchk(0).Enabled = False
        Case "N"
            Cmd(0).Enabled = False
            Cmd(1).Enabled = False
            Cmd(2).Enabled = True
            Cmd(3).Enabled = True
            Cmd(4).Enabled = False
            Cmd(7).Enabled = False
            
            NotesTxt.Enabled = True
            DateFrom.Enabled = True
            DateTo.Enabled = True
            ShowBtn.Enabled = True
            paidchk(0).Enabled = True
        Case "E"
            Cmd(0).Enabled = False
            Cmd(1).Enabled = False
            Cmd(2).Enabled = True
            Cmd(3).Enabled = True
            Cmd(4).Enabled = False
            Cmd(7).Enabled = False
            
            NotesTxt.Enabled = True
            DateFrom.Enabled = True
            DateTo.Enabled = True
            ShowBtn.Enabled = True
            paidchk(0).Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub
Function print_report(Optional NoteSerial As String, Optional sqlPrint As String, Optional printTitle As String)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
    If sqlPrint = "" Then
    MySQL = "select * from TblVATAvowal where TblVATAvowal.ID = " & ID.Text


    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\VAT\" & "RepVATAvowal.rpt"
    Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\VAT\" & "RepVATAvowalE.rpt"
    End If
  Else
  
  MySQL = sqlPrint
      If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\VAT\" & "newVatCheck.rpt"
    Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\VAT\" & "newVatCheck.rpt"
    End If
    
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
 
         StrReportTitle = printTitle

 

    'xReport.ParameterFields(3).AddCurrentValue user_name

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
 Function createVoucher() As Boolean
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
If SystemOptions.UserInterface = ArabicInterface Then
            des = "   «·«ﬁ—«— «·÷—Ì»Ì " & ID.Text
Else
            des = "   VAT Form " & ID.Text
End If
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
Dim Sql As String
tablename = "TblVATAvowal"
Filedname = "ID"
NoteSerial1 = val(ID.Text)
Notevalue = 0
 notytype = 9081
Notevalue = 1
BranchID = Current_branch
NoteDate = (RecoredDate.value)
 
If Notevalue > 0 Or val(TxtOldVat) <> 0 Then
                              
                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des
                                              TxtNoteID.Text = NoteID
                                                    TxtNoteSerial.Text = NoteSerial
                            
Cn.Execute " update tblbookingrequest2 set UserVouchID=" & user_id & " where ID=" & val(ID.Text) & " "

If Not CREATE_VOUCHER_GE(val(TxtNoteID.Text), BranchID, user_id, NoteDate) Then createVoucher = False: Exit Function


rs.Resync adAffectCurrent
 
updateNotesValueAndNobytext val(TxtNoteID.Text)
     End If
     createVoucher = True
End Function
Function CheckPayed() As Boolean
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
Sql = " select * from Notes where VATVowalNo =" & val(ID.Text) & ""
Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs2.RecordCount > 0 Then
CheckPayed = True
Else
CheckPayed = False
End If
End Function
Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date) As Boolean

 Dim Notevalue As Single
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempCustomerCode As String
    Dim StrTempCustomerCodeInsuranceAccount  As String
    
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
 Dim valuee As Double
 Dim StrSQL As String
 Dim AccountVATCreit As String
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
       LngDevNO = 0
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„Ì‰
    my_branch = BranchID
     StrTempAccountCode = get_account_code_branch(145, my_branch)
     If SystemOptions.UserInterface = ArabicInterface Then
               StrTempDes = "«·«ﬁ—«— «·÷—Ì»Ì —ﬁ„   " & ID.Text
     Else
     StrTempDes = " VAT FORM NO:  " & ID.Text
     End If
    valuee = val(VATSalesTotal.Text)
    Dim manyDes As String
             LngDevNO = LngDevNO + 1
    If valuee > 0 Then
      
 GetValueAddedAccount DateFrom.value, , AccountVATCreit, 1, 21
 If SystemOptions.UserInterface = ArabicInterface Then
        manyDes = "      Õ”«» «·ﬁÌ„… «·„÷«›… „»Ì⁄«  "
 Else
        manyDes = "      VAT On Sales "
 End If
 If Trim(AccountVATCreit) = "" Then
    MsgBox "„‰ ›÷·ﬂ ﬁ„ » ⁄—Ì›" & manyDes
    GoTo ErrTrap
 End If
' AccountVATCreit = StrTempAccountCode
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, valuee, 0, StrTempDes & manyDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
             LngDevNO = LngDevNO + 1
 
 
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, valuee, 1, StrTempDes & manyDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            LngDevNO = LngDevNO + 1
        End If
    ''///////////////„—œÊœ«  «·„»Ì⁄« 
    valuee = 0 'val(TxtRetSalesVAT.Text)
        If valuee > 0 Then
                 If SystemOptions.UserInterface = ArabicInterface Then
                manyDes = "      ﬁÌ„Â „÷«›Â „—œÊœ«  „»Ì⁄«  "
                Else
                       manyDes = " VAT On Return Sales "
                End If
        
                GetValueAddedAccount DateFrom.value, AccountVATCreit, , 1, 9
                If Trim(AccountVATCreit) = "" Then
                    MsgBox "„‰ ›÷·ﬂ ﬁ„ » ⁄—Ì›" & manyDes
                    GoTo ErrTrap
                End If
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, valuee, 0, StrTempDes & manyDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            
             LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, valuee, 1, StrTempDes & manyDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            LngDevNO = LngDevNO + 1
    End If

    
            ''/////////////// «·„‘ —Ì« 
    valuee = val(VATPurchasesTotal.Text)
        If valuee > 0 Then
                         
            If SystemOptions.UserInterface = ArabicInterface Then
                       manyDes = "      ﬁÌ„Â „÷«›Â   «·„‘ —Ì«  "
                Else
                       manyDes = " VAT On Purchase "
                End If


       GetValueAddedAccount DateFrom.value, AccountVATCreit, , 1, 22
             
                        If Trim(AccountVATCreit) = "" Then
                            MsgBox "„‰ ›÷·ﬂ ﬁ„ » ⁄—Ì›" & manyDes
                            GoTo ErrTrap
                        End If
                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, valuee, 0, StrTempDes & manyDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
                  LngDevNO = LngDevNO + 1
                If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, valuee, 1, StrTempDes & manyDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                    End If
               LngDevNO = LngDevNO + 1
       ElseIf valuee < 0 Then
       valuee = Abs(valuee)
       
                   If SystemOptions.UserInterface = ArabicInterface Then
                       manyDes = "      ﬁÌ„Â „÷«›Â   «·„‘ —Ì«  »«·”«·» "
                Else
                       manyDes = " VAT On Purchase "
                End If


       GetValueAddedAccount DateFrom.value, AccountVATCreit, , 1, 22
             
                        If Trim(AccountVATCreit) = "" Then
                            MsgBox "„‰ ›÷·ﬂ ﬁ„ » ⁄—Ì›" & manyDes
                            GoTo ErrTrap
                        End If
                
                If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, valuee, 0, StrTempDes & manyDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                    End If
               LngDevNO = LngDevNO + 1
                    
                    
                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, valuee, 1, StrTempDes & manyDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If
                  LngDevNO = LngDevNO + 1
           
    End If
    
                ''///////////////„—œÊœ«  «·„‘ —Ì« 
    valuee = 0
        If valuee > 0 Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                manyDes = "      ﬁÌ„Â „÷«›Â „—œÊœ«   «·„‘ —Ì«  "
                Else
                       manyDes = " VAT On  Return Purchase "
                End If
                
       GetValueAddedAccount DateFrom.value, , AccountVATCreit, 1, 5
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, valuee, 0, StrTempDes & manyDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
             LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, valuee, 1, StrTempDes & manyDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            LngDevNO = LngDevNO + 1
    End If
    
     StrTempAccountCode = get_account_code_branch(145, my_branch)
     If SystemOptions.UserInterface = ArabicInterface Then
               StrTempDes = "«  ÷—Ì»… «·ﬁÌ„… «·„÷«›… «· Ì  „  —ÕÌ·Â« „‰ «·› —…/ «·› —«  «·”«»ﬁ…  ··«ﬁ—«— —ﬁ„  " & ID.Text
     Else
     StrTempDes = " VAT FORM NO:  " & ID.Text
     End If
    valuee = val(TxtVatADD.Text)
     
             LngDevNO = LngDevNO + 1
    If valuee > 0 Then
      
 GetValueAddedAccount DateFrom.value, , AccountVATCreit, 1, 21
 If SystemOptions.UserInterface = ArabicInterface Then
        manyDes = "     "
 Else
        manyDes = "       "
 End If
 
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, valuee, 0, StrTempDes & manyDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
             LngDevNO = LngDevNO + 1
 
 
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, valuee, 1, StrTempDes & manyDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            LngDevNO = LngDevNO + 1
        End If
    
    
    
    
    
    
    
                ''/////////////// «·„‘ —Ì«   ÷—Ì»… «·ﬁÌ„… «·„÷«›… «· Ì  „  —ÕÌ·Â« „‰ «·› —…/ «·› —«  «·”«»ﬁ…  ··«ﬁ—«—
    valuee = val(TxtVatDis.Text)
        If valuee > 0 Then
               If SystemOptions.UserInterface = ArabicInterface Then
                   manyDes = "    ÷—Ì»… «·ﬁÌ„… «·„÷«›… «· Ì  „  —ÕÌ·Â« „‰ «·› —…/ «·› —«  «·”«»ﬁ…  ··«ﬁ—«— " & ID.Text
                Else
                       manyDes = " VAT On Purchase " & ID.Text
                End If


       GetValueAddedAccount DateFrom.value, AccountVATCreit, , 1, 22
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, valuee, 0, StrTempDes & manyDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
             LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, valuee, 1, StrTempDes & manyDes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            LngDevNO = LngDevNO + 1
    End If


''///////////////
   CREATE_VOUCHER_GE = True
    Exit Function
    
    
Exit Function 'ahmed salimmmmmmmmmmmmmmmmmmm

valuee = val(TxtContractVAT.Text)
    If valuee > 0 Then
      
 GetValueAddedAccount DateFrom.value, , AccountVATCreit, 1, 8
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, valuee, 0, StrTempDes & "      Õ”«» «·ﬁÌ„… «·„÷«›… ⁄ﬁœ «ÌÃ«—  Ã«—Ì ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
             LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, valuee, 1, StrTempDes & "    Õ”«» ÂÌ∆… «·“ﬂ«… ⁄ﬁœ «ÌÃ«—  Ã«—Ì ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            LngDevNO = LngDevNO + 1
    End If
    
    valuee = val(TxtProjCusVAT.Text)
    If valuee > 0 Then
      
 GetValueAddedAccount DateFrom.value, , AccountVATCreit, 1, 6
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, valuee, 0, StrTempDes & "      Õ”«» «·ﬁÌ„… «·„÷«›… „” Œ·’«  „‘«—Ì⁄ ··⁄„Ì· ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
             LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, valuee, 1, StrTempDes & "    Õ”«» ÂÌ∆… «·“ﬂ«… „” Œ·’«  „‘«—Ì⁄ ··⁄„Ì· ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            LngDevNO = LngDevNO + 1
    End If
        valuee = val(TxtOmraVAT.Text)
    If valuee > 0 Then
      
 GetValueAddedAccount DateFrom.value, , AccountVATCreit, 1, 3
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, valuee, 0, StrTempDes & "      Õ”«» «·ﬁÌ„… «·„÷«›… «·⁄„— ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
             LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, valuee, 1, StrTempDes & "    Õ”«» ÂÌ∆… «·“ﬂ«… «·⁄„— ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            LngDevNO = LngDevNO + 1
    End If
            valuee = val(TxtHajjVAT.Text)
    If valuee > 0 Then
      
 GetValueAddedAccount DateFrom.value, , AccountVATCreit, 1, 4
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, valuee, 0, StrTempDes & "      Õ”«» «·ﬁÌ„… «·„÷«›… «·ÕÃ ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
             LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, valuee, 1, StrTempDes & "    Õ”«» ÂÌ∆… «·“ﬂ«… «·ÕÃ ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            LngDevNO = LngDevNO + 1
    End If
                valuee = val(TxtMinisterVAT.Text)
    If valuee > 0 Then
      
 GetValueAddedAccount DateFrom.value, , AccountVATCreit, 1, 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, valuee, 0, StrTempDes & "      Õ”«» «·ﬁÌ„… «·„÷«›…  «” Õﬁ«ﬁ «·Ê“«—… ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
             LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, valuee, 1, StrTempDes & "    Õ”«» ÂÌ∆… «·“ﬂ«…«” Õﬁ«ﬁ «·Ê“«—… ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            LngDevNO = LngDevNO + 1
    End If
    
        valuee = val(TxtProjConVAT.Text)
        If valuee > 0 Then
       GetValueAddedAccount DateFrom.value, AccountVATCreit, , 1, 7
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, valuee, 0, StrTempDes & "     Õ”«» ÂÌ∆… «·“ﬂ«… „” Œ·’«  «·„‘«—Ì⁄ ··„ﬁ«Ê·", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
             LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, valuee, 1, StrTempDes & "    Õ”«» «·ﬁÌ„… «·„÷«›…  „” Œ·’«  «·„‘«—Ì⁄ ··„ﬁ«Ê·  ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            LngDevNO = LngDevNO + 1
    End If
    
            valuee = val(TxtReqConVAT.Text)
        If valuee > 0 Then
       GetValueAddedAccount DateFrom.value, AccountVATCreit, , 1, 2
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, valuee, 0, StrTempDes & "     Õ”«» ÂÌ∆… «·“ﬂ«… «” Õﬁ«ﬁ «·„ ⁄ÂœÌ‰", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
             LngDevNO = LngDevNO + 1
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, valuee, 1, StrTempDes & "    Õ”«» «·ﬁÌ„… «·„÷«›…  „” Œ·’«  «” Õﬁ«ﬁ «·„ ⁄ÂœÌ‰  ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
            LngDevNO = LngDevNO + 1
    End If
    
    
    '55555555555555555555
    
    CREATE_VOUCHER_GE = True
    Exit Function
    
ErrTrap:
CREATE_VOUCHER_GE = False
Exit Function

End Function
Public Sub Retrive(Optional Lngid As Long = 0)

    'On Error GoTo ErrTrap
    
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If
    
     If rs.EOF Or rs.BOF Then
        Exit Sub
    Else
        If Lngid <> 0 Then
            rs.Find "ID =" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If
    TxtNoteSerial.Text = IIf(IsNull(rs.Fields("NoteSerial").value), "", rs.Fields("NoteSerial").value)
    Me.TxtNoteID.Text = IIf(IsNull(rs.Fields("NoteID").value), "", rs.Fields("NoteID").value)
    
    ID.Text = IIf(IsNull(rs("ID").value), "", (rs("ID").value))
    Me.TxtSalesVAT.Text = IIf(IsNull(rs.Fields("SalesVAT").value), 0, rs.Fields("SalesVAT").value)
    Me.tztAdvBill.Text = IIf(IsNull(rs.Fields("tztAdvBill").value), 0, rs.Fields("tztAdvBill").value)
          
        
Me.Text2.Text = IIf(IsNull(rs.Fields("RetSaleWithOutVat").value), 0, rs.Fields("RetSaleWithOutVat").value)
Me.Text1(0).Text = IIf(IsNull(rs.Fields("SaleWithOutVat").value), 0, rs.Fields("SaleWithOutVat").value)
    
    Me.TxtRetSalesVAT.Text = IIf(IsNull(rs.Fields("RetSalesVAT").value), 0, rs.Fields("RetSalesVAT").value)
    Me.TxtBuyVAT.Text = IIf(IsNull(rs.Fields("BuyVAT").value), 0, rs.Fields("BuyVAT").value)
    Me.TxtRetBuyVAT.Text = IIf(IsNull(rs.Fields("RetBuyVAT").value), 0, rs.Fields("RetBuyVAT").value)
    RecoredDate.value = IIf(IsNull(rs("RecoredDate").value), Date, rs("RecoredDate").value)
    NotesTxt.Text = IIf(IsNull(rs("Notes").value), "", Trim(rs("Notes").value))
    DateFrom.value = IIf(IsNull(rs("DateFrom").value), Date, rs("DateFrom").value)
    DateTo.value = IIf(IsNull(rs("DateTo").value), Date, rs("DateTo").value)
    
    If Not IsNull(rs("Paid").value) Then
        If rs("Paid").value = True Then
            paidchk(0).value = vbChecked
        Else
            paidchk(0).value = vbUnchecked
        End If
    Else
        paidchk(0).value = vbUnchecked
    End If
    
    
    If Not IsNull(rs("ChkIsFree").value) Then
        If rs("ChkIsFree").value = True Then
            ChkIsFree.value = vbChecked
        Else
            ChkIsFree.value = vbUnchecked
        End If
    Else
        ChkIsFree.value = vbUnchecked
    End If
    
    
    
  If Not IsNull(rs("HideakarExitVat1").value) Then
        If rs("HideakarExitVat1").value = True Then
            paidchk(1).value = vbChecked
        Else
            paidchk(1).value = vbUnchecked
        End If
    Else
        paidchk(1).value = vbUnchecked
    End If
    
    
  If Not IsNull(rs("HideakarExitVat2").value) Then
        If rs("HideakarExitVat2").value = True Then
            paidchk(2).value = vbChecked
        Else
            paidchk(2).value = vbUnchecked
        End If
    Else
        paidchk(2).value = vbUnchecked
    End If
    
   DcBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", Trim(rs("BranchID").value))
    
    Sales1.Text = IIf(IsNull(rs("Sales1").value), "", Trim(rs("Sales1").value))
    Sales2.Text = IIf(IsNull(rs("Sales2").value), "", Trim(rs("Sales2").value))
    Sales3.Text = IIf(IsNull(rs("Sales3").value), "", Trim(rs("Sales3").value))
    Sales4.Text = IIf(IsNull(rs("Sales4").value), "", Trim(rs("Sales4").value))
    Sales5.Text = IIf(IsNull(rs("Sales5").value), "", Trim(rs("Sales5").value))
    SalesTotal.Text = IIf(IsNull(rs("SalesTotal").value), "", Trim(rs("SalesTotal").value))
    
    RSales1.Text = IIf(IsNull(rs("RSales1").value), "", Trim(rs("RSales1").value))
    RSales2.Text = IIf(IsNull(rs("RSales2").value), "", Trim(rs("RSales2").value))
    RSales3.Text = IIf(IsNull(rs("RSales3").value), "", Trim(rs("RSales3").value))
    RSales4.Text = IIf(IsNull(rs("RSales4").value), "", Trim(rs("RSales4").value))
    RSales5.Text = IIf(IsNull(rs("RSales5").value), "", Trim(rs("RSales5").value))
    RSalesTotal.Text = IIf(IsNull(rs("RSalesTotal").value), "", Trim(rs("RSalesTotal").value))
    
    VATSales1.Text = IIf(IsNull(rs("VATSales1").value), "", Trim(rs("VATSales1").value))
    'VATSales2.Text = IIf(IsNull(rs("VATSales2").value), "", Trim(rs("VATSales2").value))
    'VATSales3.Text = IIf(IsNull(rs("VATSales3").value), "", Trim(rs("VATSales3").value))
    'VATSales4.Text = IIf(IsNull(rs("VATSales4").value), "", Trim(rs("VATSales4").value))
    'VATSales5.Text = IIf(IsNull(rs("VATSales5").value), "", Trim(rs("VATSales5").value))
    VATSalesTotal.Text = IIf(IsNull(rs("VATSalesTotal").value), "", Trim(rs("VATSalesTotal").value))
    
    Me.TxtDept(0).Text = IIf(IsNull(rs("ValueDept").value), "", Trim(rs("ValueDept").value))
    Me.TxtDept(1).Text = IIf(IsNull(rs("ValueDept2").value), "", Trim(rs("ValueDept2").value))
    Me.TxtDept(2).Text = IIf(IsNull(rs("ValueCredit").value), "", Trim(rs("ValueCredit").value))
    Me.TxtDept(3).Text = IIf(IsNull(rs("ValueCredit2").value), "", Trim(rs("ValueCredit2").value))
    
    Me.TxtDept(4).Text = IIf(IsNull(rs("ValueNew4").value), "", Trim(rs("ValueNew4").value))
    Me.TxtDept(5).Text = IIf(IsNull(rs("ValueNew5").value), "", Trim(rs("ValueNew5").value))
    Me.TxtDept(6).Text = IIf(IsNull(rs("ValueNew6").value), "", Trim(rs("ValueNew6").value))
    Me.TxtDept(7).Text = IIf(IsNull(rs("ValueNew7").value), "", Trim(rs("ValueNew7").value))
    
    
        
    Purchases1.Text = IIf(IsNull(rs("Purchases1").value), "", Trim(rs("Purchases1").value))
    Purchases2.Text = IIf(IsNull(rs("Purchases2").value), "", Trim(rs("Purchases2").value))
    Purchases3.Text = IIf(IsNull(rs("Purchases3").value), "", Trim(rs("Purchases3").value))
    Purchases4.Text = IIf(IsNull(rs("Purchases4").value), "", Trim(rs("Purchases4").value))
    Purchases5.Text = IIf(IsNull(rs("Purchases5").value), "", Trim(rs("Purchases5").value))
    PurchasesTotal.Text = IIf(IsNull(rs("PurchasesTotal").value), "", Trim(rs("PurchasesTotal").value))
    
    RPurchases1.Text = IIf(IsNull(rs("RPurchases1").value), "", Trim(rs("RPurchases1").value))
    RPurchases2.Text = IIf(IsNull(rs("RPurchases2").value), "", Trim(rs("RPurchases2").value))
    RPurchases3.Text = IIf(IsNull(rs("RPurchases3").value), "", Trim(rs("RPurchases3").value))
    RPurchases4.Text = IIf(IsNull(rs("RPurchases4").value), "", Trim(rs("RPurchases4").value))
    RPurchases5.Text = IIf(IsNull(rs("RPurchases5").value), "", Trim(rs("RPurchases5").value))
    RPurchasesTotal.Text = IIf(IsNull(rs("RPurchasesTotal").value), "", Trim(rs("RPurchasesTotal").value))
    
    VATPurchases1.Text = IIf(IsNull(rs("VATPurchases1").value), "", Trim(rs("VATPurchases1").value))
    VATPurchases2.Text = IIf(IsNull(rs("VATPurchases2").value), "", Trim(rs("VATPurchases2").value))
    VATPurchases3.Text = IIf(IsNull(rs("VATPurchases3").value), "", Trim(rs("VATPurchases3").value))
    'VATPurchases4.Text = IIf(IsNull(rs("VATPurchases4").value), "", Trim(rs("VATPurchases4").value))
    'VATPurchases5.Text = IIf(IsNull(rs("VATPurchases5").value), "", Trim(rs("VATPurchases5").value))
    VATPurchasesTotal.Text = IIf(IsNull(rs("VATPurchasesTotal").value), "", Trim(rs("VATPurchasesTotal").value))
    
    DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", Trim(rs("UserID").value))
    ''////////
Text1(1).Text = IIf(IsNull(rs("ContractVaueHousing").value), 0, Trim(rs("ContractVaueHousing").value)) '”ﬂ‰Ì „⁄›Ì
    TxtContractVaue.Text = IIf(IsNull(rs("ContractVaue").value), 0, Trim(rs("ContractVaue").value))
    
    DB_CreateField "TblVATAvowal", "FaBuy", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblVATAvowal", "FaBuy2", adDouble, adColNullable, , , "    ", False, True
    DB_CreateField "TblVATAvowal", "FaBuy3", adDouble, adColNullable, , , "    ", False, True
   
   txtFaBuy = IIf(IsNull(rs("FaBuy").value), 0, Trim(rs("FaBuy").value))
   txtFaBuy2 = IIf(IsNull(rs("FaBuy3").value), 0, Trim(rs("FaBuy2").value))
   txtFaBuy3 = IIf(IsNull(rs("FaBuy3").value), 0, Trim(rs("FaBuy3").value))
    
    TxtContractReVaue.Text = IIf(IsNull(rs("ContractReVaue").value), 0, Trim(rs("ContractReVaue").value))
    TxtContractVAT.Text = IIf(IsNull(rs("ContractVAT").value), 0, Trim(rs("ContractVAT").value))
    TxtProjCusValue.Text = IIf(IsNull(rs("ProjCusValue").value), 0, Trim(rs("ProjCusValue").value))
    TxtProjCusReValue.Text = IIf(IsNull(rs("ProjCusReValue").value), 0, Trim(rs("ProjCusReValue").value))
    TxtProjCusVAT.Text = IIf(IsNull(rs("ProjCusVAT").value), 0, Trim(rs("ProjCusVAT").value))
    TxtOmraValue.Text = IIf(IsNull(rs("OmraValue").value), 0, Trim(rs("OmraValue").value))
    TxtOmraReValue.Text = IIf(IsNull(rs("OmraReValue").value), 0, Trim(rs("OmraReValue").value))
    TxtOmraVAT.Text = IIf(IsNull(rs("OmraVAT").value), 0, Trim(rs("OmraVAT").value))
    TxtHajjValue.Text = IIf(IsNull(rs("HajjValue").value), 0, Trim(rs("HajjValue").value))
    TxtHajjReValue.Text = IIf(IsNull(rs("HajjReValue").value), 0, Trim(rs("HajjReValue").value))
    TxtHajjVAT.Text = IIf(IsNull(rs("HajjVAT").value), 0, Trim(rs("HajjVAT").value))
    TxtMinisterValue.Text = IIf(IsNull(rs("MinisterValue").value), 0, Trim(rs("MinisterValue").value))
    TxtMinisterReValue.Text = IIf(IsNull(rs("MinisterReValue").value), 0, Trim(rs("MinisterReValue").value))
    TxtMinisterVAT.Text = IIf(IsNull(rs("MinisterVAT").value), 0, Trim(rs("MinisterVAT").value))
    TxtMaintCarValue.Text = IIf(IsNull(rs("MaintCarValue").value), 0, Trim(rs("MaintCarValue").value))
    TxtMaintCarReValue.Text = IIf(IsNull(rs("MaintCarReValue").value), 0, Trim(rs("MaintCarReValue").value))
    TxtMaintCarVAT.Text = IIf(IsNull(rs("MaintCarVAT").value), 0, Trim(rs("MaintCarVAT").value))
    TxtTotalPayValue.Text = IIf(IsNull(rs("TotalPayValue").value), 0, Trim(rs("TotalPayValue").value))
    TxtTotalRePayValue.Text = IIf(IsNull(rs("TotalRePayValue").value), 0, Trim(rs("TotalRePayValue").value))
    TxtTotalPayVAT.Text = IIf(IsNull(rs("TotalPayVAT").value), 0, Trim(rs("TotalPayVAT").value))
    TxtProjConValue.Text = IIf(IsNull(rs("ProjConValue").value), 0, Trim(rs("ProjConValue").value))
    TxtProjConReValue.Text = IIf(IsNull(rs("ProjConReValue").value), 0, Trim(rs("ProjConReValue").value))
    TxtProjConVAT.Text = IIf(IsNull(rs("ProjConVAT").value), 0, Trim(rs("ProjConVAT").value))
    TxtReqConValue.Text = IIf(IsNull(rs("ReqConValue").value), 0, Trim(rs("ReqConValue").value))
    TxtReqConReValue.Text = IIf(IsNull(rs("ReqConReValue").value), 0, Trim(rs("ReqConReValue").value))
    TxtReqConVAT.Text = IIf(IsNull(rs("ReqConVAT").value), 0, Trim(rs("ReqConVAT").value))
    TxtAssestValue.Text = IIf(IsNull(rs("AssestValue").value), 0, Trim(rs("AssestValue").value))
    TxtAssestReValue.Text = IIf(IsNull(rs("AssestReValue").value), 0, Trim(rs("AssestReValue").value))
    TxtAssestVAT.Text = IIf(IsNull(rs("AssestVAT").value), 0, Trim(rs("AssestVAT").value))
    TxtTotalReceValue.Text = IIf(IsNull(rs("TotalReceValue").value), 0, Trim(rs("TotalReceValue").value))
    TxtTotalReceReValue.Text = IIf(IsNull(rs("TotalReceReValue").value), 0, Trim(rs("TotalReceReValue").value))
    TxtTotalReceVAT.Text = IIf(IsNull(rs("TotalReceVAT").value), 0, Trim(rs("TotalReceVAT").value))
    TotalNetTxt.Text = IIf(IsNull(rs("TotalNetTxt").value), 0, Trim(rs("TotalNetTxt").value))
   
   
   TxtCorrect1.Text = IIf(IsNull(rs("TxtCorrect1").value), 0, Trim(rs("TxtCorrect1").value))
   TxtOldVat.Text = IIf(IsNull(rs("TxtOldVat").value), 0, Trim(rs("TxtOldVat").value))
   TxtNetVat.Text = IIf(IsNull(rs("TxtNetVat").value), 0, Trim(rs("TxtNetVat").value))
   PurcahseRemarks1.Text = IIf(IsNull(rs("PurcahseRemarks1").value), 0, Trim(rs("PurcahseRemarks1").value))
   
   
    TxtBillVstReverse.Text = IIf(IsNull(rs("TxtBillVstReverse").value), 0, Trim(rs("TxtBillVstReverse").value))
    TxtBillVstReverseREt.Text = IIf(IsNull(rs("TxtBillVstReverseREt").value), 0, Trim(rs("TxtBillVstReverseREt").value))
     
     
   PurcahseRemarks1Ret.Text = IIf(IsNull(rs("PurcahseRemarks1Ret").value), 0, Trim(rs("PurcahseRemarks1Ret").value))
   TotalReturnPurchaseZero.Text = IIf(IsNull(rs("TotalReturnPurchaseZero").value), 0, Trim(rs("TotalReturnPurchaseZero").value))
   PurcahseRemarks2.Text = IIf(IsNull(rs("PurcahseRemarks2").value), 0, Trim(rs("PurcahseRemarks2").value))
   PurcahseRemarks2Ret.Text = IIf(IsNull(rs("PurcahseRemarks2Ret").value), 0, Trim(rs("PurcahseRemarks2Ret").value))
   TxtRemarks1All.Text = IIf(IsNull(rs("TxtRemarks1All").value), 0, Trim(rs("TxtRemarks1All").value))
   TxtRemarks1All2.Text = IIf(IsNull(rs("TxtRemarks1All2").value), 0, Trim(rs("TxtRemarks1All2").value))
   
   Dim j As Integer
       For j = 0 To 6
       txtExpenses(j).Text = IIf(IsNull(rs("txtExpenses" & j).value), 0, Trim(rs("txtExpenses" & j).value))
       txtExpensesVat(j).Text = IIf(IsNull(rs("txtExpensesVat" & j).value), 0, Trim(rs("txtExpensesVat" & j).value))
    
        
        Next j
        
TxtMaintCarValue1.Text = IIf(IsNull(rs("TxtMaintCarValue1").value), 0, Trim(rs("TxtMaintCarValue1").value))
TxtMaintCarValue2.Text = IIf(IsNull(rs("TxtMaintCarValue2").value), 0, Trim(rs("TxtMaintCarValue2").value))
TxtMaintCarReValue1.Text = IIf(IsNull(rs("TxtMaintCarReValue1").value), 0, Trim(rs("TxtMaintCarReValue1").value))
TxtMaintCarReValue2.Text = IIf(IsNull(rs("TxtMaintCarReValue2").value), 0, Trim(rs("TxtMaintCarReValue2").value))



        
   txtTotalExpenses.Text = IIf(IsNull(rs("txtTotalExpenses").value), 0, Trim(rs("txtTotalExpenses").value))
   TxtTotalVatExpenses.Text = IIf(IsNull(rs("TxtTotalVatExpenses").value), 0, Trim(rs("TxtTotalVatExpenses").value))
   
   
   manulaSAlesZero.Text = IIf(IsNull(rs("manulaSAlesZero").value), 0, Trim(rs("manulaSAlesZero").value))
   manulaSAlesZeroRet.Text = IIf(IsNull(rs("manulaSAlesZeroRet").value), 0, Trim(rs("manulaSAlesZeroRet").value))
   SalesZero.Text = IIf(IsNull(rs("SalesZero").value), 0, Trim(rs("SalesZero").value))
   SalesRetZero.Text = IIf(IsNull(rs("SalesRetZero").value), 0, Trim(rs("SalesRetZero").value))
   TxtMinisterValuez.Text = IIf(IsNull(rs("TxtMinisterValuez").value), 0, Trim(rs("TxtMinisterValuez").value))
   
   
   TotalSalesZero.Text = IIf(IsNull(rs("TotalSalesZero").value), 0, Trim(rs("TotalSalesZero").value))
   TotalRetSalesZero.Text = IIf(IsNull(rs("TotalRetSalesZero").value), 0, Trim(rs("TotalRetSalesZero").value))
   txtmanulPurcahsezero.Text = IIf(IsNull(rs("txtmanulPurcahsezero").value), 0, Trim(rs("txtmanulPurcahsezero").value))
   txtmanulPurcahsezeroRetur.Text = IIf(IsNull(rs("txtmanulPurcahsezeroRetur").value), 0, Trim(rs("txtmanulPurcahsezeroRetur").value))
   TxtPurchaseZero.Text = IIf(IsNull(rs("TxtPurchaseZero").value), 0, Trim(rs("TxtPurchaseZero").value))
   TxtPurchaseZeroRet.Text = IIf(IsNull(rs("TxtPurchaseZeroRet").value), 0, Trim(rs("TxtPurchaseZeroRet").value))
   Txtprojectsupp.Text = IIf(IsNull(rs("Txtprojectsupp").value), 0, Trim(rs("Txtprojectsupp").value))
    
    
    TxtprojectsuppRet.Text = IIf(IsNull(rs("TxtprojectsuppRet").value), 0, Trim(rs("TxtprojectsuppRet").value))
   TotalPurchaseZero.Text = IIf(IsNull(rs("TotalPurchaseZero").value), 0, Trim(rs("TotalPurchaseZero").value))
   manulaEntey5.Text = IIf(IsNull(rs("manulaEntey5").value), 0, Trim(rs("manulaEntey5").value))
   manulaEnteyRet5.Text = IIf(IsNull(rs("manulaEnteyRet5").value), 0, Trim(rs("manulaEnteyRet5").value))
   manulaEntey5Vat.Text = IIf(IsNull(rs("manulaEntey5Vat").value), 0, Trim(rs("manulaEntey5Vat").value))
   SalesT5.Text = IIf(IsNull(rs("SalesT5").value), 0, Trim(rs("SalesT5").value))
   SalesRet5.Text = IIf(IsNull(rs("SalesRet5").value), 0, Trim(rs("SalesRet5").value))
   SalesTVAT.Text = IIf(IsNull(rs("SalesTVAT").value), 0, Trim(rs("SalesTVAT").value))
   TxtContractVaue.Text = IIf(IsNull(rs("TxtContractVaue").value), 0, Trim(rs("TxtContractVaue").value))
   
   txtFaBuy.Text = IIf(IsNull(rs("FaBuy").value), 0, Trim(rs("FaBuy").value))
   txtFaBuy2.Text = IIf(IsNull(rs("FaBuy2").value), 0, Trim(rs("FaBuy2").value))
   txtFaBuy3.Text = IIf(IsNull(rs("FaBuy3").value), 0, Trim(rs("FaBuy3").value))
   
   TxtContractReVaue.Text = IIf(IsNull(rs("TxtContractReVaue").value), 0, Trim(rs("TxtContractReVaue").value))
   TxtContractVAT.Text = IIf(IsNull(rs("TxtContractVAT").value), 0, Trim(rs("TxtContractVAT").value))
    
 TxtProjCusValue.Text = IIf(IsNull(rs("TxtProjCusValue").value), 0, Trim(rs("TxtProjCusValue").value))
   TxtProjCusReValue.Text = IIf(IsNull(rs("TxtProjCusReValue").value), 0, Trim(rs("TxtProjCusReValue").value))
   TxtProjCusVAT.Text = IIf(IsNull(rs("TxtProjCusVAT").value), 0, Trim(rs("TxtProjCusVAT").value))
     
     
  TxtOmraValue.Text = IIf(IsNull(rs("TxtOmraValue").value), 0, Trim(rs("TxtOmraValue").value))
   TxtOmraReValue.Text = IIf(IsNull(rs("TxtOmraReValue").value), 0, Trim(rs("TxtOmraReValue").value))
   TxtOmraVAT.Text = IIf(IsNull(rs("TxtOmraVAT").value), 0, Trim(rs("TxtOmraVAT").value))
     
     
  TxtHajjValue.Text = IIf(IsNull(rs("TxtHajjValue").value), 0, Trim(rs("TxtHajjValue").value))
   TxtHajjReValue.Text = IIf(IsNull(rs("TxtHajjReValue").value), 0, Trim(rs("TxtHajjReValue").value))
   TxtHajjVAT.Text = IIf(IsNull(rs("TxtHajjVAT").value), 0, Trim(rs("TxtHajjVAT").value))
     
     
  TxtMinisterValue.Text = IIf(IsNull(rs("TxtMinisterValue").value), 0, Trim(rs("TxtMinisterValue").value))
   TxtMinisterReValue.Text = IIf(IsNull(rs("TxtMinisterReValue").value), 0, Trim(rs("TxtMinisterReValue").value))
   TxtMinisterVAT.Text = IIf(IsNull(rs("TxtMinisterVAT").value), 0, Trim(rs("TxtMinisterVAT").value))
   
     TxtMaintCarValue.Text = IIf(IsNull(rs("TxtMaintCarValue").value), 0, Trim(rs("TxtMaintCarValue").value))
   TxtMaintCarReValue.Text = IIf(IsNull(rs("TxtMaintCarReValue").value), 0, Trim(rs("TxtMaintCarReValue").value))
   TxtMaintCarVAT.Text = IIf(IsNull(rs("TxtMaintCarVAT").value), 0, Trim(rs("TxtMaintCarVAT").value))
   
   TxtServiceInvoice5.Text = IIf(IsNull(rs("TxtServiceInvoice5").value), 0, Trim(rs("TxtServiceInvoice5").value))
   TxtServiceInvoice5REt.Text = IIf(IsNull(rs("TxtServiceInvoice5REt").value), 0, Trim(rs("TxtServiceInvoice5REt").value))
   TxtServiceInvoice5Vat.Text = IIf(IsNull(rs("TxtServiceInvoice5Vat").value), 0, Trim(rs("TxtServiceInvoice5Vat").value))
   
   
   
   TxtTotalPayValue.Text = IIf(IsNull(rs("TxtTotalPayValue").value), 0, Trim(rs("TxtTotalPayValue").value))
   TxtTotalRePayValue.Text = IIf(IsNull(rs("TxtTotalRePayValue").value), 0, Trim(rs("TxtTotalRePayValue").value))
   TxtTotalPayVAT.Text = IIf(IsNull(rs("TxtTotalPayVAT").value), 0, Trim(rs("TxtTotalPayVAT").value))
    
    
    
    txtManulaEntryP5.Text = IIf(IsNull(rs("txtManulaEntryP5").value), 0, Trim(rs("txtManulaEntryP5").value))
   txtManulaEntryP5Ret.Text = IIf(IsNull(rs("txtManulaEntryP5Ret").value), 0, Trim(rs("txtManulaEntryP5Ret").value))
   txtManulaEntryP5Vat.Text = IIf(IsNull(rs("txtManulaEntryP5Vat").value), 0, Trim(rs("txtManulaEntryP5Vat").value))
    
    
    
    PurchasesT5.Text = IIf(IsNull(rs("PurchasesT5").value), 0, Trim(rs("PurchasesT5").value))
   PurchasesRett5.Text = IIf(IsNull(rs("PurchasesRett5").value), 0, Trim(rs("PurchasesRett5").value))
   Purchasest5vat.Text = IIf(IsNull(rs("Purchasest5vat").value), 0, Trim(rs("Purchasest5vat").value))
    
    
    
    TxtProjConValue.Text = IIf(IsNull(rs("TxtProjConValue").value), 0, Trim(rs("TxtProjConValue").value))
   TxtProjConReValue.Text = IIf(IsNull(rs("TxtProjConReValue").value), 0, Trim(rs("TxtProjConReValue").value))
   TxtProjConVAT.Text = IIf(IsNull(rs("TxtProjConVAT").value), 0, Trim(rs("TxtProjConVAT").value))
    
   
    TxtReqConValue.Text = IIf(IsNull(rs("TxtReqConValue").value), 0, Trim(rs("TxtReqConValue").value))
   TxtReqConReValue.Text = IIf(IsNull(rs("TxtReqConReValue").value), 0, Trim(rs("TxtReqConReValue").value))
   TxtReqConVAT.Text = IIf(IsNull(rs("TxtReqConVAT").value), 0, Trim(rs("TxtReqConVAT").value))
       
        TxtAssestValue.Text = IIf(IsNull(rs("TxtAssestValue").value), 0, Trim(rs("TxtAssestValue").value))
   TxtAssestReValue.Text = IIf(IsNull(rs("TxtAssestReValue").value), 0, Trim(rs("TxtAssestReValue").value))
   TxtAssestVAT.Text = IIf(IsNull(rs("TxtAssestVAT").value), 0, Trim(rs("TxtAssestVAT").value))
        
        Expenses.Text = IIf(IsNull(rs("Expenses").value), 0, Trim(rs("Expenses").value))
   Expensesvat.Text = IIf(IsNull(rs("Expensesvat").value), 0, Trim(rs("Expensesvat").value))
   TxtTotalReceValue.Text = IIf(IsNull(rs("TxtTotalReceValue").value), 0, Trim(rs("TxtTotalReceValue").value))
   
   TxtTotalReceVAT.Text = IIf(IsNull(rs("TxtTotalReceVAT").value), 0, Trim(rs("TxtTotalReceVAT").value))
  
  transport5.Text = IIf(IsNull(rs("transport5").value), 0, Trim(rs("transport5").value))
  transport5re.Text = IIf(IsNull(rs("transport5re").value), 0, Trim(rs("transport5re").value))
  transport5vat.Text = IIf(IsNull(rs("transport5vat").value), 0, Trim(rs("transport5vat").value))
  
  TxtPReVatTotal5.Text = IIf(IsNull(rs("TxtPReVatTotal5").value), 0, Trim(rs("TxtPReVatTotal5").value))
  TxtPReVatVAT5.Text = IIf(IsNull(rs("TxtPReVatVAT5").value), 0, Trim(rs("TxtPReVatVAT5").value))
  
       
       
  TxtPReVatTotal5V.Text = IIf(IsNull(rs("TxtPReVatTotal5v").value), 0, Trim(rs("TxtPReVatTotal5v").value))
  TxtPReVatVAT5V.Text = IIf(IsNull(rs("TxtPReVatVAT5v").value), 0, Trim(rs("TxtPReVatVAT5v").value))
  
         TxtProjCusValuezero.Text = IIf(IsNull(rs("TxtProjCusValuezero").value), 0, Trim(rs("TxtProjCusValuezero").value))
  TxtProjCusValueRetzero.Text = IIf(IsNull(rs("TxtProjCusValueRetzero").value), 0, Trim(rs("TxtProjCusValueRetzero").value))
  
  
  TxtVatDis.Text = IIf(IsNull(rs("TxtVatDis").value), 0, Trim(rs("TxtVatDis").value))
  TxtVatADD.Text = IIf(IsNull(rs("TxtVatADD").value), 0, Trim(rs("TxtVatADD").value))
       
              
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub



Private Sub TxtPurchaseZero_Click()
Dim X As String
X = getTotalVATItemsValue(22, False, 0, , True, " Õ·Ì·Ì «·„‘ —Ì«  «·’›—Ì…")
End Sub

Private Sub TxtPurchaseZeroRet_Click()
Dim X As String
X = getTotalVATItemsValue(5, False, 0, , True, " ⁄œÌ·«  «·„‘ —Ì«  «·’›—Ì…")
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
Private Sub SaveData()

    Dim Msg As String
    Dim BeginTrans As Boolean
    
   ' On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
        
        Select Case Me.TxtModFlg.Text
           Case "E"
           Cn.Execute "Delete from DOUBLE_ENTREY_VOUCHERS where Notes_ID=" & val(TxtNoteID.Text) & ""
           Cn.Execute "Delete from Notes where NoteID=" & val(TxtNoteID.Text) & ""
           Case "N"
                rs.AddNew
                ID.Text = CStr(new_id("TblVATAvowal", "ID", "", True))
        End Select

        rs("ID").value = val(ID.Text)
        rs("RecoredDate").value = RecoredDate.value
        rs("Notes").value = IIf(NotesTxt.Text = "", Null, NotesTxt.Text)
        rs("SalesVAT").value = val(TxtSalesVAT.Text)
        rs("tztAdvBill").value = val(tztAdvBill.Text)
        
        
        rs("RetSalesVAT").value = val(TxtRetSalesVAT.Text)
        rs("BuyVAT").value = val(TxtBuyVAT.Text)
        rs("RetBuyVAT").value = val(TxtRetBuyVAT.Text)
        
        rs("SaleWithOutVat").value = val(Text1(0).Text)
        rs("RetSaleWithOutVat").value = val(Text2.Text)
        
        
        rs("DateFrom").value = DateFrom.value
        rs("DateTo").value = DateTo.value
          If paidchk(0).value = xtpChecked Then
            rs("Paid").value = True
        Else
            rs("Paid").value = False
        End If
        
        
          If ChkIsFree.value = xtpChecked Then
            rs("ChkIsFree").value = True
        Else
            rs("ChkIsFree").value = False
        End If
                
        
        
        
         If paidchk(1).value = xtpChecked Then
            rs("HideakarExitVat1").value = True
        Else
            rs("HideakarExitVat1").value = False
        End If
        
        
        If paidchk(2).value = xtpChecked Then
            rs("HideakarExitVat2").value = True
        Else
            rs("HideakarExitVat2").value = False
        End If
              
              
              
        'rs("BranchID").value = IIf(BranchDC.BoundText = "", Null, BranchDC.BoundText)
        
        rs("Sales1").value = IIf(Sales1.Text = "", Null, Sales1.Text)
        rs("Sales2").value = IIf(Sales2.Text = "", Null, Sales2.Text)
        rs("Sales3").value = IIf(Sales3.Text = "", Null, Sales3.Text)
        rs("Sales4").value = IIf(Sales4.Text = "", Null, Sales4.Text)
        rs("Sales5").value = IIf(Sales5.Text = "", Null, Sales5.Text)
        rs("SalesTotal").value = IIf(SalesTotal.Text = "", Null, SalesTotal.Text)
        
        rs("RSales1").value = IIf(RSales1.Text = "", Null, RSales1.Text)
        rs("RSales2").value = IIf(RSales2.Text = "", Null, RSales2.Text)
        rs("RSales3").value = IIf(RSales3.Text = "", Null, RSales3.Text)
        rs("RSales4").value = IIf(RSales4.Text = "", Null, RSales4.Text)
        rs("RSales5").value = IIf(RSales5.Text = "", Null, RSales5.Text)
        rs("RSalesTotal").value = IIf(RSalesTotal.Text = "", Null, RSalesTotal.Text)
        
        rs("VATSales1").value = IIf(VATSales1.Text = "", Null, VATSales1.Text)
        'rs("VATSales2").value = IIf(VATSales2.Text = "", Null, VATSales2.Text)
        'rs("VATSales3").value = IIf(VATSales3.Text = "", Null, VATSales3.Text)
        'rs("VATSales4").value = IIf(VATSales4.Text = "", Null, VATSales4.Text)
        'rs("VATSales5").value = IIf(VATSales5.Text = "", Null, VATSales5.Text)
        rs("VATSalesTotal").value = IIf(VATSalesTotal.Text = "", Null, VATSalesTotal.Text)
        
        
        rs("Purchases1").value = IIf(Purchases1.Text = "", Null, Purchases1.Text)
        rs("Purchases2").value = IIf(Purchases2.Text = "", Null, Purchases2.Text)
        rs("Purchases3").value = IIf(Purchases3.Text = "", Null, Purchases3.Text)
        rs("Purchases4").value = IIf(Purchases4.Text = "", Null, Purchases4.Text)
        rs("Purchases5").value = IIf(Purchases5.Text = "", Null, Purchases5.Text)
        rs("PurchasesTotal").value = IIf(PurchasesTotal.Text = "", Null, PurchasesTotal.Text)
        
        rs("RPurchases1").value = IIf(RPurchases1.Text = "", Null, RPurchases1.Text)
        rs("RPurchases2").value = IIf(RPurchases2.Text = "", Null, RPurchases2.Text)
        rs("RPurchases3").value = IIf(RPurchases3.Text = "", Null, RPurchases3.Text)
        rs("RPurchases4").value = IIf(RPurchases4.Text = "", Null, RPurchases4.Text)
        rs("RPurchases5").value = IIf(RPurchases5.Text = "", Null, RPurchases5.Text)
        rs("RPurchasesTotal").value = IIf(RPurchasesTotal.Text = "", Null, RPurchasesTotal.Text)
        rs("ValueDept").value = val(TxtDept(0).Text)
        rs("ValueDept2").value = val(TxtDept(1).Text)
        rs("ValueCredit").value = val(TxtDept(2).Text)
        rs("ValueCredit2").value = val(TxtDept(3).Text)
        
        
        
        
         
rs("ValueNew4").value = val(TxtDept(4).Text)
rs("ValueNew5").value = val(TxtDept(5).Text)
rs("ValueNew6").value = val(TxtDept(6).Text)
rs("ValueNew7").value = val(TxtDept(7).Text)
       
        
        
        
        
        
        
        rs("VATPurchases1").value = IIf(VATPurchases1.Text = "", Null, VATPurchases1.Text)
        rs("VATPurchases2").value = IIf(VATPurchases2.Text = "", Null, VATPurchases2.Text)
        rs("VATPurchases3").value = IIf(VATPurchases3.Text = "", Null, VATPurchases3.Text)
        'rs("VATPurchases4").value = IIf(VATPurchases4.Text = "", Null, VATPurchases4.Text)
        'rs("VATPurchases5").value = IIf(VATPurchases5.Text = "", Null, VATPurchases5.Text)
        rs("VATPurchasesTotal").value = IIf(VATPurchasesTotal.Text = "", Null, VATPurchasesTotal.Text)
        rs("BranchID").value = IIf(DcBranch.BoundText = "", Null, val(DcBranch.BoundText))
        rs("UserID").value = IIf(DCboUserName.BoundText = "", Null, val(DCboUserName.BoundText))
        ''/////NewData
        rs("ContractVaue").value = IIf(TxtContractVaue.Text = "", 0, val(TxtContractVaue.Text))
        
        rs("FaBuy").value = IIf(txtFaBuy.Text = "", 0, val(txtFaBuy.Text))
        rs("FaBuy2").value = IIf(txtFaBuy2.Text = "", 0, val(txtFaBuy2.Text))
        rs("FaBuy3").value = IIf(txtFaBuy3.Text = "", 0, val(txtFaBuy3.Text))
        
       
        
        rs("ContractVaueHousing").value = IIf(Text1(1).Text = "", 0, val(Text1(1).Text))
        
        rs("ContractReVaue").value = IIf(TxtContractReVaue.Text = "", 0, val(TxtContractReVaue.Text))
        rs("ContractVAT").value = IIf(TxtContractVAT.Text = "", 0, val(TxtContractVAT.Text))
        rs("ProjCusValue").value = IIf(TxtProjCusValue.Text = "", 0, val(TxtProjCusValue.Text))
        rs("ProjCusReValue").value = IIf(TxtProjCusReValue.Text = "", 0, val(TxtProjCusReValue.Text))
        rs("ProjCusVAT").value = IIf(TxtProjCusVAT.Text = "", 0, val(TxtProjCusVAT.Text))
        rs("OmraValue").value = IIf(TxtOmraValue.Text = "", 0, val(TxtOmraValue.Text))
        rs("OmraReValue").value = IIf(TxtOmraReValue.Text = "", 0, val(TxtOmraReValue.Text))
        rs("OmraVAT").value = IIf(TxtOmraVAT.Text = "", 0, val(TxtOmraVAT.Text))
        rs("HajjValue").value = IIf(TxtHajjValue.Text = "", 0, val(TxtHajjValue.Text))
        rs("HajjReValue").value = IIf(TxtHajjReValue.Text = "", 0, val(TxtHajjReValue.Text))
        rs("HajjVAT").value = IIf(TxtHajjVAT.Text = "", 0, val(TxtHajjVAT.Text))
        rs("MinisterValue").value = IIf(TxtMinisterValue.Text = "", 0, val(TxtMinisterValue.Text))
        rs("MinisterReValue").value = IIf(TxtMinisterReValue.Text = "", 0, val(TxtMinisterReValue.Text))
        rs("MinisterVAT").value = IIf(TxtMinisterVAT.Text = "", 0, val(TxtMinisterVAT.Text))
        rs("MaintCarValue").value = IIf(TxtMaintCarValue.Text = "", 0, val(TxtMaintCarValue.Text))
        rs("MaintCarReValue").value = IIf(TxtMaintCarReValue.Text = "", 0, val(TxtMaintCarReValue.Text))
        rs("MaintCarVAT").value = IIf(TxtMaintCarVAT.Text = "", 0, val(TxtMaintCarVAT.Text))
        rs("TotalPayValue").value = IIf(TxtTotalPayValue.Text = "", 0, val(TxtTotalPayValue.Text))
        rs("TotalRePayValue").value = IIf(TxtTotalRePayValue.Text = "", 0, val(TxtTotalRePayValue.Text))
        rs("TotalPayVAT").value = IIf(TxtTotalPayVAT.Text = "", 0, val(TxtTotalPayVAT.Text))
        rs("ProjConValue").value = IIf(TxtProjConValue.Text = "", 0, val(TxtProjConValue.Text))
        rs("ProjConReValue").value = IIf(TxtProjConReValue.Text = "", 0, val(TxtProjConReValue.Text))
        rs("ProjConVAT").value = IIf(TxtProjConVAT.Text = "", 0, val(TxtProjConVAT.Text))
        rs("ReqConValue").value = IIf(TxtReqConValue.Text = "", 0, val(TxtReqConValue.Text))
        rs("ReqConReValue").value = IIf(TxtReqConReValue.Text = "", 0, val(TxtReqConReValue.Text))
        rs("ReqConVAT").value = IIf(TxtReqConVAT.Text = "", 0, val(TxtReqConVAT.Text))
        rs("AssestValue").value = IIf(TxtAssestValue.Text = "", 0, val(TxtAssestValue.Text))
        rs("AssestReValue").value = IIf(TxtAssestReValue.Text = "", 0, val(TxtAssestReValue.Text))
        rs("AssestVAT").value = IIf(TxtAssestVAT.Text = "", 0, val(TxtAssestVAT.Text))
        rs("TotalReceValue").value = IIf(TxtTotalReceValue.Text = "", 0, val(TxtTotalReceValue.Text))
        rs("TotalReceReValue").value = IIf(TxtTotalReceReValue.Text = "", 0, val(TxtTotalReceReValue.Text))
        rs("TotalReceVAT").value = IIf(TxtTotalReceVAT.Text = "", 0, val(TxtTotalReceVAT.Text))
        rs("TotalNetTxt").value = IIf(TotalNetTxt.Text = "", 0, val(TotalNetTxt.Text))
        
        rs("TxtCorrect1").value = IIf(TxtCorrect1.Text = "", 0, val(TxtCorrect1.Text))
        rs("TxtOldVat").value = IIf(TxtOldVat.Text = "", 0, val(TxtOldVat.Text))
        rs("TxtNetVat").value = IIf(TxtNetVat.Text = "", 0, val(TxtNetVat.Text))
        
        
        rs("PurcahseRemarks1").value = IIf(PurcahseRemarks1.Text = "", 0, val(PurcahseRemarks1.Text))
        rs("PurcahseRemarks1Ret").value = IIf(PurcahseRemarks1Ret.Text = "", 0, val(PurcahseRemarks1Ret.Text))
        
        
        rs("TxtBillVstReverse").value = IIf(TxtBillVstReverse.Text = "", 0, val(TxtBillVstReverse.Text))
        rs("TxtBillVstReverseREt").value = IIf(TxtBillVstReverseREt.Text = "", 0, val(TxtBillVstReverseREt.Text))
        
        
        
        rs("PurcahseRemarks2").value = IIf(PurcahseRemarks2.Text = "", 0, val(PurcahseRemarks2.Text))
        rs("PurcahseRemarks2Ret").value = IIf(PurcahseRemarks2Ret.Text = "", 0, val(PurcahseRemarks2Ret.Text))
        
        rs("TxtRemarks1All").value = IIf(TxtRemarks1All.Text = "", 0, val(TxtRemarks1All.Text))
        
        rs("TxtRemarks1All2").value = IIf(TxtRemarks1All2.Text = "", 0, val(TxtRemarks1All2.Text))
        
       Dim j As Integer
       For j = 0 To 6
        rs("txtExpenses" & j).value = IIf(txtExpenses(j).Text = "", 0, val(txtExpenses(j).Text))
        rs("txtExpensesVat" & j).value = IIf(txtExpensesVat(j).Text = "", 0, val(txtExpensesVat(j).Text))
        
        
        Next j
        

        rs("TxtMaintCarValue1").value = IIf(TxtMaintCarValue1.Text = "", 0, val(TxtMaintCarValue1.Text))
        rs("TxtMaintCarValue2").value = IIf(TxtMaintCarValue2.Text = "", 0, val(TxtMaintCarValue2.Text))
        rs("TxtMaintCarReValue1").value = IIf(TxtMaintCarReValue1.Text = "", 0, val(TxtMaintCarReValue1.Text))
        rs("TxtMaintCarReValue2").value = IIf(TxtMaintCarReValue2.Text = "", 0, val(TxtMaintCarReValue2.Text))
        
        rs("txtTotalExpenses").value = IIf(txtTotalExpenses.Text = "", 0, val(txtTotalExpenses.Text))
        rs("TxtTotalVatExpenses").value = IIf(TxtTotalVatExpenses.Text = "", 0, val(TxtTotalVatExpenses.Text))
        
        rs("manulaSAlesZero").value = IIf(manulaSAlesZero.Text = "", 0, val(manulaSAlesZero.Text))
        rs("manulaSAlesZeroRet").value = IIf(manulaSAlesZeroRet.Text = "", 0, val(manulaSAlesZeroRet.Text))
        
            rs("SalesZero").value = IIf(SalesZero.Text = "", 0, val(SalesZero.Text))
                rs("SalesRetZero").value = IIf(SalesRetZero.Text = "", 0, val(SalesRetZero.Text))
                    rs("TxtMinisterValuez").value = IIf(TxtMinisterValuez.Text = "", 0, val(TxtMinisterValuez.Text))
                    
                        rs("TotalSalesZero").value = IIf(TotalSalesZero.Text = "", 0, val(TotalSalesZero.Text))
                            rs("TotalRetSalesZero").value = IIf(TotalRetSalesZero.Text = "", 0, val(TotalRetSalesZero.Text))
                            
                                rs("txtmanulPurcahsezero").value = IIf(txtmanulPurcahsezero.Text = "", 0, val(txtmanulPurcahsezero.Text))
                                
                                    rs("txtmanulPurcahsezeroRetur").value = IIf(txtmanulPurcahsezeroRetur.Text = "", 0, val(txtmanulPurcahsezeroRetur.Text))
                                    
                                        rs("TxtPurchaseZero").value = IIf(TxtPurchaseZero.Text = "", 0, val(TxtPurchaseZero.Text))
                                            rs("TxtPurchaseZeroRet").value = IIf(TxtPurchaseZeroRet.Text = "", 0, val(TxtPurchaseZeroRet.Text))
                                                rs("Txtprojectsupp").value = IIf(Txtprojectsupp.Text = "", 0, val(Txtprojectsupp.Text))
                                                
     
        
rs("TxtprojectsuppRet").value = IIf(TxtprojectsuppRet.Text = "", 0, val(TxtprojectsuppRet.Text))
rs("TotalPurchaseZero").value = IIf(TotalPurchaseZero.Text = "", 0, val(TotalPurchaseZero.Text))
rs("TotalReturnPurchaseZero").value = IIf(TotalReturnPurchaseZero.Text = "", 0, val(TotalReturnPurchaseZero.Text))
 
 
 rs("manulaEntey5").value = IIf(manulaEntey5.Text = "", 0, val(manulaEntey5.Text))
 rs("manulaEnteyRet5").value = IIf(manulaEnteyRet5.Text = "", 0, val(manulaEnteyRet5.Text))
 rs("manulaEntey5Vat").value = IIf(manulaEntey5Vat.Text = "", 0, val(manulaEntey5Vat.Text))
 
 
 rs("SalesT5").value = IIf(SalesT5.Text = "", 0, val(SalesT5.Text))
 rs("SalesRet5").value = IIf(SalesRet5.Text = "", 0, val(SalesRet5.Text))
 rs("SalesTVAT").value = IIf(SalesTVAT.Text = "", 0, val(SalesTVAT.Text))
 rs("TxtContractVaue").value = IIf(TxtContractVaue.Text = "", 0, val(TxtContractVaue.Text))
 rs("TxtContractReVaue").value = IIf(TxtContractReVaue.Text = "", 0, val(TxtContractReVaue.Text))
 rs("TxtContractVAT").value = IIf(TxtContractVAT.Text = "", 0, val(TxtContractVAT.Text))
  
  rs("TxtProjCusValue").value = IIf(TxtProjCusValue.Text = "", 0, val(TxtProjCusValue.Text))
 rs("TxtProjCusReValue").value = IIf(TxtProjCusReValue.Text = "", 0, val(TxtProjCusReValue.Text))
  rs("TxtProjCusVAT").value = IIf(TxtProjCusVAT.Text = "", 0, val(TxtProjCusVAT.Text))
  rs("TxtOmraValue").value = IIf(TxtOmraValue.Text = "", 0, val(TxtOmraValue.Text))
  rs("TxtOmraReValue").value = IIf(TxtOmraReValue.Text = "", 0, val(TxtOmraReValue.Text))
  
  rs("TxtOmraVAT").value = IIf(TxtOmraVAT.Text = "", 0, val(TxtOmraVAT.Text))
  
  rs("TxtHajjValue").value = IIf(TxtHajjValue.Text = "", 0, val(TxtHajjValue.Text))
  
  rs("TxtHajjReValue").value = IIf(TxtHajjReValue.Text = "", 0, val(TxtHajjReValue.Text))
  rs("TxtHajjVAT").value = IIf(TxtHajjVAT.Text = "", 0, val(TxtHajjVAT.Text))
  
  rs("TxtMinisterValue").value = IIf(TxtMinisterValue.Text = "", 0, val(TxtMinisterValue.Text))
  rs("TxtMinisterReValue").value = IIf(TxtMinisterReValue.Text = "", 0, val(TxtMinisterReValue.Text))
  rs("TxtMinisterVAT").value = IIf(TxtMinisterVAT.Text = "", 0, val(TxtMinisterVAT.Text))
  
  rs("TxtMaintCarValue").value = IIf(TxtMaintCarValue.Text = "", 0, val(TxtMaintCarValue.Text))
  rs("TxtMaintCarReValue").value = IIf(TxtMaintCarReValue.Text = "", 0, val(TxtMaintCarReValue.Text))
  
  rs("TxtMaintCarVAT").value = IIf(TxtMaintCarVAT.Text = "", 0, val(TxtMaintCarVAT.Text))
  
  rs("TxtServiceInvoice5").value = IIf(TxtServiceInvoice5.Text = "", 0, val(TxtServiceInvoice5.Text))
  rs("TxtServiceInvoice5REt").value = IIf(TxtServiceInvoice5REt.Text = "", 0, val(TxtServiceInvoice5REt.Text))
  rs("TxtServiceInvoice5Vat").value = IIf(TxtServiceInvoice5Vat.Text = "", 0, val(TxtServiceInvoice5Vat.Text))
  
  rs("TxtTotalPayValue").value = IIf(TxtTotalPayValue.Text = "", 0, val(TxtTotalPayValue.Text))
  rs("TxtTotalRePayValue").value = IIf(TxtTotalRePayValue.Text = "", 0, val(TxtTotalRePayValue.Text))
  
  rs("TxtTotalPayVAT").value = IIf(TxtTotalPayVAT.Text = "", 0, val(TxtTotalPayVAT.Text))
  
  
  rs("txtManulaEntryP5").value = IIf(txtManulaEntryP5.Text = "", 0, val(txtManulaEntryP5.Text))
  rs("txtManulaEntryP5Ret").value = IIf(txtManulaEntryP5Ret.Text = "", 0, val(txtManulaEntryP5Ret.Text))
  rs("txtManulaEntryP5Vat").value = IIf(txtManulaEntryP5Vat.Text = "", 0, val(txtManulaEntryP5Vat.Text))
  
  rs("PurchasesT5").value = IIf(PurchasesT5.Text = "", 0, val(PurchasesT5.Text))
  
  rs("PurchasesRett5").value = IIf(PurchasesRett5.Text = "", 0, val(PurchasesRett5.Text))
  rs("Purchasest5vat").value = IIf(Purchasest5vat.Text = "", 0, val(Purchasest5vat.Text))
  
  rs("TxtProjConValue").value = IIf(TxtProjConValue.Text = "", 0, val(TxtProjConValue.Text))
  rs("TxtProjConReValue").value = IIf(TxtProjConReValue.Text = "", 0, val(TxtProjConReValue.Text))
  rs("TxtProjConVAT").value = IIf(TxtProjConVAT.Text = "", 0, val(TxtProjConVAT.Text))
  
    rs("TxtReqConValue").value = IIf(TxtReqConValue.Text = "", 0, val(TxtReqConValue.Text))
  rs("TxtReqConReValue").value = IIf(TxtReqConReValue.Text = "", 0, val(TxtReqConReValue.Text))
  rs("TxtReqConVAT").value = IIf(TxtReqConVAT.Text = "", 0, val(TxtReqConVAT.Text))
  
  rs("TxtAssestValue").value = IIf(TxtAssestValue.Text = "", 0, val(TxtAssestValue.Text))
  rs("TxtAssestReValue").value = IIf(TxtAssestReValue.Text = "", 0, val(TxtAssestReValue.Text))
  
    rs("TxtAssestVAT").value = IIf(TxtAssestVAT.Text = "", 0, val(TxtAssestVAT.Text))
    
    
  rs("Expenses").value = IIf(Expenses.Text = "", 0, val(Expenses.Text))
  rs("Expensesvat").value = IIf(Expensesvat.Text = "", 0, val(Expensesvat.Text))
  rs("TxtTotalReceValue").value = IIf(TxtTotalReceValue.Text = "", 0, val(TxtTotalReceValue.Text))
  rs("TxtTotalReceReValue").value = IIf(TxtTotalReceReValue.Text = "", 0, val(TxtTotalReceReValue.Text))
  
 rs("TxtTotalReceVAT").value = IIf(TxtTotalReceVAT.Text = "", 0, val(TxtTotalReceVAT.Text))
  
  
  rs("transport5").value = IIf(transport5.Text = "", 0, val(transport5.Text))
  rs("transport5re").value = IIf(transport5re.Text = "", 0, val(transport5re.Text))
  rs("transport5vat").value = IIf(transport5vat.Text = "", 0, val(transport5vat.Text))
  
  rs("TxtPReVatTotal5").value = IIf(TxtPReVatTotal5.Text = "", 0, val(TxtPReVatTotal5.Text))
  rs("TxtPReVatVAT5").value = IIf(TxtPReVatVAT5.Text = "", 0, val(TxtPReVatVAT5.Text))
  
  rs("TxtPReVatTotal5v").value = IIf(TxtPReVatTotal5V.Text = "", 0, val(TxtPReVatTotal5V.Text))
  rs("TxtPReVatVAT5v").value = IIf(TxtPReVatVAT5V.Text = "", 0, val(TxtPReVatVAT5V.Text))
  

rs("TxtProjCusValuezero").value = IIf(TxtProjCusValuezero.Text = "", 0, val(TxtProjCusValuezero.Text))
rs("TxtProjCusValueRetzero").value = IIf(TxtProjCusValueRetzero.Text = "", 0, val(TxtProjCusValueRetzero.Text))

rs("TxtVatDis").value = IIf(TxtVatDis.Text = "", 0, val(TxtVatDis.Text))
rs("TxtVatADD").value = IIf(TxtVatADD.Text = "", 0, val(TxtVatADD.Text))


        

              
        
        rs.update

        Dim StrDes As String

        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        If Not createVoucher Then GoTo ErrTrap
        Select Case Me.TxtModFlg.Text
            Case "N"
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ «·»Ì«‰«  " & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Saved" & CHR(13)
                    Msg = Msg + "Do you want enter another One"
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            Case "E"
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
        End Select
        TxtModFlg.Text = "R"
        Retrive
    End If
    Exit Sub
ErrTrap:

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
            Msg = "Data Can't be saved " & CHR(13)
            Msg = Msg & "Invalid data values was entered" & CHR(13)
            Msg = Msg & "Please make sure of the entered data and try again"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & CHR(13)
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
            rs.Find " ID='" & val(ID.Text) & "'", , adSearchForward, adBookmarkFirst
            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.Text = "R"
                Exit Sub
            End If
            Retrive
            Me.TxtModFlg.Text = "R"
    End Select
    Exit Sub
ErrTrap:
End Sub
Private Sub Del_Action()
  
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
 
    'On Error GoTo ErrTrap
            
        If ID.Text <> "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "”Ì „ Õ–› »Ì«‰«  «·”Ã· —ﬁ„ " & CHR(13)
                Msg = Msg + (ID.Text) & CHR(13)
                Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
            Else
                Msg = "Delete Recored File No. ?" & CHR(13)
                Msg = Msg + (ID.Text) & CHR(13)
                Msg = Msg + "  Are you sure you want to delete ?"
            End If
        
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
                If Not rs.RecordCount < 1 Then
                    StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
                    Cn.Execute StrSQL, , adExecuteNoRecords
                      StrSQL = "Delete From Notes Where NoteID=" & val(Me.TxtNoteID.Text)
                    Cn.Execute StrSQL, , adExecuteNoRecords
                    StrSQL = "delete From TblVATAvowal where  ID =" & val(ID.Text)
                    Cn.Execute StrSQL, , adExecuteNoRecords
                 
                    rs.MoveFirst
                    
                    StrSQL = "SELECT  *  From TblVATAvowal "
                    rs.Close
                    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If rs.RecordCount < 1 Then
                        clear_all Me
                        TxtModFlg_Change
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
                Msg = "this process Not Aailable"
            End If
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtModFlg_Change
        Exit Sub
    End If
    TxtModFlg_Change
    Exit Sub
ErrTrap:
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ…  ﬁÌÌ„ «·„ÊŸ›Ì‰ "
        Else
            Msg = "Sorry can't delete data"
        End If
        Msg = Msg & CHR(13) & Err.description
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
End Sub
Function getTotalVATItemsValue(TransType As Integer, withVAT As Boolean, Optional Vatyo As Variant = Null, Optional Remarks As Integer = -1, Optional PrintReport As Boolean = False, Optional reporttitle As String, Optional VstReverse As Integer) As Double
    Dim Sql As String
    Dim RsVAT As ADODB.Recordset
    Set RsVAT = New ADODB.Recordset
    
  '  sql = "SELECT sum(Transaction_Details.showPrice * Transaction_Details.ShowQty * isnull (Transactions.Currency_rate,1)) as sumVAT "
  If TransType = 22 Then
  Sql = "SELECT  sum((Transaction_Details.showPrice * Transaction_Details.ShowQty * ISNULL(dbo.Transactions.Currency_rate, 1)    )-isnull (TotalDiscountPerLine,0 * ISNULL(dbo.Transactions.Currency_rate, 1)  ) *  Transaction_Details.ShowQty  -  CASE WHEN ItemDiscountType = 1 THEN 0 WHEN ItemDiscountType = 2 THEN ((ItemDiscount * ISNULL(dbo.Transactions.Currency_rate, 1)))            WHEN ItemDiscountType = 3 THEN ((Transaction_Details.showqty * Transaction_Details.showPrice) * ((ItemDiscount / 100)) * ISNULL(dbo.Transactions.Currency_rate, 1))            WHEN ItemDiscountType = 4 THEN (((Transaction_Details.showqty * Transaction_Details.showPrice) * ISNULL(dbo.Transactions.Currency_rate, 1))) ELSE 0 END ) AS sumVAT"
  Else
  Sql = "SELECT  sum((Transaction_Details.showPrice * Transaction_Details.ShowQty * ISNULL(dbo.Transactions.Currency_rate, 1)    )-isnull (TotalDiscountPerLine,0) * ISNULL(dbo.Transactions.Currency_rate, 1)    -  CASE WHEN ItemDiscountType = 1 THEN 0 WHEN ItemDiscountType = 2 THEN ((ItemDiscount * ISNULL(dbo.Transactions.Currency_rate, 1)))            WHEN ItemDiscountType = 3 THEN ((Transaction_Details.showqty * Transaction_Details.showPrice ) * ((ItemDiscount / 100)) * ISNULL(dbo.Transactions.Currency_rate, 1))            WHEN ItemDiscountType = 4 THEN (((Transaction_Details.showqty * Transaction_Details.showPrice) * ISNULL(dbo.Transactions.Currency_rate, 1))) ELSE 0 END ) AS sumVAT"
  End If
  
If PrintReport = True Then 'print

  If TransType = 22 Then
  Sql = "SELECT Transactions.Transaction_Date  ,   dbo.TblItems.ItemCode,  dbo.TblItems.Fullcode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee ,  Transactions.NoteSerial1, ((Transaction_Details.showPrice * Transaction_Details.ShowQty * ISNULL(dbo.Transactions.Currency_rate, 1)  )-isnull (TotalDiscountPerLine * ISNULL(dbo.Transactions.Currency_rate, 1)  ,0) *  Transaction_Details.ShowQty  -  CASE WHEN ItemDiscountType = 1 THEN 0 WHEN ItemDiscountType = 2 THEN ((ItemDiscount * ISNULL(dbo.Transactions.Currency_rate, 1)))            WHEN ItemDiscountType = 3 THEN ((Transaction_Details.showqty * Transaction_Details.showPrice) * ((ItemDiscount / 100)) * ISNULL(dbo.Transactions.Currency_rate, 1))            WHEN ItemDiscountType = 4 THEN (((Transaction_Details.showqty * Transaction_Details.showPrice) * ISNULL(dbo.Transactions.Currency_rate, 1))) ELSE 0 END ) AS sumVAT"
  Else
  Sql = "SELECT Transactions.Transaction_Date ,  dbo.TblItems.ItemCode, dbo.TblItems.Fullcode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee , Transactions.NoteSerial1, ((Transaction_Details.showPrice * Transaction_Details.ShowQty * ISNULL(dbo.Transactions.Currency_rate, 1)    )-isnull (TotalDiscountPerLine,0) * ISNULL(dbo.Transactions.Currency_rate, 1)    -  CASE WHEN ItemDiscountType = 1 THEN 0 WHEN ItemDiscountType = 2 THEN ((ItemDiscount * ISNULL(dbo.Transactions.Currency_rate, 1)))            WHEN ItemDiscountType = 3 THEN ((Transaction_Details.showqty * Transaction_Details.showPrice) * ((ItemDiscount / 100)) * ISNULL(dbo.Transactions.Currency_rate, 1))            WHEN ItemDiscountType = 4 THEN (((Transaction_Details.showqty * Transaction_Details.showPrice) * ISNULL(dbo.Transactions.Currency_rate, 1))) ELSE 0 END ) AS sumVAT"
  End If
  
End If

    Sql = Sql & " FROM VatTypes INNER JOIN "
    Sql = Sql & " Transactions ON VatTypes.ID = Transactions.Transaction_Type LEFT OUTER JOIN "
    Sql = Sql & " TblStore ON Transactions.StoreID = TblStore.StoreID LEFT OUTER JOIN "
    Sql = Sql & " TblBranchesData ON Transactions.BranchId = TblBranchesData.branch_id FULL OUTER JOIN "
    Sql = Sql & " TblUnites RIGHT OUTER JOIN "
    Sql = Sql & " TblItemsUnits ON TblUnites.UnitID = TblItemsUnits.UnitID RIGHT OUTER JOIN "
    Sql = Sql & " TblItems ON TblItemsUnits.JunckID = TblItems.ItemID RIGHT OUTER JOIN "
    Sql = Sql & " Transaction_Details ON TblItems.ItemID = Transaction_Details.Item_ID ON Transactions.Transaction_ID = Transaction_Details.Transaction_ID "
    Sql = Sql & " Where VatTypes.ID =  " & TransType & " "
    If val(DcBranch.BoundText) <> 0 Then
    Sql = Sql & " and  (Transactions.BranchId = " & val(DcBranch.BoundText) & ")"
   End If
   
If SystemOptions.PriceWithVAT = False Then
    If withVAT = True Then
        Sql = Sql & " AND (Vatyo <> 0 Or Vatyo Is Not Null) "
    Else
        Sql = Sql & " AND (Vatyo = 0 Or Vatyo Is Null) "
    End If
Else
End If


  If SystemOptions.PriceWithVAT = False Then
     If Remarks <> -1 Then
      Sql = Sql & " AND (Vatyo <> 0 Or Vatyo Is Not Null) "
      Sql = Sql & " AND  Typ=" & Remarks & ""
      GoTo ll:
     End If
 If IsNull(Vatyo) Then
      Sql = Sql & " AND (Vatyo Is   Null) "
      If TransType = 5 Or TransType = 22 Then
      Sql = Sql & " AND  Typ=-1"
      End If
      
     Else
     If Vatyo = 0 Then
      Sql = Sql & " AND ( Vatyo = " & Vatyo & " ) "
      Else
      Sql = Sql & " AND (Vatyo =15 or Vatyo = " & Vatyo & " ) "
      End If
      
      If TransType = 5 Or TransType = 22 Then
      Sql = Sql & " AND  Typ=-1  "
      End If
      
  End If
  ElseIf TransType <> 21 And TransType <> 9 Then
    If Remarks <> -1 Then
   Sql = Sql & " AND (Vatyo <> 0 Or Vatyo Is Not Null) "
      Sql = Sql & " AND  Typ=" & Remarks & ""
  GoTo ll:
  End If
 If IsNull(Vatyo) Then
     Sql = Sql & " AND (Vatyo Is   Null) "
           If TransType = 5 Or TransType = 22 Then
      Sql = Sql & " AND  Typ=-1"
      End If
      
     Else
      Sql = Sql & " AND (Vatyo=15 or Vatyo = " & Vatyo & " ) "
      If TransType = 5 Or TransType = 22 Then
      Sql = Sql & " AND  Typ=-1"
      End If
      
  End If
 End If
ll:
  
    If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & " and Transactions.Transaction_Date >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & " and Transactions.Transaction_Date <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
    
    
  ' If VstReverse = 0 Then
  '   sql = sql & " and isnull(Transactions.VstReverse,0) =0 "
  ' Else
  ' sql = sql & " and isnull(Transactions.VstReverse,0) =1 "
    
  ' End If
     If Vatyo = 66 And SystemOptions.PriceWithVAT Then
            If TransType = 21 Or TransType = 9 Then
                Sql = Sql & " AND IsNull(Transactions.chkTaxExempt,0) = 1"
            End If
    Else
        Sql = Sql & " AND IsNull(Transactions.chkTaxExempt,0) = 0"
    End If
    
     RsVAT.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If RsVAT.RecordCount > 0 Then
        getTotalVATItemsValue = IIf(IsNull(RsVAT("sumVAT").value), 0, RsVAT("sumVAT").value)
        getTotalVATItemsValue = Round(getTotalVATItemsValue, 2)
     If TransType = 21 Or TransType = 9 Then
     If SystemOptions.PriceWithVAT = True Then
        If Vatyo <> 66 Then
            getTotalVATItemsValue = getTotalVATItemsValue / (1 + intervalVat)    '1.05
        End If
      End If
     End If
    Else
        getTotalVATItemsValue = 0
    End If
    
If PrintReport = True Then
 
print_report , Sql, reporttitle
End If
End Function
Function getTotalVATItemsValue222(TransType As Integer, withVAT As Boolean, Optional Vatyo As Variant = Null, Optional Remarks As Integer = -1, Optional PrintReport As Boolean = False, Optional reporttitle As String) As Double
    Dim Sql As String
    Dim RsVAT As ADODB.Recordset
    Set RsVAT = New ADODB.Recordset
 
  Sql = " SELECT        SUM(dbo.Transactions.Transaction_NetValue - dbo.Transactions.VAT) AS sumVAT"
  Sql = Sql & "      FROM            dbo.VatTypes INNER JOIN"
  Sql = Sql & "                       dbo.Transactions ON dbo.VatTypes.ID = dbo.Transactions.Transaction_Type LEFT OUTER JOIN"
  Sql = Sql & "                       dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
  Sql = Sql & "                       dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"

    Sql = Sql & " Where VatTypes.ID =  " & TransType & " "
    If val(DcBranch.BoundText) <> 0 Then
    Sql = Sql & " and  (Transactions.BranchId = " & val(DcBranch.BoundText) & ")"
   End If


  
    If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & " and Transactions.Transaction_Date >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & " and Transactions.Transaction_Date <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
    
     RsVAT.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If RsVAT.RecordCount > 0 Then
        getTotalVATItemsValue222 = IIf(IsNull(RsVAT("sumVAT").value), 0, RsVAT("sumVAT").value)
        getTotalVATItemsValue222 = Round(getTotalVATItemsValue222, 2)
     If TransType = 21 Or TransType = 9 Then
     If SystemOptions.PriceWithVAT = True Then
        getTotalVATItemsValue222 = getTotalVATItemsValue222 / (intervalVat + 1)  '1.05
      End If
     End If
    Else
        getTotalVATItemsValue222 = 0
    End If
    
If PrintReport = True Then
 
print_report , Sql, reporttitle
End If
End Function
Public Sub GetVAT_Omrah(Optional ByRef Vat As Double, Optional ByRef netvalue As Double)
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
Sql = " SELECT     SUM(FATValue) AS VAT, SUM(Total) AS NetValue"
Sql = Sql & " From dbo.tblbookingrequest2"
Sql = Sql & " where 1=1"
If val(DcBranch.BoundText) <> 0 Then
Sql = Sql & " and  BranchID=" & val(Me.DcBranch.BoundText) & " "
End If

    If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & " and SDate >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & " and SDate <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
  Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  If Rs2.RecordCount > 0 Then
  Vat = IIf(IsNull(Rs2("VAT").value), 0, Rs2("VAT").value)
  netvalue = IIf(IsNull(Rs2("NetValue").value), 0, Rs2("NetValue").value)
  Else
  netvalue = 0
  Vat = 0
  End If
End Sub
Public Sub GetVAT_Haijj(Optional ByRef Vat As Double, Optional ByRef netvalue As Double)
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
Sql = " SELECT     SUM(FATValue) AS VAT, SUM(NetValue) AS NetValue"
Sql = Sql & " From dbo.TblDetailsAdoption"
Sql = Sql & " where 1=1"
If val(DcBranch.BoundText) <> 0 Then
Sql = Sql & " and  (BranchID = " & val(DcBranch.BoundText) & ")"
End If

    If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & " and dbo.TblDetailsAdoption.RecordDate  >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & " and dbo.TblDetailsAdoption.RecordDate    <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
  Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  If Rs2.RecordCount > 0 Then
  Vat = IIf(IsNull(Rs2("VAT").value), 0, Rs2("VAT").value)
  netvalue = IIf(IsNull(Rs2("NetValue").value), 0, Rs2("NetValue").value)
  Else
  netvalue = 0
  Vat = 0
  End If
End Sub
Public Sub GetVAT_ReqCon(Optional ByRef Vat As Double, Optional ByRef netvalue As Double)
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
Sql = "  SELECT     SUM(dbo.TblExchangeReques_Detailst.Vat2) AS Vat2 ,SUM( dbo.TblExchangeReques_Detailst.TotalNet) AS TotalNet"
Sql = Sql & " FROM         dbo.TblExchangeReques_Detailst INNER JOIN"
Sql = Sql & "                       dbo.TblExchangeRequest ON dbo.TblExchangeReques_Detailst.HID = dbo.TblExchangeRequest.ID"
Sql = Sql & "  Where (   not(       NoteID is null   ) and   (dbo.TblExchangeReques_Detailst.Vat2    >0))"

 



If val(DcBranch.BoundText) <> 0 Then
Sql = Sql & " and  (dbo.TblExchangeRequest.BranchID = " & val(DcBranch.BoundText) & ")"
End If
    If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & " and  dbo.TblExchangeRequest.EntryDate >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & " and  dbo.TblExchangeRequest.EntryDate <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
  Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  If Rs2.RecordCount > 0 Then
  Vat = IIf(IsNull(Rs2("VAT2").value), 0, Rs2("VAT2").value)
 netvalue = IIf(IsNull(Rs2("TotalNet").value), 0, Rs2("TotalNet").value)
  Else
  netvalue = 0
  Vat = 0
  End If
'  rs2.Close
End Sub


Public Sub GetVAT_Hand(Optional ByRef Vat As Double, Optional ByRef netvalue As Double, Optional ByVal mType As Integer = 0)
Dim Sql As String, StrSQL As String
Dim Rs2 As ADODB.Recordset
Vat = 0
netvalue = 0
If mType = 1 Or mType = 0 Then
    Set Rs2 = New ADODB.Recordset
    Sql = "  SELECT     SUM(dbo.TblHandWages.Vat2) AS Vat2 ,SUM( dbo.TblHandWages.Net) AS TotalNet"
    Sql = Sql & " FROM         dbo.TblHandWages"
    Sql = Sql & "  Where (   not(       NoteID is null   ) and   (dbo.TblHandWages.Vat2    >0))"
    
     
    
    
    
    If val(DcBranch.BoundText) <> 0 Then
    Sql = Sql & " and  (dbo.TblHandWages.BranchID = " & val(DcBranch.BoundText) & ")"
    End If
        If Not IsNull(Me.DateFrom.value) Then
            Sql = Sql & " and  dbo.TblHandWages.RecordDate >= " & SQLDate(Me.DateFrom.value, True) & " "
        End If
        If Not IsNull(Me.DateTo.value) Then
            Sql = Sql & " and  dbo.TblHandWages.RecordDate <= " & SQLDate(Me.DateTo.value, True) & " "
        End If
      Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
      If Rs2.RecordCount > 0 Then
        Vat = IIf(IsNull(Rs2("VAT2").value), 0, Rs2("VAT2").value)
        netvalue = IIf(IsNull(Rs2("TotalNet").value), 0, Rs2("TotalNet").value) - Vat
      Else
        netvalue = 0
        Vat = 0
      End If
  
ElseIf mType = 2 Or mType = 0 Then
    Set Rs2 = New ADODB.Recordset
   
       
        StrSQL = " SELECT     Sum(t.Vat) Vat2,  Sum(t.netvalue) netvalue"
        
        
        StrSQL = StrSQL & " FROM         "
        
        StrSQL = StrSQL & "                      dbo.Transactions t "
        StrSQL = StrSQL & " Where (t.Transaction_Type = 21) And  CBoBasedON = 7 and  IsNull(t.order_no,'') <> ''"
        
     
         Dim mTotal As Double
      
    
    
    
    If val(DcBranch.BoundText) <> 0 Then
    StrSQL = StrSQL & " and  (t.BranchID = " & val(DcBranch.BoundText) & ")"
    End If
        If Not IsNull(Me.DateFrom.value) Then
            StrSQL = StrSQL & " and  t.Transaction_Date >= " & SQLDate(Me.DateFrom.value, True) & " "
        End If
        If Not IsNull(Me.DateTo.value) Then
            StrSQL = StrSQL & " and  t.Transaction_Date <= " & SQLDate(Me.DateTo.value, True) & " "
        End If
      Rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
      If Rs2.RecordCount > 0 Then
            Vat = Vat + Round(val(Rs2!Vat2 & ""), 2)
            'netvalue = netvalue + Round(val(rs2!netvalue & "") + val(rs2!Vat2 & ""), 2)
            netvalue = netvalue + Round(val(Rs2!netvalue & ""), 2)
      Else
            'netvalue = 0
            'Vat = 0
      End If
 End If
  
'  rs2.Close
End Sub

Public Sub GetVAT_Minis(Optional ByRef Vat As Double, Optional ByRef netvalue As Double, Optional FATYou As Double)
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
Sql = " SELECT     sum(dbo.TblMinistryContract_Installment.[Value]) AS NetValue ,sum(dbo.TblMinistryContract_Installment.FATValue) AS VAT"
Sql = Sql & " FROM         dbo.TblMinistryContract_Installment INNER JOIN"
Sql = Sql & "                      dbo.TblMinistryContract ON dbo.TblMinistryContract_Installment.IDMC = dbo.TblMinistryContract.IDMC"
'sql = sql & " Where dbo.TblMinistryContract_Installment.Due_Date>'31/5/501'"
'sql = sql & " WHERE 1=1"
Sql = Sql & " WHERE     (dbo.TblMinistryContract_Installment.Due_Date > CONVERT(DATETIME, '2018-05-31 00:00:00', 102))"
If val(DcBranch.BoundText) <> 0 Then
Sql = Sql & " and  (dbo.TblMinistryContract.BranchID = " & val(DcBranch.BoundText) & ")"
End If

    If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & " and  dbo.TblMinistryContract_Installment.Due_Date >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & " and  dbo.TblMinistryContract_Installment.Due_Date <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
    
  '  If FATYou <> 0 Then
  '      sql = sql & " and  FATYou=" & FATYou
  '  End If
    
     If Not IsNull(FATYou) Then
     Sql = Sql & " and  FATYou=" & FATYou
     End If
     
    
  Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  If Rs2.RecordCount > 0 Then
  Vat = IIf(IsNull(Rs2("VAT").value), 0, Rs2("VAT").value)
  netvalue = IIf(IsNull(Rs2("NetValue").value), 0, Rs2("NetValue").value)
  Else
  netvalue = 0
  Vat = 0
  End If
End Sub
Public Sub GetVAT_Transportations(Optional ByRef Vat As Double, Optional ByRef netvalue As Double, Optional FATYou As Double)
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
Sql = " SELECT     SUM(VAT) AS Vat, SUM(TotalValue) AS NetValue From dbo.TblTravDueK"
Sql = Sql & " Where 1=1"
If val(DcBranch.BoundText) <> 0 Then
Sql = Sql & " and  (dbo.TblTravDueK.BranchID = " & val(DcBranch.BoundText) & ")"
End If

    If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & " and  dbo.TblTravDueK.recordDate >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & " and  dbo.TblTravDueK.recordDate <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
    
  '  If FATYou <> 0 Then
  '      sql = sql & " and  FATYou=" & FATYou
  '  End If
 
    
  Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  If Rs2.RecordCount > 0 Then
  Vat = IIf(IsNull(Rs2("VAT").value), 0, Rs2("VAT").value)
  netvalue = IIf(IsNull(Rs2("NetValue").value), 0, Rs2("NetValue").value)
  Else
  netvalue = 0
  Vat = 0
  End If
End Sub
Public Sub GetVAT_PREVAT(Optional ByRef Vat As Double, Optional ByRef netvalue As Double, Optional FATYou As Double)
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
Sql = " SELECT     SUM(VAT) AS VAT, SUM(Note_Value) AS TotalValue FROM         dbo.Notes"
Sql = Sql & " Where 1=1"
Sql = Sql & " and  (NCashingType = 3)    AND (NoteType = 4) AND (VAT > 0)"

If val(DcBranch.BoundText) <> 0 Then
Sql = Sql & " and  (dbo.Notes.branch_no = " & val(DcBranch.BoundText) & ")"
End If

    If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & " and  dbo.Notes.NoteDate >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & " and  dbo.Notes.NoteDate <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
   Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  If Rs2.RecordCount > 0 Then
  Vat = IIf(IsNull(Rs2("VAT").value), 0, Rs2("VAT").value)
  netvalue = IIf(IsNull(Rs2("TotalValue").value), 0, Rs2("TotalValue").value)
  Else
  netvalue = 0
  Vat = 0
  End If
End Sub


Public Sub GetVAT_PREVAT1(Optional ByRef Vat As Double, Optional ByRef netvalue As Double, Optional FATYou As Double)
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
 
Sql = "SELECT     SUM(dbo.Notes.PreVAT) AS VAT, sum(dbo.Notes.Note_Value2  ) AS TotalValue"
Sql = Sql & "  FROM         dbo.Notes INNER JOIN"
Sql = Sql & "                       dbo.TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id"
                      
Sql = Sql & " Where 1=1"
Sql = Sql & "  and CashingType<>7   AND (NoteType = 5) AND (PreVAT > 0)"

If val(DcBranch.BoundText) <> 0 Then
Sql = Sql & " and  (dbo.Notes.branch_no = " & val(DcBranch.BoundText) & ")"
End If

    If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & " and  dbo.Notes.NoteDate >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & " and  dbo.Notes.NoteDate <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
   Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  If Rs2.RecordCount > 0 Then
  Vat = IIf(IsNull(Rs2("VAT").value), 0, Rs2("VAT").value)
  netvalue = IIf(IsNull(Rs2("TotalValue").value), 0, Rs2("TotalValue").value)
  Else
  netvalue = 0
  Vat = 0
  End If
End Sub

Public Sub GetVAT_ManualNotes(Optional ByRef TotalVat As Double, Optional ByRef netvalue As Double, Optional Credit_Or_Debit As Double, Optional NoteType As Double)

 
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset

Sql = "SELECT     SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value]) AS TotalValue, SUM(dbo.DOUBLE_ENTREY_VOUCHERS.Vat) AS Vat"
Sql = Sql & "  FROM         dbo.Notes INNER JOIN"
Sql = Sql & " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID"
Sql = Sql & " WHERE  dbo.DOUBLE_ENTREY_VOUCHERS.Vat > 0"
Sql = Sql & "  and  dbo.Notes.NoteType =" & NoteType

Sql = Sql & "  and dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = " & Credit_Or_Debit
 
If val(DcBranch.BoundText) <> 0 Then
Sql = Sql & " and  (dbo.DOUBLE_ENTREY_VOUCHERS.branch_id = " & val(DcBranch.BoundText) & ")"
End If

    If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & " and  dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & " and  dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
   Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  If Rs2.RecordCount > 0 Then
  TotalVat = IIf(IsNull(Rs2("VAT").value), 0, Rs2("VAT").value)
  netvalue = IIf(IsNull(Rs2("TotalValue").value), 0, Rs2("TotalValue").value)
  Else
  TotalVat = 0
  netvalue = 0
  End If
   
End Sub

Public Sub GetVAT_Notes(Optional ByRef TotalVat As Double, Optional ByRef netvalue As Double, Optional docType As Double)
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset

Sql = " SELECT     SUM(VAT) AS TotalVat, SUM(Note_Value) AS netvalue"
Sql = Sql & " From dbo.Notes where 1=1 "
If docType = 9089 Or docType = 9090 Then
Else

Sql = Sql & " and (VATYou > 0)"
End If
If val(DcBranch.BoundText) <> 0 Then
Sql = Sql & " and  ( dbo.Notes.branch_no = " & val(DcBranch.BoundText) & ")"
End If

    If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & " and  dbo.Notes.NoteDate >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & " and  dbo.Notes.NoteDate <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
    
'If (docType = 9082) Then
' Sql = Sql & " and  dbo.Notes.notetype =" & 9082
' ElseIf (docType = 9083) Then
' Sql = Sql & " and  dbo.Notes.notetype =" & 9083
'
'
'End If
Sql = Sql & " and  dbo.Notes.notetype =" & docType

  Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  If Rs2.RecordCount > 0 Then
  TotalVat = IIf(IsNull(Rs2("TotalVat").value), 0, Rs2("TotalVat").value)
  netvalue = IIf(IsNull(Rs2("NetValue").value), 0, Rs2("NetValue").value)
  TotalVat = Round(TotalVat, 2)
  netvalue = Round(netvalue, 2)
  
  Else
  netvalue = 0
  TotalVat = 0
  End If
End Sub
Public Sub GetVAT_Expenses(Optional ByRef TotalVat As Double, Optional ByRef netvalue As Double, Optional docType As Double)
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset

Sql = "  SELECT     SUM(dbo.DOUBLE_ENTREY_VOUCHERS.[Value]) AS netvalue, SUM(dbo.DOUBLE_ENTREY_VOUCHERS.Vat) AS TotalVat"
Sql = Sql & "                       FROM         dbo.Notes INNER JOIN"
Sql = Sql & "                                             dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
Sql = Sql & "                                             dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType INNER JOIN"
 Sql = Sql & "                                            dbo.notes_all ON dbo.Notes.notes_all = dbo.notes_all.NoteID"
Sql = Sql & "                       Where 1=1   and dbo.DOUBLE_ENTREY_VOUCHERS.hideline is null   And (dbo.DOUBLE_ENTREY_VOUCHERS.Vatyo > 0)"

 
If val(DcBranch.BoundText) <> 0 Then
Sql = Sql & " and  ( dbo.DOUBLE_ENTREY_VOUCHERS.branch_id = " & val(DcBranch.BoundText) & ")"
End If

    If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & " and  dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & " and  dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
    
If (docType = 3) Then
 Sql = Sql & " and  dbo.Notes.notetype =" & 3 '„’—Ê›« 
ElseIf (docType = 80) Then
 Sql = Sql & " and  (bill_type =0 or  bill_type =1) and dbo.Notes.notetype =" & 80 '›« Ê—… „«·Ì…
ElseIf (docType = 802) Then
 Sql = Sql & " and   bill_type =2 and dbo.Notes.notetype =" & 80 '›« Ê—… «’Ê·

ElseIf (docType = 85) Then
 Sql = Sql & " and    dbo.notes_all.notetype =" & 85 '›« Ê—… Œœ„Ì…
 Sql = Sql & " and  isnull(AkarPayCheck,0) =0"
ElseIf (docType = 350) Then
 Sql = Sql & " and    dbo.notes_all.notetype =" & 350 '⁄Âœ…

End If


  Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  If Rs2.RecordCount > 0 Then
  TotalVat = IIf(IsNull(Rs2("TotalVat").value), 0, Rs2("TotalVat").value)
  netvalue = IIf(IsNull(Rs2("NetValue").value), 0, Rs2("NetValue").value)
  
  TotalVat = Round(TotalVat, 2)
   netvalue = Round(netvalue, 2)
   
  
  Else
  netvalue = 0
  TotalVat = 0
  End If
End Sub
Public Sub PreVAT(Optional ByRef TotalVat As Double, Optional ByRef netvalue As Double, Optional docType As Double)
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset

 
 Sql = "SELECT     SUM(PreVAT) AS TotalVat, SUM(Note_Value2) AS netvalue"
Sql = Sql & " From dbo.Notes"
Sql = Sql & " where   (  CashingType =7 and  (PREVAT) > 0)"

If val(DcBranch.BoundText) <> 0 Then
Sql = Sql & " and  ( dbo.Notes.branch_no = " & val(DcBranch.BoundText) & ")"
End If

    If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & " and  dbo.Notes.NoteDate >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & " and  dbo.Notes.NoteDate <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
    
 

  Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  If Rs2.RecordCount > 0 Then
  TotalVat = IIf(IsNull(Rs2("TotalVat").value), 0, Rs2("TotalVat").value)
   netvalue = IIf(IsNull(Rs2("netvalue").value), 0, Rs2("netvalue").value)
  Else
  netvalue = 0
  TotalVat = 0
  End If
End Sub

Public Sub NetCommisionVat(Optional ByRef TotalValue As Double, Optional ByRef VATValue As Double, Optional docType As Double)
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset

  Sql = "SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.Value AS TotalValue, ROUND(dbo.DOUBLE_ENTREY_VOUCHERS.TotalValue, 2) AS  vatValue "
 Sql = Sql & "  FROM         dbo.Notes INNER JOIN"
Sql = Sql & "                       dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID"
Sql = Sql & " WHERE     (dbo.Notes.NoteType = 170) AND (dbo.DOUBLE_ENTREY_VOUCHERS.TotalValue > 0)  "
Sql = Sql & "           "

If val(DcBranch.BoundText) <> 0 Then
Sql = Sql & " and  ( dbo.Notes.branch_no = " & val(DcBranch.BoundText) & ")"
End If
   If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & " and  dbo.Notes.NoteDate >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & " and  dbo.Notes.NoteDate <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
    
 

  Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  If Rs2.RecordCount > 0 Then
  TotalValue = IIf(IsNull(Rs2("TotalValue").value), 0, Rs2("TotalValue").value)
   VATValue = IIf(IsNull(Rs2("vatValue").value), 0, Rs2("vatValue").value)
  Else
  TotalValue = 0
  VATValue = 0
  End If
End Sub

Public Sub TransferVAT(Optional ByRef TotalVat As Double, Optional ByRef netvalue As Double, Optional docType As Double)
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset

 

Sql = "select sum (TotalTransferVale) netvalue , sum (VatForExpenses)TotalVat"

Sql = Sql & "         From"
Sql = Sql & " ("
Sql = Sql & " SELECT     ISNULL(TransferExpensesBranch, 0) + ISNULL(TransferExpenses, 0) AS TotalTransferVale, (ISNULL(TransferExpensesBranch, 0) + ISNULL(TransferExpenses, 0)"
'sql = sql & "        )                       / 1.05 AS VatForExpenses"
Sql = Sql & "        )                       / " & (intervalVat + 1) & " AS VatForExpenses"

Sql = Sql & "         From dbo.Notes"
Sql = Sql & "         Where (IncludVAT = 1)"
If val(DcBranch.BoundText) <> 0 Then
Sql = Sql & " and  ( dbo.Notes.branch_no = " & val(DcBranch.BoundText) & ")"
End If

    If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & " and  dbo.Notes.NoteDate >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & " and  dbo.Notes.NoteDate <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
    
    
Sql = Sql & "  )z"




 

  Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  If Rs2.RecordCount > 0 Then
  TotalVat = IIf(IsNull(Rs2("TotalVat").value), 0, Rs2("TotalVat").value)
   netvalue = IIf(IsNull(Rs2("netvalue").value), 0, Rs2("netvalue").value) - TotalVat
  Else
  netvalue = 0
  TotalVat = 0
  End If
End Sub

Public Sub GetVAT_Customs(Optional ByRef TotalVat As Double, Optional ByRef netvalue As Double, Optional docType As Double)
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset

Sql = " SELECT       SUM(VATCustoms)  as TotalVat From dbo.notes_all where 1=1"

 
If val(DcBranch.BoundText) <> 0 Then
Sql = Sql & " and  ( dbo.notes_all.branch_no = " & val(DcBranch.BoundText) & ")"
End If

    If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & " and  dbo.notes_all.NoteDate >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & " and  dbo.notes_all.NoteDate <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
    
 

  Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  If Rs2.RecordCount > 0 Then
  TotalVat = IIf(IsNull(Rs2("TotalVat").value), 0, Rs2("TotalVat").value)
  'netvalue = IIf(IsNull(Rs2("NetValue").value), 0, Rs2("NetValue").value)
  Else
  netvalue = 0
  TotalVat = 0
  End If
End Sub

Public Sub GetVAT_FABuy(Optional ByRef Vat As Double, Optional ByRef netvalue As Double, Optional ByRef FATYou, Optional ComResid As Integer)
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
Sql = " SELECT     SUM(dbo.notes_all.VAT2) AS FATValue, SUM(dbo.notes_all.FASalesPrice ) AS NetValue"
Sql = Sql & "   FROM         dbo.notes_all "
                      
Sql = Sql & "  where notes_all.NoteType = 8028 and   IsNull(FATypeOp,0) = 0 "


If val(DcBranch.BoundText) <> 0 Then
Sql = Sql & " and  (notes_all.Branch_NO = " & val(DcBranch.BoundText) & ")"
End If


    If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & " and NoteDate >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & " and NoteDate <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
    If Not IsNull(FATYou) Then
     'Sql = Sql & " and FATYou=" & FATYou
    End If
    
  Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  If Rs2.RecordCount > 0 Then
  Vat = IIf(IsNull(Rs2("FATValue").value), 0, Rs2("FATValue").value)
  netvalue = IIf(IsNull(Rs2("NetValue").value), 0, Rs2("NetValue").value)
  netvalue = netvalue '- Vat
  Else
  netvalue = 0
  Vat = 0
  End If
End Sub


Public Sub GetVAT_Contract(Optional ByRef Vat As Double, Optional ByRef netvalue As Double, Optional ByRef FATYou, Optional ComResid As Integer)
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
Sql = " SELECT     SUM(dbo.TblContractInstallments.VATValue) AS FATValue, SUM(dbo.TblContractInstallments.installValue +dbo.TblContractInstallments.NpayedValue -  dbo.TblContractInstallments.Insurance  )AS NetValue"
Sql = Sql & "   FROM         dbo.TblContract INNER JOIN"
Sql = Sql & "                        dbo.TblContractInstallments ON dbo.TblContract.ContNo = dbo.TblContractInstallments.ContNo INNER JOIN"
Sql = Sql & "                        dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID INNER JOIN"
Sql = Sql & "  dbo.TblBranchesData ON dbo.TblContract.Branch_NO = dbo.TblBranchesData.branch_id"
                      
Sql = Sql & "  where  1=1  "
If ComResid = 0 Then
Sql = Sql & " and (dbo.TblContract.ComResid = 0 )" '”ﬂ‰Ì
If paidchk(2).value = vbChecked Then
Sql = Sql & "  and (dbo.TblContract.EndContract  is null )"
 End If

Else
Sql = Sql & "  and (dbo.TblContract.ComResid = 1) " ' Ã«—Ì
If paidchk(1).value = vbChecked Then
Sql = Sql & "  and (dbo.TblContract.EndContract  is null )"
 End If
End If

If val(DcBranch.BoundText) <> 0 Then
Sql = Sql & " and  (TblContract.Branch_NO = " & val(DcBranch.BoundText) & ")"
End If
 





    If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & " and Installdate >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & " and Installdate <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
    If Not IsNull(FATYou) Then
     'Sql = Sql & " and FATYou=" & FATYou
    End If
    
  Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  If Rs2.RecordCount > 0 Then
  Vat = IIf(IsNull(Rs2("FATValue").value), 0, Rs2("FATValue").value)
  netvalue = IIf(IsNull(Rs2("NetValue").value), 0, Rs2("NetValue").value)
  netvalue = netvalue - Vat
  Else
  netvalue = 0
  Vat = 0
  End If
End Sub
Public Sub GetVAT_SalesAdvancedPayment(Optional ByRef Vat As Double, Optional ByRef netvalue As Double, Optional Cus_Cons As Integer = 0, Optional FATYou As Double, Optional ByRef advancedPayment As Double, Optional PreVAT As Integer = 0)
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
Sql = " SELECT     SUM(SumVATLine) AS VAT, SUM(SumValueLine) AS  AdvancedPayment "
Sql = Sql & " From dbo.Transactions where 1=1"
 
If val(DcBranch.BoundText) <> 0 Then
Sql = Sql & " and  BranchId=" & val(Me.DcBranch.BoundText) & " "
End If
    If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & " and Transaction_Date >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & " and Transaction_Date <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
    
 
    
    
  Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  If Rs2.RecordCount > 0 Then
  Vat = IIf(IsNull(Rs2("VAT").value), 0, Rs2("VAT").value)
  advancedPayment = IIf(IsNull(Rs2("AdvancedPayment").value), 0, Rs2("AdvancedPayment").value)
  
  Else
  netvalue = 0
  Vat = 0
  
  advancedPayment = 0
  End If
End Sub

Public Sub GetVAT_ProjectAdvancedPayment(Optional ByRef Vat As Double, Optional ByRef netvalue As Double, Optional Cus_Cons As Integer = 0, Optional FATYou As Double, Optional ByRef advancedPayment As Double, Optional PreVAT As Integer = 0)
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
Sql = " SELECT     SUM(FATValue-PreVAT) AS VAT, SUM(total) AS NetValue,sum(AdvancedPayment) as AdvancedPayment"
Sql = Sql & " From dbo.project_billl"
Sql = Sql & " where  bill_to=" & Cus_Cons & ""
If val(DcBranch.BoundText) <> 0 Then
Sql = Sql & " and  branch_no=" & val(Me.DcBranch.BoundText) & " "
End If
    If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & " and bill_date >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & " and bill_date <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
    
        If Not IsNull(FATYou) Then
     Sql = Sql & " and (FATYou=15 or  FATYou=" & FATYou & ")"
    End If
    
            If PreVAT > 0 Then
     Sql = Sql & " and PreVAT>0"
     Else
     Sql = Sql & " and (PreVAT=0 or PreVAT is null)"
    End If
    
    
    
  Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  If Rs2.RecordCount > 0 Then
  Vat = IIf(IsNull(Rs2("VAT").value), 0, Rs2("VAT").value)
  netvalue = IIf(IsNull(Rs2("NetValue").value), 0, Rs2("NetValue").value)
  advancedPayment = IIf(IsNull(Rs2("AdvancedPayment").value), 0, Rs2("AdvancedPayment").value)
  
  Else
  netvalue = 0
  Vat = 0
  
  advancedPayment = 0
  End If
End Sub
Public Sub GetVAT_Project(Optional ByRef Vat As Double, Optional ByRef netvalue As Double, Optional Cus_Cons As Integer = 0, Optional FATYou As Double, Optional ByRef advancedPayment As Double)
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
Sql = " SELECT     SUM(FATValue-PreVAT) AS VAT, SUM(total )+SUM(PerforValue) AS NetValue,sum(AdvancedPayment) as AdvancedPayment"
Sql = Sql & " From dbo.project_billl"
Sql = Sql & " where  bill_to=" & Cus_Cons & ""
If val(DcBranch.BoundText) <> 0 Then
Sql = Sql & " and  branch_no=" & val(Me.DcBranch.BoundText) & " "
End If
    If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & " and bill_date >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & " and bill_date <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
        If FATYou = 0 Then
            Sql = Sql & " and (FATYou=" & FATYou & ")"
        Else
        If Not IsNull(FATYou) Then
            'sql = sql & " and FATYou=" & FATYou 'xx
            Sql = Sql & " and (FATYou=15 or  FATYou=" & FATYou & ")"
            
        End If
        End If
    
  Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  If Rs2.RecordCount > 0 Then
  Vat = IIf(IsNull(Rs2("VAT").value), 0, Rs2("VAT").value)
  netvalue = IIf(IsNull(Rs2("NetValue").value), 0, Rs2("NetValue").value)
  advancedPayment = IIf(IsNull(Rs2("AdvancedPayment").value), 0, Rs2("AdvancedPayment").value)
  
  Else
  netvalue = 0
  Vat = 0
  
  advancedPayment = 0
  End If
End Sub

Sub ClaCulte()


End Sub
Function getItemsVATValue(TransType As Integer) As Double
    Dim Sql As String
    Dim RsVAT As ADODB.Recordset
    Set RsVAT = New ADODB.Recordset
    'sql = "SELECT sum(Transaction_Details.showPrice * Transaction_Details.ShowQty * isnull (Transactions.Currency_rate,1) * (Vatyo/100)) as sumVAT "
Sql = "     SELECT  sum("
  Sql = Sql & "  ("
  Sql = Sql & " Transaction_Details.showPrice * Transaction_Details.ShowQty"
  If TransType = 22 Then
        Sql = Sql & "  -isnull (TotalDiscountPerLine,0) * Transaction_Details.ShowQty -"
  Else
  Sql = Sql & "  -isnull (TotalDiscountPerLine,0)-"
         
  End If
  If TransType = 21 Or TransType = 9 Then
 If SystemOptions.PriceWithVAT = True Then
  Sql = Sql & " CASE WHEN ItemDiscountType = 1 THEN 0"
       Sql = Sql & " WHEN ItemDiscountType = 2 THEN (    (  ItemDiscount *  ISNULL(dbo.Transactions.Currency_rate, 1)   )  )"
       Sql = Sql & " WHEN ItemDiscountType = 3 THEN ((Transaction_Details.showqty * Transaction_Details.showPrice) * ((ItemDiscount / 100)) * ISNULL(dbo.Transactions.Currency_rate, 1))"
       Sql = Sql & " WHEN ItemDiscountType = 4 THEN (((Transaction_Details.showqty * Transaction_Details.showPrice) * ISNULL(dbo.Transactions.Currency_rate, 1))) ELSE 0 END"
  Sql = Sql & " )/  (1+" & intervalVat & ")" '''
Else
  Sql = Sql & " CASE WHEN ItemDiscountType = 1 THEN 0"
       Sql = Sql & " WHEN ItemDiscountType = 2 THEN (    (  ItemDiscount *  ISNULL(dbo.Transactions.Currency_rate, 1)   )  )"
       Sql = Sql & " WHEN ItemDiscountType = 3 THEN ((Transaction_Details.showqty * Transaction_Details.showPrice) * ((ItemDiscount / 100)) * ISNULL(dbo.Transactions.Currency_rate, 1))"
       Sql = Sql & " WHEN ItemDiscountType = 4 THEN (((Transaction_Details.showqty * Transaction_Details.showPrice) * ISNULL(dbo.Transactions.Currency_rate, 1))) ELSE 0 END"
  Sql = Sql & " )*  (Vatyo/100)"
End If
Else
  Sql = Sql & " CASE WHEN ItemDiscountType = 1 THEN 0"
       Sql = Sql & " WHEN ItemDiscountType = 2 THEN (    (  ItemDiscount *  ISNULL(dbo.Transactions.Currency_rate, 1)   )  )"
       Sql = Sql & " WHEN ItemDiscountType = 3 THEN ((Transaction_Details.showqty * Transaction_Details.showPrice) * ((ItemDiscount / 100)) * ISNULL(dbo.Transactions.Currency_rate, 1))"
       Sql = Sql & " WHEN ItemDiscountType = 4 THEN (((Transaction_Details.showqty * Transaction_Details.showPrice) * ISNULL(dbo.Transactions.Currency_rate, 1))) ELSE 0 END"
  Sql = Sql & " )*  (Vatyo/100)"
End If
   Sql = Sql & "  ) AS sumVAT"
 
  '  sql = "SELECT  sum((   (  Transaction_Details.showPrice * Transaction_Details.ShowQty)-TotalDiscountPerLine-  CASE WHEN ItemDiscountType = 1 THEN 0 WHEN ItemDiscountType = 2 THEN ((ItemDiscount * ISNULL(dbo.Transactions.Currency_rate, 1)))            WHEN ItemDiscountType = 3 THEN ((Transaction_Details.showqty * Transaction_Details.showPrice) * ((ItemDiscount / 100)) * ISNULL(dbo.Transactions.Currency_rate, 1))            WHEN ItemDiscountType = 4 THEN (((Transaction_Details.showqty * Transaction_Details.showPrice) * ISNULL(dbo.Transactions.Currency_rate, 1))) ELSE 0 END )   AS sumVAT"
     Sql = Sql & " FROM VatTypes INNER JOIN "
    Sql = Sql & " Transactions ON VatTypes.ID = Transactions.Transaction_Type LEFT OUTER JOIN "
    Sql = Sql & " TblStore ON Transactions.StoreID = TblStore.StoreID LEFT OUTER JOIN "
    Sql = Sql & " TblBranchesData ON Transactions.BranchId = TblBranchesData.branch_id FULL OUTER JOIN "
    Sql = Sql & " TblUnites RIGHT OUTER JOIN "
    Sql = Sql & " TblItemsUnits ON TblUnites.UnitID = TblItemsUnits.UnitID RIGHT OUTER JOIN "
    Sql = Sql & " TblItems ON TblItemsUnits.JunckID = TblItems.ItemID RIGHT OUTER JOIN "
    Sql = Sql & " Transaction_Details ON TblItems.ItemID = Transaction_Details.Item_ID ON Transactions.Transaction_ID = Transaction_Details.Transaction_ID "
    Sql = Sql & " Where VatTypes.ID =  " & TransType & " "
    If SystemOptions.PriceWithVAT = True Then
    If TransType = 21 Or TransType = 9 Then
   ' Sql = Sql & " AND (Vatyo =5   ) "
    Else
    Sql = Sql & " AND (Vatyo =5  or Vatyo =15  ) "
    End If
    Else
    Sql = Sql & " AND (Vatyo =5 or Vatyo =15  ) "
    End If
       
     If TransType = 5 Or TransType = 22 Then
      Sql = Sql & " AND  Typ=-1"
      End If
      
       If val(DcBranch.BoundText) <> 0 Then
    Sql = Sql & " and  (Transactions.BranchId  = " & val(DcBranch.BoundText) & ")"
    End If
    
    If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & " and Transactions.Transaction_Date >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & " and Transactions.Transaction_Date <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
    
    RsVAT.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If RsVAT.RecordCount > 0 Then
        getItemsVATValue = IIf(IsNull(RsVAT("sumVAT").value), 0, RsVAT("sumVAT").value)
    getItemsVATValue = Round(getItemsVATValue, 2)
      If TransType = 21 Or TransType = 9 Then
      If SystemOptions.PriceWithVAT = True Then
      getItemsVATValue = getItemsVATValue * (intervalVat + 1)
     End If
     End If
    Else
        getItemsVATValue = 0
    End If
    
    
End Function
Function getItemsVATValue222(TransType As Integer) As Double
    Dim Sql As String
    Dim RsVAT As ADODB.Recordset
    Set RsVAT = New ADODB.Recordset
Sql = "  SELECT        SUM(dbo.Transactions.VAT) AS sumVAT"
Sql = Sql & " FROM            dbo.VatTypes INNER JOIN"
Sql = Sql & "                         dbo.Transactions ON dbo.VatTypes.ID = dbo.Transactions.Transaction_Type LEFT OUTER JOIN"
Sql = Sql & "                         dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
Sql = Sql & "                         dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
Sql = Sql & " Where VatTypes.ID =  " & TransType & " "
      
    If val(DcBranch.BoundText) <> 0 Then
    Sql = Sql & " and  (Transactions.BranchId  = " & val(DcBranch.BoundText) & ")"
    End If
    
    If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & " and Transactions.Transaction_Date >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & " and Transactions.Transaction_Date <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
    
    RsVAT.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If RsVAT.RecordCount > 0 Then
      getItemsVATValue222 = IIf(IsNull(RsVAT("sumVAT").value), 0, RsVAT("sumVAT").value)
      getItemsVATValue222 = Round(getItemsVATValue222, 2)

    Else
        getItemsVATValue222 = 0
    End If
    
    
End Function
'1
Function getManuTotalVATItemsValue(TransType As Integer, Optional VATPer As Integer) As Double
    Dim Sql As String
    Dim RsVAT As ADODB.Recordset
    Set RsVAT = New ADODB.Recordset
    
    Sql = "SELECT SUM(TblVATSettingsDet.Value) AS ManuVAT"
    Sql = Sql & " FROM TblVATSettings INNER JOIN"
    Sql = Sql & " TblVATSettingsDet ON TblVATSettings.ID = TblVATSettingsDet.VATSettingsID INNER JOIN"
    Sql = Sql & " VatTypes ON TblVATSettings.TransIndx = VatTypes.ID"
    Sql = Sql & " Where VatTypes.ID =  " & TransType & " "
    If val(DcBranch.BoundText) <> 0 Then
    Sql = Sql & " and  (TblVATSettingsDet.BranchID = " & val(DcBranch.BoundText) & ")"
    End If
    
    If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & "and TblVATSettingsDet.DocDate  >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & "and TblVATSettingsDet.DocDate  <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
    
    If Not IsNull(VATPer) Then
     Sql = Sql & "and  VATPer= " & VATPer
    End If
    RsVAT.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If RsVAT.RecordCount > 0 Then
        getManuTotalVATItemsValue = IIf(IsNull(RsVAT("ManuVAT").value), 0, RsVAT("ManuVAT").value)
    Else
        getManuTotalVATItemsValue = 0
    End If
    
End Function
Function getManuVATItemsValue(TransType As Integer) As Double
    Dim Sql As String
    Dim RsVAT As ADODB.Recordset
    Set RsVAT = New ADODB.Recordset
    
    Sql = "SELECT SUM(TblVATSettingsDet.VATValue) AS ManuVAT"
    Sql = Sql & " FROM TblVATSettings INNER JOIN"
    Sql = Sql & " TblVATSettingsDet ON TblVATSettings.ID = TblVATSettingsDet.VATSettingsID INNER JOIN"
    Sql = Sql & " VatTypes ON TblVATSettings.TransIndx = VatTypes.ID"
    Sql = Sql & " Where VatTypes.ID =  " & TransType & ""
    If val(DcBranch.BoundText) <> 0 Then
    Sql = Sql & " and  (TblVATSettingsDet.BranchID = " & val(DcBranch.BoundText) & ")"
    End If
   
    If Not IsNull(Me.DateFrom.value) Then
        Sql = Sql & "and TblVATSettingsDet.Docdate  >= " & SQLDate(Me.DateFrom.value, True) & " "
    End If
    If Not IsNull(Me.DateTo.value) Then
        Sql = Sql & "and TblVATSettingsDet.Docdate  <= " & SQLDate(Me.DateTo.value, True) & " "
    End If
    
    RsVAT.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If RsVAT.RecordCount > 0 Then
        getManuVATItemsValue = IIf(IsNull(RsVAT("ManuVAT").value), 0, RsVAT("ManuVAT").value)
    Else
        getManuVATItemsValue = 0
    End If
    
End Function
Private Sub ShowBtn_Click()

    Dim Msg As String
    
    If IsNull(Me.DateFrom.value) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "«·—Ã«¡ «Œ Ì«—  «—ÌŒ »œ«Ì… «·› —… "
        Else
            Msg = "Please select the starting date"
        End If
         MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         Exit Sub
    End If
    If IsNull(Me.DateTo.value) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "«·—Ã«¡ «Œ Ì«—  «—ÌŒ ‰Â«Ì… «·› —… "
        Else
            Msg = "Please select the ending date"
        End If
         MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         Exit Sub
    End If
Screen.MousePointer = vbArrowHourglass
EleHeader.backcolor = &HC0C0FF
ShowBtn.Caption = "Ã«—Ì «· ÕÌ·..."
DoEvents
    Dim Vat As Double
    Dim netvalue As Double
    Dim advancedPayment As Double


   Dim advvat As Double
   Dim AdvPayment As Double
   
  GetVAT_SalesAdvancedPayment advvat, , , , AdvPayment
  tztAdvBill = AdvPayment
  
  
  
 GetVAT_Notes Vat, netvalue, 9083
TxtDept(0).Text = netvalue
TxtDept(1).Text = Vat * -1

''////////
 GetVAT_Notes Vat, netvalue, 9082
TxtDept(2).Text = netvalue
TxtDept(3).Text = Vat * -1



 GetVAT_Notes Vat, netvalue, 9089
 
TxtVatADD.Text = netvalue


 GetVAT_Notes Vat, netvalue, 9090
 
TxtVatDis.Text = netvalue

TxtCorrect1.Text = val(TxtVatADD.Text) - val(TxtVatDis.Text)
TxtOldVat = 0

Dim Account_code4 As String
Dim Balance As String
Account_code4 = get_account_code_branch(145, 0)
If Check1.value = vbChecked Then
WriteCustomerBalPublic Account_code4, Balance, , , , , , , DTPicker1.value, 1
If Balance > 0 Then
      TxtOldVat.Text = Balance * -1
 Else
 TxtOldVat.Text = 0
 End If
Else
TxtOldVat.Text = 0

End If

'************************************ﬁÌœ «· ”ÊÌ… «·ÌœÊÌ
  GetVAT_ManualNotes Vat, netvalue, 0, 57
TxtDept(7).Text = netvalue
TxtDept(6).Text = Vat

 GetVAT_ManualNotes Vat, netvalue, 1, 57
TxtDept(5).Text = netvalue
TxtDept(4).Text = Vat

 '************************************ﬁÌœ «· ”ÊÌ… «·ÌœÊÌ
'******************************* Õ·Ì· «·„’—Ê›« 
 GetVAT_Expenses Vat, netvalue, 3
txtExpenses(0) = netvalue
txtExpensesVat(0) = Vat


GetVAT_Expenses Vat, netvalue, 80
txtExpenses(1) = netvalue
txtExpensesVat(1) = Vat



GetVAT_Customs Vat
txtExpensesVat(2).Text = 0 ' Vat
txtExpenses(2).Text = 0 ' Vat / 0.05


'salimmmmmmmmmmmmm



If Format(DateFrom.value, "yyyy-mm-dd") >= Format("01/01/2018", "yyyy-mm-dd") And Format(DateFrom.value, "yyyy-mm-dd") <= Format("30/06/2020", "yyyy-mm-dd") Then
intervalVat = 0.05
ElseIf Format(DateFrom.value, "yyyy-mm-dd") >= Format("01/07/2020", "yyyy-mm-dd") And Format(DateFrom.value, "yyyy-mm-dd") <= Format("31/12/2030", "yyyy-mm-dd") Then
intervalVat = 0.15
 
Else
intervalVat = 1
End If


VATPurchases2.Text = Round(Vat, 2)
Purchases2.Text = Round(Vat / intervalVat, 2)




PreVAT Vat, netvalue
 
txtExpenses(3).Text = Round(netvalue, 2) ' '  ··œ›⁄«  «·„ﬁœ„…
txtExpensesVat(3).Text = Round(Vat, 2) '«·›«  ··œ›⁄«  «·„ﬁœ„…


TransferVAT Vat, netvalue
 
txtExpenses(4).Text = Round(Vat, 2) ' '   ﬁÌ„… «·ÕÊ«·« 
txtExpensesVat(4).Text = Round(netvalue, 2)  '«·›«   «·ÕÊ«·« 


GetVAT_Expenses Vat, netvalue, 350 ' ’›Ì… «·⁄Âœ…
txtExpenses(5) = Round(netvalue, 2)
txtExpensesVat(5) = Round(Vat, 2)



NetCommisionVat netvalue, Vat, 350  '÷—Ì»Â ⁄„Ê·« 
txtExpenses(6) = Round(netvalue, 2)
txtExpensesVat(6) = Round(Vat, 2)



GetVAT_Expenses Vat, netvalue, 85
TxtServiceInvoice5 = netvalue
TxtServiceInvoice5Vat = Vat
TxtServiceInvoice5REt = 0


txtTotalExpenses = val(txtExpenses(0)) + val(txtExpenses(1)) + val(txtExpenses(2)) + val(txtExpenses(3)) + val(txtExpenses(4)) + val(txtExpenses(5)) + val(txtExpenses(6))
TxtTotalVatExpenses = val(txtExpensesVat(0)) + val(txtExpensesVat(1)) + val(txtExpensesVat(2)) + val(txtExpensesVat(3)) + val(txtExpensesVat(4)) + val(txtExpensesVat(5)) + val(txtExpensesVat(6))

Expenses = txtTotalExpenses
Expensesvat = TxtTotalVatExpenses

'******************************* Õ·Ì· «·„’—Ê›« 

'TxtContractVAT.Text = vat

ClaCulte
'tab 5%
'›Ê« Ì— „»Ì⁄«  5%
If SystemOptions.AllItemInVAT = True Then
    SalesT5.Text = getTotalVATItemsValue222(21, True, 5)
Else
    SalesT5 = getTotalVATItemsValue(21, True, 5)
    Text1(0).Text = getTotalVATItemsValue(21, False, 66)
    Me.Text2.Text = getTotalVATItemsValue(9, False, 66)
    
    Sales5 = Text1(0).Text
    RSales5 = Me.Text2.Text

End If
        
 '⁄ﬁœ «ÌÃ«— 5%
GetVAT_Contract Vat, netvalue, 5, 1
TxtContractVaue.Text = netvalue
TxtContractVAT.Text = Vat
 
 '⁄ﬁœ «ÌÃ«— „⁄›Ì%
GetVAT_Contract Vat, netvalue, 5, 0
Text1(1).Text = netvalue
 
GetVAT_FABuy Vat, netvalue, 5, 0
txtFaBuy = netvalue
txtFaBuy3 = Vat


'„‘«—Ì⁄ 5 „” Œ·’«  ··⁄„Ì·%
GetVAT_Project Vat, netvalue, 0, 5, advancedPayment
TxtProjCusValue.Text = netvalue
TxtProjCusVAT.Text = Vat


GetVAT_ProjectAdvancedPayment Vat, netvalue, 0, 5, advancedPayment, 1

TxtProjCusReValue = advancedPayment


GetVAT_ProjectAdvancedPayment Vat, netvalue, 1, 5, advancedPayment, 1

 TxtProjConReValue = advancedPayment

'⁄„—… 5%
GetVAT_Omrah Vat, netvalue
TxtOmraValue.Text = netvalue
TxtOmraVAT.Text = Vat
'ÕÃ 5%
GetVAT_Haijj Vat, netvalue
TxtHajjValue.Text = netvalue
TxtHajjVAT.Text = Vat
'Ê“«—… 5 %
GetVAT_Minis Vat, netvalue, 5
TxtMinisterValue.Text = netvalue
TxtMinisterVAT.Text = Vat

'·„  ›⁄· Õ Ì «·«‰
TxtMaintCarValue.Text = 0
'·«»œ „‰ ⁄„· Œ«‰«  «·«œŒ«· «·ÌœÊÌ
'«·Õ—ﬂ«  «·ÌœÊÌ… 5%
manulaEntey5.Text = getManuTotalVATItemsValue(21, 5)
'ﬂ· «·„»Ì⁄«  5 %


'tab 5 % Returns
'«œŒ·«  „—œÊœ«  ÌœÊÌ…
manulaEnteyRet5 = getManuTotalVATItemsValue(9, 5)
' ›Ê« Ì— „—œÊœ« 5%
  If SystemOptions.AllItemInVAT = True Then
    
 SalesRet5 = getTotalVATItemsValue222(9, True, 5)
  Else
   SalesRet5 = getTotalVATItemsValue(9, True, 5)
  End If

 
    TxtContractReVaue = 0
   
    TxtOmraReValue = 0
    TxtHajjReValue = 0
    TxtMinisterReValue = 0
    TxtMaintCarReValue = 0
    
'tab  5 % purcahse
'„‘ —Ì«  ÌœÊÌ… 5%
txtManulaEntryP5 = getManuTotalVATItemsValue(5, 5)
'„‘ —Ì«  5 %
PurchasesT5 = getTotalVATItemsValue(22, True, 5)
'„—œÊœ«  ÌœÊÌ… 5%
    txtManulaEntryP5Ret = getManuTotalVATItemsValue(5, 5)
'„—œÊœ«  5%
    PurchasesRett5 = getTotalVATItemsValue(5, True, 5)
    
    '5%„‘ —Ì« 
    '*****************************************
    Purchasest5vat = 1
    txtManulaEntryP5Vat = 1
    
    
    '*****************************************
    
    

  
  
  

  txtManulaEntryP5Vat = (getManuVATItemsValue(22) - getManuVATItemsValue(5))
  Purchasest5vat = (getItemsVATValue(22) - getItemsVATValue(5))
  
 GetVAT_Project Vat, netvalue, 1, 5
TxtProjConValue.Text = netvalue
TxtProjConVAT.Text = Vat
 
 GetVAT_Project Vat, netvalue, 1, 0
Txtprojectsupp.Text = netvalue
 
  
GetVAT_Hand Vat, netvalue, 0
 TxtMaintCarValue = netvalue
 TxtMaintCarVAT = Vat
 
 GetVAT_Hand Vat, netvalue, 2
 TxtMaintCarValue1 = netvalue
 TxtMaintCarReValue1 = Vat
 
 
  GetVAT_Hand Vat, netvalue, 1
 TxtMaintCarValue2 = netvalue
 TxtMaintCarReValue2 = Vat
 
 
 GetVAT_ReqCon Vat, netvalue
TxtReqConValue.Text = netvalue
TxtReqConVAT.Text = Vat


GetVAT_Expenses Vat, netvalue, 802
TxtAssestValue = netvalue
TxtAssestVAT = Vat

 GetVAT_PREVAT1 Vat, netvalue, 0
TxtPReVatTotal5V.Text = Round(netvalue, 2) 'supplier
TxtPReVatVAT5V.Text = Round(Vat, 2)



TxtTotalReceValue.Text = val(TxtDept(7).Text) + val(TxtProjConValue.Text) + val(TxtReqConValue.Text) + val(TxtAssestValue.Text) + val(txtManulaEntryP5) + val(PurchasesT5) + val(Expenses) + val(TxtPReVatTotal5V.Text)
TxtTotalReceVAT.Text = val(TxtDept(6).Text) + val(TxtDept(3).Text) + val(TxtProjConVAT.Text) + val(TxtReqConVAT.Text) + val(TxtAssestVAT.Text) + val(txtManulaEntryP5Vat) + val(Purchasest5vat) + val(0) + val(Expensesvat) + val(TxtPReVatVAT5V.Text)
TxtTotalReceReValue = val(TxtDept(2).Text) + val(TxtProjConReValue) + val(TxtReqConReValue) + val(TxtAssestReValue) + val(txtManulaEntryP5Ret) + val(0) + val(PurchasesRett5)
      
      
  '  VATPurchases1.Text = (getItemsVATValue(22) - getItemsVATValue(5)) + (getManuVATItemsValue(22) - getManuVATItemsValue(5)) + val(TxtTotalReceVAT.Text)
   
    '«Ã„«·Ì „‘ —Ì«  5 %
   Purchases1.Text = val(TxtTotalReceValue.Text)

'«Ã„«·Ì „—œÊœ«  5%
    RPurchases1.Text = TxtTotalReceReValue
  '«Ã„«·Ì vat 5% ··„‘ —Ì« 
   
   
  VATPurchases1.Text = Round(val(TxtTotalReceVAT), 2)
  

'tab2 Zeros

'›Ê« Ì— „»Ì⁄«  0%

manulaSAlesZero.Text = getManuTotalVATItemsValue(21, 0)
SalesZero = getTotalVATItemsValue(21, True, 0)

       If SystemOptions.PriceWithVAT = True Then
      SalesZero.Text = 0
   End If
   
   
   
 GetVAT_Minis Vat, netvalue, 0
TxtMinisterValuez.Text = netvalue



   
   
' ›Ê« Ì— „—œÊœ« 0%
  manulaSAlesZeroRet.Text = getManuTotalVATItemsValue(9, 0)
    SalesRetZero = getTotalVATItemsValue(9, False, 0)
    
      If SystemOptions.PriceWithVAT = True Then
      SalesRetZero.Text = 0
   End If
   

    
manulaEntey5Vat = (getManuVATItemsValue(21) - getManuVATItemsValue(9))
If SystemOptions.AllItemInVAT = True Then
   SalesTVAT = (getItemsVATValue222(21) - getItemsVATValue222(9))
  Else
   SalesTVAT = (getItemsVATValue(21) - getItemsVATValue(9))
End If
   If SystemOptions.PriceWithVAT = True Then
   SalesTVAT = (SalesT5 - SalesRet5) * (intervalVat)
   End If
   SalesTVAT = SalesTVAT - advvat
   '„‘«—Ì⁄ 0 „” Œ·’«  ··⁄„Ì·%
GetVAT_Project Vat, netvalue, 0, 0
TxtProjCusValuezero.Text = netvalue
 TxtProjCusValueRetzero.Text = 0
 

' Ã„Ì⁄ ﬂ· «·„»Ì⁄«  «·’›—Ì
TotalSalesZero = val(manulaSAlesZero) + val(SalesZero) + val(TxtMinisterValuez) + val(TxtProjCusValuezero.Text)
Sales3.Text = TotalSalesZero
' Ã„Ì⁄ ﬂ· «·„—œÊœ«  «·’›—Ì
TotalRetSalesZero = val(manulaSAlesZeroRet) + val(SalesRetZero) + val(TxtMinisterReValuez)
      RSales3.Text = TotalRetSalesZero
    'VATSales3.Text = 0
    
    

 GetVAT_Transportations Vat, netvalue, 0
transport5.Text = Round(netvalue, 2)
transport5vat.Text = Round(Vat, 2)

 GetVAT_PREVAT Vat, netvalue, 0
TxtPReVatTotal5.Text = Round(netvalue, 2) 'customer
TxtPReVatVAT5.Text = Round(Vat, 2)


    ' Ã„Ì⁄ ﬂ· ‰”»Â «·5% ··„»Ì⁄« 
TxtTotalPayValue.Text = val(TxtDept(5).Text) + val(txtFaBuy) + val(TxtContractVaue.Text) + val(TxtProjCusValue.Text) + val(TxtOmraValue.Text) + val(TxtHajjValue.Text) + val(TxtMinisterValue.Text) + val(TxtMaintCarValue.Text) + val(SalesT5.Text) + val(manulaEntey5.Text) + val(TxtServiceInvoice5) + val(transport5.Text) + val(TxtPReVatTotal5)
TxtTotalPayVAT.Text = val(TxtDept(4).Text) + val(txtFaBuy3) + val(TxtDept(1).Text) + val(TxtContractVAT.Text) + val(TxtProjCusVAT.Text) + val(TxtOmraVAT.Text) + val(TxtHajjVAT.Text) + val(TxtMinisterVAT.Text) + val(TxtMaintCarVAT.Text) + val(SalesTVAT.Text) + val(manulaEntey5Vat) + val(TxtServiceInvoice5Vat) + val(transport5vat.Text) + val(TxtPReVatVAT5)

TxtTotalRePayValue.Text = val(TxtDept(0).Text) + val(SalesRet5) + val(TxtContractReVaue) + val(TxtProjCusReValue) + val(TxtOmraReValue) + val(TxtHajjReValue) + val(TxtMinisterReValue) + val(TxtMaintCarReValue) + val(manulaEnteyRet5) + val(transport5re.Text)

    '„»Ì⁄« 5%
Sales1.Text = val(TxtTotalPayValue.Text)   ' «Ã„«Ì·Ì Õ—ﬂ«   5 %€Ì— «·»Ì⁄
     'ﬂ· «·„—œÊœ«  5 %
 RSales1.Text = Round(val(TxtTotalRePayValue), 2)
 
     '„⁄·ﬁ «·«‰
    VATSales1.Text = Round(val(TxtTotalPayVAT.Text), 2)   ' «Ã„«Ì·Ì Õ—ﬂ«  €Ì— «·»Ì⁄
    

    
    

  '«·„»Ì⁄«  «·„⁄›«…
    If SystemOptions.AllItemInVAT = True Then
 Sales5.Text = 0
 RSales5.Text = 0
    Else
    If Not SystemOptions.PriceWithVAT Then
        Sales5.Text = getTotalVATItemsValue(21, False) + val(Text1(1).Text) '„÷«› «·”ﬂ‰
        RSales5.Text = getTotalVATItemsValue(9, False)
    End If
    End If
    
    
'       If SystemOptions.PriceWithVAT = True Then
'        Sales5.Text = 0
'      RSales5.Text = 0
'   End If
   
   
   
    'VATSales5.Text = 0
    
  '  TxtSalesVAT.Text = getItemsVATValue(21) + getManuVATItemsValue(21)
  '  TxtRetSalesVAT.Text = getItemsVATValue(9) + getManuVATItemsValue(9)


    SalesTotal.Text = val(Sales1.Text) + val(Sales2.Text) + val(Sales3.Text) + val(Sales4.Text) + IIf(ChkIsFree.value, 0, val(Sales5.Text))
    RSalesTotal.Text = val(RSales1.Text) + val(RSales2.Text) + val(RSales3.Text) + val(RSales4.Text) + IIf(ChkIsFree.value, 0, val(RSales5.Text))
    RSalesTotal.Text = Round(RSalesTotal.Text, 2)
    
    VATSalesTotal.Text = val(VATSales1.Text) + val(0) + val(0) + val(0) + val(0)
    
 
 
    txtmanulPurcahsezero = getManuTotalVATItemsValue(22, 0)
     txtmanulPurcahsezeroRetur = getManuTotalVATItemsValue(5, 0)
     TxtPurchaseZero = getTotalVATItemsValue(22, False, 0)

     TxtPurchaseZeroRet = getTotalVATItemsValue(5, False, 0)
     
          If SystemOptions.PriceWithVAT = True Then
     TxtPurchaseZero = 0
     TxtPurchaseZeroRet = 0
     End If
     
    TotalPurchaseZero.Text = val(txtmanulPurcahsezero) + val(TxtPurchaseZero) + val(Txtprojectsupp)
    TotalReturnPurchaseZero.Text = val(txtmanulPurcahsezeroRetur) + val(TxtPurchaseZeroRet) + val(TxtprojectsuppRet)
      
      
    Purchases4.Text = val(txtmanulPurcahsezero) + val(TxtPurchaseZero) + val(Txtprojectsupp)
    RPurchases4.Text = val(txtmanulPurcahsezeroRetur) + val(TxtPurchaseZeroRet) + val(TxtprojectsuppRet)
    
    'VATPurchases4.Text = 0
    
    Purchases5.Text = getTotalVATItemsValue(22, False)
    RPurchases5.Text = getTotalVATItemsValue(5, False)
    
    
     PurcahseRemarks1.Text = getTotalVATItemsValue(22, False, , 0)
     PurcahseRemarks2.Text = getTotalVATItemsValue(22, False, , 1)
    PurcahseRemarks1Ret.Text = getTotalVATItemsValue(5, False, , 0)
    PurcahseRemarks2Ret.Text = getTotalVATItemsValue(5, False, , 1)
    
    
    
    TxtBillVstReverse.Text = getTotalVATItemsValue(22, False, , 2)
       Purchases3.Text = val(TxtBillVstReverse.Text)
    
     
    TxtBillVstReverseREt.Text = getTotalVATItemsValue(5, False, , 2)
    RPurchases3.Text = val(TxtBillVstReverseREt.Text)
    
 
     
    TxtRemarks1All.Text = val(PurcahseRemarks1) + val(PurcahseRemarks2) + val(TxtBillVstReverse)
    
        TxtRemarks1All2.Text = val(PurcahseRemarks1Ret) + val(PurcahseRemarks1Ret) + val(TxtBillVstReverseREt)
        
    'VATPurchases5.Text = 0
    
    TxtBuyVAT.Text = getItemsVATValue(22) + getManuVATItemsValue(22)
    TxtRetBuyVAT.Text = getItemsVATValue(5) + getManuVATItemsValue(5)
    

    
    PurchasesTotal.Text = val(Purchases1.Text) + val(Purchases2.Text) + val(Purchases3.Text) + val(Purchases4.Text) + val(Purchases5.Text)
    RPurchasesTotal.Text = val(RPurchases1.Text) + val(RPurchases2.Text) + val(RPurchases3.Text) + val(RPurchases4.Text) + val(RPurchases5.Text)
    VATPurchasesTotal.Text = val(VATPurchases1.Text) + val(VATPurchases2.Text) + val(VATPurchases3.Text) + val(0) + val(0)
    VATPurchasesTotal.Text = Round(VATPurchasesTotal.Text, 2)
    
    TotalNetTxt = 0
 '   If val(VATSalesTotal) = 0 Then
 '    TotalNetTxt.Text = Round(val(VATPurchasesTotal.Text), 2)
 '   Else
 '   TotalNetTxt.Text = Round(val(VATSalesTotal.Text) - val(VATPurchasesTotal.Text), 2)
 '   End If
    
    TotalNetTxt.Text = Round(val(VATSalesTotal.Text) - val(VATPurchasesTotal.Text), 2)
    CalnNet
Screen.MousePointer = vbDefault
EleHeader.backcolor = &HFFFFFF
    
    ShowBtn.Caption = "⁄—÷"
End Sub
