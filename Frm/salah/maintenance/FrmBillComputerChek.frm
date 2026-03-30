VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmBillComputerChek 
   BackColor       =   &H00E2E9E9&
   Caption         =   "›« Ê—… ›Õ’ ﬂ„»ÌÊ —"
   ClientHeight    =   10365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17295
   Icon            =   "FrmBillComputerChek.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10365
   ScaleWidth      =   17295
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   10365
      Index           =   2
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   17295
      _cx             =   30506
      _cy             =   18283
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
      Begin VB.TextBox oldtxtNoteSerial1 
         Height          =   240
         Left            =   15600
         TabIndex        =   80
         Top             =   690
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   240
         Left            =   15690
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   1305
         Visible         =   0   'False
         Width           =   1275
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   495
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   17295
         _cx             =   30506
         _cy             =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   24
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
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "›« Ê—… ›Õ’ ﬂ„»ÌÊ — "
         Align           =   1
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
            Left            =   1185
            TabIndex        =   2
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillComputerChek.frx":038A
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
            Left            =   120
            TabIndex        =   3
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillComputerChek.frx":0724
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
            Left            =   1710
            TabIndex        =   4
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillComputerChek.frx":0ABE
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
            Left            =   645
            TabIndex        =   5
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillComputerChek.frx":0E58
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin VB.Image ImgFavorites 
            Height          =   390
            Left            =   6960
            Picture         =   "FrmBillComputerChek.frx":11F2
            Stretch         =   -1  'True
            Top             =   0
            Width           =   525
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H000000FF&
            Height          =   555
            Index           =   27
            Left            =   2280
            TabIndex        =   6
            Top             =   0
            Width           =   2205
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   2220
         Index           =   0
         Left            =   -585
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   510
         Width           =   17685
         _cx             =   31194
         _cy             =   3916
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
         Begin VB.TextBox TxtClient 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   12315
            RightToLeft     =   -1  'True
            TabIndex        =   94
            Top             =   1140
            Width           =   3300
         End
         Begin VB.TextBox TxtMobile 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   6150
            RightToLeft     =   -1  'True
            TabIndex        =   93
            Top             =   1140
            Width           =   4530
         End
         Begin VB.TextBox TxtPlateNO 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   12315
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   1845
            Width           =   3300
         End
         Begin VB.ComboBox DcbYearFact 
            Height          =   315
            Left            =   6150
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   1500
            Width           =   4530
         End
         Begin VB.TextBox TxtClientCode 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   17910
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   2130
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   345
            Left            =   14550
            TabIndex        =   89
            Top             =   780
            Width           =   1065
         End
         Begin VB.CommandButton BtImage 
            Caption         =   " ÕœÌœ «·„·«ÕŸ«  "
            Height          =   375
            Left            =   930
            Picture         =   "FrmBillComputerChek.frx":4E5A
            TabIndex        =   88
            Top             =   360
            Width           =   1425
         End
         Begin VB.TextBox XPTxtID 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   13740
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   -270
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   345
            Left            =   2550
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   60
            Width           =   2010
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   14550
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   60
            Width           =   1065
         End
         Begin VB.TextBox TxtNoteID 
            Height          =   285
            Left            =   19845
            TabIndex        =   11
            Top             =   1350
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   14550
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   360
            Width           =   1065
         End
         Begin VB.ComboBox CboPayMentType 
            Height          =   315
            Left            =   6855
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   420
            Width           =   1995
         End
         Begin VB.TextBox TxtPaymentValue 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   4155
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   420
            Width           =   1605
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   12285
            TabIndex        =   15
            Top             =   30
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            Format          =   207618049
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmBillComputerChek.frx":8DA1
            Height          =   315
            Left            =   5460
            TabIndex        =   16
            Top             =   60
            Width           =   4470
            _ExtentX        =   7885
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
         Begin MSDataListLib.DataCombo DcboEmpName 
            Height          =   315
            Left            =   12345
            TabIndex        =   17
            Top             =   360
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   315
            Index           =   10
            Left            =   2550
            TabIndex        =   18
            Top             =   420
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… «·ﬁÌœ"
            BackColor       =   14871017
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
         Begin MSDataListLib.DataCombo DcbCarType 
            Bindings        =   "FrmBillComputerChek.frx":8DB6
            Height          =   315
            Left            =   1305
            TabIndex        =   95
            Top             =   1140
            Width           =   2955
            _ExtentX        =   5212
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
         Begin MSDataListLib.DataCombo DcbModel 
            Bindings        =   "FrmBillComputerChek.frx":8DCB
            Height          =   315
            Left            =   12315
            TabIndex        =   96
            Top             =   1500
            Width           =   3300
            _ExtentX        =   5821
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
         Begin MSComCtl2.DTPicker DtEnd 
            Height          =   315
            Left            =   1305
            TabIndex        =   97
            Top             =   1845
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   556
            _Version        =   393216
            Format          =   207618049
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker DTStart 
            Height          =   360
            Left            =   6150
            TabIndex        =   98
            Top             =   1845
            Width           =   4530
            _ExtentX        =   7990
            _ExtentY        =   635
            _Version        =   393216
            Format          =   207618049
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcbColor 
            Bindings        =   "FrmBillComputerChek.frx":8DE0
            Height          =   315
            Left            =   1305
            TabIndex        =   99
            Top             =   1500
            Width           =   2955
            _ExtentX        =   5212
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
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   6150
            TabIndex        =   100
            Top             =   780
            Width           =   8400
            _ExtentX        =   14817
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            BoundColumn     =   ""
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboBoxStor 
            Height          =   315
            Left            =   1305
            TabIndex        =   101
            Top             =   780
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì· «·‰ﬁœÌ"
            Height          =   285
            Index           =   2
            Left            =   15180
            TabIndex        =   112
            Top             =   1140
            Width           =   2010
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÃÊ«· «·⁄„Ì·"
            Height          =   225
            Index           =   3
            Left            =   10590
            TabIndex        =   111
            Top             =   1140
            Width           =   1530
         End
         Begin VB.Label lblcar 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·„⁄œÂ/«·”Ì«—…"
            Height          =   330
            Left            =   4185
            RightToLeft     =   -1  'True
            TabIndex        =   110
            Top             =   1140
            Width           =   1815
         End
         Begin VB.Label lblmodel 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÿ—«“ «·„⁄œÂ/«·”Ì«—…"
            Height          =   240
            Left            =   15405
            RightToLeft     =   -1  'True
            TabIndex        =   109
            Top             =   1500
            Width           =   1785
         End
         Begin VB.Label lblyear 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "”‰… «·’‰⁄"
            Height          =   195
            Left            =   10575
            RightToLeft     =   -1  'True
            TabIndex        =   108
            Top             =   1530
            Width           =   1545
         End
         Begin VB.Label LblColor 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "·Ê‰ «·„⁄œÂ/«·”Ì«—…"
            Height          =   300
            Left            =   4185
            RightToLeft     =   -1  'True
            TabIndex        =   107
            Top             =   1500
            Width           =   1815
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " √—ÌŒ «·Œ—ÊÃ"
            Height          =   330
            Index           =   9
            Left            =   4365
            TabIndex        =   106
            Top             =   1845
            Width           =   1635
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «··ÊÕ…"
            Height          =   210
            Index           =   5
            Left            =   15570
            TabIndex        =   105
            Top             =   1935
            Width           =   1620
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·œŒÊ·"
            Height          =   330
            Index           =   10
            Left            =   10680
            TabIndex        =   104
            Top             =   1905
            Width           =   1575
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì·"
            Height          =   270
            Index           =   14
            Left            =   15060
            TabIndex        =   103
            Top             =   825
            Width           =   2130
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’‰œÊﬁ"
            Height          =   240
            Index           =   16
            Left            =   4185
            TabIndex        =   102
            Top             =   855
            Width           =   1815
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   1
            Left            =   13845
            TabIndex        =   25
            Top             =   60
            Width           =   690
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·›« Ê—…"
            Height          =   285
            Index           =   4
            Left            =   16050
            TabIndex        =   24
            Top             =   60
            Width           =   1140
         End
         Begin VB.Label lblBr 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            Height          =   255
            Left            =   9870
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   60
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Â‰œ”"
            Height          =   285
            Index           =   15
            Left            =   16050
            TabIndex        =   22
            Top             =   375
            Width           =   1140
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   255
            Left            =   4425
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   60
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·”œ«œ"
            Height          =   255
            Index           =   13
            Left            =   8640
            TabIndex        =   20
            Top             =   465
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„œ›Ê⁄"
            Height          =   315
            Index           =   19
            Left            =   5625
            TabIndex        =   19
            Top             =   420
            Width           =   885
         End
         Begin VB.Image img 
            Height          =   855
            Left            =   2400
            Picture         =   "FrmBillComputerChek.frx":8DF5
            Stretch         =   -1  'True
            Top             =   30
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.Image imgnul 
            Height          =   1095
            Left            =   2400
            Top             =   30
            Width           =   810
         End
      End
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   6795
         Left            =   -15
         TabIndex        =   26
         Top             =   2640
         Width           =   17085
         _cx             =   30136
         _cy             =   11986
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
         ForeColor       =   0
         FrontTabColor   =   14871017
         BackTabColor    =   12648447
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   16711680
         Caption         =   "»Ì«‰«  «·›Õ’|Õ«·Â «·«⁄ „«œ"
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
         Picture(0)      =   "FrmBillComputerChek.frx":93A3
         Flags(1)        =   2
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   6330
            Left            =   17730
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   45
            Width           =   16995
            _cx             =   29977
            _cy             =   11165
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
            Begin VSFlex8UCtl.VSFlexGrid GRID2 
               Height          =   3630
               Left            =   120
               TabIndex        =   28
               Tag             =   "1"
               Top             =   240
               Width           =   13230
               _cx             =   23336
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
               AllowBigSelection=   -1  'True
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   3
               Cols            =   8
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmBillComputerChek.frx":973D
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
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
               Height          =   255
               Left            =   9000
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   4080
               Width           =   3375
            End
            Begin VB.Label Label1100 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
               Height          =   255
               Left            =   9960
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   4560
               Width           =   3375
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6330
            Index           =   15
            Left            =   45
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   45
            Width           =   16995
            _cx             =   29977
            _cy             =   11165
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial (Arabic)"
               Size            =   12
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
            BackColor       =   14871017
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   8
            BorderWidth     =   1
            ChildSpacing    =   1
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
            GridRows        =   1
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmBillComputerChek.frx":9889
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   6330
               Index           =   16
               Left            =   0
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   0
               Width           =   16995
               _cx             =   29977
               _cy             =   11165
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
               Appearance      =   5
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
               Begin VB.Frame gimage 
                  BackColor       =   &H80000005&
                  Height          =   6750
                  Left            =   585
                  TabIndex        =   33
                  Top             =   -15
                  Visible         =   0   'False
                  Width           =   11235
                  Begin VB.CommandButton bClose 
                     BackColor       =   &H000000FF&
                     Caption         =   "X"
                     Height          =   375
                     Left            =   9360
                     Style           =   1  'Graphical
                     TabIndex        =   34
                     Top             =   120
                     Width           =   375
                  End
                  Begin VB.Image img8 
                     Height          =   615
                     Left            =   8940
                     Picture         =   "FrmBillComputerChek.frx":98BF
                     Stretch         =   -1  'True
                     Top             =   4500
                     Width           =   705
                  End
                  Begin VB.Image img14 
                     Height          =   615
                     Left            =   750
                     Picture         =   "FrmBillComputerChek.frx":9E6D
                     Stretch         =   -1  'True
                     Top             =   4260
                     Width           =   705
                  End
                  Begin VB.Image img12 
                     Height          =   615
                     Left            =   3750
                     Picture         =   "FrmBillComputerChek.frx":A41B
                     Stretch         =   -1  'True
                     Top             =   4260
                     Width           =   705
                  End
                  Begin VB.Image img11 
                     Height          =   615
                     Left            =   5310
                     Picture         =   "FrmBillComputerChek.frx":A9C9
                     Stretch         =   -1  'True
                     Top             =   4260
                     Width           =   705
                  End
                  Begin VB.Image img10 
                     Height          =   615
                     Left            =   7110
                     Picture         =   "FrmBillComputerChek.frx":AF77
                     Stretch         =   -1  'True
                     Top             =   4380
                     Width           =   705
                  End
                  Begin VB.Image img9 
                     Height          =   615
                     Left            =   8070
                     Picture         =   "FrmBillComputerChek.frx":B525
                     Stretch         =   -1  'True
                     Top             =   4140
                     Width           =   705
                  End
                  Begin VB.Image img13 
                     Height          =   615
                     Left            =   2310
                     Picture         =   "FrmBillComputerChek.frx":BAD3
                     Stretch         =   -1  'True
                     Top             =   4110
                     Width           =   705
                  End
                  Begin VB.Image imag4 
                     Height          =   615
                     Left            =   5700
                     Picture         =   "FrmBillComputerChek.frx":C081
                     Stretch         =   -1  'True
                     Top             =   1290
                     Width           =   705
                  End
                  Begin VB.Image img7 
                     Height          =   615
                     Left            =   780
                     Picture         =   "FrmBillComputerChek.frx":C62F
                     Stretch         =   -1  'True
                     Top             =   1530
                     Width           =   705
                  End
                  Begin VB.Image img6 
                     Height          =   615
                     Left            =   2580
                     Picture         =   "FrmBillComputerChek.frx":CBDD
                     Stretch         =   -1  'True
                     Top             =   1650
                     Width           =   705
                  End
                  Begin VB.Image imag5 
                     Height          =   615
                     Left            =   4260
                     Picture         =   "FrmBillComputerChek.frx":D18B
                     Stretch         =   -1  'True
                     Top             =   1530
                     Width           =   705
                  End
                  Begin VB.Image imag3 
                     Height          =   615
                     Left            =   7140
                     Picture         =   "FrmBillComputerChek.frx":D739
                     Stretch         =   -1  'True
                     Top             =   1770
                     Width           =   705
                  End
                  Begin VB.Image imag2 
                     Height          =   615
                     Left            =   7980
                     Picture         =   "FrmBillComputerChek.frx":DCE7
                     Stretch         =   -1  'True
                     Top             =   1530
                     Width           =   705
                  End
                  Begin VB.Image imag1 
                     Height          =   615
                     Left            =   8820
                     Picture         =   "FrmBillComputerChek.frx":E295
                     Stretch         =   -1  'True
                     Top             =   1770
                     Width           =   705
                  End
                  Begin VB.Shape Shape5 
                     BorderColor     =   &H00FF0000&
                     BorderWidth     =   5
                     FillColor       =   &H000000FF&
                     Height          =   612
                     Left            =   8880
                     Top             =   4440
                     Width           =   732
                  End
                  Begin VB.Shape Shape6 
                     BorderColor     =   &H00FF0000&
                     BorderWidth     =   5
                     FillColor       =   &H000000FF&
                     Height          =   612
                     Left            =   8040
                     Top             =   4200
                     Width           =   732
                  End
                  Begin VB.Shape Shape7 
                     BorderColor     =   &H00FF0000&
                     BorderWidth     =   5
                     FillColor       =   &H000000FF&
                     Height          =   612
                     Left            =   7080
                     Top             =   4440
                     Width           =   732
                  End
                  Begin VB.Shape Shape8 
                     BorderColor     =   &H00FF0000&
                     BorderWidth     =   5
                     FillColor       =   &H000000FF&
                     Height          =   612
                     Left            =   720
                     Top             =   4320
                     Width           =   732
                  End
                  Begin VB.Shape Shape13 
                     BorderColor     =   &H00FF0000&
                     BorderWidth     =   5
                     FillColor       =   &H000000FF&
                     Height          =   612
                     Left            =   5280
                     Top             =   4320
                     Width           =   732
                  End
                  Begin VB.Shape Shape14 
                     BorderColor     =   &H00FF0000&
                     BorderWidth     =   5
                     FillColor       =   &H000000FF&
                     Height          =   612
                     Left            =   3720
                     Top             =   4320
                     Width           =   732
                  End
                  Begin VB.Shape Shape1 
                     BorderColor     =   &H00FF0000&
                     BorderWidth     =   5
                     FillColor       =   &H000000FF&
                     Height          =   612
                     Left            =   8760
                     Top             =   1800
                     Width           =   732
                  End
                  Begin VB.Shape Shape2 
                     BorderColor     =   &H00FF0000&
                     BorderWidth     =   5
                     FillColor       =   &H000000FF&
                     Height          =   612
                     Index           =   1
                     Left            =   7920
                     Top             =   1560
                     Width           =   732
                  End
                  Begin VB.Shape Shape3 
                     BorderColor     =   &H00FF0000&
                     BorderWidth     =   5
                     FillColor       =   &H000000FF&
                     Height          =   612
                     Index           =   1
                     Left            =   7080
                     Top             =   1800
                     Width           =   732
                  End
                  Begin VB.Shape Shape4 
                     BorderColor     =   &H00FF0000&
                     BorderWidth     =   5
                     FillColor       =   &H000000FF&
                     Height          =   612
                     Left            =   5640
                     Top             =   1320
                     Width           =   732
                  End
                  Begin VB.Shape Shape9 
                     BorderColor     =   &H00FF0000&
                     BorderWidth     =   5
                     FillColor       =   &H000000FF&
                     Height          =   612
                     Left            =   4200
                     Top             =   1560
                     Width           =   732
                  End
                  Begin VB.Shape Shape10 
                     BorderColor     =   &H00FF0000&
                     BorderWidth     =   5
                     FillColor       =   &H000000FF&
                     Height          =   612
                     Left            =   2520
                     Top             =   1680
                     Width           =   732
                  End
                  Begin VB.Shape Shape11 
                     BorderColor     =   &H00FF0000&
                     BorderWidth     =   5
                     FillColor       =   &H000000FF&
                     Height          =   612
                     Left            =   720
                     Top             =   1560
                     Width           =   732
                  End
                  Begin VB.Shape Shape12 
                     BorderColor     =   &H00FF0000&
                     BorderWidth     =   5
                     FillColor       =   &H000000FF&
                     Height          =   615
                     Left            =   2250
                     Top             =   4140
                     Width           =   735
                  End
                  Begin VB.Image Image6 
                     Height          =   5775
                     Left            =   60
                     Picture         =   "FrmBillComputerChek.frx":E843
                     Stretch         =   -1  'True
                     Top             =   300
                     Width           =   9735
                  End
               End
               Begin ImpulseButton.ISButton Accredit 
                  Height          =   555
                  Left            =   2925
                  TabIndex        =   35
                  Top             =   6525
                  Width           =   2670
                  _ExtentX        =   4710
                  _ExtentY        =   979
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "«—”«· ··«⁄ „«œ"
                  BackColor       =   14871017
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
                  Height          =   285
                  Index           =   8
                  Left            =   15840
                  TabIndex        =   36
                  Top             =   5925
                  Width           =   825
                  _ExtentX        =   1455
                  _ExtentY        =   503
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
                  ButtonImage     =   "FrmBillComputerChek.frx":2BF93
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   315
                  Index           =   12
                  Left            =   14520
                  TabIndex        =   37
                  Top             =   5925
                  Width           =   1125
                  _ExtentX        =   1984
                  _ExtentY        =   556
                  ButtonPositionImage=   1
                  Caption         =   "≈œ—«Ã"
                  BackColor       =   14871017
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
               Begin VSFlex8Ctl.VSFlexGrid fg 
                  Height          =   5760
                  Left            =   330
                  TabIndex        =   38
                  Top             =   90
                  Width           =   16440
                  _cx             =   28998
                  _cy             =   10160
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
                  Cols            =   7
                  FixedRows       =   1
                  FixedCols       =   0
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmBillComputerChek.frx":2C52D
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
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„ »ﬁÌ"
                  Height          =   435
                  Index           =   18
                  Left            =   2715
                  TabIndex        =   43
                  Top             =   5835
                  Visible         =   0   'False
                  Width           =   1290
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Height          =   435
                  Index           =   17
                  Left            =   525
                  TabIndex        =   42
                  Top             =   5835
                  Visible         =   0   'False
                  Width           =   1275
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Height          =   345
                  Index           =   12
                  Left            =   10860
                  TabIndex        =   41
                  Top             =   5805
                  Visible         =   0   'False
                  Width           =   1350
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«Ã„«·Ì"
                  Height          =   450
                  Index           =   11
                  Left            =   11940
                  TabIndex        =   40
                  Top             =   5805
                  Visible         =   0   'False
                  Width           =   1590
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   3720
                  Index           =   62
                  Left            =   3330
                  RightToLeft     =   -1  'True
                  TabIndex        =   39
                  Top             =   1965
                  Width           =   675
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   6300
               Index           =   9
               Left            =   15
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   15
               Width           =   16965
               _cx             =   29924
               _cy             =   11113
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
               Appearance      =   5
               MousePointer    =   0
               Version         =   801
               BackColor       =   14871017
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   ""
               Align           =   0
               AutoSizeChildren=   7
               BorderWidth     =   0
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
               Begin VB.TextBox Text8 
                  Alignment       =   1  'Right Justify
                  Height          =   4905
                  Left            =   4425
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   46
                  Top             =   1125
                  Width           =   720
               End
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì»… «·„»Ì⁄« "
                  Height          =   3225
                  Left            =   5460
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   1950
                  Width           =   1560
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   3225
                  Index           =   67
                  Left            =   3225
                  RightToLeft     =   -1  'True
                  TabIndex        =   49
                  Top             =   1950
                  Width           =   735
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   3195
                  Index           =   68
                  Left            =   5145
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   2310
                  Width           =   165
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "%"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   3720
                  Index           =   69
                  Left            =   3960
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   1950
                  Width           =   465
               End
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   825
         Index           =   1
         Left            =   0
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   9540
         Width           =   17295
         _cx             =   30506
         _cy             =   1455
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
         Align           =   2
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
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            Height          =   735
            Left            =   6480
            TabIndex        =   51
            Top             =   30
            Width           =   8865
            Begin VB.CommandButton menue 
               DownPicture     =   "FrmBillComputerChek.frx":2C61F
               Height          =   555
               Index           =   11
               Left            =   300
               Picture         =   "FrmBillComputerChek.frx":33951
               Style           =   1  'Graphical
               TabIndex        =   63
               Top             =   120
               Width           =   735
            End
            Begin VB.CommandButton menue 
               DownPicture     =   "FrmBillComputerChek.frx":344E5
               Height          =   555
               Index           =   10
               Left            =   6690
               Picture         =   "FrmBillComputerChek.frx":34ACC
               Style           =   1  'Graphical
               TabIndex        =   62
               Top             =   120
               Width           =   735
            End
            Begin VB.CommandButton menue 
               DownPicture     =   "FrmBillComputerChek.frx":350B3
               Height          =   555
               Index           =   9
               Left            =   4530
               Picture         =   "FrmBillComputerChek.frx":3C3E5
               Style           =   1  'Graphical
               TabIndex        =   61
               Top             =   120
               Width           =   735
            End
            Begin VB.CommandButton menue 
               Height          =   555
               Index           =   8
               Left            =   1650
               Picture         =   "FrmBillComputerChek.frx":3C905
               Style           =   1  'Graphical
               TabIndex        =   60
               Top             =   120
               Width           =   735
            End
            Begin VB.CommandButton menue 
               DownPicture     =   "FrmBillComputerChek.frx":3CDEA
               Height          =   555
               Index           =   7
               Left            =   3810
               Picture         =   "FrmBillComputerChek.frx":4411C
               Style           =   1  'Graphical
               TabIndex        =   59
               Top             =   120
               Width           =   735
            End
            Begin VB.CommandButton menue 
               DownPicture     =   "FrmBillComputerChek.frx":449AC
               Height          =   555
               Index           =   6
               Left            =   5970
               Picture         =   "FrmBillComputerChek.frx":4BCDE
               Style           =   1  'Graphical
               TabIndex        =   58
               Top             =   120
               Width           =   735
            End
            Begin VB.CommandButton menue 
               DownPicture     =   "FrmBillComputerChek.frx":4C17F
               Height          =   555
               Index           =   0
               Left            =   7410
               Picture         =   "FrmBillComputerChek.frx":534B1
               Style           =   1  'Graphical
               TabIndex        =   57
               Top             =   120
               Width           =   735
            End
            Begin VB.CommandButton menue 
               Height          =   555
               Index           =   1
               Left            =   3240
               Picture         =   "FrmBillComputerChek.frx":53A58
               Style           =   1  'Graphical
               TabIndex        =   56
               Top             =   840
               Width           =   735
            End
            Begin VB.CommandButton menue 
               Height          =   555
               Index           =   2
               Left            =   5250
               Picture         =   "FrmBillComputerChek.frx":53EF9
               Style           =   1  'Graphical
               TabIndex        =   55
               Top             =   120
               Width           =   735
            End
            Begin VB.CommandButton menue 
               Height          =   555
               Index           =   3
               Left            =   3090
               Picture         =   "FrmBillComputerChek.frx":543C9
               Style           =   1  'Graphical
               TabIndex        =   54
               Top             =   120
               Width           =   735
            End
            Begin VB.CommandButton menue 
               Height          =   555
               Index           =   4
               Left            =   2370
               Picture         =   "FrmBillComputerChek.frx":54882
               Style           =   1  'Graphical
               TabIndex        =   53
               Top             =   120
               Width           =   735
            End
            Begin VB.CommandButton menue 
               Height          =   555
               Index           =   5
               Left            =   1050
               Picture         =   "FrmBillComputerChek.frx":54DDA
               Style           =   1  'Graphical
               TabIndex        =   52
               Top             =   120
               Width           =   735
            End
            Begin ImpulseButton.ISButton CmdAttach 
               Height          =   540
               Left            =   8190
               TabIndex        =   64
               Top             =   135
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   953
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
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   540
            Index           =   3
            Left            =   30
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   30
            Width           =   6405
            _cx             =   11298
            _cy             =   953
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
               Height          =   375
               Index           =   0
               Left            =   5610
               TabIndex        =   66
               Top             =   105
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   661
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
               Height          =   375
               Index           =   1
               Left            =   4905
               TabIndex        =   67
               Top             =   105
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   661
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
               Height          =   375
               Index           =   2
               Left            =   4170
               TabIndex        =   68
               Top             =   105
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   661
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
               Height          =   375
               Index           =   3
               Left            =   3420
               TabIndex        =   69
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
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
               Height          =   375
               Index           =   4
               Left            =   2655
               TabIndex        =   70
               Top             =   105
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   661
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
               Height          =   375
               Index           =   6
               Left            =   0
               TabIndex        =   71
               Top             =   105
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   661
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
               Height          =   375
               Index           =   5
               Left            =   1890
               TabIndex        =   72
               Top             =   105
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   661
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
               Height          =   375
               Index           =   9
               Left            =   1260
               TabIndex        =   73
               Top             =   105
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   661
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
               Height          =   375
               Index           =   11
               Left            =   570
               TabIndex        =   74
               Top             =   105
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   661
               ButtonPositionImage=   1
               Caption         =   "«„— ›Õ’"
               BackColor       =   14871017
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
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   9990
            TabIndex        =   75
            Top             =   1080
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   315
            Index           =   6
            Left            =   1110
            TabIndex        =   78
            Top             =   420
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   315
            Index           =   7
            Left            =   3660
            TabIndex        =   77
            Top             =   300
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   270
            Index           =   8
            Left            =   12180
            TabIndex        =   76
            Top             =   1095
            Width           =   900
         End
      End
      Begin MSDataListLib.DataCombo DcboBox 
         Height          =   315
         Left            =   15405
         TabIndex        =   81
         Top             =   1560
         Visible         =   0   'False
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   315
         Index           =   7
         Left            =   15690
         TabIndex        =   82
         Top             =   225
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
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
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·Œ“‰…"
         Height          =   240
         Index           =   0
         Left            =   14535
         TabIndex        =   87
         Top             =   390
         Width           =   1020
      End
      Begin VB.Label XPTxtCurrent 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   240
         Left            =   1695
         TabIndex        =   86
         Top             =   4365
         Width           =   495
      End
      Begin VB.Label XPTxtCount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   240
         Left            =   0
         TabIndex        =   85
         Top             =   4365
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ﬁÌœ:"
         Height          =   270
         Index           =   30
         Left            =   14970
         RightToLeft     =   -1  'True
         TabIndex        =   84
         Top             =   0
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Â–… «·‘«‘…  ﬁÊ„ » ”ÃÌ· ÿ·» ”›… ‰ﬁœÌ… ÊÌ „ «Õ ”«» ﬁÌ„… «·œ›⁄ «·Ì«"
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
         Height          =   555
         Index           =   25
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   83
         Top             =   2535
         Width           =   5535
      End
   End
End
Attribute VB_Name = "FrmBillComputerChek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim TTD As clstooltipdemand
Dim Employee_account As String
Dim RsNotesGeneral As ADODB.Recordset
Dim general_noteid  As Long
Dim Ch As Integer

Private Sub Accredit_Click()
    Dim BeginTrans As Boolean

    Cn.BeginTrans
    BeginTrans = True

    If IsNull(rs("Posted")) Then
        rs("Posted") = user_id
        rs("PostedDate") = Time
    Else
        rs("Posted") = Null
       rs("PostedDate") = Time
    End If
   
    rs.update
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If

    Cn.CommitTrans
    BeginTrans = False
FillApprovedTable
    Retrive (val(Me.XPTxtID.text))
End Sub

Private Sub bClose_Click()

BtImage.Visible = True
gimage.Visible = False

End Sub

Private Sub BtImage_Click()
        BtImage.Visible = False
        gimage.Visible = True
End Sub

Private Sub CboPayMentType_Change()
    TxtPaymentValue.Enabled = False
    If val(Me.CboPayMentType.ListIndex) = 0 Then
        DcboBoxStor.Enabled = True
        TxtPaymentValue.Enabled = True
    Else
        TxtPaymentValue.text = 0
        DcboBoxStor.BoundText = ""
        DcboBoxStor.Enabled = False
    End If
    TxtPaymentValue.Enabled = True
End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub Cmd_Click(index As Integer)

    ' On Error GoTo ErrTrap
    Select Case index
        Case 10
            ShowGL_cc Me.TxtNoteSerial.text, , 200, val(Me.TxtNoteID.text)
        Case 12
           
            Dim Rs3 As ADODB.Recordset
            Set Rs3 = New ADODB.Recordset
            Dim sql As String
                
            sql = "SELECT TblComputerChek.Id, "
            sql = sql & "       TblComputerChek.name, "
            sql = sql & "       TblComputerChek.namee, "
            sql = sql & "       TblComputerChek.Manualcode, "
            sql = sql & "       TblComputerChek.MainId, "
            sql = sql & "       maintbl.name MainName, "
            sql = sql & "       maintbl.namee MainNamee "
            sql = sql & "FROM TblComputerChek "
            sql = sql & "    LEFT OUTER JOIN TblComputerChek maintbl "
            sql = sql & "        ON TblComputerChek.MainId = maintbl.Id "
            sql = sql & " where     "
            sql = sql & "  isnull(TblComputerChek.IsMain , 0 )  =  0 Order by TblComputerChek.Id,maintbl.Id"
         
            Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If Not Rs3.EOF Then
                Fg.Clear flexClearScrollable, flexClearEverything
                Fg.rows = 2
                Fg.Enabled = True
            End If
            ' fg.AddItem ""
          
            Do While Not Rs3.EOF
                Fg.TextMatrix(Fg.rows - 1, Fg.ColIndex("Code")) = Rs3!Manualcode & ""
                FG_AfterEdit Fg.rows - 1, Fg.ColIndex("Code")
                
                Rs3.MoveNext
               
            Loop
            Dim myrow    As Integer
            Dim Lastmain As String
            Dim crntMain As String
            For myrow = 1 To Fg.rows - 1
            
                crntMain = Fg.TextMatrix(myrow, Fg.ColIndex("MainType"))
                If crntMain = Lastmain Then
                    Fg.TextMatrix(myrow, Fg.ColIndex("MainType")) = ""
                Else
                       Lastmain = crntMain
                End If
          
            Next
        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
            Fg.Clear flexClearScrollable, flexClearEverything
            Fg.rows = 2
            Fg.Enabled = True
            ' lbl(20).Caption = "0"
            '   lbl(21).Caption = "0"
            'lbl(22).Caption = "0"
            'lbl(23).Caption = "0"
            
            GRID2.Clear flexClearScrollable, flexClearEverything
            GRID2.rows = 1
            Me.DCboUserName.BoundText = user_id
            '  TxtPaymentCounts.text = 1
            dcBranch.BoundText = Current_branch
            'XPDtbTrans.SetFocus
            
            Accredit.Enabled = True
            If SystemOptions.UserInterface = ArabicInterface Then
                Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
            Else
                Accredit.Caption = " send to Approval   "
            End If
                                               
        Case 1
            If ChekClodePeriod(XPDtbTrans.value) = True Then
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
            Fg.rows = Fg.rows + 1
            Fg.Enabled = True
            TxtModFlg.text = "E"
            Me.DCboUserName.BoundText = user_id

        Case 2
        
            If CboPayMentType.ListIndex = 0 Then 'cash
          
                If (DcboBoxStor.BoundText) = "" Then
                
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·«»œ „‰  ÕœÌœ ’‰œÊﬁ"
                    Else
                        MsgBox "Please Define Box"
                    End If
                    Exit Sub
              
                End If
            Else
          
                If (DBCboClientName.BoundText) = "" Or val(DBCboClientName.BoundText) = 2 Then
                
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·«»œ „‰  ÕœÌœ ⁄„Ì·"
                    Else
                        MsgBox "Please Define Box"
                    End If
                    Exit Sub
              
                End If
                
            End If
          
            If ChekClodePeriod(XPDtbTrans.value) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                Else
                    MsgBox "Please Change Date Becouse This is Period is Closed"
                End If
                Exit Sub
            End If
              
            Dim Msg As String

            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch"
                Else
                    Msg = "Õœœ «·›—⁄ "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                dcBranch.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = Me.dcBranch.BoundText

            my_branch = Me.dcBranch.BoundText
                       
            If TxtNoteSerial.text = "" Then
                If Notes_coding(val(my_branch), XPDtbTrans.value) = "error" Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
                Else
                       
                    If Notes_coding(val(my_branch), XPDtbTrans.value) = "" Then
                        MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                    Else
                        '                       TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbBill.value)
                    End If
                End If
            End If

            Dim TxtNoteSerial1str As String

            If TxtNoteSerial1.text = "" Then
                TxtNoteSerial1str = Voucher_coding(val(my_branch), XPDtbTrans.value, 52, 5252)
    
                If TxtNoteSerial1str = "error" Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›…   ›« Ê—… ’Ì«‰…  ÃœÌœ… ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                Else
                               
                    If TxtNoteSerial1str = "" Then
                        MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ›« Ê—…  «·’Ì«‰…  ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                    Else
                        '             txtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbBill.value, 7, 170, , 21, DCPreFix.text)
                    End If
                End If
            End If

            If val(CboPayMentType.ListIndex) = -1 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Ì—ÃÏ «Œ Ì«— ÿ—Ìﬁ… «·œ›⁄"
                Else
                    MsgBox "Please Enter Type Payment"
                End If
                Exit Sub
            End If

            If val(CboPayMentType.ListIndex) = 0 Then
                If val(TxtPaymentValue.text) <= 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "Ì—ÃÏ «œŒ«· «·ﬁÌ„…"
                    Else
                        MsgBox "Please enter value"
                    End If
                    Exit Sub
                End If
                '                If val(lbl(17).Caption) < 0 Then
                '                    If SystemOptions.UserInterface = ArabicInterface Then
                '                        MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ «·ﬁÌ„… «ﬂ»— „‰ «·«Ã„«·Ì"
                '                    Else
                '                        MsgBox " Can not be Value larger than Total"
                '                    End If
                '                    Exit Sub
                '                End If
            End If
            If val(lbl(17).Caption) > 0 Then
                If val(DBCboClientName.BoundText) = 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "Ì—ÃÏ «Œ Ì«— «·⁄„Ì·"
                    Else
                        MsgBox "Please select Customer"
                    End If
                    Exit Sub
                End If
            End If
            SaveData

        Case 3
            Undo

        Case 4
            If ChekClodePeriod(XPDtbTrans.value) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                Else
                    MsgBox "Please Change Date Becouse This is Period is Closed"
                End If
                Exit Sub
            End If
              
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5
            Load FrmSearchComputercheck
            FrmSearchComputercheck.show

        Case 6
            Unload Me

        Case 7
            ShowGL_cc Me.TxtNoteSerial.text, , 200

        Case 8
            ' CalCulateParts
            RemoveGridRow
            
        Case 9
            Ch = 0
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.text) <> 0 Then
                print_report val(Me.XPTxtID.text)
        
            End If
        Case 11
            Ch = 1
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.text) <> 0 Then
                print_report val(Me.XPTxtID.text)
        
            End If
        
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub RemoveGridRow()

    With Me.Fg

        If .row <= 0 Then Exit Sub
        .RemoveItem .row
    End With

    ReLineGrid
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


 MySQL = " SELECT     dbo.TblBillComputerChek.ID, dbo.TblBillComputerChekDetails.Code, dbo.TblBillComputerChekDetails.Remarks, dbo.TblBillComputerChekDetails.Type,TblBillComputerChekDetails.MainType, "
 MySQL = MySQL & "                     dbo.TblBillComputerChekDetails.Amount, dbo.TblBillComputerChek.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
 MySQL = MySQL & "                     dbo.TblBillComputerChek.ClientName, dbo.TblBillComputerChek.Mobile, dbo.TblBillComputerChek.RecordDate, dbo.TblBillComputerChek.EndDate,"
 MySQL = MySQL & "                     dbo.TblBillComputerChek.PlateNo, dbo.TblBillComputerChek.CarID, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee, dbo.TblBillComputerChek.ColorID,"
 MySQL = MySQL & "                      dbo.TblColor.name AS Color, dbo.TblColor.namee AS Colore, dbo.TblBillComputerChek.ModelID, dbo.TblCarModels.Model, dbo.TblEmployee.Emp_Name,"
 MySQL = MySQL & "                     dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4,"
 MySQL = MySQL & "                     dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4,"
 MySQL = MySQL & "                     dbo.TblBillComputerChek.EmpID, dbo.TblEmployee.Emp_ID, dbo.TblBillComputerChek.StartDate, dbo.TblBillComputerChek.YarFact, dbo.TblYearFact.name AS Year,"
 MySQL = MySQL & "                     dbo.TblYearFact.namee AS Yeare, dbo.TblBillComputerChek.Cus_ID1, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode,"
 MySQL = MySQL & "                     dbo.TblBillComputerChek.PayMentType, dbo.TblBillComputerChek.BoxStorID, dbo.TblBoxesData.BoxName, dbo.TblBoxesData.BoxNameE,"
 MySQL = MySQL & "                     dbo.TblBillComputerChek.remainvalue , dbo.TblBillComputerChek.TotalValue, dbo.TblBillComputerChek.PaymentValue,"
 
 MySQL = MySQL & " TblBillComputerChek.subcar1,TblBillComputerChek.subcar2,TblBillComputerChek.subcar3,TblBillComputerChek.subcar4,TblBillComputerChek.subcar5,TblBillComputerChek.subcar6,"
 MySQL = MySQL & " TblBillComputerChek.subcar7,TblBillComputerChek.subcar8,TblBillComputerChek.subcar9,TblBillComputerChek.subcar10,TblBillComputerChek.subcar11,TblBillComputerChek.subcar12,TblBillComputerChek.subcar13,TblBillComputerChek.subcar14"
 MySQL = MySQL & " FROM         dbo.TblBillComputerChek INNER JOIN"
 MySQL = MySQL & "                      dbo.TblBillComputerChekDetails ON dbo.TblBillComputerChek.ID = dbo.TblBillComputerChekDetails.ID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblBoxesData ON dbo.TblBillComputerChek.BoxStorID = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
 MySQL = MySQL & "                      dbo.TblCustemers ON dbo.TblBillComputerChek.Cus_ID1 = dbo.TblCustemers.CusID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblYearFact ON dbo.TblBillComputerChek.YarFact = dbo.TblYearFact.Id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblEmployee ON dbo.TblBillComputerChek.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblCarModels ON dbo.TblBillComputerChek.ModelID = dbo.TblCarModels.Id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblColor ON dbo.TblBillComputerChek.ColorID = dbo.TblColor.Id LEFT OUTER JOIN"
 MySQL = MySQL & "                    dbo.TBLCarTypes ON dbo.TblBillComputerChek.CarID = dbo.TBLCarTypes.id LEFT OUTER JOIN"
 MySQL = MySQL & "                     dbo.TblBranchesData ON dbo.TblBillComputerChek.BranchID = dbo.TblBranchesData.branch_id"
 
 
 
 
MySQL = " SELECT        TblBillComputerChek.ID, TblBillComputerChekDetails.Code, TblBillComputerChekDetails.Remarks, TblBillComputerChekDetails.Type, TblBillComputerChekDetails.MainType, TblBillComputerChekDetails.Amount,"
MySQL = MySQL & "                                             TblBillComputerChek.BranchID, TblBranchesData.branch_name, TblBranchesData.branch_namee, TblBillComputerChek.ClientName, TblBillComputerChek.Mobile, TblBillComputerChek.RecordDate,"
MySQL = MySQL & "                                              TblBillComputerChek.EndDate, TblBillComputerChek.PlateNo, TblBillComputerChek.CarID, TBLCarTypes.name, TBLCarTypes.namee, TblBillComputerChek.ColorID, TblColor.name AS Color, TblColor.namee AS Colore,"
MySQL = MySQL & "                                              TblBillComputerChek.ModelID, TblCarModels.Model, TblEmployee.Emp_Name, TblEmployee.Emp_Code, TblEmployee.Emp_Name1, TblEmployee.Emp_Name2, TblEmployee.Emp_Name3, TblEmployee.Emp_Name4,"
MySQL = MySQL & "                                              TblEmployee.Emp_Namee, TblEmployee.Emp_Namee1, TblEmployee.Emp_Namee2, TblEmployee.Emp_Namee3, TblEmployee.Emp_Namee4, TblBillComputerChek.EmpID, TblEmployee.Emp_ID,"
MySQL = MySQL & "                                              TblBillComputerChek.StartDate, TblBillComputerChek.YarFact, TblYearFact.name AS Year, TblYearFact.namee AS Yeare, TblBillComputerChek.Cus_ID1, TblCustemers.CusName, TblCustemers.CusNamee,"
MySQL = MySQL & "                                              TblCustemers.Fullcode, TblBillComputerChek.PayMentType, TblBillComputerChek.BoxStorID, TblBoxesData.BoxName, TblBoxesData.BoxNameE, TblBillComputerChek.RemainValue, TblBillComputerChek.TotalValue,"
MySQL = MySQL & "                                              TblBillComputerChek.PaymentValue, TblBillComputerChek.subcar1, TblBillComputerChek.subcar2, TblBillComputerChek.subcar3, TblBillComputerChek.subcar4, TblBillComputerChek.subcar5, TblBillComputerChek.subcar6,"
MySQL = MySQL & "                                              TblBillComputerChek.subcar7, TblBillComputerChek.subcar8, TblBillComputerChek.subcar9, TblBillComputerChek.subcar10, TblBillComputerChek.subcar11, TblBillComputerChek.subcar12, TblBillComputerChek.subcar13,"
MySQL = MySQL & "                                              TblBillComputerChek.subcar14, TblComputerChek_1.name AS MainName,TblComputerChek.Name AS BasicName ,TblComputerChek.namee AS BasicNamee,TblComputerChek_1.ID AS MainID"
 
MySQL = MySQL & "   FROM            TblBillComputerChek INNER JOIN"
MySQL = MySQL & "                            TblBillComputerChekDetails ON TblBillComputerChek.ID = TblBillComputerChekDetails.ID INNER JOIN"
MySQL = MySQL & "                            TblComputerChek ON TblBillComputerChekDetails.Code = TblComputerChek.Manualcode INNER JOIN"
MySQL = MySQL & "                            TblComputerChek AS TblComputerChek_1 ON TblComputerChek.MainID = TblComputerChek_1.Id LEFT OUTER JOIN"
MySQL = MySQL & "                            TblBoxesData ON TblBillComputerChek.BoxStorID = TblBoxesData.BoxID LEFT OUTER JOIN"
MySQL = MySQL & "                            TblCustemers ON TblBillComputerChek.Cus_ID1 = TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                            TblYearFact ON TblBillComputerChek.YarFact = TblYearFact.Id LEFT OUTER JOIN"
MySQL = MySQL & "                            TblEmployee ON TblBillComputerChek.EmpID = TblEmployee.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                            TblCarModels ON TblBillComputerChek.ModelID = TblCarModels.Id LEFT OUTER JOIN"
MySQL = MySQL & "                            TblColor ON TblBillComputerChek.ColorID = TblColor.Id LEFT OUTER JOIN"
MySQL = MySQL & "                            TBLCarTypes ON TblBillComputerChek.CarID = TBLCarTypes.id LEFT OUTER JOIN"
MySQL = MySQL & "                            TblBranchesData ON TblBillComputerChek.BranchID = TblBranchesData.branch_id"

MySQL = MySQL & "      Where (dbo.TblBillComputerChek.id =" & val(XPTxtID.text) & ")"

MySQL = MySQL & "      ORDER BY TblBillComputerChekDetails.ID"
 If Ch = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBillComputerChek.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBillComputerChek.rpt"
        End If
Else
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBillComputerChekWi.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepBillComputerChekWi.rpt"
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
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
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
        xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(lbl(12).Caption), "0.00"), 0, True, ".")
  '      xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
      '   xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
'  xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
 '  xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
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


 
  
 
End Function

Private Sub CmdAttach_Click()
 On Error Resume Next
ShowAttachments TxtNoteSerial1, "0612212092023"
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub DBCboClientName_Change()
DBCboClientName_Click (0)
End Sub

Private Sub DBCboClientName_Click(Area As Integer)
    Dim Fullcode As String
    GetCustomersDetail val(DBCboClientName.BoundText), , Fullcode, 1
    Text1.text = Fullcode
End Sub

Private Sub DcbCarType_Click(Area As Integer)
Dim Dcombos As ClsDataCombos
      Set Dcombos = New ClsDataCombos
    
      If val(Me.DcbCarType.BoundText) <> 0 Then
   Dcombos.GetTblCarModels Me.DcbModel, , val(Me.DcbCarType.BoundText)
   End If
   
 
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub FG_AfterEdit(ByVal row As Long, ByVal Col As Long)
    Dim StrAccountCode  As String
    Dim StrAccountCode1 As String
    Dim Msg             As String
    Dim rs              As New ADODB.Recordset
    Dim StrSQL          As String
    Dim ClsAcc          As New ClsAccounts
    Dim LngRow          As Long
    Dim StrComboList    As String
    
    With Fg

        Select Case .ColKey(Col)
 
            Case "Type"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Code"), False, True)
                .TextMatrix(row, .ColIndex("Code")) = StrAccountCode
            Case "Code"

                Dim Rs3 As ADODB.Recordset
                Set Rs3 = New ADODB.Recordset
                Dim sql As String
                
                sql = "SELECT TblComputerChek.Id, "
                sql = sql & "       name = TblComputerChek.name + '-' +  TblComputerChek.namee , "
                sql = sql & "       TblComputerChek.namee, "
                sql = sql & "       TblComputerChek.Manualcode, "
                sql = sql & "       TblComputerChek.MainId, "
                sql = sql & "       maintbl.name MainName, "
                sql = sql & "       maintbl.namee MainNamee ,maintbl.Manualcode MainCode "
                sql = sql & "FROM TblComputerChek "
                sql = sql & "    LEFT OUTER JOIN TblComputerChek maintbl "
                sql = sql & "        ON TblComputerChek.MainId = maintbl.Id "
                sql = sql & " where TblComputerChek.Manualcode=" & val(.TextMatrix(row, .ColIndex("Code"))) & ""
                sql = sql & " And isnull(TblComputerChek.ismain , 0 )  =  0 "
         
                Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
                ' IIf(Rs3!MainName & "" <> "", Rs3!MainName & " - ", " ") &
                If Rs3.RecordCount > 0 Then
                    .TextMatrix(row, .ColIndex("Type")) = Rs3!Name & ""
                    .TextMatrix(row, .ColIndex("MainType")) = Rs3!MainCode & "-" & Rs3!MainName & "" & "-" & Rs3!MainNamee & ""
                    
                    .TextMatrix(row, .ColIndex("Amount")) = 1
                    
                    
                Else
                    MsgBox "«·ﬂÊœ «·„œŒ· €Ì— ’ÕÌÕ"
                    .TextMatrix(row, .ColIndex("Code")) = ""
                    .TextMatrix(row, .ColIndex("MainType")) = ""
                    
                    Exit Sub
                End If
                
                '  MsgBox StrAccountCode
        End Select
   
        If row = .rows - 1 Then
    
            .rows = .rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid
End Sub
Private Sub ReLineGrid()
    Dim i          As Integer
    Dim IntCounter As Integer
    Dim amountTo   As Integer
    IntCounter = 0
    amountTo = 0
    With Fg

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("Type")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("serial")) = IntCounter
                amountTo = amountTo + val(.TextMatrix(i, .ColIndex("Amount")))
            End If
 
        Next i
 
    End With
    lbl(12).Caption = amountTo

End Sub

Private Sub FG_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
 With Fg

        '   If Row > .FixedRows Then
        '       If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
        '           Cancel = True
        '       End If
        '   End If
        Select Case .ColKey(Col)
            
            Case "Code"
       Fg.ComboList = ""
         '  Case "JobName"
          '     Fg.ComboList = ""
                'Cancel = True
                Case "Remarks"
              Fg.ComboList = ""
                'Cancel = True
                 Case "Amount"
               Fg.ComboList = ""
        End Select

    End With

    
End Sub

Private Sub fg_StartEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs             As New ADODB.Recordset
    Dim StrSQL         As String
    Dim StrAccountType As String
    Dim StrComboList   As String
    Dim Msg            As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Fg

        Select Case .ColKey(Col)
            Case "MainType"
                Cancel = True
            Case "Type"
                StrSQL = "select Name + '-' +  ISNULL(namee,'') as NN ,Manualcode from TblComputerChek"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg.BuildComboList(rs, "NN", "Manualcode")
                Else
                    StrComboList = Fg.BuildComboList(rs, "NN", "Manualcode")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                .ComboList = StrComboList
            Case "Code"

        End Select

    End With

End Sub

'Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
'    Dim EmpID As Integer
'
'    If KeyAscii = vbKeyReturn Then
'        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
'        DcboEmpName.BoundText = EmpID
'    End If
'
'End Sub

 

Private Sub DcboEmpName_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 3
        FrmEmployeeSearch.show
  
    End If

End Sub

Private Sub DcboEmpName_Click(Area As Integer)
'    On Error Resume Next
       If val(DcboEmpName.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.text = EmpCode
    
   If Me.TxtModFlg = "R" Then Exit Sub
   
   
    Dim StrSQL As String

 
        
        
        Dim IssueDate As Date
        Dim DepID As Double
        Dim specid As Double
        Dim JobTypeID As Double
        Dim gradeID As Double
        Dim Account_code2 As String
           Dim Account_code  As String
        Dim Balance As String
        Dim endContractPerMonth As Double
        get_employee_information val(Me.DcboEmpName.BoundText), IssueDate, DepID, specid, JobTypeID, gradeID, Account_code2, Account_code, endContractPerMonth
        
  '        WriteCustomerBalPublic Account_code2, Balance
          
  'lbl(22).Caption = val(Balance)

  '        WriteCustomerBalPublic Account_Code, Balance
          
  'lbl(21).Caption = val(Balance)
  'lbl(20).Caption = IIf(endContractPerMonth > 0, endContractPerMonth, 0)
      '  DBIssueDate.value = issuedate
      '  DcboEmpDepartments.BoundText = depid
      '  DcboSpecifications.BoundText = gradeID
      '  DcboJobsType.BoundText = JobTypeID
      '  lbl(23).Caption = GetEmployeeSalaryAccordingToComponent(val(Me.DcboEmpName.BoundText), "")
      '
    'End If

End Sub


Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption

End Sub



Private Sub imag1_Click()

If Me.imag1.Picture = Me.imgnul.Picture Then
Me.imag1.Picture = Me.Img.Picture
Else
 Me.imag1.Picture = Me.imgnul.Picture
 End If
 
End Sub

Private Sub imag2_Click()

If Me.imag2.Picture = Me.imgnul.Picture Then
Me.imag2.Picture = Me.Img.Picture
Else
 Me.imag2.Picture = Me.imgnul.Picture
 End If
 
End Sub

Private Sub imag3_Click()

If Me.imag3.Picture = Me.imgnul.Picture Then
Me.imag3.Picture = Me.Img.Picture
Else
 Me.imag3.Picture = Me.imgnul.Picture
 End If
End Sub

Private Sub imag4_Click()

If Me.imag4.Picture = Me.imgnul.Picture Then
Me.imag4.Picture = Me.Img.Picture
Else
 Me.imag4.Picture = Me.imgnul.Picture
 End If
End Sub

Private Sub imag5_Click()

If Me.imag5.Picture = Me.imgnul.Picture Then
Me.imag5.Picture = Me.Img.Picture
Else
 Me.imag5.Picture = Me.imgnul.Picture
 End If
 
End Sub

Private Sub img10_Click()
If Me.img10.Picture = Me.imgnul.Picture Then
Me.img10.Picture = Me.Img.Picture
Else
 Me.img10.Picture = Me.imgnul.Picture
 End If

End Sub

Private Sub img11_Click()
If Me.img11.Picture = Me.imgnul.Picture Then
Me.img11.Picture = Me.Img.Picture
Else
 Me.img11.Picture = Me.imgnul.Picture
 End If
End Sub

Private Sub img12_Click()
If Me.img12.Picture = Me.imgnul.Picture Then
Me.img12.Picture = Me.Img.Picture
Else
 Me.img12.Picture = Me.imgnul.Picture
 End If
End Sub

Private Sub img13_Click()
If Me.img13.Picture = Me.imgnul.Picture Then
Me.img13.Picture = Me.Img.Picture
Else
 Me.img13.Picture = Me.imgnul.Picture
 End If
End Sub

Private Sub img14_Click()
If Me.img14.Picture = Me.imgnul.Picture Then
Me.img14.Picture = Me.Img.Picture
Else
 Me.img14.Picture = Me.imgnul.Picture
 End If
End Sub

Private Sub img6_Click()

If Me.img6.Picture = Me.imgnul.Picture Then
Me.img6.Picture = Me.Img.Picture
Else
 Me.img6.Picture = Me.imgnul.Picture
 End If
End Sub

Private Sub img7_Click()

If Me.img7.Picture = Me.imgnul.Picture Then
Me.img7.Picture = Me.Img.Picture
Else
 Me.img7.Picture = Me.imgnul.Picture
 End If
End Sub

Private Sub img8_Click()

If Me.img8.Picture = Me.imgnul.Picture Then
Me.img8.Picture = Me.Img.Picture
Else
 Me.img8.Picture = Me.imgnul.Picture
 End If
 

End Sub

Private Sub img9_Click()

If Me.img9.Picture = Me.imgnul.Picture Then
Me.img9.Picture = Me.Img.Picture
Else
 Me.img9.Picture = Me.imgnul.Picture
 End If
End Sub



Private Sub lbl_Change(index As Integer)
If Me.TxtModFlg.text <> "R" Then
If index = 12 Then
lbl(17).Caption = val(lbl(12).Caption) - val(Me.TxtPaymentValue.text)
End If
End If
End Sub


Private Sub menue_Click(index As Integer)
showsforms index
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , Text1.text, 1
        DBCboClientName.BoundText = CUSTID
    End If
End Sub

Private Sub TxtClient_Change()
If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
'                Dim CUSTID As Double
'createCustomer TxtClient, TxtClient, val(Dcbranch.BoundText), CUSTID
'            TxtClientCode.text = CUSTID

End If

End Sub

Private Sub TxtPaymentValue_Change()
If Me.TxtModFlg.text <> "R" Then
lbl(17).Caption = val(lbl(12).Caption) - val(Me.TxtPaymentValue.text)
End If
End Sub

Private Sub XPDtbTrans_Change()

    If Trim(TxtNoteSerial1.text) <> "" Then
        oldtxtNoteSerial1.text = TxtNoteSerial1.text
    End If

    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""

End Sub

Private Sub Dcbranch_Click(Area As Integer)
 
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim StrSQL  As String
    
    Dim GrdBack As ClsBackGroundPic
    Set Dcombos = New ClsDataCombos
    ' Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetBoxes Me.DcboBoxStor
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetTblCarsDataGroup Me.DcbCarType
    Dcombos.GetTblColor Me.DcbColor
    Dcombos.GetTblCarModels Me.DcbModel
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "  select CusID,CusName from TblCustemers  where  type =1   order by CusName"
    Else
        StrSQL = "  select CusID,CusNamee from TblCustemers  where  type =1   order by CusNamee"
    End If

    fill_combo DBCboClientName, StrSQL
    
    Dim i As Integer
    For i = 1995 To 2100
        Me.DcbYearFact.AddItem (i)
    Next i
    On Error GoTo ErrTrap
    Set GrdBack = New ClsBackGroundPic
    If SystemOptions.UserInterface = ArabicInterface Then
        With CboPayMentType
            .Clear
            .AddItem "‰ﬁœ«"
            .AddItem "¬Ã·"
        End With
    Else

        With CboPayMentType
            .Clear
            .AddItem "Cash"
            .AddItem "Credit"
        End With
    End If
    ' With Me.Fg
    '     .RowHeightMin = 300
    '     .WallPaper = GrdBack.Picture
    '     .AutoSize 0, .Cols - 1, False
    ' End With

    Set TTD = New clstooltipdemand
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    'Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Resize_Form Me
    AddTip
      
    'Dcombos.GetTblYearFact Me.DcbYearFact
    ' Dcombos.GetEmpDepartments Me.DcboEmpDepartments
    ' Dcombos.GetEmpJobsTypes Me.DcboJobsType

    'Dcombos.GetEmpGrades Me.DcboSpecifications
    
    If SystemOptions.usertype <> UserAdminAll Then
        Me.dcBranch.Enabled = True
    End If

    SetDtpickerDate Me.XPDtbTrans
    ' YearMonth
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblBillComputerChek     Order By ID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPDtbTrans.value = Date
    Me.TxtModFlg.text = "R"
    Retrive

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
      Fg.ColComboList(Fg.ColIndex("Amount")) = "#1;”·Ì„|#0;„⁄ÿ·"
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
  '  Label1.Visible = False
  lbl(18).Caption = "Remaining"
  Label1.Caption = "No GL"
  lbl(16).Caption = "Box"
  lbl(14).Caption = "Customer"
lbl(19).Caption = "Payment"
lbl(13).Caption = "Type "
Cmd(10).Caption = "Print GL"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
 Cmd(9).Caption = "Prient"
 Cmd(11).Caption = "Order Computer Check"
    Cmd(6).Caption = "Exit"
    Cmd(8).Caption = "Delete"
'    CmdHelp.Caption = "Help"
Cmd(9).Caption = "Prient"
XPTab301.Caption = "Data Check"

    Me.Caption = "Bill Computer check"""
    EleHeader.Caption = Me.Caption
    lbl(4).Caption = "OPR#"
    lbl(1).Caption = "Date"
    lblBr.Caption = "Branch"
    lbl(3).Caption = "Mobile"
    lbl(5).Caption = "PlateNo"
    LblCar.Caption = "Car Type"
    lblModel.Caption = "Model"
    LblYear.Caption = "Year of Fact"
    lblColor.Caption = "Colour"
    lbl(15).Caption = "Eng Name"
    lbl(2).Caption = "ClientName"
    lbl(10).Caption = "Date Entry"
    lbl(9).Caption = "Date Exit"
    lbl(11).Caption = "Total"
    
   ' lbl(3).Caption = "Employee"
   ' lbl(2).Caption = "value"
   ' lbl(0).Caption = "Box"
   ' Fra(0).Caption = "payments Method"
   ' lbl(9).Caption = "Count"
   ' lbl(10).Caption = "Start"
   ' lbl(11).Caption = "Month"
   ' lbl(12).Caption = "Year"
   ' Cmd(8).Caption = "Calc Dates"
   ' ChkSaleryDis.Caption = "Auto Discount"
    lbl(8).Caption = "By"
    lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"

    With Me.Fg
        .TextMatrix(0, .ColIndex("Type")) = "Type"
        .TextMatrix(0, .ColIndex("serial")) = "Serial"
        .TextMatrix(0, .ColIndex("Amount")) = "Amount"
         .TextMatrix(0, .ColIndex("Code")) = "Code"
          .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"

    End With

End Sub

'Private Sub YearMonth()

  '  Dim i As Integer
  '  Dim IntDefIndex As Integer
'
'    CmbMonth.Clear
'
'    For i = 1 To 12
'        CmbMonth.AddItem MonthName(i)
'    Next
'
'    CmbMonth.ListIndex = Month(Date) - 1
'    CboYear.Clear
'
 '   For i = 2010 To 2050
''        CboYear.AddItem i
'
'        If i = year(Date) Then
'            IntDefIndex = CboYear.NewIndex
'        End If

'    Next

'    CboYear.ListIndex = IntDefIndex
'End Sub

Private Sub Form_Paint()
    TTD.Destroy
End Sub

Private Sub Form_Resize()
    TTD.Destroy
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
        Set rs = Nothing
    End If

    Set TTP = Nothing
    'Set EmpReport = Nothing
    TTD.Destroy
    Exit Sub
ErrTrap:
End Sub



Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
            '        Me.Caption = "”·› «·„ÊŸ›Ì‰"
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
         '   TxtAdvanceValue.Locked = True
            Me.DcboBox.locked = True
            XPDtbTrans.Enabled = False

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If

        Case "N"
            '        Me.Caption = "”·› «·„ÊŸ›Ì‰( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            '      Me.XPBtnMove(0).Enabled = False
            '      Me.XPBtnMove(1).Enabled = False
            '      Me.XPBtnMove(2).Enabled = False
            '      Me.XPBtnMove(3).Enabled = False
           ' TxtAdvanceValue.Locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            XPDtbTrans.value = Date

        Case "E"
            '        Me.Caption = "”·› «·„ÊŸ›Ì‰(  ⁄œÌ· )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
           ' TxtAdvanceValue.Locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub



Private Sub XPBtnMove_Click(index As Integer)
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
    End If

    Select Case index

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

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim i         As Integer
    Dim StrSQL    As String
    Fg.Clear flexClearScrollable, flexClearEverything
    Fg.rows = 2
    Fg.Enabled = True
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
            rs.Find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If

    XPTxtID.text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    dcBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
    Me.DcboEmpName.BoundText = IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value)
    Me.TxtClient.text = IIf(IsNull(rs("ClientName").value), "", rs("ClientName").value)
    Me.TxtMobile.text = IIf(IsNull(rs("Mobile").value), "", rs("Mobile").value)
    Me.TxtPlateNo.text = IIf(IsNull(rs("PlateNo").value), "", rs("PlateNo").value)
    Me.DtEnd.value = IIf(IsNull(rs("EndDate").value), Date, rs("EndDate").value)
    Me.DtStart.value = IIf(IsNull(rs("StartDate").value), Date, rs("StartDate").value)
    Me.DcbCarType.BoundText = IIf(IsNull(rs("CarID").value), "", rs("CarID").value)
    Me.DcbColor.BoundText = IIf(IsNull(rs("ColorID").value), "", rs("ColorID").value)
    Me.DcbModel.BoundText = IIf(IsNull(rs("ModelID").value), "", rs("ModelID").value)
    Me.DcbYearFact.text = IIf(IsNull(rs("YarFact").value), "", rs("YarFact").value)
   
    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    Me.TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
    Me.TxtNoteID.text = IIf(IsNull(rs("NoteID").value), "", rs("NoteID").value)
    Me.DBCboClientName.BoundText = IIf(IsNull(rs("Cus_ID1").value), "", rs("Cus_ID1").value)
    Me.DcboBoxStor.BoundText = IIf(IsNull(rs("BoxStorID").value), "", rs("BoxStorID").value)
    Me.CboPayMentType.ListIndex = IIf(IsNull(rs("PayMentType").value), -1, rs("PayMentType").value)
   
    lbl(17).Caption = IIf(IsNull(rs("RemainValue").value), 0, rs("RemainValue").value)
    lbl(12).Caption = IIf(IsNull(rs("TotalValue").value), 0, rs("TotalValue").value)
    TxtPaymentValue.text = IIf(IsNull(rs("PaymentValue").value), 0, rs("PaymentValue").value)

    
    
    If IIf(IsNull(rs("subcar1").value), False, rs("subcar1").value) Then
        Me.imag1.Picture = Me.Img.Picture
    Else
        Me.imag1.Picture = Me.imgnul.Picture
    End If
     
    If IIf(IsNull(rs("subcar2").value), False, rs("subcar2").value) Then
        Me.imag2.Picture = Me.Img.Picture
    Else
        Me.imag2.Picture = Me.imgnul.Picture
    End If
     
'
    If IIf(IsNull(rs("subcar3").value), False, rs("subcar3").value) Then
        Me.imag3.Picture = Me.Img.Picture
    Else
        Me.imag3.Picture = Me.imgnul.Picture
    End If

    If IIf(IsNull(rs("subcar4").value), False, rs("subcar4").value) Then
        Me.imag4.Picture = Me.Img.Picture
    Else
        Me.imag4.Picture = Me.imgnul.Picture
    End If

    If IIf(IsNull(rs("subcar5").value), False, rs("subcar5").value) Then
        Me.imag5.Picture = Me.Img.Picture
    Else
        Me.imag5.Picture = Me.imgnul.Picture
    End If

    If IIf(IsNull(rs("subcar6").value), False, rs("subcar6").value) Then
        Me.img6.Picture = Me.Img.Picture
    Else
        Me.img6.Picture = Me.imgnul.Picture
    End If

    If IIf(IsNull(rs("subcar7").value), False, rs("subcar7").value) Then
        Me.img7.Picture = Me.Img.Picture
    Else
        Me.img7.Picture = Me.imgnul.Picture
    End If

    If IIf(IsNull(rs("subcar8").value), False, rs("subcar8").value) Then
        Me.img8.Picture = Me.Img.Picture
    Else
        Me.img8.Picture = Me.imgnul.Picture
    End If

    If IIf(IsNull(rs("subcar9").value), False, rs("subcar9").value) Then
        Me.img9.Picture = Me.Img.Picture
    Else
        Me.img9.Picture = Me.imgnul.Picture
    End If

    If IIf(IsNull(rs("subcar10").value), False, rs("subcar10").value) Then
        Me.img10.Picture = Me.Img.Picture
    Else
        Me.img10.Picture = Me.imgnul.Picture
    End If

    If IIf(IsNull(rs("subcar11").value), False, rs("subcar11").value) Then
        Me.img11.Picture = Me.Img.Picture
    Else
        Me.img11.Picture = Me.imgnul.Picture
    End If

    If IIf(IsNull(rs("subcar12").value), False, rs("subcar12").value) Then
        Me.img12.Picture = Me.Img.Picture
    Else
        Me.img12.Picture = Me.imgnul.Picture
    End If

    If IIf(IsNull(rs("subcar13").value), False, rs("subcar13").value) Then
        Me.img13.Picture = Me.Img.Picture
    Else
        Me.img13.Picture = Me.imgnul.Picture
    End If



    If IIf(IsNull(rs("subcar14").value), False, rs("subcar14").value) Then
        Me.img14.Picture = Me.Img.Picture
    Else
        Me.img14.Picture = Me.imgnul.Picture
    End If

'
'.TextMatrix(i, .ColIndex("subcar1")) = IIf(IsNull(RsDetails("subcar1").value), 0, RsDetails("subcar1").value)
'.TextMatrix(i, .ColIndex("subcar2")) = IIf(IsNull(RsDetails("subcar2").value), 0, RsDetails("subcar2").value)
'.TextMatrix(i, .ColIndex("subcar3")) = IIf(IsNull(RsDetails("subcar3").value), 0, RsDetails("subcar3").value)
'.TextMatrix(i, .ColIndex("subcar4")) = IIf(IsNull(RsDetails("subcar4").value), 0, RsDetails("subcar4").value)
'.TextMatrix(i, .ColIndex("subcar5")) = IIf(IsNull(RsDetails("subcar5").value), 0, RsDetails("subcar5").value)
'.TextMatrix(i, .ColIndex("subcar6")) = IIf(IsNull(RsDetails("subcar6").value), 0, RsDetails("subcar6").value)
'.TextMatrix(i, .ColIndex("subcar7")) = IIf(IsNull(RsDetails("subcar7").value), 0, RsDetails("subcar7").value)
'.TextMatrix(i, .ColIndex("subcar8")) = IIf(IsNull(RsDetails("subcar8").value), 0, RsDetails("subcar8").value)
'.TextMatrix(i, .ColIndex("subcar9")) = IIf(IsNull(RsDetails("subcar9").value), 0, RsDetails("subcar9").value)
'.TextMatrix(i, .ColIndex("subcar10")) = IIf(IsNull(RsDetails("subcar10").value), 0, RsDetails("subcar10").value)
'.TextMatrix(i, .ColIndex("subcar11")) = IIf(IsNull(RsDetails("subcar11").value), 0, RsDetails("subcar11").value)
'.TextMatrix(i, .ColIndex("subcar12")) = IIf(IsNull(RsDetails("subcar12").value), 0, RsDetails("subcar12").value)
'.TextMatrix(i, .ColIndex("subcar13")) = IIf(IsNull(RsDetails("subcar13").value), 0, RsDetails("subcar13").value)
'.TextMatrix(i, .ColIndex("subcar14")) = IIf(IsNull(RsDetails("subcar14").value), 0, RsDetails("subcar14").value)



  


    ' TxtFromName.text = IIf(IsNull(rs("FromName").value), "", rs("FromName").value)
    ' TxtPersonalDept.text = IIf(IsNull(rs("PersonalDept").value), "", rs("PersonalDept").value)
    '  DcboEmpDepartments.BoundText = IIf(IsNull(rs("DeparmentID").value), "", rs("DeparmentID").value)

    'DcboSpecifications.BoundText = IIf(IsNull(rs("gradeID").value), "", rs("gradeID").value)

    'DcboJobsType.BoundText = IIf(IsNull(rs("JobTypeID").value), "", rs("JobTypeID").value)

    'lbl(23).Caption = IIf(IsNull(rs("basicSalary").value), "", rs("basicSalary").value)
 
    ' lbl(22).Caption = IIf(IsNull(rs("EmpDue").value), "", rs("EmpDue").value)
    'lbl(20).Caption = IIf(IsNull(rs("Contractvalid").value), "", rs("Contractvalid").value)
    'lbl(21).Caption = IIf(IsNull(rs("oldAdvance").value), "", rs("oldAdvance").value)
 
    'TxtDiscount.text = IIf(IsNull(rs("Discount").value), "", rs("Discount").value)
    'txtDiscountDES.text = IIf(IsNull(rs("DiscountDES").value), "", rs("DiscountDES").value)

    '    Me.DcboEmpName.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
    '    TxtAdvanceValue.text = IIf(IsNull(rs("AdvanceValue").value), "", rs("AdvanceValue").value)
    '  Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
    '    Me.TxtPaymentCounts.text = IIf(IsNull(rs("PaymentCounts").value), "", rs("PaymentCounts").value)
 
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    If IsNull(rs("posted").value) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
        Else
            Accredit.Caption = " send to Approval   "
        End If
        Accredit.Enabled = True
    Else
        If SystemOptions.UserInterface = ArabicInterface Then
            Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
        Else
            Accredit.Caption = " sent to Approval   "
        End If
        Accredit.Enabled = False
    End If
   
    Set RsDetails = New ADODB.Recordset
    Dim s As String
  
   
    s = "SELECT TblBillComputerChekDetails.ID2, "
    s = s & "       TblBillComputerChekDetails.ID, "
    s = s & "       TblBillComputerChekDetails.Code, "
    s = s & "       TblBillComputerChekDetails.Remarks, "
    s = s & "       TblBillComputerChekDetails.Type, "
    s = s & "       TblBillComputerChekDetails.Amount,TblBillComputerChekDetails.MainType ,"
    s = s & "       TblComputerChek.name TypeName, "
    s = s & "       maintbl.name MainName, "
    s = s & "       maintbl.Manualcode MainCode "
    s = s & "FROM TblBillComputerChekDetails "
    s = s & "    LEFT OUTER JOIN TblComputerChek "
    s = s & "        ON TblBillComputerChekDetails.Code = TblComputerChek.Manualcode "
    s = s & "    LEFT OUTER JOIN TblComputerChek maintbl "
    s = s & "        ON TblComputerChek.MainId = maintbl.Id  "
    s = s & " Where TblBillComputerChekDetails.ID=" & val(XPTxtID.text)
    'StrSQL = "Select * From  TblBillComputerChekDetails Where ID=" & val(XPTxtID.text)
    RsDetails.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Fg.Clear flexClearScrollable, flexClearEverything
    Fg.rows = Fg.FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
        Fg.rows = Fg.FixedRows + RsDetails.RecordCount

        For i = Me.Fg.FixedRows To Fg.rows - 1
            Fg.TextMatrix(i, Fg.ColIndex("Code")) = RsDetails("Code").value
            Fg.TextMatrix(i, Fg.ColIndex("Type")) = RsDetails("Type").value
            Fg.TextMatrix(i, Fg.ColIndex("Amount")) = RsDetails("Amount").value
            Fg.TextMatrix(i, Fg.ColIndex("Remarks")) = RsDetails("Remarks").value
            Fg.TextMatrix(i, Fg.ColIndex("ID")) = RsDetails("id").value
            Fg.TextMatrix(i, Fg.ColIndex("MainType")) = RsDetails!MainType & ""
            RsDetails.MoveNext
        Next i

    End If

    RsDetails.Close
    Set RsDetails = Nothing
    ReLineGrid
    fillapprovData
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub
Public Function CREATE_VOUCHER1_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date, Optional NoteVal As Double, Optional AccountDeb_Creadit As String, Optional Account_Code_dynamic As String, Optional PayNotevalue As Double, Optional DebitAccount2 As String, Optional Msg As String)
Dim BasicSalaryAccount As String
Dim StrSQL As String
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords



    Dim i As Integer
    Dim LngDevID As Long
    
   
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim x As Integer
    Dim rs As New ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Dim j As Integer
    Dim ColumnName As String
    Dim SalaryAccount As String
    Dim BonusAccount As String
    Dim DiscountAccount As String
    
 notes_id = general_noteid
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")

    Dim line_no As Integer
    line_no = 1
                
     
    Dim CValue As Double
    Dim Branch As Integer
    Dim ProjectID As Integer
    
    BranchID = 1

line_no = 1
    BranchID = val(dcBranch.BoundText)
    
       If DebitAccount2 <> "" Then
             If ModAccounts.AddNewDev(LngDevID, line_no, DebitAccount2, NoteVal, 0, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
             
                If ModAccounts.AddNewDev(LngDevID, line_no, AccountDeb_Creadit, PayNotevalue, 0, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
            Else
                If ModAccounts.AddNewDev(LngDevID, line_no, AccountDeb_Creadit, NoteVal, 0, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
          End If
                
                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, NoteVal + PayNotevalue, 1, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                
            
    updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
   End Function
Private Sub SaveData()
    Dim Msg            As String
    Dim RsTemp         As New ADODB.Recordset
    Dim StrSQL         As String
    Dim BeginTrans     As Boolean
    Dim RsDetails      As ADODB.Recordset
    Dim i              As Integer
    Dim LngDevID       As Long
    Dim LngDevLineNo   As Long
    Dim StrAccountCode As String

    'On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
        If Me.DcboEmpName.BoundText = "" Then
            Msg = "ÌÃ»  ÕœÌœ «”„ «·„Â‰œ”..!! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Me.DcboEmpName.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If

        If Me.TxtClient.text = "" Then
            Msg = "ÌÃ» «œŒ«· «”„ «·⁄„Ì·..!! "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Me.TxtClient.SetFocus
            '   SendKeys "{F4}"
            Exit Sub
        End If

        Dim CUSTID As Double
        'createCustomer TxtClient, TxtClient, val(Dcbranch.BoundText), CUSTID
        CUSTID = val(DBCboClientName.BoundText)
        TxtClientCode.text = CUSTID

        Dim Account_Code_dynamic As String

        Account_Code_dynamic = get_account_code_branch(77, my_branch)
        
        If Account_Code_dynamic = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Else
                MsgBox "branch Not Created", vbCritical
            End If

            GoTo ErrTrap
        Else

            If Account_Code_dynamic = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»    „»Ì⁄«  ’Ì«‰… ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                Else
                    MsgBox "Sales Cost Account Not Defined in this Branch", vbCritical
                End If

                GoTo ErrTrap
         
            End If
        End If
            
        If TxtClientCode.text = "" Then
            '          If SystemOptions.UserInterface = ArabicInterface Then
            '             MsgBox "ÌÊÃœ Œÿ√ €Ì Õ”«» Â–« «·⁄„Ì·", vbCritical
            '         Else
            '             MsgBox "Customer  Account Have an Error", vbCritical
            '         End If
            '
            '                    GoTo ErrTrap
        End If
        
        Cn.BeginTrans
        BeginTrans = True

        If TxtModFlg.text = "N" Then

            XPTxtID.text = CStr(new_id("TblBillComputerChek", "ID", "", True))
            '     TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
            '     Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
        
            rs.AddNew
        ElseIf Me.TxtModFlg.text = "E" Then
            StrSQL = "Delete From TblBillComputerChekDetails Where ID=" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords

            StrSQL = "delete From Notes where NoteID=" & val(Me.TxtNoteID.text) ' Val(rs("Transaction_ID").value)
            Cn.Execute StrSQL, , adExecuteNoRecords

        End If
        
        Set RsNotesGeneral = New ADODB.Recordset
        RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
        RsNotesGeneral.AddNew
        RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
        general_noteid = RsNotesGeneral("NoteID").value
        TxtNoteID.text = general_noteid
        ' RsNotesGeneral("Transaction_ID").value = Val(XPTxtBillID.text)
        RsNotesGeneral("NoteDate").value = XPDtbTrans.value
        RsNotesGeneral("NoteType").value = 5252
        RsNotesGeneral("Note_Value").value = val(lbl(12).Caption)
        '  my_branch = val(Me.Dcbranch.BoundText)
                       
        If TxtNoteSerial.text = "" Then
            TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbTrans.value)
        End If
                      
        If TxtNoteSerial1.text = "" Then
        
            TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbTrans.value, 52, 5252)
    
        End If
        
        RsNotesGeneral("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
        RsNotesGeneral("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
        RsNotesGeneral("remark").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
        
        '    RsNotesGeneral("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '

        RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
        RsNotesGeneral("numbering_type1").value = sand_numbering_type(52) '  ›« Ê—… »Ì⁄
        RsNotesGeneral("sanad_year").value = year(XPDtbTrans.value)
        RsNotesGeneral("sanad_month").value = Month(XPDtbTrans.value)
        RsNotesGeneral("branch_no").value = val(Me.dcBranch.BoundText)
        RsNotesGeneral("note_value_by_characters").value = WriteNo(val(Me.lbl(12).Caption), 0, True)
        RsNotesGeneral.update
        
        rs("ID").value = val(XPTxtID.text)
        rs("RecordDate").value = XPDtbTrans.value
        rs("BranchID").value = IIf(Me.dcBranch.BoundText = "", Null, Me.dcBranch.BoundText)
        rs("EmpID").value = val(IIf(Me.DcboEmpName.BoundText = "", Null, Me.DcboEmpName.BoundText))
        rs("EndDate").value = Me.DtEnd.value
        rs("ClientName").value = Me.TxtClient.text
        rs("Mobile").value = Me.TxtMobile.text
        rs("PlateNo").value = Me.TxtPlateNo.text
        rs("CarID").value = IIf(Me.DcbCarType.BoundText = "", Null, Me.DcbCarType.BoundText)
        rs("ModelID").value = IIf(Me.DcbModel.BoundText = "", Null, Me.DcbModel.BoundText)
        rs("ColorID").value = IIf(Me.DcbColor.BoundText = "", Null, Me.DcbColor.BoundText)
        rs("UserID").value = Me.DCboUserName.BoundText
        rs("StartDate").value = Me.DtStart.value
        rs("YarFact").value = IIf(Me.DcbYearFact.text = "", Null, val(Me.DcbYearFact.text))
        rs("NoteSerial").value = Me.TxtNoteSerial.text
        rs("NoteSerial1").value = Me.TxtNoteSerial1.text
        rs("NoteID").value = Me.TxtNoteID.text
        rs("PaymentValue").value = val(TxtPaymentValue.text)
        rs("TotalValue").value = val(lbl(12).Caption)
        rs("RemainValue").value = val(lbl(17).Caption)

        rs("CusId").value = val(TxtClientCode.text)

        rs("PayMentType").value = IIf(val(Me.CboPayMentType.ListIndex) = -1, Null, val(Me.CboPayMentType.ListIndex))
        rs("Cus_ID1").value = IIf(Me.DBCboClientName.BoundText = "", Null, Me.DBCboClientName.BoundText)
        rs("BoxStorID").value = IIf(Me.DcboBoxStor.BoundText = "", Null, Me.DcboBoxStor.BoundText)
        
        
            If Me.imag1.Picture <> 0 Then
                rs("subcar1").value = 1
           Else
                rs("subcar1").value = 0
           End If


            If Me.imag2.Picture <> 0 Then
                rs("subcar2").value = 1
           Else
                rs("subcar2").value = 0
           End If

            If Me.imag3.Picture <> 0 Then
                rs("subcar3").value = 1
           Else
                rs("subcar3").value = 0
           End If

            If Me.imag4.Picture <> 0 Then
                rs("subcar4").value = 1
           Else
                rs("subcar4").value = 0
           End If

            If Me.imag5.Picture <> 0 Then
                rs("subcar5").value = 1
           Else
                rs("subcar5").value = 0
           End If

            If Me.img6.Picture <> 0 Then
                rs("subcar6").value = 1
           Else
                rs("subcar6").value = 0
           End If

            If Me.img7.Picture <> 0 Then
                rs("subcar7").value = 1
           Else
                rs("subcar7").value = 0
           End If

            If Me.img8.Picture <> 0 Then
                rs("subcar8").value = 1
           Else
                rs("subcar8").value = 0
           End If

            If Me.img9.Picture <> 0 Then
                rs("subcar9").value = 1
           Else
                rs("subcar9").value = 0
           End If

            If Me.img10.Picture <> 0 Then
                rs("subcar10").value = 1
           Else
                rs("subcar10").value = 0
           End If

            If Me.img11.Picture <> 0 Then
                rs("subcar11").value = 1
           Else
                rs("subcar11").value = 0
           End If

            If Me.img12.Picture <> 0 Then
                rs("subcar12").value = 1
           Else
                rs("subcar12").value = 0
           End If

            If Me.img13.Picture <> 0 Then
                rs("subcar13").value = 1
           Else
                rs("subcar13").value = 0
           End If

            If Me.img14.Picture <> 0 Then
                rs("subcar14").value = 1
           Else
                rs("subcar14").value = 0
           End If

'        rs("subcar1").value = val(.TextMatrix(i, .ColIndex("subcar1")))
'        rs("subcar2").value = val(.TextMatrix(i, .ColIndex("subcar2")))
'        rs("subcar3").value = val(.TextMatrix(i, .ColIndex("subcar3")))
'        rs("subcar4").value = val(.TextMatrix(i, .ColIndex("subcar4")))
'        rs("subcar5").value = val(.TextMatrix(i, .ColIndex("subcar5")))
'        rs("subcar6").value = val(.TextMatrix(i, .ColIndex("subcar6")))
'        rs("subcar7").value = val(.TextMatrix(i, .ColIndex("subcar7")))
'        rs("subcar8").value = val(.TextMatrix(i, .ColIndex("subcar8")))
'        rs("subcar9").value = val(.TextMatrix(i, .ColIndex("subcar9")))
'        rs("subcar10").value = val(.TextMatrix(i, .ColIndex("subcar10")))
'        rs("subcar11").value = val(.TextMatrix(i, .ColIndex("subcar11")))
'        rs("subcar12").value = val(.TextMatrix(i, .ColIndex("subcar12")))
'        rs("subcar13").value = val(.TextMatrix(i, .ColIndex("subcar13")))
'        rs("subcar14").value = val(.TextMatrix(i, .ColIndex("subcar14")))
'
        
        rs.update
        Set RsDetails = New ADODB.Recordset
        StrSQL = "SELECT     *  from dbo.TblBillComputerChekDetails Where (1 = -1)"
        RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        ' RsDetails.Open "TblBillComputerChekDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

        For i = Me.Fg.FixedRows To Fg.rows - 1
            If Fg.TextMatrix(i, Fg.ColIndex("Type")) <> "" Then
                RsDetails.AddNew
                RsDetails("ID").value = val(XPTxtID.text)
                RsDetails("Type").value = Fg.TextMatrix(i, Fg.ColIndex("Type"))
                RsDetails("Amount").value = val(Fg.TextMatrix(i, Fg.ColIndex("Amount")))
                RsDetails("Code").value = val(Fg.TextMatrix(i, Fg.ColIndex("Code")))
                RsDetails("MainType").value = Trim(Fg.TextMatrix(i, Fg.ColIndex("MainType")))
                
                RsDetails("Remarks").value = Fg.TextMatrix(i, Fg.ColIndex("Remarks"))
                RsDetails.update
            End If
        Next i
    
        Cn.CommitTrans
        BeginTrans = False
        RsDetails.Close
        Set RsDetails = Nothing
        
        Dim DebitAccount  As String
        Dim CreditAccount As String
        Dim des           As String
       
        If val(CboPayMentType.ListIndex) = 0 Then 'cash
            DebitAccount = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBoxStor.BoundText))
        Else
            DebitAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code")

        End If
        Dim DebitAccount2 As String
        des = "   ›« Ê—… ›Õ’ ﬂ„»ÌÊ — »—ﬁ„ " & TxtNoteSerial & "   ··⁄„Ì· " & TxtClient
         
        If val(CboPayMentType.ListIndex) = 0 Then
            If val(TxtPaymentValue.text) > 0 And val(lbl(17).Caption) > 0 Then
                DebitAccount2 = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText), "Account_code")
                CREATE_VOUCHER1_GE val(TxtNoteID.text), val(Me.dcBranch.BoundText), user_id, XPDtbTrans.value, val(TxtPaymentValue.text), DebitAccount, Account_Code_dynamic, val(TxtPaymentValue.text), DebitAccount2

            Else
                CREATE_VOUCHER1_GE val(TxtNoteID.text), val(Me.dcBranch.BoundText), user_id, XPDtbTrans.value, val(TxtPaymentValue.text), DebitAccount, Account_Code_dynamic
            End If
        Else
            CREATE_VOUCHER1_GE val(TxtNoteID.text), val(Me.dcBranch.BoundText), user_id, XPDtbTrans.value, val(TxtPaymentValue.text), DebitAccount, Account_Code_dynamic
        End If
        'CREATE_VOUCHER1_GE val(TxtNoteID.text), val(Me.Dcbranch.BoundText), user_id, XPDtbTrans.value, val(lbl(17).Caption), DebitAccount, Account_Code_dynamic, val(TxtPaymentValue.text), DebitAccount2, des
  
        ''//////////////////

        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    
        Select Case Me.TxtModFlg.text

            Case "N"
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
                Retrive
            Case "E"
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        End Select
        Retrive
        TxtModFlg.text = "R"
    End If

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

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find "ID=" & val(XPTxtID.text) & "", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
Dim StrSQL1 As String
    On Error GoTo ErrTrap

    If XPTxtID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
        StrSQL1 = "Delete From TblBillComputerChekDetails Where id=" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL1, , adExecuteNoRecords
                
          StrSQL = "delete From Notes where NoteID=" & val(Me.TxtNoteID.text) ' Val(rs("Transaction_ID").value)
        Cn.Execute StrSQL, , adExecuteNoRecords
                
                
                rs.delete
              ' StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where AdvanceID=" & val(Me.XPTxtID.text)
            '    Cn.Execute StrSQL, , adExecuteNoRecords
                rs.MoveFirst
                
                If rs.RecordCount < 1 Then
                 
            Fg.Clear flexClearScrollable, flexClearEverything
            Fg.rows = 2
            Fg.Enabled = True
            End If
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
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs.CancelUpdate
End Sub



Function FillApprovedTable()
 Dim RSApproval  As New ADODB.Recordset
   Set RSApproval = New ADODB.Recordset
   Dim currentdate As Date
   RSApproval.Open "[ApprovalData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable


 Dim sql As String
  Dim Rs1 As New ADODB.Recordset
 Dim i As Integer
    sql = "SELECT     TOP 100 PERCENT dbo.TblApprovalDef.ScreenName, dbo.TblApprovalDefDetails.PlainMessageID AS levelo, dbo.TbllevelWorker.EmpID, "
  sql = sql & " dbo.TblApprovalDefDetails.id AS levelorder, dbo.TbllevelWorker.id AS currorder"
  sql = sql & " FROM         dbo.TblApprovalDef INNER JOIN"
  sql = sql & " dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID INNER JOIN"
  sql = sql & "  dbo.TbllevelWorker ON dbo.TblApprovalDefDetails.PlainMessageID = dbo.TbllevelWorker.LevelID"
sql = sql & " WHERE     (dbo.TblApprovalDef.ScreenName = N'" & Me.Name & "')"
sql = sql & " ORDER BY dbo.TblApprovalDefDetails.id, dbo.TbllevelWorker.id  "

    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
            currentdate = Now
            For i = 1 To Rs1.RecordCount
              RSApproval.AddNew
                RSApproval("ScreenName").value = Me.Name
                RSApproval("levelo").value = IIf(IsNull(Rs1("levelo").value), Null, Rs1("levelo").value)
               RSApproval("EmpID").value = IIf(IsNull(Rs1("EmpID").value), Null, Rs1("EmpID").value)
                RSApproval("levelorder").value = IIf(IsNull(Rs1("levelorder").value), Null, Rs1("levelorder").value)
                 RSApproval("currorder").value = IIf(IsNull(Rs1("currorder").value), Null, Rs1("currorder").value)
                  RSApproval("Transaction_ID").value = val(Me.XPTxtID.text)
                   RSApproval("NoteSerial").value = val(Me.XPTxtID.text)
                RSApproval("Transaction_Date").value = Date
                
                  RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.Name), currentdate)
               RSApproval("SendTime").value = currentdate

                 If i = 1 Then
                        RSApproval("Currcursor").value = 1
                         RSApproval("FromUser").value = user_name
                End If
                
                RSApproval.update
                Rs1.MoveNext
            Next i

    End If
    
    

End Function



Function fillapprovData()
Dim Num As Integer
 Dim RsDetails As New ADODB.Recordset
 Dim StrSQL As String
 
 
 StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
StrSQL = StrSQL + " FROM         dbo.ApprovalData INNER JOIN"
StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

 If Not (RsDetails.EOF Or RsDetails.BOF) Then
        GRID2.rows = RsDetails.RecordCount + 1
 

        For Num = 1 To RsDetails.RecordCount
        
       GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
    If GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = "1" Then
   GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
   Else
    GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
    End If
    
        GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
           If SystemOptions.UserInterface = ArabicInterface Then
            GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
          Else
             GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
          End If
            If SystemOptions.UserInterface = ArabicInterface Then
            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            Else
            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            End If
            GRID2.TextMatrix(Num, GRID2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
          GRID2.TextMatrix(Num, GRID2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 
 
RsDetails.MoveNext
If Num = RsDetails.RecordCount Then

        If GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) <> "" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                      Label11.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
                                 Else
                                       Label11.Caption = "Approved"
                                 End If
                            Label11.backcolor = &H80FF80
        Else
                             If SystemOptions.UserInterface = ArabicInterface Then
                                     Label11.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                            Else
                                     Label11.Caption = "Currently required Approve"
                            End If
                 Label11.backcolor = &HFFFFC0
        End If

End If

        Next Num
Else
 GRID2.rows = 1
    End If
RsDetails.Close

End Function


Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            Cmd_Click (0)
        Else
            Sendkeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.text = "R" Then
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

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip
 With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘… «·«Ê«„— «·„› ÊÕ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(7), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
     With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘… «· ”·Ì„ ··⁄„Ì·", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(3), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…  «· ‰»ÌÂ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(4), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
     With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…  «· ﬁ«—Ì—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(5), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
      With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…  ’—› ﬁÿ⁄ «·€Ì«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(2), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

       With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘… ÿ·» ›Õ’ ﬂ„»ÌÊ —  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(6), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
         With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…    ÿ·» ’Ì«‰…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(0), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
        .Create Me.hWnd, " ›« Ê—… ›Õ’ ﬂ„»ÌÊ —", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "›« Ê—… ›Õ’ ﬂ„»ÌÊ —", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "›« Ê—… ›Õ’ ﬂ„»ÌÊ —", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "›« Ê—… ›Õ’ ﬂ„»ÌÊ —", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "›« Ê—… ›Õ’ ﬂ„»ÌÊ —", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "›« Ê—… ›Õ’ ﬂ„»ÌÊ —", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hWnd, "›« Ê—… ›Õ’ ﬂ„»ÌÊ —", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "›« Ê—… ›Õ’ ﬂ„»ÌÊ —", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "›« Ê—… ›Õ’ ﬂ„»ÌÊ —", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "›« Ê—… ›Õ’ ﬂ„»ÌÊ —", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "›« Ê—… ›Õ’ ﬂ„»ÌÊ —", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        '.AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With
    
    
          With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…   «·⁄„Ê·«  «·„” Õﬁ…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(9), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
          With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…   „·› «·⁄„·«¡  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(10), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
          With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…   ﬁ«—Ì— «·⁄„Ê·«   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(11), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    Exit Sub
ErrTrap:
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

                ' btnSave
            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub



