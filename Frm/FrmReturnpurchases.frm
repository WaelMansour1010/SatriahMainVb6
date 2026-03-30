VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmReturnpurchases 
   Caption         =   "„—œÊœ«  «·„‘ —Ì« "
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14685
   HelpContextID   =   240
   Icon            =   "FrmReturnpurchases.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   14685
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8550
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   14685
      _cx             =   25903
      _cy             =   15081
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   435
         Index           =   4
         Left            =   15
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   7710
         Width           =   14655
         _cx             =   25850
         _cy             =   767
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
         BorderWidth     =   0
         ChildSpacing    =   0
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
         Begin VB.TextBox XPTxtSum 
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
            Height          =   360
            Left            =   5205
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   60
            Visible         =   0   'False
            Width           =   315
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   3180
            TabIndex        =   31
            Top             =   75
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label LblValueAdded 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   8880
            RightToLeft     =   -1  'True
            TabIndex        =   137
            Top             =   0
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„… «·„÷«›…"
            Height          =   255
            Index           =   42
            Left            =   9735
            RightToLeft     =   -1  'True
            TabIndex        =   136
            Top             =   60
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«›Ì"
            Height          =   315
            Index           =   34
            Left            =   8265
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   120
            Width           =   570
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·ﬂ„Ì…"
            Height          =   315
            Index           =   32
            Left            =   6405
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   120
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label LblTotalQty 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   5550
            TabIndex        =   82
            Top             =   0
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«Ã„«·Ì"
            Height          =   315
            Index           =   63
            Left            =   13845
            TabIndex        =   81
            Top             =   0
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Œ’„"
            Height          =   255
            Index           =   23
            Left            =   11925
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   90
            Width           =   435
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œ’Ê„« "
            Height          =   255
            Index           =   11
            Left            =   14730
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   90
            Width           =   690
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·≈Ã„«·Ï"
            Height          =   255
            Index           =   10
            Left            =   12765
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   660
            Width           =   1305
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·≈Ã„«·Ï"
            Height          =   255
            Index           =   0
            Left            =   17085
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   90
            Width           =   780
         End
         Begin VB.Label LblDiscountsTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   10950
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   30
            Width           =   915
         End
         Begin VB.Label LblTotalAll 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   12390
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   30
            Width           =   1425
         End
         Begin VB.Label LblTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   7065
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   30
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„” Œœ„"
            Height          =   315
            Index           =   8
            Left            =   4725
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   75
            Width           =   450
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Left            =   -645
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   105
            Width           =   1425
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   270
            Left            =   1230
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   105
            Width           =   555
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "/"
            Height          =   240
            Index           =   2
            Left            =   750
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   120
            Width           =   240
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”Ã·"
            Height          =   270
            Index           =   1
            Left            =   2310
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   90
            Width           =   645
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic8 
         Height          =   2205
         Index           =   0
         Left            =   0
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   645
         Width           =   14655
         _cx             =   25850
         _cy             =   3889
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
         BorderWidth     =   0
         ChildSpacing    =   0
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
         Begin VB.TextBox XPTxtDiscountVal 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   0
            TabIndex        =   149
            Top             =   0
            Visible         =   0   'False
            Width           =   2235
         End
         Begin VB.ComboBox XPCboDiscountType 
            Height          =   315
            Left            =   3345
            Style           =   2  'Dropdown List
            TabIndex        =   148
            Top             =   45
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.TextBox txt_Currency_rate 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   144
            Text            =   "1"
            Top             =   1095
            Width           =   960
         End
         Begin VB.TextBox TxtVATNO 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   142
            Top             =   1080
            Width           =   1305
         End
         Begin VB.TextBox TXtResonVAT 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5400
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   141
            Top             =   1920
            Width           =   4440
         End
         Begin VB.ComboBox Dcbtyp 
            Height          =   315
            ItemData        =   "FrmReturnpurchases.frx":058A
            Left            =   10515
            List            =   "FrmReturnpurchases.frx":058C
            RightToLeft     =   -1  'True
            TabIndex        =   138
            Text            =   "Dcbtyp"
            Top             =   1920
            Width           =   3030
         End
         Begin VB.TextBox Transporter 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7680
            RightToLeft     =   -1  'True
            TabIndex        =   126
            Top             =   1560
            Width           =   2175
         End
         Begin VB.TextBox TxtEmployeeID 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   9450
            TabIndex        =   122
            Top             =   1170
            Width           =   360
         End
         Begin VB.TextBox TxtReasonReturns 
            Alignment       =   1  'Right Justify
            Height          =   450
            Left            =   90
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   118
            Top             =   1410
            Width           =   2100
         End
         Begin VB.TextBox TxtBillSupplier 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   117
            Top             =   1080
            Width           =   1185
         End
         Begin VB.TextBox TxtOrderSupply 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   116
            Top             =   1485
            Width           =   1185
         End
         Begin VB.TextBox TxtStoreID 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   12720
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   1125
            Width           =   825
         End
         Begin VB.TextBox TxtManualNo1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   101
            Top             =   1485
            Width           =   1305
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3120
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   1860
            Width           =   1305
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   465
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   -1530
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   11145
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   0
            Width           =   2400
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   -135
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   -1410
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.ComboBox CboPayMentType 
            Height          =   315
            Left            =   10515
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   1485
            Width           =   3030
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   12795
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   -105
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   225
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   -1200
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            Height          =   225
            Left            =   225
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   -1155
            Visible         =   0   'False
            Width           =   1110
         End
         Begin C1SizerLibCtl.C1Elastic TxtManualNO 
            Height          =   1020
            Index           =   6
            Left            =   1260
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   0
            Width           =   6300
            _cx             =   11113
            _cy             =   1799
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
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
            ForeColor       =   128
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   "‰Ê⁄ ⁄„·Ì… «·√— Ã«⁄"
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
            Begin VB.TextBox TxtManualNO11 
               Alignment       =   1  'Right Justify
               Height          =   330
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   115
               Top             =   240
               Width           =   2055
            End
            Begin VB.ComboBox CboRetrunType 
               Height          =   315
               Left            =   3405
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   48
               Top             =   240
               Width           =   2505
            End
            Begin VB.TextBox TxtInvID 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   4455
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   -210
               Visible         =   0   'False
               Width           =   1590
            End
            Begin VB.TextBox TxtInvSerial 
               Alignment       =   1  'Right Justify
               Height          =   360
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   600
               Width           =   2055
            End
            Begin ImpulseButton.ISButton CmdSearchTrans 
               Height          =   360
               Left            =   990
               TabIndex        =   49
               Top             =   570
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   635
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
               ButtonImage     =   "FrmReturnpurchases.frx":058E
            End
            Begin ImpulseButton.ISButton CmdOpenTrans 
               Height          =   360
               Left            =   120
               TabIndex        =   50
               Top             =   570
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   635
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
               ButtonImage     =   "FrmReturnpurchases.frx":0928
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·—ﬁ„ «·ÌœÊÌ"
               Height          =   270
               Index           =   36
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   114
               Top             =   255
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ ›« Ê—… «·‘—«¡"
               Height          =   315
               Index           =   9
               Left            =   3870
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   630
               Width           =   2265
            End
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   10515
            TabIndex        =   54
            Top             =   750
            Width           =   3030
            _ExtentX        =   5345
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   10515
            TabIndex        =   55
            Top             =   1140
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   300
            Left            =   11130
            TabIndex        =   57
            Top             =   405
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   529
            _Version        =   393216
            Format          =   240975873
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   7635
            TabIndex        =   83
            Top             =   0
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            CausesValidation=   0   'False
            Height          =   420
            Index           =   10
            Left            =   1890
            TabIndex        =   92
            Top             =   1860
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   741
            ButtonStyle     =   1
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   7635
            TabIndex        =   93
            Top             =   780
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCDocTypes 
            Height          =   315
            Left            =   7635
            TabIndex        =   95
            Top             =   375
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            CausesValidation=   0   'False
            Height          =   420
            Index           =   8
            Left            =   480
            TabIndex        =   100
            Top             =   1860
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   741
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "⁄—÷ ”‰œ «·’—›"
            BackColor       =   14871017
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
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   315
            Left            =   7620
            TabIndex        =   123
            Top             =   1170
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcCurrency 
            Height          =   315
            Left            =   1110
            TabIndex        =   145
            Top             =   1080
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„…"
            Height          =   270
            Index           =   46
            Left            =   2160
            TabIndex        =   151
            Top             =   90
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Œ’„"
            Height          =   255
            Index           =   45
            Left            =   4515
            TabIndex        =   150
            Top             =   45
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "««·⁄„·…"
            Height          =   405
            Index           =   44
            Left            =   2145
            RightToLeft     =   -1  'True
            TabIndex        =   146
            Top             =   1155
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ VAT"
            Height          =   315
            Index           =   43
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   143
            Top             =   1095
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·”»»"
            Height          =   285
            Index           =   78
            Left            =   9600
            RightToLeft     =   -1  'True
            TabIndex        =   140
            Top             =   1920
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„… «·„÷«›…"
            Height          =   285
            Index           =   77
            Left            =   13290
            RightToLeft     =   -1  'True
            TabIndex        =   139
            Top             =   1920
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‰«ﬁ·"
            Height          =   300
            Index           =   41
            Left            =   9720
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Top             =   1560
            Width           =   660
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„‰œÊ»"
            Height          =   300
            Index           =   40
            Left            =   9750
            RightToLeft     =   -1  'True
            TabIndex        =   125
            Top             =   1170
            Width           =   660
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·„‰œÊ»"
            Height          =   195
            Index           =   72
            Left            =   10965
            RightToLeft     =   -1  'True
            TabIndex        =   124
            Top             =   1380
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„  ›« Ê—… «·„Ê—œ"
            Height          =   495
            Index           =   39
            Left            =   6600
            RightToLeft     =   -1  'True
            TabIndex        =   121
            Top             =   1080
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„  «„— «· Ê—Ìœ"
            Height          =   375
            Index           =   38
            Left            =   6600
            RightToLeft     =   -1  'True
            TabIndex        =   120
            Top             =   1560
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”»» «·«—Ã«⁄"
            Height          =   405
            Index           =   37
            Left            =   2145
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   1500
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «–‰ «·’—› "
            Height          =   315
            Index           =   35
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   1500
            Width           =   1740
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·”‰œ"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   9600
            TabIndex        =   96
            Top             =   375
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·Œ“‰…"
            Height          =   285
            Index           =   22
            Left            =   9750
            RightToLeft     =   -1  'True
            TabIndex        =   94
            Top             =   810
            Width           =   765
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„  «·ﬁÌœ"
            Height          =   375
            Index           =   25
            Left            =   4020
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   1860
            Width           =   1275
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   390
            Left            =   9435
            TabIndex        =   84
            Top             =   0
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   285
            Index           =   3
            Left            =   12915
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   1155
            Width           =   1650
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ÿ—Ìﬁ… «·ﬁ»÷"
            Height          =   270
            Index           =   4
            Left            =   12915
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   1500
            Width           =   1650
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«· «—ÌŒ "
            Height          =   270
            Index           =   6
            Left            =   12915
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   435
            Width           =   1650
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·⁄„·Ì…"
            Height          =   270
            Index           =   7
            Left            =   12915
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   75
            Width           =   1650
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«”„ «·„Ê—œ"
            Height          =   285
            Index           =   5
            Left            =   12915
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   780
            Width           =   1650
         End
      End
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   4830
         Left            =   15
         TabIndex        =   10
         Top             =   2865
         Width           =   14625
         _cx             =   25797
         _cy             =   8520
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
         Caption         =   "«·√’‰«›|«·√Ê—«ﬁ «·„«·Ì…|«·ﬁÌ„… «·„÷«›…"
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
         Picture(0)      =   "FrmReturnpurchases.frx":0CC2
         Picture(1)      =   "FrmReturnpurchases.frx":105C
         Flags(1)        =   2
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4365
            Index           =   1
            Left            =   45
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   45
            Width           =   14535
            _cx             =   25638
            _cy             =   7699
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
            AutoSizeChildren=   8
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
            GridRows        =   6
            GridCols        =   6
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmReturnpurchases.frx":13F6
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin MSComctlLib.Toolbar TBar 
               Height          =   630
               Left            =   510
               TabIndex        =   63
               Top             =   3975
               Width           =   13485
               _ExtentX        =   23786
               _ExtentY        =   1111
               ButtonWidth     =   609
               ButtonHeight    =   1005
               Appearance      =   1
               _Version        =   393216
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   930
               Index           =   5
               Left            =   30
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   30
               Width           =   14475
               _cx             =   25532
               _cy             =   1640
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
               Begin VB.TextBox TxtShortName 
                  Height          =   240
                  Left            =   4785
                  TabIndex        =   128
                  Top             =   0
                  Width           =   6825
               End
               Begin VB.TextBox TxtPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   255
                  Left            =   915
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   135
                  Top             =   645
                  Width           =   1680
               End
               Begin VB.TextBox TxtSerial 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   4455
                  MaxLength       =   20
                  RightToLeft     =   -1  'True
                  TabIndex        =   133
                  Top             =   645
                  Width           =   2235
               End
               Begin VB.TextBox TxtQuantity 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   2640
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   134
                  Top             =   645
                  Width           =   1770
               End
               Begin VB.ComboBox CboItemCase 
                  Height          =   315
                  Left            =   6735
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   132
                  Top             =   645
                  Width           =   1965
               End
               Begin MSDataListLib.DataCombo DCboItemsName 
                  Height          =   315
                  Left            =   8700
                  TabIndex        =   131
                  Top             =   645
                  Width           =   2820
                  _ExtentX        =   4974
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCboItemsCode 
                  Height          =   315
                  Left            =   11565
                  TabIndex        =   130
                  Top             =   645
                  Width           =   2910
                  _ExtentX        =   5133
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdAdd 
                  Height          =   405
                  Left            =   90
                  TabIndex        =   0
                  Top             =   585
                  Width           =   735
                  _ExtentX        =   1296
                  _ExtentY        =   714
                  ButtonStyle     =   1
                  ButtonPositionImage=   4
                  Caption         =   ""
                  BackColor       =   14871017
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BackStyle       =   0
                  ButtonImage     =   "FrmReturnpurchases.frx":1494
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   -2147483637
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   -2147483637
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·»ÕÀ «·”—Ì⁄"
                  Height          =   240
                  Index           =   97
                  Left            =   11970
                  TabIndex        =   129
                  Top             =   0
                  Width           =   1275
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”⁄—"
                  Height          =   300
                  Index           =   26
                  Left            =   1005
                  RightToLeft     =   -1  'True
                  TabIndex        =   43
                  Top             =   285
                  Width           =   1590
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂ„Ì…"
                  Height          =   300
                  Index           =   27
                  Left            =   2910
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   285
                  Width           =   1455
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ì—Ì«·"
                  Height          =   300
                  Index           =   28
                  Left            =   4545
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   285
                  Width           =   2145
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·… «·’‰›"
                  Height          =   300
                  Index           =   29
                  Left            =   6960
                  RightToLeft     =   -1  'True
                  TabIndex        =   40
                  Top             =   285
                  Width           =   1740
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈”„ «·’‰›"
                  Height          =   300
                  Index           =   30
                  Left            =   9015
                  RightToLeft     =   -1  'True
                  TabIndex        =   39
                  Top             =   285
                  Width           =   2550
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·’‰›"
                  Height          =   300
                  Index           =   31
                  Left            =   11475
                  RightToLeft     =   -1  'True
                  TabIndex        =   38
                  Top             =   285
                  Width           =   2640
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid FG 
               Height          =   2985
               Left            =   30
               TabIndex        =   44
               Top             =   975
               Width           =   14475
               _cx             =   25532
               _cy             =   5265
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
               Rows            =   2
               Cols            =   18
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmReturnpurchases.frx":182E
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
            Begin VB.Label LblItemsCount 
               Alignment       =   2  'Center
               BackColor       =   &H00404040&
               ForeColor       =   &H0000FFFF&
               Height          =   240
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   4095
               Width           =   450
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4365
            Index           =   2
            Left            =   15270
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   45
            Width           =   14535
            _cx             =   25638
            _cy             =   7699
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
            AutoSizeChildren=   8
            BorderWidth     =   0
            ChildSpacing    =   0
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
            GridRows        =   12
            GridCols        =   6
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmReturnpurchases.frx":1B1B
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.Frame Fram 
               BackColor       =   &H00E2E9E9&
               BorderStyle     =   0  'None
               Height          =   585
               Index           =   0
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   510
               Width           =   4065
               Begin VB.CheckBox XPChkPayType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰ﬁœ«"
                  Height          =   360
                  Index           =   0
                  Left            =   7440
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   120
                  Width           =   885
               End
               Begin VB.TextBox XPTxtSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   0
                  Left            =   3600
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   25
                  Top             =   150
                  Width           =   1185
               End
               Begin VB.TextBox XPTxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   0
                  Left            =   5610
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   1
                  Top             =   150
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·”·"
                  Height          =   210
                  Index           =   12
                  Left            =   4980
                  RightToLeft     =   -1  'True
                  TabIndex        =   27
                  Top             =   180
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   210
                  Index           =   13
                  Left            =   6900
                  RightToLeft     =   -1  'True
                  TabIndex        =   26
                  Top             =   210
                  Width           =   465
               End
            End
            Begin VB.Frame Fram 
               BackColor       =   &H00E2E9E9&
               BorderStyle     =   0  'None
               Height          =   1440
               Index           =   1
               Left            =   4065
               RightToLeft     =   -1  'True
               TabIndex        =   19
               Top             =   510
               Width           =   8340
               Begin VB.CheckBox XPChkPayType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "¬Ã· "
                  Height          =   360
                  Index           =   1
                  Left            =   3960
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   0
                  Width           =   885
               End
               Begin VB.TextBox XPTxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   1
                  Left            =   2130
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   2
                  Top             =   90
                  Width           =   1185
               End
               Begin VB.TextBox XPTxtSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   1
                  Left            =   150
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   20
                  Top             =   90
                  Width           =   1185
               End
               Begin MSComCtl2.DTPicker DtpDelayDate 
                  Height          =   330
                  Left            =   150
                  TabIndex        =   3
                  Top             =   480
                  Width           =   1545
                  _ExtentX        =   2725
                  _ExtentY        =   582
                  _Version        =   393216
                  Format          =   241041409
                  CurrentDate     =   38784
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·”·"
                  Height          =   210
                  Index           =   14
                  Left            =   1500
                  RightToLeft     =   -1  'True
                  TabIndex        =   23
                  Top             =   150
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   210
                  Index           =   15
                  Left            =   3300
                  RightToLeft     =   -1  'True
                  TabIndex        =   22
                  Top             =   150
                  Width           =   465
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
                  Height          =   210
                  Index           =   21
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   21
                  Top             =   540
                  Width           =   1155
               End
            End
            Begin VB.Frame Fram 
               BackColor       =   &H00E2E9E9&
               BorderStyle     =   0  'None
               Height          =   945
               Index           =   2
               Left            =   4065
               RightToLeft     =   -1  'True
               TabIndex        =   14
               Top             =   1950
               Width           =   8340
               Begin VB.CheckBox XPChkPayType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‘Ìﬂ"
                  Height          =   255
                  Index           =   2
                  Left            =   3720
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   0
                  Width           =   1125
               End
               Begin VB.TextBox XPTxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Index           =   2
                  Left            =   2970
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   6
                  Top             =   705
                  Width           =   975
               End
               Begin VB.TextBox XPTxtChqueNum 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2970
                  MaxLength       =   40
                  RightToLeft     =   -1  'True
                  TabIndex        =   4
                  Top             =   315
                  Width           =   975
               End
               Begin MSDataListLib.DataCombo DCboBankName 
                  Height          =   315
                  Left            =   60
                  TabIndex        =   5
                  Top             =   330
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _Version        =   393216
                  Locked          =   -1  'True
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker XPDTPDueDate 
                  Height          =   345
                  Left            =   60
                  TabIndex        =   7
                  Top             =   705
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   609
                  _Version        =   393216
                  Format          =   241041409
                  CurrentDate     =   38784
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   210
                  Index           =   16
                  Left            =   4215
                  RightToLeft     =   -1  'True
                  TabIndex        =   18
                  Top             =   765
                  Width           =   465
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·‘Ìﬂ"
                  Height          =   210
                  Index           =   18
                  Left            =   4080
                  RightToLeft     =   -1  'True
                  TabIndex        =   17
                  Top             =   375
                  Width           =   735
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·»‰ﬂ"
                  Height          =   210
                  Index           =   17
                  Left            =   1875
                  RightToLeft     =   -1  'True
                  TabIndex        =   16
                  Top             =   375
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
                  Height          =   210
                  Index           =   19
                  Left            =   1620
                  RightToLeft     =   -1  'True
                  TabIndex        =   15
                  Top             =   765
                  Width           =   1155
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
               Height          =   510
               Index           =   20
               Left            =   4065
               RightToLeft     =   -1  'True
               TabIndex        =   78
               Top             =   0
               Width           =   8340
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   4365
            Left            =   15570
            TabIndex        =   153
            TabStop         =   0   'False
            Top             =   45
            Width           =   14535
            _cx             =   25638
            _cy             =   7699
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
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   375
               Left            =   240
               TabIndex        =   159
               Top             =   120
               Width           =   1215
            End
            Begin VB.TextBox TxtValueAdded 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   5730
               TabIndex        =   155
               Top             =   3960
               Width           =   2265
            End
            Begin VB.CheckBox ChecVAT 
               Alignment       =   1  'Right Justify
               Caption         =   " ÕœÌœ «·ﬂ·"
               Height          =   210
               Left            =   11415
               RightToLeft     =   -1  'True
               TabIndex        =   154
               Top             =   105
               Width           =   1050
            End
            Begin VSFlex8UCtl.VSFlexGrid VatGrid 
               Height          =   3285
               Left            =   135
               TabIndex        =   156
               Tag             =   "1"
               Top             =   600
               Width           =   14265
               _cx             =   25162
               _cy             =   5794
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
               Cols            =   9
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmReturnpurchases.frx":1C01
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
               BackStyle       =   0  'Transparent
               Caption         =   "«œŒ«· «·‰”»… «·ÌœÊÌ…"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   270
               Index           =   148
               Left            =   1440
               TabIndex        =   160
               Top             =   240
               Width           =   1800
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               Caption         =   "«’‰«› «·ﬁÌ„… «·„÷«›…"
               ForeColor       =   &H00FF0000&
               Height          =   180
               Left            =   11145
               TabIndex        =   158
               Top             =   105
               Width           =   3120
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   " «·«Ã„«·Ì"
               Height          =   300
               Index           =   104
               Left            =   8265
               TabIndex        =   157
               Top             =   3960
               Width           =   1080
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   615
         Left            =   15
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   15
         Width           =   14655
         _cx             =   25850
         _cy             =   1085
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
         Caption         =   "„—œÊœ«  «·„‘ —Ì«  "
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   0
         ChildSpacing    =   0
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
         Begin VB.CheckBox chkIsBranch 
            Caption         =   "»œÊ‰ ﬁÌœ"
            Height          =   225
            Index           =   1
            Left            =   4560
            TabIndex        =   164
            Top             =   120
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.CheckBox chkIsBranch 
            Caption         =   "»«·›—⁄"
            Height          =   225
            Index           =   0
            Left            =   5820
            TabIndex        =   163
            Top             =   120
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox txtPassword 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   3720
            PasswordChar    =   "*"
            TabIndex        =   162
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton cmdReSave 
            Caption         =   "÷»ÿ «·Õ—ﬂ« "
            Height          =   285
            Left            =   9900
            TabIndex        =   161
            Top             =   120
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.TextBox TxtItemsIDes 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   152
            Top             =   0
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox TXTNoteID 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   8535
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   0
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox oldtxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5355
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   0
            Visible         =   0   'False
            Width           =   1875
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1980
            TabIndex        =   73
            Top             =   120
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   609
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
            ButtonImage     =   "FrmReturnpurchases.frx":1D33
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
            Left            =   1050
            TabIndex        =   74
            Top             =   120
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   609
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
            ButtonImage     =   "FrmReturnpurchases.frx":20CD
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
            Left            =   2910
            TabIndex        =   75
            Top             =   120
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   609
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
            ButtonImage     =   "FrmReturnpurchases.frx":2467
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
            Left            =   150
            TabIndex        =   76
            Top             =   120
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   609
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
            ButtonImage     =   "FrmReturnpurchases.frx":2801
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin MSComCtl2.DTPicker txtToDateReSave 
            Height          =   315
            Left            =   6930
            TabIndex        =   165
            Top             =   120
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   556
            _Version        =   393216
            Format          =   239468545
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker txtFromDateReSave 
            Height          =   315
            Left            =   8520
            TabIndex        =   166
            Top             =   120
            Visible         =   0   'False
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            Format          =   239468545
            CurrentDate     =   38784
         End
         Begin VB.Label LBLGross 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   375
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   147
            Top             =   120
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Image ImgFavorites 
            Height          =   390
            Left            =   10320
            Picture         =   "FrmReturnpurchases.frx":2B9B
            Stretch         =   -1  'True
            Top             =   0
            Width           =   525
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   435
            Index           =   24
            Left            =   3795
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   0
            Width           =   7245
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   495
         Index           =   0
         Left            =   15
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   8040
         Width           =   14655
         _cx             =   25850
         _cy             =   873
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
         AutoSizeChildren=   8
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
         GridRows        =   1
         GridCols        =   18
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmReturnpurchases.frx":6803
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin ImpulseButton.ISButton Cmd 
            Height          =   315
            Index           =   0
            Left            =   13080
            TabIndex        =   105
            Top             =   90
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   556
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
            ColorToggledText=   -2147483631
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   315
            Index           =   1
            Left            =   11400
            TabIndex        =   106
            Top             =   90
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   556
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
            Height          =   315
            Index           =   2
            Left            =   9795
            TabIndex        =   107
            Top             =   90
            Width           =   1485
            _ExtentX        =   2619
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
            Height          =   315
            Index           =   3
            Left            =   8295
            TabIndex        =   108
            Top             =   90
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
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
            Height          =   315
            Index           =   4
            Left            =   6450
            TabIndex        =   109
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
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
            Height          =   315
            Index           =   5
            Left            =   4905
            TabIndex        =   110
            Top             =   90
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
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
            Height          =   315
            Index           =   6
            Left            =   90
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   90
            Width           =   1425
            _ExtentX        =   2514
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
            Height          =   315
            Index           =   7
            Left            =   3255
            TabIndex        =   112
            Top             =   90
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
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
            Height          =   315
            Left            =   1620
            TabIndex        =   113
            Top             =   90
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   556
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
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«Ã„«·Ì «·ﬂ„Ì…"
      Height          =   315
      Index           =   33
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   98
      Top             =   0
      Width           =   990
   End
End
Attribute VB_Name = "FrmReturnpurchases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim NewGrid As New ClsGrid
Dim ReturnReport As ClsReturnBackReport
Dim cSearchDcbo(3) As clsDCboSearch
Dim usedaccount As Integer
Public BolPrint As Boolean
Dim general_noteid As Long
Dim RsNotesGeneral As ADODB.Recordset
Dim CurrentVoucherNo As String
Dim CurrentVoucherSerialNo As String
Dim DateChanged As Boolean
Dim TxtNoteSerial1V As String
Dim IsSaveWithOutMsg As Boolean
Public Sub RetriveSerials(ItemID As String, _
                          ItemName As String, _
                          seriallist As String, _
                          currentrow As Long, Optional Price As Double, Optional UnitID As Double = 1, Optional UnitName As String)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    On Error GoTo ErrTrap
    Dim strInputString As String
    Dim strFilterText As String
    Dim astrSplitItems() As String
    Dim astrFilteredItems() As String
    Dim strFilteredString As String
    Dim intX As Integer
    strInputString = seriallist
    strFilterText = ","
 
    astrSplitItems = Split(strInputString, strFilterText)
    Dim i As Integer
 
   
    Num = currentrow

    '  For Num = currentrow To UBound(astrSplitItems)+currentrow
    
    Dim CurrentSerial As String
 
   
    '*****************************************************
    For intX = 0 To UBound(astrSplitItems)
   FG.cell(flexcpData, Num, FG.ColIndex("Code")) = ItemID
   FG.TextMatrix(Num, FG.ColIndex("Code")) = ItemID
   
        FG.TextMatrix(Num, FG.ColIndex("Name")) = ItemID
        
        
         FG.TextMatrix(Num, FG.ColIndex("UnitID")) = ItemID
        FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = 1
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = 0
            FG.TextMatrix(Num, FG.ColIndex("guaranteeTime")) = 0
    
           'FG.TextMatrix(I, FG.ColIndex("HaveSerial")) = True
         
        FG.TextMatrix(Num, FG.ColIndex("Count")) = 1
        FG.TextMatrix(Num, FG.ColIndex("Serial")) = astrSplitItems(intX)
        
        FG.cell(flexcpData, Num, FG.ColIndex("UnitID")) = UnitID
        FG.TextMatrix(Num, FG.ColIndex("UnitID")) = UnitName
             FG.TextMatrix(i, FG.ColIndex("HaveSerial")) = True
             
        
If val(Price) > 0 Then
            FG.TextMatrix(Num, FG.ColIndex("price")) = Price
        End If
        
        '      RsDetails.MoveNext
        '      Debug.Print Num
        FG.rows = FG.rows + 1
 
        Num = Num + 1
    If intX = UBound(astrSplitItems) Then
    NewGrid.Calculate Num
    NewGrid.bassprofit = True
    NewGrid.DtpBillDate_Change
        End If
    Next
     
     
    TxtFillData.text = "F"
    TxtFillData_Change
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub


Public Sub RetriveSerialsx(ItemID As String, _
                          ItemName As String, _
                          seriallist As String, _
                                                    currentrow As Long, Optional Price As Double, Optional UnitID As Double = 1, Optional UnitName As String)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    On Error GoTo ErrTrap
    Dim strInputString As String
    Dim strFilterText As String
    Dim astrSplitItems() As String
    Dim astrFilteredItems() As String
    Dim strFilteredString As String
    Dim intX As Integer
    strInputString = seriallist
    strFilterText = ","
 
    astrSplitItems = Split(strInputString, strFilterText)
    Dim i As Integer
    ' For i = 1 To Fg.Rows - 2
    '        If Fg.TextMatrix(i, Fg.ColIndex("Code")) = ItemID Then
    '         Me.Fg.RemoveItem (i)
    '         i = 1
    '        End If
    'NewGrid.Grid_AfterEdit Num, Fg.ColIndex("Code")
    ' Next i
   
    Num = currentrow

    '  For Num = currentrow To UBound(astrSplitItems)+currentrow
    For intX = 0 To UBound(astrSplitItems)
   
        FG.TextMatrix(Num, FG.ColIndex("Code")) = ItemID
        NewGrid.Grid_AfterEdit Num, FG.ColIndex("Code")
        ' FG.TextMatrix(Num, FG.ColIndex("Name")) = itemname
        FG.TextMatrix(Num, FG.ColIndex("Count")) = 1
        FG.TextMatrix(Num, FG.ColIndex("Serial")) = astrSplitItems(intX)
  
If val(Price) > 0 Then
            FG.TextMatrix(Num, FG.ColIndex("price")) = Price
        End If
        


        '      RsDetails.MoveNext
        '      Debug.Print Num
        FG.rows = FG.rows + 1
 
        Num = Num + 1
    Next
 
    TxtFillData.text = "F"
    TxtFillData_Change
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub
Function SaveItemsData(Optional Transaction_ID As String = 0)
If SystemOptions.WorkWithItemsDetails = False Then Exit Function
       Dim RsgGrantee    As New ADODB.Recordset
    Dim strInputString As String
    Dim strFilterText As String
    Dim astrSplitItems() As String
    Dim astrFilteredItems() As String
    Dim strFilteredString As String
    Dim intX As Integer
    Dim AllDes As String
    Dim RowNum As Integer
    Dim StrSQL As String
    strFilterText = ","
    Set RsgGrantee = New ADODB.Recordset
    Cn.Execute "delete ItemsDetails   where Transaction_ID= " & (Me.XPTxtBillID.text)
    
  '  RsgGrantee.Open "TBLRegularMaint", Cn, adOpenStatic, adLockOptimistic, adCmdTable

   StrSQL = "SELECT    * from  ItemsDetails Where (1 = -1)"
   RsgGrantee.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     
 
    Dim strFilterText1 As String
      Dim UnitName As String
    Dim ttypename As String
     Dim typename As String
 
 
 
 
    Dim inty As Integer
    Dim intervalstr As String
Dim Name As String
Dim NameE As String
Dim Remarks As String
Dim NooFRows As Double
    
     Dim astrSplitItems1() As String
 
    strFilterText = "&&"
         strFilterText1 = "@@"
     
    For RowNum = 1 To FG.rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            
           If FG.TextMatrix(RowNum, FG.ColIndex("ItemsDetailsNewidea")) <> "" Then
                AllDes = FG.TextMatrix(RowNum, FG.ColIndex("ItemsDetailsNewidea"))
                astrSplitItems = Split(AllDes, strFilterText)
         NooFRows = UBound(astrSplitItems) + 1
                For intX = 0 To NooFRows - 2
             
                
                          RsgGrantee.AddNew
                         astrSplitItems1 = Split(astrSplitItems(intX), strFilterText1)
                         RsgGrantee("ItemDetailedCode").value = (astrSplitItems1(0))
                         RsgGrantee("ParrtNoCode").value = (astrSplitItems1(1))
                         RsgGrantee("count").value = val(astrSplitItems1(2))
                         RsgGrantee("unitid").value = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", 1, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))  ' val(astrSplitItems1(3))
                         RsgGrantee("ColorID").value = val(astrSplitItems1(4))
                         RsgGrantee("sizeid").value = val(astrSplitItems1(5))
                         RsgGrantee("ClassId").value = val(astrSplitItems1(6))
                         RsgGrantee("ProductionDate").value = IIf(IsDate((astrSplitItems1(7))), astrSplitItems1(7), Null)
                         RsgGrantee("ExpireDate").value = IIf(IsDate((astrSplitItems1(8))), astrSplitItems1(8), Null)
                        RsgGrantee("Transaction_ID").value = val(Me.XPTxtBillID.text)
                        RsgGrantee("ItemId").value = FG.TextMatrix(RowNum, FG.ColIndex("Code"))
                       RsgGrantee("EffectN").value = -1
                    RsgGrantee.update
                                    Next intX
                Else
                
                If FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode")) <> "" Then
                RsgGrantee.AddNew
              RsgGrantee("ParrtNoCode").value = FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode"))
            RsgGrantee("count").value = FG.TextMatrix(RowNum, FG.ColIndex("Count"))
            RsgGrantee("unitid").value = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
          RsgGrantee("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
            RsgGrantee("sizeid").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
            RsgGrantee("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
            RsgGrantee("Transaction_ID").value = val(Me.XPTxtBillID.text)
           RsgGrantee("ItemId").value = FG.TextMatrix(RowNum, FG.ColIndex("Code"))
          RsgGrantee("ItemDetailedCode").value = FG.TextMatrix(RowNum, FG.ColIndex("ItemDetailedCode"))
          RsgGrantee("EffectN").value = -1
           RsgGrantee.update
                  
         End If
         
                  
                   
                   End If
                   

 
                
  
                    
            End If

       

    Next RowNum


End Function
Private Sub C1Elastic6_DblClick()
    On Error GoTo ErrTrap

    If Me.WindowState = vbNormal Then
        Me.WindowState = vbMaximized
    Else
        Me.WindowState = vbNormal
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub CboRetrunType_Change()

    If CboRetrunType.ListIndex = 0 Then
        lbl(9).Enabled = True
        Me.TxtInvSerial.Enabled = True
        Me.CmdOpenTrans.Enabled = True
        Me.CmdSearchTrans.Enabled = True
    ElseIf Me.CboRetrunType.ListIndex = 1 Then
        lbl(9).Enabled = False
        Me.TxtInvSerial.Enabled = False
        Me.CmdOpenTrans.Enabled = False
        Me.CmdSearchTrans.Enabled = False
        TxtInvSerial.text = ""
    End If
    If Me.TxtModFlg.text <> "R" Then
    If val(CboRetrunType.ListIndex) = 0 Then
 NewGrid.ReturnTyp = 2
 Else
 NewGrid.ReturnTyp = 1
 VatGrid.Clear flexClearScrollable, flexClearEverything
           VatGrid.rows = 1
 End If
 End If
  NewGrid.DtpBillDate_Change
NewGrid.Calculate 1, , , True
End Sub

Private Sub CboRetrunType_Click()
    CboRetrunType_Change
End Sub

Private Sub ChecVAT_Click()
  Dim i As Integer
If Me.TxtModFlg.text <> "R" Then
    If ChecVAT.value = vbChecked Then

        With Me.VatGrid
 
            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("select")) = True
            Next i

        End With

    Else

        With Me.VatGrid

            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("select")) = False
            Next i

        End With

    End If
    RelinVatGrid
    End If
End Sub

Private Sub Cmd_Click(Index As Integer)
    Dim AskOption As Boolean
    Dim intDef As Integer
    Dim Msg As String

    BolPrint = True
' On Error GoTo ErrTrap

    Select Case Index
    
        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            clear_all Me
            TxtModFlg.text = "N"
           VatGrid.Clear flexClearScrollable, flexClearEverything
           VatGrid.rows = 1
            XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            TxtTransSerial.text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=5"))
            Me.DCboUserName.BoundText = user_id
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSup", 1)
            DBCboClientName.BoundText = intDef
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultPurchaseStore", 1)
          '  DCboStoreName.BoundText = intDef
            Me.DcboBox.BoundText = 1
            NewGrid.GridDefaultValue 1
            XPTab301.CurrTab = 0
            'FG.SetFocus
            FG.Col = FG.ColIndex("Code")
            FG.Row = FG.rows - 1
            CboRetrunType.ListIndex = 1
            
            '·«»œ „‰ «· ›⁄Ì·
            'If Voucher_coding(Val(my_branch), XPDtbBill.value, X, Y) = "" And Val(my_branch) <> 0 Then
            'TxtNoteSerial1.locked = False
            'Else
            ' TxtNoteSerial1.locked = True
            '
            'End If
Dim dstore As Integer
            Dim dBox As Integer
            Dim usertype As Integer
            Dim EmpID As Integer
            Dim userbranchid As Integer
            Dim StoreId1 As Integer
            Dim boxid1 As Integer
            Dim dstore1 As Integer
            Dim CUSTID1 As Integer
            'GetBranchData branch_id, dstore, dBox
                 
            GetUserData user_id, usertype, userbranchid, , dBox, , EmpID, , , dstore1, CUSTID1, boxid1
     
            If usertype <> 0 Then 'admin
                dcBranch.Enabled = False
 
                DCboStoreName.Enabled = True
              '  TxtStoreID.Enabled = False
                Me.DCboStoreName.BoundText = dstore1
                Me.DcboBox.BoundText = boxid1
                Me.DBCboClientName.BoundText = CUSTID1
                
            Else
                dcBranch.Enabled = True
 
                DCboStoreName.Enabled = True
 
                Me.dcBranch.BoundText = ""
                Me.DCboStoreName.BoundText = ""
'                TxtStoreID.Enabled = True
            End If
                    
                    
        

      If SystemOptions.usertype <> UserAdminAll Then
                            If checkmanyBranches = False Then
                                   Me.dcBranch.Enabled = True
                                   Else
                                    Me.dcBranch.Enabled = True
                             End If
                    
                      If checkmanyStores = False Then
                                   Me.DCboStoreName.Enabled = True
                                    
                                   Else
                                   Me.DCboStoreName.Enabled = True
 
                             End If
                                  
           End If
            Me.dcBranch.BoundText = Current_branch
            TxtNoteSerial1V = ""
            DcCurrency.BoundText = MainCurrency()
        
         If SystemOptions.DefaultIsCreditPurchaseRet = False Then
            CboPayMentType.ListIndex = 0
   Else
   CboPayMentType.ListIndex = 1
    End If
   
   
        Case 1
        
        
                             If ChekClodePeriod(XPDtbBill.value) = True Then
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

            If SystemOptions.usertype = UserNormal Then
'                Msg = "·Ì” ·ﬂ Õﬁ  ⁄œÌ· ›Ï «·›Ê« Ì—"
'                MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.title
'                Exit Sub
            End If

            TxtModFlg.text = "E"
        If Trim(txtPassword) <> "Alex2025" Then
            Me.DCboUserName.BoundText = user_id
         End If
            CuurentLogdata

If val(CboRetrunType.ListIndex) = 0 Then
 NewGrid.ReturnTyp = 2
 Else
 NewGrid.ReturnTyp = 1
 End If
  NewGrid.CboDiscount_Type_Change
        Case 2
        
        
       '***********************
                 If (val(DBCboClientName.BoundText) = 1 Or val(DBCboClientName.BoundText) = 2) And CboPayMentType.ListIndex = 1 Then
                 
                                If SystemOptions.UserInterface = ArabicInterface Then
                                                  MsgBox "·« Ì„ﬂ‰ «Œ Ì«— „Ê—œ ‰ﬁœÌ …«·«—Ã«⁄ «Ã· "
                                        Else
                                               MsgBox "Please Change  Payment Type"
                                 End If
              Exit Sub
            
              
                End If
       '******************
        
        
                             If ChekClodePeriod(XPDtbBill.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
                  
                  
            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Departement"
                Else
                    Msg = "Õœœ «·›—⁄ «Ê·« "
                End If
              
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                dcBranch.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
    
            my_branch = Me.dcBranch.BoundText
      
            '            If Me.TxtModFlg.text = "N" Then
             
            '             End If
If val(Me.TxtValueAdded.text) > 0 Then
If GetValueAddedAccount(XPDtbBill.value, , , 1, 5) = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›…"
Else
MsgBox "Value added account not specified"
End If
Exit Sub
End If
End If

            SaveData

        Case 3
            Call Undo

        Case 4
                             If ChekClodePeriod(XPDtbBill.value) = True Then
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

      '      If SystemOptions.usertype = UserNormal Then
      '          Msg = "·Ì” ·ﬂ Õﬁ Õ–› ›Ï «·›Ê« Ì—"
      '          MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.title
      '          Exit Sub
      '      End If

            Del_TransAction

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            FrmBuySearch.DealingForm = Returntransaction
            FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ „— Ã⁄ «·„‘ —Ì« "
            FrmBuySearch.show vbModal

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowMe", False)

            If AskOption = False Then
                FrmPrintOptions.show vbModal
            End If

            If BolPrint = False Then
                Exit Sub
            End If

            PrintReport

        Case 6
            Unload Me

        Case 8
     
        FrmOut.XPBtnMove_Click (2)
            FrmOut.Retrive val(Me.Text1.text)
            

        Case 10
            ShowGL_cc TxtNoteSerial.text, , 200, val(Me.TXTNoteID.text)

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub CmdOpenTrans_Click()
    Dim Msg As String
    Dim FrmNewSales As FrmBillBuy

    If val(Me.TxtInvSerial.text) = 0 Then
        Msg = "»—Ã«¡ ﬂ «»… —ﬁ„ «·›« Ê—… ·Ì „ ⁄—÷Â«..!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    ''Me.TxtInvID.text = GetTransIDSerial(1, , Trim(Me.TxtInvSerial.text), 1, , Val(TxtInvID.text))
    ''If Val(Me.TxtInvID.text) = 0 Then
    ''    Msg = "·« ÊÃœ ›« Ê—… »Â–« «·—ﬁ„ ..!!"
    ''    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    ''    Exit Sub
    'Else
    Set FrmNewSales = New FrmBillBuy
    FrmNewSales.show
    FrmNewSales.Retrive val(Me.TxtInvID.text)
    FrmNewSales.ZOrder 0
    'End If
End Sub

Private Sub cmdReSave_Click()
   Dim s As String
   Dim i As Double
     XPBtnMove_Click (2)
    DoEvents
For i = 1 To rs.RecordCount
  Cmd_Click (1)
  DoEvents
  NewGrid.DtpBillDate_Change
  DoEvents
  DoEvents
  IsSaveWithOutMsg = True
  DoEvents
  Cmd_Click (2)
  DoEvents
  XPBtnMove_Click (0)
Next i
    IsSaveWithOutMsg = False
    MsgBox " „ «·Õ›Ÿ"
    Cmd(2).Enabled = True
End Sub

Private Sub CmdSearchTrans_Click()
    ' ›« Ê—… „»Ì⁄« 
 If CboRetrunType.ListIndex = 0 Then
   
    FrmBuySearch.DealingForm = PurchaseTransaction
    Set FrmBuySearch.ExtraRetrunObject = Me.TxtInvSerial
    FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ›« Ê—… ‘—«¡"
     Load FrmBuySearch
   ' FrmBuySearch.DCboClientsName.BoundText = Me.DBCboClientName.BoundText
    
    'FrmBuySearch.show
    End If
    
End Sub

 Function Retrive_Items_data1()
    Dim StrSQL  As String
    Dim row_count As Long
    Dim Num As Long
    Dim i As Long
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    StrSQL = "select * from TblItems where ItemID in(" & TxtItemsIDes.text & ")"
    rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
   If rs2.RecordCount > 0 Then
        
        If FG.TextMatrix(FG.rows - 1, FG.ColIndex("Code")) = "" Then
      FG.rows = FG.rows - 1
        End If
     With FG
     row_count = FG.rows
       rs2.MoveFirst
       .rows = rs2.RecordCount + .rows
        For Num = row_count To .rows - 1 'RsDetails.RecordCount
        .TextMatrix(Num, .ColIndex("Code")) = IIf(IsNull(rs2("ItemID").value), 0, rs2("ItemID").value)
      
        rs2.MoveNext
        Next Num
        For i = row_count To .rows - 1 'RsDetails.RecordCount
          NewGrid.Grid_AfterEdit i, .ColIndex("Code")
        Next i
        NewGrid.Grid_AfterEdit row_count, .ColIndex("Code")
    End With
    End If


End Function

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmCompanySearch.lblSearchtype.Caption = 1000
        FrmCompanySearch.show vbModal
    
    End If
    
    
        If KeyCode = vbKeyF5 Then
      '  ReloadCombos

    End If
    
End Sub

Private Sub DCboItemsCode_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then

        Load FrmItemSearch
        FrmItemSearch.RetrunType = 909
        FrmItemSearch.show vbModal
    End If
End Sub

Private Sub DCboStoreName_Change()

 TxtStoreID.text = getStoreCoding(val(DCboStoreName.BoundText))
 
    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then

     If CheckStoreCoding(val(dcBranch.BoundText), 15) = True Or CheckStoreCoding(val(dcBranch.BoundText), 10) = True Then
     TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
    CurrentVoucherNo = ""
    TxtNoteSerial1V = ""
    DateChanged = True
     End If
     
    End If
End Sub

Private Sub Dcbranch_Change()
Dim Dcombos As New ClsDataCombos
 
    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then

       Dcombos.GetDocTypebyid Me.DCDocTypes, 5, val(Me.dcBranch.BoundText)
       DateChanged = True
    End If
End Sub

Private Sub Dcbranch_Click(Area As Integer)
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
        TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
    CurrentVoucherNo = ""
    TxtNoteSerial1V = ""
    DateChanged = True
    
End Sub

Private Sub Dcbranch_GotFocus()
Dcbranch_Click (0)
End Sub

Private Sub DcbTyp_Change()
If Me.TxtModFlg.text <> "R" Then
If val(Dcbtyp.ListIndex) = -1 Then
NewGrid.DtpBillDate_Change
Else
RelinVatGrid
End If
End If
End Sub

Private Sub DcbTyp_Click()
DcbTyp_Change
End Sub

Private Sub DcCurrency_Change()
DcCurrency_Click (0)
End Sub

Private Sub DcCurrency_Click(Area As Integer)
    If Me.TxtModFlg.text = "" Or Me.TxtModFlg.text = "R" Then Exit Sub
    If Me.DcCurrency.BoundText <> "" Then
        txt_Currency_rate.text = get_currency_rate(Me.DcCurrency.BoundText)
    Else
        txt_Currency_rate.text = 1
    End If
End Sub

Public Sub FG_AfterEdit(ByVal Row As Long, _
                        ByVal Col As Long)

    'XPTxtSum.text = FG.Aggregate(flexSTSum, 1, FG.ColIndex("Price"), FG.Rows - 1, FG.ColIndex("Price"))
    If Me.TxtModFlg <> "E" Then Exit Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    If Col = FG.ColIndex("Code") Or Col = FG.ColIndex("Name") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 230
    ElseIf Col = FG.ColIndex("UnitID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("UnitID")), , , , , , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 230
    ElseIf Col = FG.ColIndex("Count") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , (FG.TextMatrix(Row, FG.ColIndex("Count"))), , , , , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 230
    ElseIf Col = FG.ColIndex("Price") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , (FG.TextMatrix(Row, FG.ColIndex("Price"))), , , , , , , val(Me.TxtNoteSerial), val(Me.TxtNoteSerial1), 230
    ElseIf Col = FG.ColIndex("ColorID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("ColorID")), , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 230
    ElseIf Col = FG.ColIndex("ItemSize") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("ItemSize")), , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 230
    ElseIf Col = FG.ColIndex("ClassId") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("ClassId")), , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 230
    ElseIf Col = FG.ColIndex("DiscountType") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("DiscountType")), , Me.TxtNoteSerial, Me.TxtNoteSerial1, 230
    ElseIf Col = FG.ColIndex("DiscountVal") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , FG.TextMatrix(Row, FG.ColIndex("DiscountVal")), Me.TxtNoteSerial, Me.TxtNoteSerial1, 230

    End If

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////

End Sub

Private Sub Fg_DblClick()
    'FrmItemsDetails.Show
End Sub

Private Sub Form_Activate()
    'TxtTransSerial.SetFocus
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption



End Sub

Private Sub TxtFillData_Change()

    If TxtFillData.text = "F" Then
        FG_AfterEdit 1, 1
    End If

End Sub

Private Sub TxtInvID_Change()

    If Me.TxtInvID.text <> "" Then
        If Me.CboRetrunType.ListIndex = 0 Then
         '   Me.TxtInvSerial.text = GetTransIDSerial(1, val(Me.TxtInvID.text), , 1, , val(TxtInvID.text))
        End If
    End If

End Sub

Private Sub TxtInvSerial_Change()
        If Me.TxtModFlg.text <> "R" Then
         
         Dim Transaction_ID As Double
             
Transaction_ID = get_transactionData("NoteSerial1", TxtInvSerial.text, "Transaction_ID", 22)
TxtInvID = Transaction_ID
 Retrive_orders_data (val(Transaction_ID))
         NewGrid.Calculate 1, , , True
         
 If val(CboRetrunType.ListIndex) = 0 Then
 RetriveValueAddedData Transaction_ID
 RelinVatGrid
 Else
    VatGrid.Clear flexClearScrollable, flexClearEverything
           VatGrid.rows = 1
           TxtInvSerial.text = ""
 End If


        End If

End Sub
Sub RetriveValueAddedData(Optional Transaction_ID As Double)
Dim sql As String
Dim i As Integer
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
    VatGrid.Clear flexClearScrollable, flexClearEverything
    VatGrid.rows = 1
sql = " SELECT     dbo.TransactionValueAdded.Transaction_Type, dbo.TransactionValueAdded.Transaction_ID, dbo.TransactionValueAdded.Vat, dbo.TransactionValueAdded.Vatyo,"
sql = sql & " dbo.TransactionValueAdded.ItemID , dbo.TblItems.itemname, dbo.TblItems.Fullcode, dbo.TblItems.ItemNamee ,dbo.TransactionValueAdded.selectd ,dbo.TransactionValueAdded.Typ ,dbo.TransactionValueAdded.Valu "
sql = sql & " FROM         dbo.TransactionValueAdded LEFT OUTER JOIN"
sql = sql & "                      dbo.TblItems ON dbo.TransactionValueAdded.ItemID = dbo.TblItems.ItemID"
sql = sql & " Where (dbo.TransactionValueAdded.Transaction_Type = 22) And (dbo.TransactionValueAdded.Transaction_ID = " & Transaction_ID & " )"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
With Me.VatGrid
rs2.MoveFirst
.rows = .rows + rs2.RecordCount
For i = 1 To .rows - 1
 .TextMatrix(i, .ColIndex("index")) = i
.TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(rs2("ItemID").value), "", rs2("ItemID").value)
.TextMatrix(i, .ColIndex("Vat")) = IIf(IsNull(rs2("Vat").value), "", rs2("Vat").value)
.TextMatrix(i, .ColIndex("Vatyo")) = IIf(IsNull(rs2("Vatyo").value), "", rs2("Vatyo").value)
.TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
.TextMatrix(i, .ColIndex("select")) = IIf(IsNull(rs2("selectd").value), 0, rs2("selectd").value)
.TextMatrix(i, .ColIndex("Typ")) = IIf(IsNull(rs2("Typ").value), "", rs2("Typ").value)
.TextMatrix(i, .ColIndex("Valu")) = IIf(IsNull(rs2("Valu").value), 0, rs2("Valu").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("ItemName").value), "", rs2("ItemName").value)
Else
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("ItemNamee").value), "", rs2("ItemNamee").value)
End If
rs2.MoveNext
Next i
End With
End If
End Sub
Function Retrive_orders_data(Transaction_ID As Double)
    Dim StrSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim row_count As Integer
    Dim Num As Integer
 FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything


    StrSQL = "Select * from transactions where  Transaction_ID=" & Transaction_ID
    
    Set rs = New ADODB.Recordset
    
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount < 1 Then
 
        Exit Function
    Else
        DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
        DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
        dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
        Me.DcCurrency.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
        txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
        TxtBillSupplier.text = IIf(IsNull(rs("ManualNO").value), "", rs("ManualNO").value)
        TxtOrderSupply.text = IIf(IsNull(rs("nots").value), "", rs("nots").value)
        TxtVATNO.text = IIf(IsNull(rs("VATNO").value), "", rs("VATNO").value)
        
        'Me.DcCurrency.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
        'txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
  '      TxtLcNo.text = IIf(IsNull(rs("LcNo").value), "", (rs("LcNo").value))
    End If

    If rs.EOF Or rs.BOF Then
        Exit Function
    End If

    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & Transaction_ID

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.text = ""

    FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        row_count = FG.rows
    
        If FG.TextMatrix(row_count - 1, FG.ColIndex("Code")) = "" Then
            row_count = row_count - 1
        End If
     
        FG.rows = RsDetails.RecordCount + row_count

        For Num = row_count To FG.rows - 1 'RsDetails.RecordCount
    
            FG.TextMatrix(Num, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no")), "", (RsDetails("order_no").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemCostPrice")) = IIf(IsNull(RsDetails("CostPrice")), "", (RsDetails("CostPrice").value))
            
     '       FG.TextMatrix(Num, FG.ColIndex("OrderArrivalDate")) = DTArrivalDate.value
         
            FG.TextMatrix(Num, FG.ColIndex("TypeVAT")) = IIf(IsNull(RsDetails("TypeVAT")), "", (RsDetails("TypeVAT").value))
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
        
            '          FG.TextMatrix(Num, FG.ColIndex("Count")) = items_qty_not_recieved_in_order(FG.TextMatrix(Num, FG.ColIndex("Code")), FG.TextMatrix(Num, FG.ColIndex("order_no")))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Showqty")), "", (RsDetails("Showqty").value))
        
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("ShowPrice")), "", (RsDetails("ShowPrice").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
        
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassId")) = IIf(IsNull(RsDetails("ClassId")), 1, (RsDetails("ClassId").value))
        
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
FG.TextMatrix(Num, FG.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial")), "", Trim(RsDetails("ItemSerial").value))
           FG.TextMatrix(Num, FG.ColIndex("guaranteeTime")) = IIf(IsNull(RsDetails("guaranteeTime")), "", (RsDetails("guaranteeTime").value))
            End If
            FG.TextMatrix(Num, FG.ColIndex("Vat")) = IIf(IsNull(RsDetails("Vat")), "", (RsDetails("Vat").value))
            FG.TextMatrix(Num, FG.ColIndex("Vatyo")) = IIf(IsNull(RsDetails("Vatyo")), "", (RsDetails("Vatyo").value))
             
            FG.cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        
            RsDetails.MoveNext
            ' Debug.Print Num
            ' If FG.Rows > 10 Then
            '     If Num = 8 Then FG.Refresh
            ' End If
        Next Num

    End If

End Function

    


 
Private Sub TxtInvSerial_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
FrmBuySearch.DealingForm = GridTransType.PurchaseTransaction
  FrmBuySearch.Index = 5
            FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ›Ê« Ì— ‘—«¡   "
            FrmBuySearch.show vbModal

End If
End Sub

Private Sub txtPassword_Change()
If Trim(txtPassword) = "Alex2025" Then
    cmdReSave.Visible = True
    txtFromDateReSave.Visible = True
    txtToDateReSave.Visible = True
    chkIsBranch(0).Visible = True
    chkIsBranch(1).Visible = True
Else
    cmdReSave.Visible = False
    txtFromDateReSave.Visible = False
    txtToDateReSave.Visible = False
   chkIsBranch(0).Visible = False
   chkIsBranch(1).Visible = False
End If
txtFromDateReSave.value = Date
txtToDateReSave.value = Date
End Sub

Private Sub TxtShortName_KeyDown(KeyCode As Integer, Shift As Integer)
'   LoadSpecificItems
SerchItems (TxtShortName.text)
DoEvents
DoEvents
DoEvents
DoEvents

        If KeyCode = vbKeyReturn Then
        
        
   DCboItemsName.SetFocus
   DCboItemsName.BoundText = ""
        Sendkeys "{F4}"
        End If
End Sub
Sub SerchItems(Optional str As String)
 
Dim sql As String
Dim SQL1 As String
   
    SerchItemspUBLIC str, sql, SQL1
    fill_combo DCboItemsCode, sql
  fill_combo DCboItemsName, SQL1
        
         
End Sub

Sub SerchItemsxx(Optional str As String)
 
Set DCboItemsCode.RowSource = Nothing
Set DCboItemsName.RowSource = Nothing
If str <> "" Then
Dim sql As String
Dim SQL1 As String
 
Dim StrWhere As String
  Dim astrSplit2tems2() As String
  Dim j As Integer
  Dim nElements As Integer
  Dim SearchString As String
StrWhere = ""
SearchString = ""
sql = " select  ItemID,barCodeNO   from  dbo.TblItems where TblItems.IsArchive=0"
If SystemOptions.UserInterface = ArabicInterface Then
SQL1 = " select  ItemID,ItemName   from  dbo.TblItems where TblItems.IsArchive=0"
Else
SQL1 = " select  ItemID,ItemNamee   from  dbo.TblItems where TblItems.IsArchive=0"
End If

          astrSplit2tems2 = Split(str, " ")
          nElements = UBound(astrSplit2tems2) - LBound(astrSplit2tems2)
          If nElements = 0 Then
                       If SystemOptions.UserInterface = ArabicInterface Then
                            StrWhere = " and (ItemName Like N'%" & Trim(str) & "%' or barCodeNO Like N'%" & Trim(str) & "%' or shortName Like N'%" & Trim(str) & "%'  or fullcode Like N'%" & Trim(str) & "%') "
                    Else
                            StrWhere = " and (ItemNamee Like N'%" & Trim(str) & "%' or barCodeNO Like N'%" & Trim(str) & "%' or shortName Like N'%" & Trim(str) & "%' or fullcode Like N'%" & Trim(str) & "%' ) "
                    End If
                    
          End If
        If nElements > 0 Then
        
     '   StrWhere = StrWhere + " and (ItemName Like N'%" & Trim(astrSplit2tems2(0)) & "%' or barCodeNO Like N'%" & Trim(astrSplit2tems2(0)) & "%' or shortName Like N'%" & Trim(astrSplit2tems2(0)) & "%') "
        SearchString = ""
        For j = 0 To nElements
        
         SearchString = SearchString & "%" & Trim(astrSplit2tems2(j))
             '     SearchString = "%" & Trim(astrSplit2tems2(j)) & SearchString
                  
        '   StrWhere = StrWhere + " and (ItemName Like N'%" & Trim(astrSplit2tems2(j)) & "%' or barCodeNO Like N'%" & Trim(astrSplit2tems2(j)) & "%' or shortName Like N'%" & Trim(astrSplit2tems2(j)) & "%') "
        '   StrWhere = StrWhere + "  and NOT (ItemName IS NULL) and NOT (shortName IS NULL) and  NOT (ItemCode IS NULL)"
         Next j
         SearchString = SearchString & "%"
                             If SystemOptions.UserInterface = ArabicInterface Then

             StrWhere = StrWhere + " and (ItemName Like '" & SearchString & "' or barCodeNO Like '" & SearchString & "' or shortName Like '" & SearchString & "') "
             Else
              StrWhere = StrWhere + " and (ItemNamee Like '" & SearchString & "' or barCodeNO Like '" & SearchString & "' or shortName Like '" & SearchString & "') "
             End If
        '-  StrWhere = StrWhere + "  and NOT (ItemName IS NULL) and NOT (shortName IS NULL) and  NOT (ItemCode IS NULL)"
      
         End If
        
    sql = sql & StrWhere
        If SystemOptions.UserInterface = ArabicInterface Then
        sql = sql + " Order BY ItemName "
    Else
        sql = sql + " Order BY ItemName "
    End If


    SQL1 = SQL1 & StrWhere
        If SystemOptions.UserInterface = ArabicInterface Then
        SQL1 = SQL1 + " Order BY ItemName "
    Else
        SQL1 = SQL1 + " Order BY ItemNamee "
    End If
    
   End If
    fill_combo DCboItemsCode, sql
        fill_combo DCboItemsName, SQL1
        DoEvents
        DoEvents
  
        
       Exit Sub
       
If str <> "" Then
'Dim Sql As String
'Dim StrWhere As String
'  Dim astrSplit2tems2() As String
'  Dim j As Integer
'  Dim nElements As Integer
StrWhere = ""
If SystemOptions.UserInterface = ArabicInterface Then
sql = " select  ItemID,ItemName   from  dbo.TblItems where TblItems.IsArchive=0"
Else
sql = " select  ItemID,ItemNamee   from  dbo.TblItems where TblItems.IsArchive=0"
End If
          astrSplit2tems2 = Split(str, " ")
          nElements = UBound(astrSplit2tems2) - LBound(astrSplit2tems2)
        If nElements > 0 Then
        StrWhere = StrWhere + " and (ItemName Like N'%" & Trim(astrSplit2tems2(0)) & "%' or barCodeNO Like N'%" & Trim(astrSplit2tems2(0)) & "%' or shortName Like N'%" & Trim(astrSplit2tems2(0)) & "%') "
        For j = 1 To nElements - 1
        
           StrWhere = StrWhere + " and (ItemName Like N'%" & Trim(astrSplit2tems2(j)) & "%' or barCodeNO Like N'%" & Trim(astrSplit2tems2(j)) & "%' or shortName Like N'%" & Trim(astrSplit2tems2(j)) & "%') "
           StrWhere = StrWhere + "  and NOT (ItemName IS NULL) and NOT (shortName IS NULL) and  NOT (ItemCode IS NULL)"
         Next j
         End If
    sql = sql & StrWhere
        If SystemOptions.UserInterface = ArabicInterface Then
        sql = sql + " Order BY ItemName "
    Else
        sql = sql + " Order BY ItemNamee "
    End If


   End If
   
        fill_combo DCboItemsName, sql
        
End Sub
Private Sub TxtStoreID_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim StoreID As Integer

    If KeyCode = vbKeyReturn Then
    StoreID = getStoreInformatin(TxtStoreID)
        DCboStoreName.BoundText = StoreID
    End If
End Sub

Private Sub TxtValueAdded_Change()
RelinVatGrid
End Sub

Private Sub VatGrid_Click()
RelinVatGrid
End Sub

Public Sub XPBtnMove_Click(Index As Integer)

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
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
        '11
DisplayRec:
         Me.TxtModFlg.text = ""
        Dim StrSQL As String
     StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=5 "
     
StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"
            If SystemOptions.usertype <> UserAdminAll Then
           '     StrSQL = StrSQL & " AND   BranchId=" & Current_branch
            End If


     If SystemOptions.usertype <> UserAdminAll Then
 
          If SystemOptions.FixedCustomer = 1 Then
            StrSQL = StrSQL & " and  UserID = " & user_id
             End If
  
  
        Me.dcBranch.Enabled = True
      
      
    End If
    
            If SystemOptions.SortInvoiceByEntry Then
                StrSQL = StrSQL & " Order by Transaction_ID"
            Else
                StrSQL = StrSQL & " Order by noteserial1"
            End If
           If cmdReSave.Visible = True Then
    
    StrSQL = " SELECT * FROM Transactions WHERE Transaction_Type = 5 "
    StrSQL = StrSQL & "   and ( Transaction_Date >= " & SQLDate(txtFromDateReSave.value, True) & " and "
    StrSQL = StrSQL & "   Transaction_Date <=   " & SQLDate(txtToDateReSave.value, True) & " )"
    
 
    
    If chkIsBranch(0).value = vbChecked And val(Me.dcBranch.BoundText) > 0 Then
        StrSQL = StrSQL & "  and BranchID =  " & val(Me.dcBranch.BoundText)
    End If
     If chkIsBranch(1).value = vbChecked Then
        StrSQL = StrSQL & "  and Transaction_ID in "
        StrSQL = StrSQL & "  ( Select Transaction_ID from Transactions where Transaction_Type=5 and NoteId not In (SELECT IsNull(notes_id,0) FROM DOUBLE_ENTREY_VOUCHERS where Credit_Or_Debit = 0))"
    End If

End If
            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText


            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveLast
            End If
            Me.TxtModFlg.text = "R"
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

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            '        Cmd_Click (0)
        Else
            '        SendKeys "{TAB}"
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

    If KeyCode = vbKeyF3 Then
        If Cmd(5).Enabled = False Then Exit Sub
        Cmd_Click (5)
    End If

    If KeyCode = vbKeyF6 Then
        If Cmd(7).Enabled = False Then Exit Sub
        Cmd_Click (7)
    End If

    If KeyCode = vbKeyF2 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
        
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
        
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeySpace Then
            If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            
            End If
        End If
    End If

    If Shift = 2 Then
        XPTab301.SetFocus

        If KeyCode = vbKeyTab Then
            If XPTab301.CurrTab = 0 Then
                XPTab301.CurrTab = 1

                If XPChkPayType(0).Enabled = True Then
                    XPChkPayType(0).SetFocus
                End If

            Else
                XPTab301.CurrTab = 0
                FG.SetFocus
            End If
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub
Sub SaveValueAdded()
Dim i As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

sql = "Select * from  TransactionValueAdded where 1=-1"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
With Me.VatGrid
For i = 1 To .rows - 1
If val(.TextMatrix(i, .ColIndex("ItemID"))) <> 0 Then
rs2.AddNew
rs2("Transaction_ID").value = val(Me.XPTxtBillID.text)
rs2("Transaction_Type").value = 5
rs2("ItemID").value = val(.TextMatrix(i, .ColIndex("ItemID")))
rs2("Vatyo").value = val(.TextMatrix(i, .ColIndex("Vatyo")))
rs2("Vat").value = val(.TextMatrix(i, .ColIndex("Vat")))
rs2("Valu").value = val(.TextMatrix(i, .ColIndex("Valu")))
rs2("Typ").value = val(.TextMatrix(i, .ColIndex("Typ")))
If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
rs2("selectd").value = 1
Else
rs2("selectd").value = 0
End If
rs2.update
End If
Next i
End With
End Sub
Sub RetriveValueAdded()
Dim sql As String
Dim i As Integer
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
    VatGrid.Clear flexClearScrollable, flexClearEverything
    VatGrid.rows = 1
sql = " SELECT     dbo.TransactionValueAdded.Transaction_Type, dbo.TransactionValueAdded.Transaction_ID, dbo.TransactionValueAdded.Vat, dbo.TransactionValueAdded.Vatyo,"
sql = sql & " dbo.TransactionValueAdded.ItemID , dbo.TblItems.itemname, dbo.TblItems.Fullcode, dbo.TblItems.ItemNamee ,dbo.TransactionValueAdded.selectd ,dbo.TransactionValueAdded.Typ ,dbo.TransactionValueAdded.Valu "
sql = sql & " FROM         dbo.TransactionValueAdded LEFT OUTER JOIN"
sql = sql & "                      dbo.TblItems ON dbo.TransactionValueAdded.ItemID = dbo.TblItems.ItemID"
sql = sql & " Where (dbo.TransactionValueAdded.Transaction_Type = 5) And (dbo.TransactionValueAdded.Transaction_ID = " & val(XPTxtBillID.text) & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
With Me.VatGrid
rs2.MoveFirst
.rows = .rows + rs2.RecordCount
For i = 1 To .rows - 1
 .TextMatrix(i, .ColIndex("index")) = i
.TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(rs2("ItemID").value), "", rs2("ItemID").value)
.TextMatrix(i, .ColIndex("Vat")) = IIf(IsNull(rs2("Vat").value), "", rs2("Vat").value)
.TextMatrix(i, .ColIndex("Vatyo")) = IIf(IsNull(rs2("Vatyo").value), "", rs2("Vatyo").value)
.TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
.TextMatrix(i, .ColIndex("select")) = IIf(IsNull(rs2("selectd").value), 0, rs2("selectd").value)
.TextMatrix(i, .ColIndex("Typ")) = IIf(IsNull(rs2("Typ").value), "", rs2("Typ").value)
.TextMatrix(i, .ColIndex("Valu")) = IIf(IsNull(rs2("Valu").value), 0, rs2("Valu").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("ItemName").value), "", rs2("ItemName").value)
Else
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("ItemNamee").value), "", rs2("ItemNamee").value)
End If
rs2.MoveNext
Next i
End With
End If
End Sub
Private Sub Form_Load()
    Dim RsClients As New ADODB.Recordset
    Dim StrSQL As String
    Dim Num As Integer
    Dim StrList As String
    Dim BGround As New ClsBackGroundPic
    Dim RsNote As New ADODB.Recordset
    Dim Dcombos As ClsDataCombos

'    On Error GoTo ErrTrap
    ScreenNameArabic = " „—œÊœ«  «·„‘ —Ì«  "
    ScreenNameEnglish = "  Return Purchase "
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 230
 If True = True Then
XPTab301.TabVisible(2) = True
Else
XPTab301.TabVisible(2) = False
End If

If SystemOptions.AllowEditVaTManulay = True Then
txtManulaVat.Enabled = True
txtManulaVat.Visible = True
Else
txtManulaVat.Enabled = False
txtManulaVat.text = 0
txtManulaVat.Visible = False
End If


   StrSQL = " select id,code from currency"
 
    fill_combo Me.DcCurrency, StrSQL
    With Me.VatGrid
    If SystemOptions.UserInterface = ArabicInterface Then
                .ColComboList(.ColIndex("Typ")) = "#1; ·„ ÌﬁÊ„ «·„Ê—œ »«÷«›… ﬁÌ„…|#2; «·„Ê—œ „⁄›Ï"
              ElseIf SystemOptions.UserInterface = EnglishInterface Then
                .ColComboList(.ColIndex("Typ")) = "#1;Supplier did not add VAT|#2;Supplier is exempt "
            End If
    End With
    If SystemOptions.UserInterface = ArabicInterface Then
    With Me.Dcbtyp
    .Clear
    .AddItem "·„ ÌﬁÊ„ «·„Ê—œ »«÷«›… ﬁÌ„…"
    .AddItem "«·„Ê—œ „⁄›Ï"
    End With
    Else
     With Me.Dcbtyp
    .Clear
    .AddItem "Supplier did not add VAT"
    .AddItem "Supplier is exempt"
    End With
    End If
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Set NewGrid.Grid = FG
       NewGrid.GridTrans = Returntransaction
    Set NewGrid.TxtInvID = Me.XPTxtBillID
    Set NewGrid.TxtNots = Me.Text1
     NewGrid.ReturnTyp = 0
     Set NewGrid.txtManulaVat = Me.txtManulaVat
     Set NewGrid.CboDiscount_Type = XPCboDiscountType
    Set NewGrid.TxtDiscount_Val = XPTxtDiscountVal
    Set NewGrid.TxtModFlag = TxtModFlg
    Set NewGrid.txtTotal = XPTxtSum
    Set NewGrid.TxtValueAdded = TxtValueAdded
    Set NewGrid.VatGrid = Me.VatGrid
    Set NewGrid.TxtValueCash = XPTxtValue(0)
    Set NewGrid.TxtValueDelay = XPTxtValue(1)
    Set NewGrid.TxtValuechque = XPTxtValue(2)
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.StoreName = Me.DCboStoreName
    Set NewGrid.DtpBillDate = Me.XPDtbBill
    Set NewGrid.LblTotalQty = Me.LblTotalQty
    Set NewGrid.TxtShortName = Me.TxtShortName
    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    Set NewGrid.DCboItemName = DCboItemsName
    Set NewGrid.DCboItemCode = DCboItemsCode
    Set NewGrid.CboItemCase = CboItemCase
    Set NewGrid.CmdAddData = CmdAdd
    Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.TxtQuantity = TxtQuantity
    Set NewGrid.TxtPrice = TxtPrice
    Set NewGrid.GrdTBar = Me.TBar
    Set NewGrid.LblItemsCount = Me.LblItemsCount
    Set NewGrid.LblDiscountsTotal = Me.LblDiscountsTotal
    Set NewGrid.LblTotalAll = Me.LblTotalAll
        Set NewGrid.LBLGross = LBLGross

Set NewGrid.Customer = Me.DBCboClientName
    NewGrid.FillGrid

    FG.WallPaper = BGround.Picture
    AddTip
    XPTab301.CurrTab = 0
    XPDtbBill.value = Date
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetCustomersSuppliers 3, Me.DBCboClientName, True
    Dcombos.GetStores Me.DCboStoreName
    Dcombos.GetBanks Me.DCboBankName
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetSalesRepDatapurchase Me.DcboEmp
Dcombos.GetDocTypebyid Me.DCDocTypes, 5, val(Me.dcBranch.BoundText)

    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DBCboClientName

    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DCboBankName

    Set cSearchDcbo(2) = New clsDCboSearch
    Set cSearchDcbo(2).Client = Me.DCboStoreName


    With Me.CboRetrunType
     '   .Clear
  If SystemOptions.UserInterface = ArabicInterface Then
             .Clear
            .AddItem "≈— Ã«⁄ „ﬁÌœ(„— »ÿ »›« Ê—… ‘—«¡)"
            .AddItem "≈— Ã«⁄ €Ì— „ﬁÌœ(€Ì— „— »ÿ »›« Ê—… ‘—«¡)"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
             .Clear
            .AddItem "With bill"
            .AddItem "With out Bill"
        End If

     
    End With


    If SystemOptions.UserInterface = EnglishInterface Then

        With CboPayMentType
            .Clear
            .AddItem "Cash"
            .AddItem "Credit"
        End With
    
    Else
    
        With CboPayMentType
            .Clear
            .AddItem "‰ﬁœ«"
            .AddItem "¬Ã·"
        End With

    End If

    SetDtpickerDate Me.XPDtbBill
    SetDtpickerDate Me.XPDTPDueDate
    SetDtpickerDate Me.DtpDelayDate
    StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=-5"
     StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"
     
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If SystemOptions.UserInterface = EnglishInterface Then
      SetInterface Me
        ChangeLang
    End If
     
     With Me.CboRetrunType
     '   .Clear
  If SystemOptions.UserInterface = ArabicInterface Then
             .Clear
            .AddItem "≈— Ã«⁄ „ﬁÌœ(„— »ÿ »›« Ê—… ‘—«¡)"
            .AddItem "≈— Ã«⁄ €Ì— „ﬁÌœ(€Ì— „— »ÿ »›« Ê—… ‘—«¡)"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
             .Clear
            .AddItem "With bill"
            .AddItem "With out Bill"
        End If

     
    End With
   
   
   ' XPBtnMove_Click 2
    InvType = 5
  '  Me.TxtModFlg.Text = "R"
    Resize_Form Me, TransactionSize
     
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
    'CmdConvert.Caption = "Convert to bill"
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
lbl(97).Caption = "Smart Search"
lbl(44).Caption = "Currency"
lbl(43).Caption = "VAT No"
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
TxtManualNO(6).Caption = "Return Type"
    Me.XPTab301.TabCaption(0) = "Items"
    lbl(40).Caption = "Person"
    lbl(41).Caption = "Transporter"
     lbl(39).Caption = "Supp Bill No."
    lbl(38).Caption = "PO No."
     '''/////////////
    ChecVAT.RightToLeft = False
    ChecVAT.Caption = "Select All"
    lbl(77).Caption = "Status VAT"
    lbl(78).Caption = "Reason"
    lbl(42).Caption = "VAT"
    lbl(104).Caption = "Total"
    Me.XPTab301.TabCaption(2) = "VAT"
    Label22.Caption = "Data of VAT"
With VatGrid
.TextMatrix(0, .ColIndex("index")) = "Serial"
.TextMatrix(0, .ColIndex("select")) = "Select"
.TextMatrix(0, .ColIndex("Code")) = "Item Code"
.TextMatrix(0, .ColIndex("Name")) = "Item Name"
.TextMatrix(0, .ColIndex("Vatyo")) = "Percentage"
.TextMatrix(0, .ColIndex("Vat")) = "Value"
.TextMatrix(0, .ColIndex("Valu")) = "Item Value"
.TextMatrix(0, .ColIndex("Typ")) = "Status"
End With
     ''//////////
         lbl(37).Caption = "Return Reason."
         
    lbl(34).Caption = "Total"
    lbl(32).Caption = "Total QTY"
    Me.XPTab301.TabCaption(1) = "Notes"
    lbl(25).Caption = "GE NO:"
    Cmd(10).Caption = "Print GE"
    Label3.Caption = "Branch"
    Label4.Caption = "Doc Type"
    lbl(20).Caption = "Payment Method"
    XPChkPayType(0).Caption = "Cash"
    XPChkPayType(1).Caption = "Credit"
    XPChkPayType(2).Caption = "Cheque"

    lbl(13).Caption = "Value"
    lbl(15).Caption = "Value"
    lbl(16).Caption = "Value"
    lbl(18).Caption = "Cheque#"

    lbl(12).Caption = "Index"
    lbl(14).Caption = "Index"
    lbl(22).Caption = "Box"
    lbl(21).Caption = " date"

    lbl(19).Caption = " Cheque date"

    lbl(17).Caption = "Bank"

    Me.Caption = "Return purchases"
    C1Elastic6.Caption = Me.Caption

    lbl(7).Caption = "ID"
    lbl(6).Caption = "Invoice Date"
    lbl(5).Caption = "Vendor Name"
    lbl(3).Caption = "Store "

    lbl(4).Caption = "Payment Type"
    lbl(9).Caption = "Invoice#"
'    Ele(6).Caption = "Return Type"
lbl(32).Visible = False
lbl(63).Visible = False

lbl(36).Caption = "Manual No."

    lbl(0).Caption = " Total:"
    lbl(11).Caption = "Disc"
    lbl(23).Caption = " Net:"

    lbl(8).Caption = " By:"
    lbl(1).Caption = "Rec.Count:"

    lbl(31).Caption = "Item Code"
    lbl(30).Caption = "Item Name"
    lbl(29).Caption = " Case"
    lbl(28).Caption = " Serial"
    lbl(27).Caption = "QTY"
    lbl(26).Caption = "Price"

    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    Me.CmdHelp.Caption = "Help"
    Me.Cmd(8).Caption = "Show Issue Vchr"
    lbl(35).Caption = "Manual NO"
    
End Sub

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·›« Ê—…   " & TxtNoteSerial1.text & CHR(13) & " «· «—ÌŒ " & XPDtbBill.value & CHR(13) & " «·Œ“Ì‰… " & DcboBox.text & CHR(13) & " «·„Œ“‰  " & DCboStoreName.text & CHR(13) & "  «·⁄„Ì· / «·„Ê—œ   " & DBCboClientName.text & CHR(13) & "‰Ê⁄ «·„—œÊœ«  " & CboRetrunType & CHR(13) & " ‰Ê⁄ «·”‰œ" & DCDocTypes & CHR(13) & "ÿ—Ìﬁ… «·œ›⁄ " & CboPayMentType & CHR(13) & "—ﬁ„ «·ﬁÌœ " & TxtNoteSerial
                     
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Bill No " & TxtNoteSerial1.text & CHR(13) & " Date " & XPDtbBill.value & CHR(13) & " Box " & DcboBox.text & CHR(13) & " Store  " & DCboStoreName.text & CHR(13) & " Supplier/Cuxtomer" & DBCboClientName.text & CHR(13) & "  Type" & CboRetrunType & CHR(13) & " Doc Type" & DCDocTypes & CHR(13) & "Payment Type" & CboPayMentType & CHR(13) & " GE NO" & TxtNoteSerial
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 230, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", , TxtNoteSerial, TxtNoteSerial1
    Else
        AddToLogFile CInt(user_id), 230, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", , TxtNoteSerial, TxtNoteSerial1
    End If
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, , 230
    Dim i As Integer

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    For i = LBound(cSearchDcbo) To UBound(cSearchDcbo)
        Set cSearchDcbo(i) = Nothing
    Next i

    Set rs = Nothing
    Set TTP = Nothing
    NewGrid.Class_Terminate
    Set NewGrid = Nothing
    Set ReturnReport = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"
            '     Me.Caption = "„— Ã⁄ «·„‘ —Ì« "
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.Cmd(7).Enabled = True
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            XPDtbBill.Enabled = False
            Me.DBCboClientName.locked = True
            Me.DCboStoreName.locked = True
            Me.DCboBankName.locked = True
            XPChkPayType(0).Enabled = False
            XPChkPayType(1).Enabled = False
            XPChkPayType(2).Enabled = False
            XPTxtValue(0).Enabled = False
            XPTxtSerial(0).Enabled = False
            XPTxtValue(1).Enabled = False
            XPTxtSerial(1).Enabled = False
            XPTxtChqueNum.Enabled = False
            DCboBankName.Enabled = False
            XPTxtValue(2).Enabled = False
            XPDTPDueDate.Enabled = False
            FG.Editable = flexEDNone

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                Me.Cmd(7).Enabled = False
            End If
        
            CboPayMentType.locked = True
            DtpDelayDate.Enabled = False
            Ele(5).Enabled = False
            Me.CboRetrunType.locked = True

        Case "N"
            '     Me.Caption = "„— Ã⁄ «·„‘ —Ì« ( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            '        Me.XPBtnMove(0).Enabled = False
            '        Me.XPBtnMove(1).Enabled = False
            '        Me.XPBtnMove(2).Enabled = False
            '        Me.XPBtnMove(3).Enabled = False

            FG.Enabled = True
            FG.rows = 2
            XPDtbBill.Enabled = True
            XPDtbBill.value = Date
            Me.DBCboClientName.locked = False
            Me.DCboStoreName.locked = False
        
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            XPChkPayType(0).value = Unchecked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            FG.Editable = flexEDKbdMouse
            CboPayMentType.ListIndex = 0
            CboPayMentType.locked = False
            DtpDelayDate.Enabled = True
            DtpDelayDate.value = Date
            XPDTPDueDate.value = Date
            Ele(5).Enabled = True
            CboItemCase.ListIndex = 0
            Me.CboRetrunType.locked = False

        Case "E"
            '     Me.Caption = "„— Ã⁄ «·„‘ —Ì« (  ⁄œÌ· )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            FG.Enabled = True
            XPDtbBill.Enabled = True
            Me.DBCboClientName.locked = False
            Me.DCboStoreName.locked = False
            Me.DCboBankName.locked = False
        
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            XPDTPDueDate.Enabled = True
            DtpDelayDate.Enabled = True

            If XPChkPayType(0).value = Checked Then
                XPChkPayType_Click (0)
            End If

            If XPChkPayType(1).value = Checked Then
                XPChkPayType_Click (1)
            End If

            If XPChkPayType(2).value = Checked Then
                XPChkPayType_Click (2)
            End If

            If CboPayMentType.ListIndex = 0 Then
                CboPayMentType_Change
            End If

            FG.Editable = flexEDKbdMouse
        
            CboPayMentType.locked = False
            DBCboClientName_Change
            Ele(5).Enabled = True
            Me.CboRetrunType.locked = False
    End Select

    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim Num As Long
    On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    If Lngid <> 0 Then
        rs.Find "Transaction_ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If

    TxtFillData.text = "T"
    Screen.MousePointer = vbArrowHourglass
    Me.Text1.text = IIf(IsNull(rs("nots").value), "", (rs("nots").value))
Me.Transporter.text = IIf(IsNull(rs("Transporter").value), "", (rs("Transporter").value))
    dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", (rs("NoteSerial").value))
    Me.TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", (rs("NoteSerial1").value))
    Me.oldtxtNoteSerial1.text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)
    Me.DcboEmp.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
    DCDocTypes.BoundText = IIf(IsNull(rs("Doctype").value), "", rs("Doctype").value)
    Me.Dcbtyp.ListIndex = IIf(IsNull(rs("Typ").value), -1, (rs("Typ").value))
    TXtResonVAT.text = IIf(IsNull(rs("ResonVAT").value), "", (rs("ResonVAT").value))
    Me.DcCurrency.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
    txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))

    lbl(24).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

    Me.TXTNoteID.text = IIf(IsNull(rs("NoteID").value), "", (rs("NoteID").value))

    If IsNull(rs("BoxID").value) Then
        Me.DcboBox.BoundText = ""
    Else
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
    End If
 TxtValueAdded.text = IIf(IsNull(rs("VAT").value), 0, (rs("VAT").value))
 LblValueAdded.Caption = IIf(IsNull(rs("VAT").value), 0, (rs("VAT").value))


    TxtManualNo1.text = IIf(IsNull(rs("ManualNo1").value), "", (rs("ManualNo1").value))
    XPTxtBillID.text = IIf(IsNull(rs("Transaction_ID").value), "", val(rs("Transaction_ID").value))
    TxtTransSerial.text = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
    XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", (rs("Transaction_Date").value))
    Me.DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    FG.Clear flexClearScrollable, flexClearEverything
    Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
    CboPayMentType.ListIndex = IIf(IsNull(rs("PaymentType").value), 0, rs("PaymentType").value)
txtManulaVat.text = IIf(IsNull(rs("txtManulaVat").value), 0, (rs("txtManulaVat").value))
txtManulaVat.text = val(txtManulaVat.text)

    If Not IsNull(rs("ReturnID").value) Then
        Me.CboRetrunType.ListIndex = 0
        Me.TxtInvID.text = rs("ReturnID").value
      Me.TxtInvSerial.text = IIf(IsNull(rs("ReturnSerial").value), "", (rs("ReturnSerial").value))
    Else
        Me.CboRetrunType.ListIndex = 1
        Me.TxtInvID.text = ""
          Me.TxtInvSerial.text = ""
    End If
''// 31 05 2015
TxtVATNO.text = IIf(IsNull(rs("VATNO").value), "", (rs("VATNO").value))
Me.TxtManualNO11.text = IIf(IsNull(rs("ManualNO").value), "", (rs("ManualNO").value))
Me.TxtOrderSupply.text = IIf(IsNull(rs("OrderSupply").value), "", (rs("OrderSupply").value))
Me.TxtBillSupplier.text = IIf(IsNull(rs("BillSupplier").value), "", (rs("BillSupplier").value))
Me.TxtReasonReturns.text = IIf(IsNull(rs("ReasonReturns").value), "", (rs("ReasonReturns").value))
    FG.rows = 2
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + "  where Transaction_ID=" & val(rs("Transaction_ID").value)
    StrSQL = StrSQL & " ORDER BY ID "

    'StrSql = "select * From Transaction_Details where Transaction_ID=" & Val(Rs("Transaction_ID").Value)
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            FG.TextMatrix(Num, FG.ColIndex("TypeVAT")) = IIf(IsNull(RsDetails("TypeVAT")), "", (RsDetails("TypeVAT").value))
            FG.TextMatrix(Num, FG.ColIndex("FoxyNo")) = IIf(IsNull(RsDetails("FoxyNo")), "", (RsDetails("FoxyNo").value))
            FG.TextMatrix(Num, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no").value), "", RsDetails("order_no").value)
            FG.TextMatrix(Num, FG.ColIndex("OrderArrivalDate")) = IIf(IsNull(RsDetails("OrderArrivalDate").value), "", RsDetails("OrderArrivalDate").value)
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))

            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial")), "", Trim(RsDetails("ItemSerial").value))
            FG.TextMatrix(Num, FG.ColIndex("Vat")) = IIf(IsNull(RsDetails("Vat")), "", (RsDetails("Vat").value))
            FG.TextMatrix(Num, FG.ColIndex("Vatyo")) = IIf(IsNull(RsDetails("Vatyo")), "", Trim(RsDetails("Vatyo").value))
            
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If

            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("SHOWQTY")), "", (RsDetails("SHOWQTY").value))
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("SHOWPrice")), "", (RsDetails("SHOWPrice").value))

            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            Else
                FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("Transaction_Details.ItemCase")), "", (RsDetails("Transaction_Details.ItemCase").value))
            End If

            FG.TextMatrix(Num, FG.ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks")), "", (RsDetails("Remarks").value))
        
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 0, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))

            FG.cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            FG.TextMatrix(Num, FG.ColIndex("ProductionDate")) = IIf(IsNull(RsDetails("ProductionDate")), "", (RsDetails("ProductionDate").value))
            FG.TextMatrix(Num, FG.ColIndex("ExpiryDate")) = IIf(IsNull(RsDetails("ExpiryDate")), "", (RsDetails("ExpiryDate").value))
            FG.TextMatrix(Num, FG.ColIndex("LotNO")) = IIf(IsNull(RsDetails("LotNO")), "", (RsDetails("LotNO").value))
        
            RsDetails.MoveNext
        Next Num

        NewGrid.Calculate 1
    End If

    XPChkPayType(0).value = Unchecked
    XPChkPayType(1).value = Unchecked
    XPChkPayType(2).value = Unchecked
    XPTxtValue(0).text = ""
    XPTxtValue(1).text = ""
    XPTxtValue(2).text = ""
    XPTxtSerial(0).text = ""
    XPTxtSerial(1).text = ""
    XPTxtChqueNum.text = ""
    DCboBankName.BoundText = ""
    XPDTPDueDate.value = Date
    DtpDelayDate.value = Date
    StrSQL = "select * From Notes where Transaction_ID=" & val(rs("Transaction_ID").value)
    RsNotes.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsNotes.EOF Or RsNotes.BOF) Then

        For Num = 1 To RsNotes.RecordCount

            If RsNotes("NoteType").value = 0 Then
                XPChkPayType(0).value = Checked
                XPChkPayType_Click (0)
                XPTxtValue(0).text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtSerial(0).text = IIf(IsNull(RsNotes("NoteSerial").value), "", Trim(RsNotes("NoteSerial").value))
                Me.DcboBox.BoundText = IIf(IsNull(RsNotes("BoxID").value), "", (RsNotes("BoxID").value))
            End If

            If RsNotes("NoteType").value = 1 Then
                XPChkPayType(1).value = Checked
                XPChkPayType_Click (1)
                XPTxtValue(1).text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtSerial(1).text = IIf(IsNull(RsNotes("NoteSerial").value), "", Trim(RsNotes("NoteSerial").value))
                DtpDelayDate.value = IIf(IsNull(RsNotes("DueDate").value), "", (RsNotes("DueDate").value))
            End If

            If RsNotes("NoteType").value = 2 Then
                XPChkPayType(2).value = Checked
                XPChkPayType_Click (2)
                XPTxtValue(2).text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtChqueNum.text = IIf(IsNull(RsNotes("ChqueNum").value), "", Trim(RsNotes("ChqueNum").value))
                Me.DCboBankName.BoundText = IIf(IsNull(RsNotes("BankID").value), "", RsNotes("BankID").value)
                XPDTPDueDate.value = IIf(IsNull(RsNotes("DueDate").value), "", (RsNotes("DueDate").value))
            End If

            RsNotes.MoveNext
        Next Num

        FG_AfterEdit 1, 1
    End If
If Me.TxtModFlg.text <> "E" And Me.TxtModFlg.text <> "N" Then
RetriveValueAdded

End If
RelinVatGrid
    Screen.MousePointer = vbDefault
    TxtFillData.text = "F"
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub
Sub RelinVatGrid()
Dim k As Integer
If FG.ColIndex("Vat") = -1 Then Exit Sub
If val(Dcbtyp.ListIndex) <> -1 Then
For k = FG.FixedRows To FG.rows - 1
FG.TextMatrix(k, FG.ColIndex("Vat")) = 0
FG.TextMatrix(k, FG.ColIndex("Vatyo")) = 0
FG.TextMatrix(k, FG.ColIndex("TypeVAT")) = 0
Next k
VatGrid.Clear flexClearScrollable, flexClearEverything
    VatGrid.rows = 2
End If
Dim i As Integer
Dim SmValu As Double
SmValu = 0
With VatGrid
For i = 1 To .rows - 1
If .cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
For k = FG.FixedRows To FG.rows - 1
If k = i And val(FG.TextMatrix(k, FG.ColIndex("Code"))) = val(.TextMatrix(i, .ColIndex("ItemID"))) And val(FG.TextMatrix(k, FG.ColIndex("Valu"))) = val(.TextMatrix(i, .ColIndex("Valu"))) Then
FG.TextMatrix(k, FG.ColIndex("Vat")) = val(.TextMatrix(i, .ColIndex("Vat")))
FG.TextMatrix(k, FG.ColIndex("Vatyo")) = val(.TextMatrix(i, .ColIndex("Vatyo")))
End If
Next k

SmValu = SmValu + val(.TextMatrix(i, .ColIndex("Vat")))
.TextMatrix(i, .ColIndex("Typ")) = ""
Else
For k = FG.FixedRows To FG.rows - 1
If k = i And val(FG.TextMatrix(k, FG.ColIndex("Code"))) = val(.TextMatrix(i, .ColIndex("ItemID"))) And val(FG.TextMatrix(k, FG.ColIndex("Valu"))) = val(.TextMatrix(i, .ColIndex("Valu"))) Then
FG.TextMatrix(k, FG.ColIndex("Vat")) = 0
FG.TextMatrix(k, FG.ColIndex("Vatyo")) = 0
FG.TextMatrix(k, FG.ColIndex("TypeVAT")) = 0
End If
Next k
If val(Me.Dcbtyp.ListIndex) > -1 Then
.TextMatrix(i, .ColIndex("Typ")) = val(Me.Dcbtyp.ListIndex) + 1
End If
End If
Next i
End With
TxtValueAdded.text = Format(SmValu, ".##")
LblValueAdded.Caption = Format(SmValu, ".##")
LblTotal.Caption = val(LblTotalAll.Caption) - val(LblDiscountsTotal.Caption) + val(LblValueAdded.Caption)
End Sub
Private Sub Undo()
    Dim Msg As String

    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                clear_all Me
                Me.TxtModFlg.text = "R"
                XPBtnMove_Click (1)
            End If

        Case "E"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ⁄œÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                rs.Find "Transaction_ID='" & val(XPTxtBillID.text) & "'", , adSearchForward, adBookmarkFirst

                If rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.text = "R"
                    Exit Sub
                End If

                If Not rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.text = "R"
                    Retrive
                End If
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_TransAction()
    Dim Msg As String
    Dim BegainTrans As Boolean
    Dim StrSQL As String
    Dim StrSqlDel  As String
    On Error GoTo ErrTrap

    If XPTxtBillID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtNoteSerial1.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                Cn.BeginTrans
                BegainTrans = True
                'StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS  " & "Where DOUBLE_ENTREY_VOUCHERS.Transaction_ID=" & rs("Transaction_ID").value
                'Cn.Execute StrSQL, , adExecuteNoRecords
                Cn.Execute "Delete from TransactionValueAdded where Transaction_ID=" & val(Me.XPTxtBillID.text) & ""
               ' StrSqlDel = "delete From Notes where Transaction_ID=" & val(Me.XPTxtBillID.Text) ' Val(rs("Transaction_ID").value)
               ' Cn.Execute StrSqlDel, , adExecuteNoRecords
        
                'StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & val(Me.XPTxtBillID.Text)
                'Cn.Execute StrSQL, , adExecuteNoRecords
        
                'StrSQL = "delete From Notes where noteid=" & val(TXTNoteID.Text)
                'Cn.Execute StrSQL, , adExecuteNoRecords
                
                                StrSQL = "Delete From Notes Where  NoteType= 230 and NoteID=" & val(Me.TXTNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
                DeleteTransactiomsVoucher val(Text1.text)
                CuurentLogdata ("D")
                rs.delete
                Cn.CommitTrans
                BegainTrans = False
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                     VatGrid.Clear flexClearScrollable, flexClearEverything
                    VatGrid.rows = 1
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
         VatGrid.Clear flexClearScrollable, flexClearEverything
           VatGrid.rows = 1
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·”Ã· "
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & Err.Description
    Msg = Msg & CHR(13) & Err.Source
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title

    If BegainTrans = True Then
        rs.CancelUpdate
        Cn.RollbackTrans
        BegainTrans = False
    End If

End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, "„— Ã⁄ «·„‘ —Ì« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›…  ⁄„·Ì… ≈—Ã«⁄ „‘ —Ì« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "„— Ã⁄ «·„‘ —Ì« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
    End With

    With TTP
        .Create Me.hWnd, "„— Ã⁄ «·„‘ —Ì« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  ⁄„·Ì… ≈—Ã«⁄ „‘ —Ì« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "„— Ã⁄ «·„‘ —Ì« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  ⁄„·Ì… ≈—Ã«⁄ «·„‘ —Ì« " & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "„— Ã⁄ «·„‘ —Ì« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… ≈—Ã«⁄ «·„‘ —Ì« " & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "„— Ã⁄ «·„‘ —Ì« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  ⁄„·Ì… ≈—Ã«⁄ «·„‘ —Ì« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "„— Ã⁄ «·„‘ —Ì« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄„·Ì… ≈—Ã«⁄ „‘ —Ì« " & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "„— Ã⁄ «·„‘ —Ì« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "„— Ã⁄ «·„‘ —Ì« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "„— Ã⁄ «·„‘ —Ì« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "„— Ã⁄ «·„‘ —Ì« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "„— Ã⁄ «·„‘ —Ì« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "„— Ã⁄ «·„‘ —Ì« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Function CreateIssueVoucher() As Boolean
On Error GoTo ErrTrap
CreateIssueVoucher = False
    'DeleteTransactiomsVoucher Val(Text1.text)

    Dim i As Long
    Dim LngCurItemID As Integer
    Dim LngUnitID As Long
    Dim UnitFactor As Double
    '›Ì Õ«·… «·«‰ «Ã «·‰„ÿÌ
 
'    With Fg

'        For i = 1 To Fg.Rows - 1

'            If Fg.TextMatrix(i, Fg.ColIndex("Code")) <> "" And val(Fg.TextMatrix(i, Fg.ColIndex("ItemType"))) <> 1 Then
                                      
'                LngCurItemID = val(Fg.TextMatrix(i, Fg.ColIndex("Code")))
'                LngUnitID = val(Fg.Cell(flexcpData, i, Fg.ColIndex("UnitID")))
            
'                GetUnitNoOfItems LngCurItemID, LngUnitID, UnitFactor
                    
                'TOTAL_COST = TOTAL_COST + (FG.TextMatrix(i, FG.ColIndex("Count")) * UnitFactor * ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod))
                    
'                If val(Fg.TextMatrix(i, Fg.ColIndex("ItemCostPrice"))) = 0 Then
'                    If SystemOptions.UserInterface = ArabicInterface Then
'                        MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ  ﬂ·›Â «·»Ì⁄ ·Â Ê·„ Ì „  ÕœÌœ À„‰ «·‘—«¡ Ê·Ì” ·Â ﬁÌ„Â —’Ìœ «›  «ÕÌ… ·–·ﬂ ·« Ì„ﬂ‰ «‰‘«¡ ”‰œ «·’—› "
'                    Else
'                        MsgBox "Item in line no " & i & "Have No Qty "
'                    End If
                            
'                    Exit Sub
'                End If
'            End If

'        Next i

'    End With

ll:
    Dim groupAccount  As String

    If detect_inventory_work_type = 3 Then
   
        With FG

            For i = 1 To FG.rows - 1

                If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
                
                    ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                    groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 0)

                    If groupAccount = "Error" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”⁄·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                        Else
                            MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                        End If
                        CreateIssueVoucher = False
                        Exit Function
                    End If
                End If

            Next i

        End With

    End If

    Dim MYWAER As String
    Dim StrSQL As String
    Dim RsNotes As ADODB.Recordset
    Dim MYinvnum As String
    Dim note_id As Long

    Dim RSTransDetails As ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
 
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
    '>>>>>>>>>>>>>>>>>>>>>>>>>

    'rs.Close
    '19 «–‰ ’—›
    '        rs.Open "select * from Transactions where nots =' " & XPTxtBillID.text & "' and Transaction_type = 19"
    '       If rs.RecordCount > 0 Then
    '        If rs!nots <> "" Then
    '        If SystemOptions.UserInterface = ArabicInterface Then
    '             Msg = "·ﬁœ  „  ÕÊÌ· Â–… «·›« Ê—… «·Ï «–‰ ’—›    .."
    '            Msg = Msg & Chr(13) & "Ê·«Ì„ﬂ‰  ÕÊÌ·… „—… «Œ—Ï  ..!!"
    '        Else
    '          Msg = "This bill already converted"
    '        End If
    '          MsgBox Msg, vbOKOnly, App.Title
    '        Exit Sub
    '        End If
        
    '        End If

    '        rs.Close
    '21 ›« Ê—… „»Ì⁄« 
    '        rs.Open "select * from Transactions where Transaction_ID = " & XPTxtBillID.text & " and Transaction_type = 21"

    '        If SystemOptions.UserInterface = ArabicInterface Then
    '        Msg = "”Ê› Ì „ «‰‘«¡ «–‰ ’—› „‰ Â–… «·›« Ê—…   .."
    '        Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
    '        Else
    '        Msg = "Create ISSUE Voucher to this bill ?"
    '        End If
    '  On Error GoTo ErrTrap
    Dim xyeas As Boolean
    xyeas = True

    If xyeas = True Then
 
        MYTEXT = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=19"))
        'mytext = TxtTransSerial.text

        '         rs!nots = mytext
        '         rs.update

        Dim Transaction_ID As Long
        
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        Dim general_noteid As Long
        Dim RsNotesGeneral As ADODB.Recordset
        Dim TxtNoteSerialV As String
            
        my_branch = Me.dcBranch.BoundText

        If TxtNoteSerialV = "" Then
            If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": CreateIssueVoucher = False: Exit Function
            Else
                       
                If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": CreateIssueVoucher = False: Exit Function
                Else
                    TxtNoteSerialV = Notes_coding(val(my_branch), XPDtbBill.value)
                End If
            End If
        End If
           Dim TxtNoteSerial1Vstr As String
        If TxtNoteSerial1V = "" Then
        TxtNoteSerial1Vstr = Voucher_coding(val(my_branch), XPDtbBill.value, 10, 180, , 19, , val(DCboStoreName.BoundText))
            If TxtNoteSerial1Vstr = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ  ’—› ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": CreateIssueVoucher = False: Exit Function
            Else

                If TxtNoteSerial1Vstr = "" Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ”‰œ «·’—›  ÌœÊÌ« ﬂ„« Õœœ   ": Exit Function
                Else
                    TxtNoteSerial1V = TxtNoteSerial1Vstr
                End If
            End If
        End If
 
        If Trim(CurrentVoucherNo) <> "" And DateChanged <> True Then
            TxtNoteSerialV = CurrentVoucherNo '—ﬁ„ «·ﬁÌœ
            TxtNoteSerial1V = Trim(CurrentVoucherSerialNo)
        End If
           
        Dim sql As String
Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
       ' Text1.text = Transaction_ID
        
         sql = "INSERT INTO  Transactions (Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,CusID,StoreID,UserID,Emp_ID,nots,nots2,NoteSerial,NoteSerial1,NoteId,BranchId,Closed,ChkReturnPurcahse,BillBasedOn,order_no,ManualNO)SELECT " & Transaction_ID & "," & MYTEXT & ",Transaction_Date,Transaction_Type = 19,CusID,StoreID,UserID,Emp_ID,nots=" & val(XPTxtBillID.text) & ",nots2='" & TxtNoteSerial1.text & " ',NoteSerial=' " & TxtNoteSerialV & "',NoteSerial1='" & TxtNoteSerial1V & "',NoteId=" & general_noteid & ",BranchId,1,1 ,3,'" & Me.TxtNoteSerial1 & "',ManualNO From Transactions Where  Transaction_ID =" & val(XPTxtBillID.text) & " And Transaction_Type = 5"
        Cn.Execute sql
        '
        
      If CboRetrunType.ListIndex = 0 Then ' »›« Ê—…  »”⁄— «·›« Ê—… ' „ﬁÌœ
      '  Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,ItemSize,UnitId,ShowQty,QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID ,ProductionDate,ExpiryDate,LotNO )SELECT  costprice,guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , Quantity, costprice/ QtyBySmalltUnit ,ColorID,ItemSize, UnitId, ShowQty, QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID,ProductionDate,ExpiryDate,LotNO  From dbo.Transaction_Details Where  Transaction_ID = " & XPTxtBillID.text
      Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice ,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price ,ColorID,ItemSize,UnitId,ShowQty,QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID ,ProductionDate,ExpiryDate,LotNO,OldQty,OldCost,NewQty,NewCost )SELECT  showPrice * " & val(txt_Currency_rate) & " ,guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , Quantity, Price/ QtyBySmalltUnit* " & val(txt_Currency_rate) & "  ,ColorID,ItemSize, UnitId, ShowQty, QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID,ProductionDate,ExpiryDate,LotNO ,OldQty,OldCost,NewQty,NewCost From dbo.Transaction_Details Where  Transaction_ID = " & XPTxtBillID.text
      
      Else ' »›« Ê—… »”⁄— «· ﬂ·€… ' ›Ì— „ﬁÌœ
      Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,ItemSize,UnitId,ShowQty,QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID ,ProductionDate,ExpiryDate,LotNO )SELECT  price  ,guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , Quantity, price/ QtyBySmalltUnit   ,ColorID,ItemSize, UnitId, ShowQty, QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID,ProductionDate,ExpiryDate,LotNO  From dbo.Transaction_Details Where  Transaction_ID = " & XPTxtBillID.text
       
      End If
      
           UpdateTransactionsCost CStr(Transaction_ID)
        
        Text1.text = Transaction_ID
        'TxtIssueSerial.text = TxtNoteSerial1V
 
        StrSQL = "UPDATE Transactions SET NOTS=" & Transaction_ID & " WHERE Transaction_ID=" & val(Me.XPTxtBillID.text)
        Cn.Execute StrSQL

        If SystemOptions.TypicalProduction = True Then
            Exit Function
        End If

        'Create big notes
        Set RsNotesGeneral = New ADODB.Recordset
'        RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
       StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

        If Me.TxtModFlg.text = "N" Then
    
        Else
        
            general_noteid = val(TXTNoteID.text)
        End If

        RsNotesGeneral.AddNew
        RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
        general_noteid = RsNotesGeneral("NoteID").value
        TXTNoteID.text = general_noteid
        RsNotesGeneral("Transaction_ID").value = Transaction_ID
        RsNotesGeneral("NoteDate").value = XPDtbBill.value
        RsNotesGeneral("NoteType").value = 180
        RsNotesGeneral("Note_Value").value = Null
        RsNotesGeneral("NoteSerial").value = IIf(Trim(TxtNoteSerialV) = "", Null, Trim(TxtNoteSerialV))
       ' RsNotesGeneral("NoteSerial1").value = IIf(Trim(TxtNoteSerial1V) = "", Null, Trim(TxtNoteSerial1V))
         RsNotesGeneral("remark").value = IIf(Trim(TxtNoteSerial1V) = "", Null, Trim(TxtNoteSerial1V))
        RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
        RsNotesGeneral("numbering_type1").value = sand_numbering_type(10) '«–‰ wvt
        RsNotesGeneral("sanad_year").value = year(XPDtbBill.value)
        RsNotesGeneral("sanad_month").value = Month(XPDtbBill.value)
        RsNotesGeneral("branch_no").value = val(Me.dcBranch.BoundText)
        'RsNotes("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
        RsNotesGeneral.update
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

        CREATE_VOUCHER_GE Transaction_ID, TxtNoteSerialV, TxtNoteSerial1V, general_noteid, val(Me.dcBranch.BoundText)
        sql = " update Transactions Set NoteID = " & general_noteid & " where Transaction_ID = " & Transaction_ID
        
        Cn.Execute sql
    End If
 
    '
    CreateIssueVoucher = True
 Exit Function
ErrTrap:
CreateIssueVoucher = False
End Function

Function CREATE_VOUCHER_GE(Transaction_ID As Long, TxtNoteSerialV As String, TxtNoteSerial1V As String, general_noteid As Long, BranchID As Integer) As Boolean
    On Error GoTo ErrTrap
    CREATE_VOUCHER_GE = False
    
    Dim usedaccount As Integer
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
    Dim TOTAL_COST As Single
    Dim LngCurItemID As Double
    Dim LngUnitID As Long
    Dim UnitFactor As Double
    Dim TOTAL_COST2 As Double
    Dim line_value2 As Double
    Dim SngTemp2 As Double
    TOTAL_COST2 = 0
    TOTAL_COST = 0
    With FG

        For i = 1 To FG.rows - 1

            If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" And val(FG.TextMatrix(i, FG.ColIndex("ItemType"))) <> 1 Then
                LngCurItemID = val(FG.TextMatrix(i, FG.ColIndex("Code")))
                LngUnitID = val(FG.cell(flexcpData, i, FG.ColIndex("UnitID")))
            
                GetUnitNoOfItems LngCurItemID, LngUnitID, UnitFactor
                    
                
            
               If CboRetrunType.ListIndex = 0 Then ' »›« Ê—…  »”⁄— «·›« Ê—… ' „ﬁÌœ
      '  Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,ItemSize,UnitId,ShowQty,QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID ,ProductionDate,ExpiryDate,LotNO )SELECT  costprice,guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , Quantity, costprice/ QtyBySmalltUnit ,ColorID,ItemSize, UnitId, ShowQty, QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID,ProductionDate,ExpiryDate,LotNO  From dbo.Transaction_Details Where  Transaction_ID = " & XPTxtBillID.text
      TOTAL_COST = TOTAL_COST + FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count")) * val(txt_Currency_rate)
      TOTAL_COST2 = TOTAL_COST2 + FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count"))
      Else ' »›« Ê—… »”⁄— «· ﬂ·€… ' ›Ì— „ﬁÌœ
      'TOTAL_COST = TOTAL_COST + (FG.TextMatrix(i, FG.ColIndex("Count")) * ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, , LngUnitID)) * val(txt_Currency_rate.text)
      ' TOTAL_COST2 = TOTAL_COST2 + (FG.TextMatrix(i, FG.ColIndex("Count")) * ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, , LngUnitID))
      
            TOTAL_COST = TOTAL_COST + val(FG.TextMatrix(i, FG.ColIndex("Count")) * val(FG.TextMatrix(i, FG.ColIndex("price")))) * val(txt_Currency_rate.text)
       TOTAL_COST2 = TOTAL_COST2 + (FG.TextMatrix(i, FG.ColIndex("Count")) * ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, , LngUnitID))

      End If
      
      
            
            End If

        Next i

    End With

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·œ«∆‰
    SngTemp = NewGrid.GetItemsTotal(ItemsGoodType) + val(LblValueAdded.Caption)
    SngTemp2 = SngTemp
    SngTemp = SngTemp * val(txt_Currency_rate.text)
    my_branch = BranchID

    If TOTAL_COST > 0 Then
   
        If detect_inventory_work_type = 1 Then

            Account_Code_dynamic = get_account_code_branch(0, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„Œ“Ê‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    GoTo ErrTrap
         
                End If
            End If

            If val(DCDocTypes.BoundText) > 0 Then
                getDocAccounts val(DCDocTypes.BoundText), , , , , StrTempAccountCode, , , , , usedaccount

                If StrTempAccountCode = "" And usedaccount = 1 Then
                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·œ«∆‰ ·”‰œ «·’—›", vbCritical
                    GoTo ErrTrap
                ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
                ElseIf usedaccount = 0 Then
                    StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…
                End If

            Else
                StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…
            End If
            
            ' StrTempAccountCode = "a1a2a5" '„Œ“Ê‰ «·»÷«⁄…
            StrTempDes = "”‰œ ’—› —ﬁ„ " & Me.TxtTransSerial.text & " »‰«¡ ⁄·Ï  „—œÊœ«  „‘ —Ì«  »—ﬁ„  " & Me.TxtNoteSerial1.text
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , TOTAL_COST2, DcCurrency.text, val(txt_Currency_rate.text), , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 2 Then
            '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
    
            Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")

            If Account_Code_dynamic = "" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                GoTo ErrTrap
            End If
    
            If val(DCDocTypes.BoundText) > 0 Then
                getDocAccounts val(DCDocTypes.BoundText), , , , , StrTempAccountCode, , , , , usedaccount

                If StrTempAccountCode = "" And usedaccount = 1 Then
                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·œ«∆‰ ·”‰œ «·’—›", vbCritical
                    GoTo ErrTrap
                ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
                ElseIf usedaccount = 0 Then
                    StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…
                End If

            Else
                StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…
            End If

            '            StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰
            ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "”‰œ    ’—› —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï  „—œÊœ«  „‘ —Ì«  »—ﬁ„  " & Me.TxtNoteSerial1.text
            Else
                StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V & "   Return Purchase NO  " & Me.TxtNoteSerial1.text
            End If

            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , TOTAL_COST2, DcCurrency.text, val(txt_Currency_rate.text), , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value As Single

            With FG

                For i = 1 To FG.rows - 1

                    If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                        ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                        groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 0)

                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”⁄·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        line_value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value) * FG.TextMatrix(i, FG.ColIndex("Count"))
                        line_value2 = line_value
                       line_value = line_value * val(txt_Currency_rate.text)
                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "”‰œ    ’—› —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï  „—œÊœ«  „‘ —Ì«  »—ﬁ„  " & Me.TxtNoteSerial1.text
                        Else
                            StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V & " Return Purcahse No: " & Me.TxtNoteSerial1.text
                        End If
            
                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , line_value2, DcCurrency.text, val(txt_Currency_rate.text), , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

                Next i

            End With

        End If

        '«·ÿ—› «·„œÌ‰
        SngTemp = NewGrid.GetItemsTotal(ItemsGoodType) + val(LblValueAdded.Caption)
        SngTemp = SngTemp * val(txt_Currency_rate.text)
       SngTemp2 = NewGrid.GetItemsTotal(ItemsGoodType) + val(LblValueAdded.Caption)
        If TOTAL_COST > 0 Then
            If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Then

                Account_Code_dynamic = get_account_code_branch(5, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ   „—œÊœ«  «·„‘ —Ì«  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
         
                    End If
                End If

                If val(DCDocTypes.BoundText) > 0 Then
                    getDocAccounts val(DCDocTypes.BoundText), , , , StrTempAccountCode, , , , , usedaccount

                    If StrTempAccountCode = "" And usedaccount = 1 Then
                        MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·„œÌ‰ «·Œ«’ »”‰œ ’—› «·„Ê«œ", vbCritical
                        GoTo ErrTrap
                    ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
                    ElseIf usedaccount = 0 Then
                        StrTempAccountCode = Account_Code_dynamic '„—œÊœ«  «·„‘ —Ì« 
        
                    End If

                Else
                    StrTempAccountCode = Account_Code_dynamic '  „—œÊœ«  «·„‘ —Ì« 
                End If
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ    ’—› —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï  „—œÊœ«  „‘ —Ì«  »—ﬁ„  " & Me.TxtNoteSerial1.text
                Else
                    StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V & " Return Purchase No " & Me.TxtNoteSerial1.text
                End If
            
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , TOTAL_COST2, DcCurrency.text, val(txt_Currency_rate.text), , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
         
            ElseIf detect_inventory_work_type = 3 Then

                With FG

                    For i = 1 To FG.rows - 1

                        If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                            groupAccount = get_item_group_account_in_branch(FG.TextMatrix(i, FG.ColIndex("Code")), val(my_branch), 1)

                            '  groupAccount = get_item_group_account_inventory(FG.TextMatrix(I, FG.ColIndex("Code")), DCboStoreName.BoundText, 4)
                            If groupAccount = "Error" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»    ﬂ·›… «·„»Ì⁄«    ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                                Else
                                    MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                                End If

                                GoTo ErrTrap
                            End If

                            line_value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value) * FG.TextMatrix(i, FG.ColIndex("Count"))
                             line_value2 = line_value
                             line_value = line_value * val(txt_Currency_rate.text)
                             If SystemOptions.UserInterface = ArabicInterface Then
                                StrTempDes = "”‰œ    ’—› —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï  „—œÊœ«  „‘ —Ì«  »—ﬁ„  " & Me.TxtNoteSerial1.text
                            Else
                                StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V & "  Retuen Purcahse NO  " & Me.TxtNoteSerial1.text
                            End If
            
                            LngDevNO = LngDevNO + 1

                            If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , line_value2, DcCurrency.text, val(txt_Currency_rate.text), , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                                GoTo ErrTrap
                            End If
    
                        End If

                    Next i

                End With

            End If
        End If
    End If

    Dim StrSQL  As String
    StrSQL = "UPDATE Transactions SET NOTS=" & Transaction_ID & " WHERE Transaction_ID=" & val(Me.XPTxtBillID.text)
    Cn.Execute StrSQL
    updateNotesValueAndNobytext CDbl(general_noteid)
    
    CREATE_VOUCHER_GE = True
    Exit Function
ErrTrap:
CREATE_VOUCHER_GE = False
End Function

Function CheckAccount() As Boolean
Dim StrTempAccountCode As String
Dim usedaccount As Integer
Dim Account_Code_dynamic As String
    CheckAccount = False
    'Dcombos.GetDocTypebyid Me.DCDocTypes, 21, val(Me.dcBranch.BoundText)

    If val(DCDocTypes.BoundText) > 0 Then
        getDocAccounts val(DCDocTypes.BoundText), , StrTempAccountCode, , , , , usedaccount

        If StrTempAccountCode = "" And usedaccount = 1 Then
            MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«»  «·œ«∆‰ «·Œ«’ »«·„—œÊœ«   ", vbCritical
            GoTo ErrTrap
        End If
               
    End If



 


 
    
   If val(DCDocTypes.BoundText) > 0 Then
        getDocAccounts val(DCDocTypes.BoundText), , , , StrTempAccountCode, , , , , usedaccount

        If StrTempAccountCode = "" And usedaccount = 1 Then
            MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·„œÌ‰ ·”‰œ  «·’—›", vbCritical
            GoTo ErrTrap
        End If
 
    End If
     
     
    
    
    
    
    CheckAccount = True
    Exit Function
ErrTrap:

    CheckAccount = False
End Function

Private Sub SaveData()
    Dim Msg As String
    Dim RowNum As Integer
    Dim RSTransDetails As ADODB.Recordset
    Dim RsNotes As ADODB.Recordset
    Dim RsTemp  As New ADODB.Recordset
    Dim RsTest As New ADODB.Recordset
    Dim RsRepeat As ADODB.Recordset
    Dim StrSQL As String
    Dim StrSqlDel As String
    Dim BeginTrans As Boolean
    Dim LngItemID As Long

  On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass

    If Me.TxtModFlg.text <> "R" Then
    If Not IsSaveWithOutMsg Then
        If DBCboClientName.text = "" Then
            Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·„Ê—œ"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DBCboClientName.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    If DcCurrency.BoundText = "" Then
    
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "«Œ — «·⁄„·… «Ê·« "
        Else
            Msg = "Select Currency First"
    
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        DcCurrency.SetFocus
        Sendkeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    
    End If
        If DCboStoreName.text = "" Then
            Msg = "ÌÃ»  ÕœÌœ «·„Œ“‰"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCboStoreName.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        If XPDtbBill.value = "" Then
            Msg = "ÌÃ»  ÕœÌœ  «—ÌŒ ⁄„·Ì… ≈—Ã«⁄ «·„‘ —Ì« "
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPDtbBill.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If CboPayMentType.ListIndex = -1 Then
            Msg = "ÌÃ»  ÕœÌœ ÿ—Ìﬁ… «·œ›⁄"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CboPayMentType.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If Me.XPChkPayType(0).value = vbChecked Then
            If Me.DcboBox.BoundText = "" Then
                Msg = "ÌÃ»  ÕœÌœ «·Œ“‰…..!!!"
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If

        If Me.XPChkPayType(1).value = vbChecked Then
            If DateDiff("d", Me.DtpDelayDate.value, Date) > 0 Then
                Msg = "ÌÃ»  ÕœÌœ  «—ÌŒ ≈” Õﬁ«ﬁ «·ﬁÌ„… «·√Ã·…"
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If

        If XPChkPayType(2).value = vbChecked Then
            If DCboBankName.BoundText = "" Then
                Screen.MousePointer = vbDefault
                MsgBox "ÌÃ»  ÕœÌœ «”„ «·»‰ﬂ", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            If Trim(Me.XPTxtChqueNum.text) = "" Then
                Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!!"
                Screen.MousePointer = vbDefault
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            If Check_CheckNum(Me.XPTxtChqueNum.text, val(Me.XPTxtBillID.text), Me.TxtModFlg.text, 0) = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If

        If val(XPTxtValue(0).text) + val(XPTxtValue(1).text) + val(XPTxtValue(2).text) > val(XPTxtSum.text) Then
            Msg = "≈Ã„«·Ì «·ﬁÌ„ «·„Õ’·… Ê«·„ƒÃ·…" & CHR(13)
            Msg = Msg + "√ﬂ»— „‰ ≈Ã„«·Ì «·›« Ê—…"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPTab301.CurrTab = 1
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        ' If Val(XPTxtValue(0).text) + Val(XPTxtValue(1).text) + Val(XPTxtValue(2).text) < Val(XPTxtSum.text) Then
        '     Msg = "≈Ã„«·Ì «·ﬁÌ„ «·„Õ’·… Ê«·„ƒÃ·…" & Chr(13)
        '     Msg = Msg + "√ﬁ· „‰ ≈Ã„«·Ì «·›« Ê—…"
        '     MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '     XPTab301.CurrTab = 1
        '     Screen.MousePointer = vbDefault
        '     Exit Sub
        ' End If
    
        If Me.Ele(2).Visible = True Then
            If Me.CboRetrunType.ListIndex = -1 Then
                Msg = "»—Ã«¡ ≈Œ Ì«— ‰Ê⁄ «·√— Ã«⁄.."
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                CboRetrunType.SetFocus
                 Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            ElseIf Me.CboRetrunType.ListIndex = 0 Then

                If Trim(Me.TxtInvSerial.text) = "" Then
                    Msg = "›Ï Õ«·… «·√— Ã«⁄ «·„ﬁÌœ »›« Ê—… «·‘—«¡ "
                    Msg = Msg & CHR(13) & "ÌÃ» ﬂ «»… —ﬁ„ ›« Ê—… «·‘—«¡"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    TxtInvSerial.SetFocus
                    Screen.MousePointer = vbDefault
                    Exit Sub
                ElseIf CheckInvData = False Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
        End If
        End If

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
        ' „—«Ã⁄Â «”⁄«— «· ﬂ·›…
 
        '    For RowNum = 1 To FG.Rows - 1
        '                        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
        '
        '                              If CboRetrunType.ListIndex = 0 Then '„ﬁÌœ »›« Ê—…
        '                                            If Val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice"))) = 0 Then
        '                                              MsgBox "«·’‰›   " & FG.TextMatrix(RowNum, FG.ColIndex("Name")) & " €Ì— „Õœœ ”⁄—  ﬂ·› Â Ê·–·ﬂ ·« Ì„ﬂ‰ « „«„ ⁄„·ÌÂ «·«—Ã«⁄ "
        '                                              Exit Sub
        '                                             End If
                                 
        '                            Else '€Ì— „ﬁÌœ »›« Ê—…
        '                                                If Val(ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(RowNum, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod)) = 0 Then
        '                                                 MsgBox "«·’‰›   " & FG.TextMatrix(RowNum, FG.ColIndex("Name")) & " €Ì— „Õœœ ”⁄—  ﬂ·› Â Ê·–·ﬂ ·« Ì„ﬂ‰ « „«„ ⁄„·ÌÂ «·«—Ã«⁄ "
        '                                                 Exit Sub
        '                                                End If
        '
        '                            End If
        '
        '                       End If
        '    Next RowNum
    
        ' „—«Ã⁄Â «”⁄«— «· ﬂ·›…
        Dim LngCurItemID  As Double
        Dim LngUnitID As Long
        Dim UnitFactor As Double
        Dim DblItemCostPrice As Double
        Dim UnitID As Long
        Dim MsgBoxResult  As Integer

        For RowNum = 1 To FG.rows - 1

            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                      
                If CboRetrunType.ListIndex = 0 Then '„ﬁÌœ »›« Ê—…
                 '   If val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice"))) = 0 Then
                 '       MsgBox "«·’‰›   " & FG.TextMatrix(RowNum, FG.ColIndex("Name")) & " €Ì— „Õœœ ”⁄—  ﬂ·› Â Ê·–·ﬂ ·« Ì„ﬂ‰ « „«„ ⁄„·ÌÂ «·«—Ã«⁄ "
                 '       Exit Sub
                 '   End If
                                 
                Else '€Ì— „ﬁÌœ »›« Ê—…

                    If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" And val(FG.TextMatrix(RowNum, FG.ColIndex("ItemType"))) <> 1 Then
                                      
                        LngCurItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                        LngUnitID = val(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
            
                        GetUnitNoOfItems LngCurItemID, LngUnitID, UnitFactor
                                            
                        UnitID = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))

                        If val(ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(RowNum, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Text1.text), UnitID)) = 0 Then
                            'If Val(ModItemCostPrice.GetCostItemPrice(Fg.TextMatrix(RowNum, Fg.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod)) = 0 Then
                            MsgBoxResult = MsgBox("«·’‰›   " & FG.TextMatrix(RowNum, FG.ColIndex("Name")) & " €Ì— „Õœœ ”⁄—  ﬂ·› Â —»„« ·⁄œ„ ÊÃÊœ ﬂ„Ì… Ê·–·ﬂ ·« Ì„ﬂ‰ « „«„ ⁄„·ÌÂ «·«—Ã«⁄ " & CHR(13) & "Â·  —Ìœ Õ”«»  ﬂ·› … ⁄·Ï «”«” «Œ— ”‰œ ’—› «‰ ÊÃœ ‰⁄„ «Ê ·« ", vbYesNo)

                            If MsgBoxResult = vbYes Then
                                FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = getLastCostPriceForItems(FG.TextMatrix(RowNum, FG.ColIndex("Code")), UnitID)
                            Else
                                MsgBoxResult = MsgBox("«·’‰›   " & FG.TextMatrix(RowNum, FG.ColIndex("Name")) & " €Ì— „Õœœ ”⁄—  ﬂ·› Â —»„« ·⁄œ„ ÊÃÊœ ﬂ„Ì… Ê·–·ﬂ ·« Ì„ﬂ‰ « „«„ ⁄„·ÌÂ «·«—Ã«⁄ " & CHR(13) & "Â·  —Ìœ Õ”«»  ﬂ·› … ⁄·Ï «”«” «‰ ÌﬂÊ‰ ‰›” ”⁄— «·„—œÊœ«  ‰⁄„ / ·« «ﬁÊ„ »√œŒ«· ”⁄— ÌœÊÌ ", vbYesNo)

                                If MsgBoxResult = vbYes Then
                                    FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = FG.TextMatrix(RowNum, FG.ColIndex("Price"))
                                Else
                                    DblItemCostPrice = InputBox("«œŒ· «·”⁄— ··’‰›" & FG.TextMatrix(RowNum, FG.ColIndex("Name")))
                                    FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = val(DblItemCostPrice)
                                End If
                                                                       
                            End If
                                                    
                        Else
                            FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = val(ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(RowNum, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Text1.text), UnitID))
If FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = 0 Then
        FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = val(FG.TextMatrix(RowNum, FG.ColIndex("Price")))
End If
                            '   Exit Sub
                        End If
                                             
                    End If
                End If
            End If

        Next RowNum
    
        'Check the Items Grid
        Me.XPTab301.CurrTab = 0

        If NewGrid.CheckDataEntered = False Then
            Exit Sub
        End If
        If CheckAccount = False Then
        Exit Sub
        End If
        

        '--------------------------------------------------------------
        If CheckRetrunInv = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        CurrentVoucherNo = ""
        CurrentVoucherSerialNo = ""
        my_branch = dcBranch.BoundText
        If TxtNoteSerial.text = "" Then
            If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
            Else
                       
                If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                Else
                    TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbBill.value)
                End If
            End If
        End If
       Dim NoteSerial1str As String
        
        If TxtNoteSerial1.text = "" Then
        NoteSerial1str = Voucher_coding(val(my_branch), XPDtbBill.value, 15, 230, , 5, , val(DCboStoreName.BoundText))
        
                        If NoteSerial1str = "error" Then
                            MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ „—œÊœ«  „‘ —Ì«   ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                        Else
                                   
                                        If NoteSerial1str = "" Then
                                            MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„  ”‰œ «·«—Õ«⁄ ÌœÊÌ«  ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                                        Else
                                            TxtNoteSerial1.text = NoteSerial1str
                                        End If
                        End If
        End If
             
        If Voucher_coding(val(my_branch), XPDtbBill.value, 10, 180, , 19, , val(DCboStoreName.BoundText)) = "" Then
                                
            If Trim$(TxtManualNo1) = "" Then
                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ”‰œ «·’—›  ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
            
            Else
                TxtNoteSerial1V = TxtManualNo1
            End If
            
        End If
              Screen.MousePointer = vbArrowHourglass
        Cn.BeginTrans
        BeginTrans = True
        'Create big notes
        Set RsNotesGeneral = New ADODB.Recordset
       ' RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
       StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText


        If Me.TxtModFlg.text = "N" Then
            Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
            TxtNoteSerial1V = ""
        Else
            'StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & Val(rs("Transaction_ID").value)
            'Cn.Execute StrSqlDel, , adExecuteNoRecords
        
            StrSqlDel = "delete From Notes where Transaction_ID=" & val(Me.XPTxtBillID.text) ' Val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
        
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & val(Me.XPTxtBillID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        
       '     StrSQL = "delete From Notes where noteid=" & val(TXTNoteID.Text)
       '     Cn.Execute StrSQL, , adExecuteNoRecords
        
        
                 StrSQL = "Delete From Notes Where  NoteType= 230 and NoteID=" & val(Me.TXTNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
        
            CurrentVoucherNo = GetVoucherGLNO(val(Text1.text), CurrentVoucherSerialNo)
            DeleteTransactiomsVoucher val(Text1.text)
        
            general_noteid = val(TXTNoteID.text)
        End If

        RsNotesGeneral.AddNew
        RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
        general_noteid = RsNotesGeneral("NoteID").value
        TXTNoteID.text = general_noteid
        ' RsNotesGeneral("Transaction_ID").value = Val(XPTxtBillID.text)
        RsNotesGeneral("NoteDate").value = XPDtbBill.value
        RsNotesGeneral("NoteType").value = 230
        RsNotesGeneral("Note_Value").value = val(LblTotal.Caption)
        RsNotesGeneral("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
        
    RsNotesGeneral("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
RsNotesGeneral("remark").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))

        RsNotesGeneral("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '

        RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
        RsNotesGeneral("numbering_type1").value = sand_numbering_type(5) '    „—œÊœ«  „‘ —Ì« 
        RsNotesGeneral("sanad_year").value = year(XPDtbBill.value)
        RsNotesGeneral("sanad_month").value = Month(XPDtbBill.value)
        RsNotesGeneral("branch_no").value = val(Me.dcBranch.BoundText)
        'RsNotes("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
     '''// 31 05 2015
   
        RsNotesGeneral.update
    
        '--------------------------Start Saving------------------------
        Set RSTransDetails = New ADODB.Recordset
     '   RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
     
 StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
   
        
        Set RsNotes = New ADODB.Recordset
        'RsNotes.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
           StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText



        If Me.TxtModFlg.text = "N" Then
            XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            TxtTransSerial.text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=5"))
            rs.AddNew
            rs("Transaction_ID").value = val(XPTxtBillID.text)
        ElseIf Me.TxtModFlg.text = "E" Then
        Cn.Execute "Delete from TransactionValueAdded where Transaction_ID=" & val(Me.XPTxtBillID.text) & ""
            StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
            StrSqlDel = "delete From Notes where Transaction_ID=" & val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & val(Me.XPTxtBillID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        End If

        If CboPayMentType.ListIndex = 0 Then
            rs("BoxID").value = IIf(DcboBox.BoundText = "", Null, val(DcboBox.BoundText))
        Else
            rs("BoxID").value = Null
      
        End If
        rs("txtManulaVat").value = val(txtManulaVat.text)
        
        rs("Currency_rate").value = IIf(Not IsNumeric(txt_Currency_rate.text), 1, txt_Currency_rate.text)
        rs("Currency_id").value = IIf(DcCurrency.BoundText = "", Null, val(DcCurrency.BoundText))
        rs("Transporter").value = IIf(Transporter.text = "", "", (Transporter.text))
        rs("VAT").value = val(TxtValueAdded.text)
        rs("Typ").value = val(Me.Dcbtyp.ListIndex)
        rs("ResonVAT").value = TXtResonVAT.text
        rs("Doctype").value = IIf(Me.DCDocTypes.BoundText = "", Null, val(DCDocTypes.BoundText))
        rs("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
        rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
        rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
        rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '
        rs("NoteId").value = val(TXTNoteID.text)
        rs("ManualNo1").value = IIf(TxtManualNo1.text = "", Null, val(TxtManualNo1.text))
        rs("Transaction_Serial").value = IIf(Trim(Me.TxtTransSerial.text) = "", Null, Trim(Me.TxtTransSerial.text))
        rs("Transaction_Date").value = XPDtbBill.value
        rs("Transaction_Type").value = 5
        rs("CusID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
        rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, val(DCboStoreName.BoundText))
        rs("UserID").value = user_id
        rs("VATNO").value = IIf(Trim(Me.TxtVATNO.text) = "", Null, Trim(Me.TxtVATNO.text))
        If CboPayMentType.ListIndex = -1 Then
            rs("PaymentType").value = 0
        Else
            rs("PaymentType").value = val(CboPayMentType.ListIndex)
        End If

        If Me.CboRetrunType.ListIndex = 0 Then
            rs("ReturnID").value = val(Me.TxtInvID.text)
                        rs("ReturnSerial").value = Me.TxtInvSerial.text
        Else
            rs("ReturnID").value = Null
                        rs("ReturnSerial").value = Null
        End If
''// 31 05 2015
 rs("Emp_ID").value = IIf(DcboEmp.BoundText = "", Null, DcboEmp.BoundText)

rs("ManualNO").value = IIf(TxtManualNO11.text = "", Null, val(TxtManualNO11.text))
  rs("OrderSupply").value = IIf(Trim(Me.TxtOrderSupply.text) = "", Null, Trim(Me.TxtOrderSupply.text))
     rs("BillSupplier").value = IIf(Trim(Me.TxtBillSupplier.text) = "", Null, Trim(Me.TxtBillSupplier.text))
     rs("ReasonReturns").value = IIf(Trim(Me.TxtReasonReturns.text) = "", Null, Trim(Me.TxtReasonReturns.text))

rs("Transaction_NetValue").value = val(LblTotal.Caption)

        rs.update

        For RowNum = 1 To FG.rows - 1

            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").value = val(XPTxtBillID.text)
                RSTransDetails("TypeVAT").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("TypeVAT")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("TypeVAT"))))
                RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
                RSTransDetails("Quantity").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))

                '            RSTransDetails("ItemName").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Name")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Name"))))
                If Not FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "" Then
                    StrSQL = "select * From TblItems where ItemID=" & FG.TextMatrix(RowNum, FG.ColIndex("Name"))
                    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (RsTemp.EOF Or RsTemp.BOF) Then
                        If RsTemp("HaveSerial").value = True Then
                            RSTransDetails("ItemSerial").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Serial")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("Serial")))
                        End If
                    End If

                    RsTemp.Close
                End If

                RSTransDetails("ItemDiscountType").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType"))))
                RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
            
                RSTransDetails("ItemDiscount").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal"))))

                RSTransDetails("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
                RSTransDetails("showPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
                RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
                RSTransDetails("Remarks").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Remarks")) = ""), Null, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("Remarks"))))
                RSTransDetails("order_no").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("order_no")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("order_no")))
                        
                RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
                RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
             
                RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
 
                RSTransDetails("UnitID").value = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))

                If CboRetrunType.ListIndex = 0 Then '„ﬁÌœ »›« Ê—…
                    RSTransDetails("CostPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice"))))
           
                Else '€Ì— „ﬁÌœ »›« Ê—…
                    RSTransDetails("CostPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice"))))
       
                End If

                RSTransDetails("ShowQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
                RSTransDetails("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
                RSTransDetails("Vat").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Vat")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("Vat"))))
                RSTransDetails("Vatyo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Vatyo")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("Vatyo"))))
  
                Dim RsUnitData As ADODB.Recordset
                'Dim LngCurItemID As Long
                'Dim LngUnitID As Long
                Dim DblQty As Double
        
                LngCurItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                LngUnitID = val(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
                DblQty = val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))

                StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
                StrSQL = StrSQL + " AND UnitID=" & LngUnitID
                Set RsUnitData = New ADODB.Recordset
                RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
        
                    RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                    RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
    
                    RSTransDetails("Price").value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
                  
                End If
            
                RSTransDetails("ProductionDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate"))), "DD/mm/YYYY"))
                RSTransDetails("ExpiryDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate"))), "DD/mm/YYYY"))
                RSTransDetails("LotNO").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("LotNO")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("LotNO")))
                         Dim OldQty As Double
             Dim OldCost As Double
              Dim NewQty As Double
               Dim NewCost As Double
               
'getItemCostData XPDtbBill.value, RSTransDetails("Item_ID").value, val(DCboStoreName.BoundText), val(Me.XPTxtBillID.Text), OldQty, OldCost, NewQty, NewCost,,LngUnitID
'       RSTransDetails("OldQty").value = NewQty
'       RSTransDetails("OldCost").value = NewCost
'
'      RSTransDetails("NewQty").value = RSTransDetails("OldQty").value - RSTransDetails("Quantity").value
'       RSTransDetails("NewCost").value = RSTransDetails("OldCost").value ' ((RSTransDetails("OldQty").value * RSTransDetails("OldCost").value) + (RSTransDetails("Quantity").value * RSTransDetails("Price").value)) / (RSTransDetails("Quantity").value + RSTransDetails("OldQty").value)
       
       
                RSTransDetails.update
            End If

            If FG.rows > 10 Then
                If RowNum = 8 Then FG.Refresh
            End If

        Next RowNum
    
        '    If Me.XPChkPayType(0).value = Checked Then
        '        RsNotes.AddNew
        '        RsNotes("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
        '        If Me.TxtModFlg.text = "N" Then
        '            RsNotes("NoteSerial").value = CStr(new_id("Notes", "NoteSerial", "", True))
        '            XPTxtSerial(0).text = RsNotes("NoteSerial").value
        '        ElseIf Trim(XPTxtSerial(0).text) <> "" Then
        '            RsNotes("NoteSerial").value = Trim(XPTxtSerial(0).text)
        '        Else
        '            RsNotes("NoteSerial").value = CStr(new_id("Notes", "NoteSerial", "", True))
        '            XPTxtSerial(0).text = RsNotes("NoteSerial").value
        '        End If
        '        RsNotes("Transaction_ID").value = Val(XPTxtBillID.text)
        '        RsNotes("NoteType").value = 0
        '        RsNotes("NoteDate").value = XPDtbBill.value
        '        RsNotes("Note_Value").value = IIf(XPTxtValue(0).text = "", Null, Val(XPTxtValue(0).text))
        '        RsNotes("Member_ID").value = IIf(DBCboClientName.BoundText = "", Null, Val(DBCboClientName.BoundText))
        '        RsNotes("BankID").value = Null
        '        RsNotes("BoxID").value = IIf(DcboBox.BoundText = "", Null, Val(DcboBox.BoundText))
        '        RsNotes("CusID").value = Null
        '        RsNotes.update
        '    End If
    
        If Me.XPChkPayType(1).value = Checked Then
            RsNotes.AddNew
            RsNotes("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
        
            'If Me.TxtModFlg.text = "N" Then
            '    RsNotes("NoteSerial").value = CStr(new_id("Notes", "NoteSerial", "", True))
            '    XPTxtSerial(1).text = RsNotes("NoteSerial").value
            'ElseIf Trim(XPTxtSerial(1).text) <> "" Then
            '    RsNotes("NoteSerial").value = Trim(XPTxtSerial(1).text)
            'Else
            '    RsNotes("NoteSerial").value = CStr(new_id("Notes", "NoteSerial", "", True))
            '    XPTxtSerial(1).text = RsNotes("NoteSerial").value
            'End If
            RsNotes("NoteSerial1").value = Me.TxtNoteSerial1.text
            RsNotes("NoteSerial").value = Null
            RsNotes("Transaction_ID").value = val(XPTxtBillID.text)
            RsNotes("NoteType").value = 1
            RsNotes("NoteDate").value = XPDtbBill.value
            RsNotes("Note_Value").value = IIf(XPTxtValue(1).text = "", Null, val(XPTxtValue(1).text))
            RsNotes("Member_ID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
            RsNotes("BankID").value = Null
            RsNotes("CusID").value = Null
            RsNotes("DueDate").value = DtpDelayDate.value
            RsNotes.update
        End If
    
        If Me.XPChkPayType(2).value = Checked Then
            RsNotes.AddNew
            RsNotes("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
            RsNotes("Transaction_ID").value = val(XPTxtBillID.text)
            RsNotes("NoteType").value = 2
            RsNotes("NoteDate").value = XPDtbBill.value
            RsNotes("Note_Value").value = IIf(XPTxtValue(2).text = "", Null, val(XPTxtValue(2).text))
            RsNotes("Member_ID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
            RsNotes("BankID").value = IIf(DCboBankName.BoundText = "", Null, val(DCboBankName.BoundText))
            RsNotes("ChqueNum").value = IIf(XPTxtChqueNum.text = "", "", Trim(XPTxtChqueNum.text))
            RsNotes("DueDate").value = XPDTPDueDate.value
            RsNotes("CusID").value = Me.DBCboClientName.BoundText
            RsNotes.update
        End If

        Dim LngDevID As Long
        Dim LngDevNO  As Integer
        Dim StrTempAccountCode As String
        Dim StrTempDes As String
        Dim SngTemp As Variant
        Dim SngTemp2 As Variant
    
        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
        '----------------
        SngTemp = NewGrid.GetItemsTotal(ItemsGoodType) - val(Me.LblDiscountsTotal.Caption) + val(LblValueAdded.Caption)
        SngTemp2 = NewGrid.GetItemsTotal(ItemsGoodType) - val(Me.LblDiscountsTotal.Caption) + val(LblValueAdded.Caption)
        SngTemp2 = SngTemp2
       SngTemp = Round(SngTemp, SystemOptions.SysDefCurrencyForamt) * val(txt_Currency_rate.text)
       SngTemp2 = Round(SngTemp2, SystemOptions.SysDefCurrencyForamt)
        '«·ÿ—› «·„œÌ‰
        '    If Me.XPChkPayType(0).value = vbChecked Then
        

            Dim DocumentDebitAccountCode As String
               If val(DCDocTypes.BoundText) > 0 Then
                     getDocAccounts val(DCDocTypes.BoundText), DocumentDebitAccountCode, StrTempAccountCode, , , , , usedaccount
                     

                If DocumentDebitAccountCode = "" And usedaccount = 1 Then
                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·„œÌ‰ ·„—œÊœ«  «·„‘ —Ì«   ", vbCritical
                    GoTo ErrTrap
            End If
            End If
                    
                    
        
        If CboPayMentType.ListIndex = 0 Then
            StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
            StrTempDes = "„— Ã⁄ „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.text
            LngDevNO = LngDevNO + 1

                    If DocumentDebitAccountCode <> "" Then
                    StrTempAccountCode = DocumentDebitAccountCode
                    End If

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , SngTemp2, DcCurrency.text, val(txt_Currency_rate.text), , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
        End If



            
            
        'If Me.XPChkPayType(1).value = vbChecked Then
        If CboPayMentType.ListIndex = 1 Then
            StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
            StrTempDes = "„— Ã⁄ „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.text
            LngDevNO = LngDevNO + 1
                    If DocumentDebitAccountCode <> "" Then
                    StrTempAccountCode = DocumentDebitAccountCode
                    End If

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , SngTemp2, DcCurrency.text, val(txt_Currency_rate.text), , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
        End If

        If Me.XPChkPayType(2).value = vbChecked Then
            StrTempAccountCode = "a1a2a4" '√Ê—«ﬁ «·ﬁ»÷
            StrTempDes = "⁄œœ 1 " & "  ‘Ìﬂ«  " & CHR(13)
            StrTempDes = "„— Ã⁄ „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.text
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , SngTemp2, DcCurrency.text, val(txt_Currency_rate.text), , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
        End If

        '    If Val(Me.LblDiscountsTotal.Caption) > 0 Then
        '        StrTempAccountCode = "a3a5" '«·Œ’„ «·„”„ÊÕ »Â
        '        StrTempDes = "„— Ã⁄ „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.text
        '        LngDevNO = LngDevNO + 1
        '        If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Val(Me.LblDiscountsTotal.Caption), _
        '            0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , Val(Me.dcBranch.BoundText)) = False Then
        '            GoTo ErrTrap
        '        End If
        '    End If
        '«·ÿ—› «·œ«∆‰
 
        Dim Account_Code_dynamic As String
        Dim i As Single
        ' LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
        '«·ÿ—› «·œ«∆‰
         my_branch = dcBranch.BoundText
        SngTemp = NewGrid.GetItemsTotal(ItemsGoodType) - val(Me.LblDiscountsTotal.Caption)
        SngTemp2 = NewGrid.GetItemsTotal(ItemsGoodType) - val(Me.LblDiscountsTotal.Caption)
        SngTemp = SngTemp * val(txt_Currency_rate.text)
        SngTemp = Round(SngTemp, SystemOptions.SysDefCurrencyForamt)
        SngTemp2 = Round(SngTemp2, SystemOptions.SysDefCurrencyForamt)
        If SngTemp > 0 Then

            If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Then
                Account_Code_dynamic = get_account_code_branch(5, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«» „—œÊœ«  «·„‘ —Ì«   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
         
                    End If
                End If

      '          StrTempAccountCode = Account_Code_dynamic '„—œÊœ«  «·„‘ —Ì« 
            
            'Dim DocumentDebitAccountCode As String
               If val(DCDocTypes.BoundText) > 0 Then
                     getDocAccounts val(DCDocTypes.BoundText), DocumentDebitAccountCode, StrTempAccountCode, , , , , usedaccount
                     

                If StrTempAccountCode = "" And usedaccount = 1 Then
                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·œ«∆‰ ·„—œÊœ«  «·„‘ —Ì«   ", vbCritical
                    GoTo ErrTrap
                ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
                ElseIf usedaccount = 0 Then
                    StrTempAccountCode = Account_Code_dynamic '„—œÊœ«  «·„‘ —Ì« 
                End If

            Else
                StrTempAccountCode = Account_Code_dynamic  '„—œÊœ«  «·„‘ —Ì« 
            End If
            
            
                StrTempDes = "„—œÊœ«  „‘ —Ì«     —ﬁ„ " & Me.TxtNoteSerial1.text
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , SngTemp2, DcCurrency.text, val(txt_Currency_rate.text), , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
'''/////////////////////
If val(TxtValueAdded.text) > 0 Then
    Dim AccountVATCreit As String
 GetValueAddedAccount XPDtbBill.value, , AccountVATCreit, 1, 5
            LngDevNO = LngDevNO + 1
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = StrTempDes & " «·ﬁÌ„… «·„÷«›…  "
                Else
                    StrTempDes = StrTempDes & " VAT "
                End If
             If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, Round(val(TxtValueAdded.text) * val(txt_Currency_rate.text), 2), 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , Round(val(TxtValueAdded.text), 2), DcCurrency.text, val(txt_Currency_rate.text), , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
 End If
 ''//////////////
            ElseIf detect_inventory_work_type = 3 Then
                Dim groupAccount As String
             
                Dim line_value As Single

                With FG

                    For i = 1 To FG.rows - 1

                        If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                            ' groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 3)
                            groupAccount = get_item_group_account_in_branch(FG.TextMatrix(i, FG.ColIndex("Code")), val(my_branch), 5)

                            If groupAccount = "Error" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«» „—œÊœ«  «·„‘ —Ì«  ·„Ã„Ê⁄ …"
                                Else
                                    MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                                End If

                                GoTo ErrTrap
                            End If

                            line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count")) * val(txt_Currency_rate.text)
                            SngTemp2 = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count"))
                            StrTempDes = "  „—œÊœ«  «·„‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.text
                            LngDevNO = LngDevNO + 1

                            If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , SngTemp2, DcCurrency.text, val(txt_Currency_rate.text), , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                                GoTo ErrTrap
                            End If
    
                        End If

                    Next i

                End With

            End If
        End If

        '    StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
        '    StrTempDes = "„— Ã⁄ „‘ —Ì«  —ﬁ„ " & Me.TxtTransSerial.text
        '    SngTemp = Val(Me.LblTotalAll)
        '    LngDevNO = LngDevNO + 1
        '    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
        '        1, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        '        GoTo ErrTrap
        '    End If
        If Not CreateIssueVoucher Then GoTo ErrTrap
        SaveItemsData
        SaveValueAdded
       updateNotesValueAndNobytext CDbl(general_noteid)

        '----------------
        Cn.CommitTrans
        BeginTrans = False
        If IsSaveWithOutMsg Then Exit Sub
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
    
        '  If SystemOptions.autoIssueVoucher = True Then
       
        'End If
        'If SystemOptions.usertype = UserAdminAll Then
  
        'End If
        '----------------------------------------------------------------
        '·√‰‰« ﬁ„‰« »≈÷«›… Õ—ﬂ… „‰ ‰Ê⁄ „Œ ·›…
        StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=5"
         
        If SystemOptions.usertype <> UserAdminAll Then
            'StrSQL = StrSQL & "  AND   BranchId=" & branch_id
        End If

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        Me.Retrive val(Me.XPTxtBillID.text)
        '----------------------------------------------------------------
        CuurentLogdata

        Select Case Me.TxtModFlg.text

            Case "N"
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            
            Case "E"
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        End Select

        TxtModFlg.text = "R"
    End If

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    Screen.MousePointer = vbDefault

    'Resume
    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Msg = Msg & Err.Description & CHR(13)
    Msg = Msg & Err.Number & CHR(13)
    Msg = Msg & Err.Source
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Function CheckInvData() As Boolean
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim Msg As String

    CheckInvData = False

    If Me.TxtInvSerial.text <> "" And cmdReSave.Visible = False Then
        StrSQL = "SELECT * From Transactions where  Transactions.Transaction_Type=1 or Transactions.Transaction_Type=22"
        StrSQL = StrSQL + " AND  NoteSerial1='" & Trim(Me.TxtInvSerial.text) & "'"
        '    strsql = strsql + " AND Transactions.Transaction_Type=1 or Transactions.Transaction_Type=22"
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If rs.BOF Or rs.EOF Then
            Msg = "·« ÊÃœ ›« Ê—… ‘—«¡ »Â–« «·—ﬁ„..!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CheckInvData = False
            rs.Close
            Set rs = Nothing
            Exit Function
        ElseIf rs("CusID").value <> Me.DBCboClientName.BoundText Then
            Msg = "«·›« Ê—… —ﬁ„ " & Trim(Me.TxtInvSerial.text)
            Msg = Msg & CHR(13) & "·Ì”  „⁄ «·„Ê—œ" & Me.DBCboClientName.text
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CheckInvData = False
            rs.Close
            Set rs = Nothing
            Exit Function
        Else
            Me.TxtInvID.text = rs("Transaction_ID").value
        End If
   

    rs.Close
    Set rs = Nothing
     End If
    CheckInvData = True
End Function

Private Function CheckRetrunInv() As Boolean
    Dim StrSQL  As String
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    '----------------------------
    StrSQL = "Select * From Transaction_Details Where Transaction_ID=" & val(Me.TxtInvID.text) & ""
    StrSQL = StrSQL + " Order  By ID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    CheckRetrunInv = False

    If Not (rs.BOF Or rs.EOF) Then

        With Me.FG

            For i = .FixedRows To .rows - 1

                If .TextMatrix(i, .ColIndex("Name")) <> "" Then
                    If rs.filter <> adFilterNone Then
                        rs.filter = adFilterNone
                    End If

                    rs.MoveFirst
                    rs.filter = "Item_ID=" & val(.TextMatrix(i, .ColIndex("Name")))

                    If rs.BOF Or rs.EOF Then
                        Msg = "«·’‰› : " & .cell(flexcpTextDisplay, i, .ColIndex("Name"))
                        Msg = Msg & CHR(13) & "Ê«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„ : " & i
                        Msg = Msg & CHR(13) & "·„ Ìﬂ‰ „ÊÃÊœ ›Ï «·›« Ê—… —ﬁ„ : " & Me.TxtInvSerial.text
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        CheckRetrunInv = False
                        rs.Close
                        Set rs = Nothing
                        Exit Function
                    ElseIf FG.cell(flexcpChecked, i, .ColIndex("HaveSerial")) = flexChecked Then
                        rs.Find "ItemSerial='" & Trim(.TextMatrix(i, .ColIndex("Serial"))) & "'", , adSearchForward, 1

                        If rs.BOF Or rs.EOF Then
                            Msg = "«·ﬁÿ⁄… –«  «·”Ì—Ì«·:  " & Trim(.TextMatrix(i, .ColIndex("Serial")))
                            Msg = Msg & CHR(13) & "„‰ «·’‰› : " & .cell(flexcpTextDisplay, i, .ColIndex("Name"))
                            Msg = Msg & CHR(13) & "Ê«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„  : " & i
                            Msg = Msg & CHR(13) & "·„ Ìﬂ‰ „ÊÃÊœ ›Ï «·›« Ê—… —ﬁ„  : " & Me.TxtInvSerial.text
                            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                            CheckRetrunInv = False
                            rs.Close
                            Set rs = Nothing
                            Exit Function
                        End If
                    End If
                End If

            Next i

        End With

    End If

    '----------------------------

    '----------------------------
    CheckRetrunInv = True
End Function

Public Function CheckTotalQty(LngRow As Long) As Double
    Dim i As Integer
    Dim LngFoundGridRow As Long
    Dim Dbl_OutQtyBySmallUnit As Double

    If LngRow = 0 Then Exit Function

    With Me.FG
        LngFoundGridRow = 0

        For i = 1 To .rows - 1
            LngFoundGridRow = .FindRow(.TextMatrix(LngRow, .ColIndex("Code")), i, .ColIndex("Code"), True, True)

            If LngFoundGridRow > 0 Then
                i = LngFoundGridRow
                Dbl_OutQtyBySmallUnit = Dbl_OutQtyBySmallUnit + val(.TextMatrix(LngFoundGridRow, .ColIndex("Count")))
            ElseIf LngFoundGridRow = -1 Then        'did not found the item entered before
                Exit For
            End If

        Next i

    End With

    CheckTotalQty = Dbl_OutQtyBySmallUnit

End Function

Private Sub XPChkPayType_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If XPChkPayType(0).value = Checked Then
                If Me.TxtModFlg.text = "N" Then
                    XPTxtValue(0).text = ""
                    XPTxtSerial(0).text = ""
                End If

                If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
                    XPTxtValue(0).Enabled = True
                    XPTxtValue(0).locked = False
                End If

            Else
                XPTxtValue(0).Enabled = False
                XPTxtValue(0).text = ""
            End If

        Case 1

            If XPChkPayType(1).value = Checked Then
                If Me.TxtModFlg.text = "N" Then
                    XPTxtValue(1).text = ""
                    XPTxtSerial(1).text = ""
                    DtpDelayDate.Enabled = True
                End If

                If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
                    XPTxtValue(1).Enabled = True
                    XPTxtValue(1).locked = False
                    DtpDelayDate.Enabled = True
                Else
                    DtpDelayDate.Enabled = False
                End If

            Else
                XPTxtValue(1).Enabled = False
                XPTxtValue(1).text = ""
            End If

        Case 2

            If XPChkPayType(2).value = Checked Then
                If Me.TxtModFlg.text = "N" Then
                    XPTxtValue(2).text = ""
                    XPTxtChqueNum.text = ""
                    XPDTPDueDate.value = Date
                    DCboBankName.BoundText = ""
                End If

                If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
                    XPTxtValue(2).Enabled = True
                    XPTxtChqueNum.Enabled = True
                    XPDTPDueDate.Enabled = True
                    XPTxtValue(2).locked = False
                    XPTxtChqueNum.locked = False
                    DCboBankName.locked = False
                    DCboBankName.Enabled = True
                End If

            Else
                XPTxtValue(2).text = ""
                XPTxtValue(2).Enabled = False
                XPTxtChqueNum.Enabled = False
                XPDTPDueDate.Enabled = False
                DCboBankName.locked = True
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub PrintReport()
    On Error GoTo ErrTrap
    Dim ShowType As Boolean
    ShowType = GetSetting(StrAppRegPath, "View_Type", "ReportType", True)

    If ShowType = True Then
        If XPTxtBillID.text <> "" Then
            Set ReturnReport = New ClsReturnBackReport
            ReturnReport.ShowReturnBack XPTxtBillID.text, False, LblTotal.Caption
        End If

    Else

        If XPTxtBillID.text <> "" Then
            Set ReturnReport = New ClsReturnBackReport
            ReturnReport.ShowReturnBackShort XPTxtBillID.text
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap
    If Trim(Me.TxtModFlg.text) = "" Then Exit Sub
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

Private Sub XPDtbBill_Change()

    If Trim(TxtNoteSerial1.text) <> "" Then
        oldtxtNoteSerial1.text = TxtNoteSerial1.text
    End If

    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
    CurrentVoucherNo = ""
    TxtNoteSerial1V = ""
    DateChanged = True
End Sub

Private Sub XPTxtSum_Change()

    If CboPayMentType.ListIndex = 0 Then
        XPChkPayType(0).value = Checked
        XPTxtValue(0).text = XPTxtSum.text
    End If

    Me.LblTotal.Caption = Me.XPTxtSum.text
End Sub

Private Sub CboPayMentType_Change()
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If CboPayMentType.ListIndex = 0 Then
            XPChkPayType(0).Enabled = False
            XPChkPayType(1).Enabled = False
            XPChkPayType(2).Enabled = False
            XPChkPayType(0).value = Checked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            XPTxtValue(0).text = XPTxtSum.text
            'DcboBox.BoundText = ""
            DcboBox.Enabled = True
        Else
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            XPChkPayType(0).value = Unchecked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            XPTxtValue(0).text = ""
            DcboBox.BoundText = ""
            DcboBox.Enabled = False
            
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub DBCboClientName_Change()
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    TxtVATNO.text = GetCustomerVAT(val(Me.DBCboClientName.BoundText))
        If DBCboClientName.BoundText <> "" Then
            If DBCboClientName.BoundText = 1 Or DBCboClientName.BoundText = 2 Then
                CboPayMentType.locked = True
                CboPayMentType.ListIndex = 0
            Else
                CboPayMentType.locked = False
            End If
            
                     If SystemOptions.DefaultIsCreditPurchaseRet = False Then
            CboPayMentType.ListIndex = 0
   Else
         CboPayMentType.ListIndex = 1
         CboPayMentType.locked = False
    End If
    
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub DBCboClientName_Click(Area As Integer)
    DBCboClientName_Change
End Sub
