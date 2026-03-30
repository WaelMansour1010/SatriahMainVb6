VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmOutProductionOrder 
   Caption         =   "”‰œ ’—› „Ê«œ Œ«„ ··«‰ «Ã"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12735
   HelpContextID   =   160
   Icon            =   "FrmOutProductionOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   12735
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8160
      Left            =   0
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   0
      Width           =   12735
      _cx             =   22463
      _cy             =   14393
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
      AutoSizeChildren=   8
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
      GridRows        =   5
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmOutProductionOrder.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1800
         Index           =   0
         Left            =   15
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   645
         Width           =   12705
         _cx             =   22410
         _cy             =   3175
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
         Begin VB.CheckBox chkIgnorDetails 
            Alignment       =   1  'Right Justify
            Caption         =   " Ã«Â· «· ›«’Ì·"
            Height          =   270
            Left            =   3060
            RightToLeft     =   -1  'True
            TabIndex        =   198
            Top             =   45
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.TextBox txtPassword 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   4470
            PasswordChar    =   "*"
            TabIndex        =   197
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox chkIsBranch 
            Caption         =   "»«·›—⁄"
            Height          =   225
            Left            =   5250
            TabIndex        =   196
            Top             =   390
            Width           =   945
         End
         Begin VB.CommandButton cmdReSave 
            Caption         =   "÷»ÿ «·Õ—ﬂ« "
            Height          =   285
            Left            =   7110
            TabIndex        =   193
            Top             =   -15
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.TextBox txtProductionOrderID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   192
            Top             =   0
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.TextBox txtMixID 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   191
            Top             =   0
            Visible         =   0   'False
            Width           =   1650
         End
         Begin VB.TextBox txtMIxCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9870
            RightToLeft     =   -1  'True
            TabIndex        =   190
            Top             =   1080
            Width           =   1650
         End
         Begin VB.TextBox txtRemark 
            Alignment       =   1  'Right Justify
            Height          =   645
            Left            =   6195
            MaxLength       =   500
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   185
            Top             =   390
            Width           =   2640
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  ﬁÌœ «·”‰œ"
            Height          =   615
            Left            =   -30
            RightToLeft     =   -1  'True
            TabIndex        =   181
            Top             =   -60
            Width           =   3075
            Begin VB.TextBox TxtNoteSerial 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   183
               Top             =   240
               Width           =   1185
            End
            Begin ImpulseButton.ISButton Cmd 
               CausesValidation=   0   'False
               Height          =   375
               Index           =   10
               Left            =   120
               TabIndex        =   182
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   661
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
         End
         Begin VB.TextBox TxtWorkOrderNO 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8955
            RightToLeft     =   -1  'True
            TabIndex        =   178
            Top             =   780
            Width           =   2565
         End
         Begin VB.TextBox Txtnots2 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   148
            Top             =   360
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox TXTNoteID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   145
            Top             =   960
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9915
            RightToLeft     =   -1  'True
            TabIndex        =   144
            Top             =   30
            Width           =   1575
         End
         Begin ALLButtonS.ALLButton CmdConvert 
            Height          =   375
            Left            =   0
            TabIndex        =   143
            Top             =   960
            Visible         =   0   'False
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   " ÕÊÌ· «·Ï ›« Ê—…"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmOutProductionOrder.frx":03F0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   2025
            RightToLeft     =   -1  'True
            TabIndex        =   141
            Top             =   1080
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   1065
            RightToLeft     =   -1  'True
            TabIndex        =   140
            Top             =   1440
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.TextBox TxtCashCustomerName 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   -1050
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   1320
            Visible         =   0   'False
            Width           =   5475
         End
         Begin VB.TextBox TxtEmployeeID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   1500
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox TxtStoreID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10830
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   1410
            Width           =   705
         End
         Begin VB.TextBox TxtCusID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   1125
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.ComboBox CboSaleType 
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1050
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   10935
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   -300
            Visible         =   0   'False
            Width           =   2610
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   735
            Index           =   8
            Left            =   4815
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   1920
            Visible         =   0   'False
            Width           =   3390
            _cx             =   5980
            _cy             =   1296
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
            Begin ImpulseButton.ISButton CmdInvProfit 
               Height          =   390
               Left            =   90
               TabIndex        =   48
               Top             =   165
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   688
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "..."
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
               ButtonImage     =   "FrmOutProductionOrder.frx":040C
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰”»… «·—»Õ"
               ForeColor       =   &H00C00000&
               Height          =   255
               Index           =   23
               Left            =   3465
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   420
               Width           =   1785
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·—»Õ"
               ForeColor       =   &H00C00000&
               Height          =   255
               Index           =   22
               Left            =   3465
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   150
               Width           =   1785
            End
            Begin VB.Label lblInvPrecent 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   1500
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   390
               Width           =   2220
            End
            Begin VB.Label LblInvProfit 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   240
               Left            =   1500
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   135
               Width           =   2220
            End
         End
         Begin VB.ComboBox XPCboDiscountType 
            Height          =   315
            Left            =   2025
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1320
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.TextBox XPTxtDiscountVal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   255
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   1200
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.ComboBox CboPayMentType 
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1230
            Visible         =   0   'False
            Width           =   3135
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   240
            TabIndex        =   3
            Top             =   960
            Visible         =   0   'False
            Width           =   4770
            _ExtentX        =   8414
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            BoundColumn     =   ""
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   7800
            TabIndex        =   6
            Top             =   1410
            Width           =   3030
            _ExtentX        =   5345
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   345
            Left            =   9765
            TabIndex        =   1
            Top             =   390
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   609
            _Version        =   393216
            Format          =   278659073
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton XPBtnNewClients 
            Height          =   360
            Left            =   4905
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   750
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   635
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
            ButtonImage     =   "FrmOutProductionOrder.frx":07A6
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   315
            Left            =   2760
            TabIndex        =   8
            Top             =   1485
            Visible         =   0   'False
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton CmdCash 
            Height          =   270
            Index           =   0
            Left            =   6555
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   1140
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   476
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
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
            BackStyle       =   0
            ButtonImage     =   "FrmOutProductionOrder.frx":0B40
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdCash 
            Height          =   270
            Index           =   1
            Left            =   6135
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   1140
            Visible         =   0   'False
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   476
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
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
            BackStyle       =   0
            ButtonImage     =   "FrmOutProductionOrder.frx":0EDA
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   8250
            TabIndex        =   146
            Top             =   0
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCDocTypes 
            Height          =   315
            Left            =   120
            TabIndex        =   186
            Top             =   600
            Width           =   3420
            _ExtentX        =   6033
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker txtFromDateReSave 
            Height          =   315
            Left            =   5820
            TabIndex        =   194
            Top             =   15
            Visible         =   0   'False
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            Format          =   278724609
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker txtToDateReSave 
            Height          =   315
            Left            =   4515
            TabIndex        =   195
            Top             =   30
            Visible         =   0   'False
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            Format          =   278724609
            CurrentDate     =   38784
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·„Ìﬂ”"
            Height          =   240
            Index           =   57
            Left            =   11190
            RightToLeft     =   -1  'True
            TabIndex        =   189
            Top             =   1140
            Width           =   1440
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·”‰œ"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3630
            TabIndex        =   187
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„·«ÕŸ« "
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   8910
            TabIndex        =   184
            Top             =   480
            Width           =   750
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»‰«¡ ⁄·Ï  ›« Ê—Â —ﬁ„"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2910
            TabIndex        =   150
            Top             =   360
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   9225
            TabIndex        =   147
            Top             =   0
            Width           =   600
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   55
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   142
            Top             =   1080
            Visible         =   0   'False
            Width           =   210
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì· «·‰ﬁœÏ"
            Height          =   300
            Index           =   33
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   1080
            Visible         =   0   'False
            Width           =   1710
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ √„— «·«‰ «Ã"
            Height          =   240
            Index           =   32
            Left            =   11190
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   780
            Width           =   1440
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„‰œÊ»"
            Height          =   255
            Index           =   25
            Left            =   5910
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   1515
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Œ’„"
            Height          =   315
            Index           =   10
            Left            =   3420
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   750
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   315
            Index           =   9
            Left            =   3270
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   870
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„…"
            Height          =   330
            Index           =   8
            Left            =   1500
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   1200
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   255
            Index           =   24
            Left            =   11400
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   1470
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì·"
            Height          =   300
            Index           =   7
            Left            =   4650
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   1140
            Visible         =   0   'False
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·«–‰"
            Height          =   285
            Index           =   6
            Left            =   10830
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   390
            Width           =   1785
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·«–‰"
            Height          =   255
            Index           =   5
            Left            =   11280
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   75
            Width           =   1335
         End
      End
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   4680
         Left            =   15
         TabIndex        =   22
         Top             =   2460
         Width           =   12705
         _cx             =   22410
         _cy             =   8255
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
         ForeColor       =   0
         FrontTabColor   =   14871017
         BackTabColor    =   12648447
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   16711680
         Caption         =   "«·√’‰«›|»Ì«‰«  ›« Ê—… «·„»Ì⁄« |≈” ﬁÿ«⁄«  ⁄·Ï «·›« Ê—…|ﬁÌÊœ «·ÌÊ„Ì…"
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
         Picture(0)      =   "FrmOutProductionOrder.frx":1274
         Picture(1)      =   "FrmOutProductionOrder.frx":160E
         Flags(1)        =   2
         Picture(2)      =   "FrmOutProductionOrder.frx":19A8
         Flags(2)        =   2
         Flags(3)        =   3
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4215
            Index           =   19
            Left            =   13950
            TabIndex        =   112
            TabStop         =   0   'False
            Top             =   45
            Width           =   12615
            _cx             =   22251
            _cy             =   7435
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
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4215
            Index           =   15
            Left            =   13650
            TabIndex        =   105
            TabStop         =   0   'False
            Top             =   45
            Width           =   12615
            _cx             =   22251
            _cy             =   7435
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
            GridRows        =   7
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmOutProductionOrder.frx":1D42
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1035
               Index           =   18
               Left            =   15
               TabIndex        =   121
               TabStop         =   0   'False
               Top             =   1245
               Width           =   12585
               _cx             =   22199
               _cy             =   1826
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
               Begin VB.TextBox TxtTaxServiceValue 
                  Alignment       =   1  'Right Justify
                  Height          =   0
                  Left            =   45
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   127
                  Top             =   30
                  Width           =   0
               End
               Begin VB.CheckBox ChkTaxSerivce 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì… Œœ„…"
                  Height          =   0
                  Left            =   60
                  RightToLeft     =   -1  'True
                  TabIndex        =   122
                  Top             =   45
                  Width           =   0
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   90
                  Index           =   54
                  Left            =   30
                  RightToLeft     =   -1  'True
                  TabIndex        =   139
                  Top             =   30
                  Width           =   15
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
                  Height          =   90
                  Index           =   47
                  Left            =   45
                  RightToLeft     =   -1  'True
                  TabIndex        =   132
                  Top             =   30
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   120
                  Index           =   43
                  Left            =   45
                  RightToLeft     =   -1  'True
                  TabIndex        =   128
                  Top             =   30
                  Width           =   15
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   705
               Index           =   17
               Left            =   15
               TabIndex        =   119
               TabStop         =   0   'False
               Top             =   525
               Width           =   12585
               _cx             =   22199
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
               Begin VB.TextBox TxtTaxStampValue 
                  Alignment       =   1  'Right Justify
                  Height          =   0
                  Left            =   45
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   126
                  Top             =   315
                  Width           =   0
               End
               Begin VB.CheckBox ChkTaxStamp 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "œ„€…"
                  Height          =   0
                  Left            =   60
                  RightToLeft     =   -1  'True
                  TabIndex        =   120
                  Top             =   465
                  Width           =   0
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   1335
                  Index           =   53
                  Left            =   30
                  RightToLeft     =   -1  'True
                  TabIndex        =   138
                  Top             =   465
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "$"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1335
                  Index           =   48
                  Left            =   45
                  RightToLeft     =   -1  'True
                  TabIndex        =   133
                  Top             =   465
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   1500
                  Index           =   41
                  Left            =   45
                  RightToLeft     =   -1  'True
                  TabIndex        =   124
                  Top             =   465
                  Width           =   15
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   240
               Index           =   16
               Left            =   15
               TabIndex        =   117
               TabStop         =   0   'False
               Top             =   525
               Width           =   12585
               _cx             =   22199
               _cy             =   423
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
               Begin VB.TextBox TxtTaxAddValue 
                  Alignment       =   1  'Right Justify
                  Height          =   0
                  Left            =   45
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   125
                  Top             =   15
                  Width           =   0
               End
               Begin VB.CheckBox ChkTaxAdd 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì»… Œ’„ Ê≈÷«›… (√—»«Õ  Ã«—Ì…)"
                  Height          =   90
                  Left            =   45
                  RightToLeft     =   -1  'True
                  TabIndex        =   118
                  Top             =   0
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   45
                  Index           =   52
                  Left            =   30
                  RightToLeft     =   -1  'True
                  TabIndex        =   137
                  Top             =   15
                  Width           =   15
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
                  Height          =   45
                  Index           =   46
                  Left            =   45
                  RightToLeft     =   -1  'True
                  TabIndex        =   131
                  Top             =   15
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   60
                  Index           =   39
                  Left            =   45
                  RightToLeft     =   -1  'True
                  TabIndex        =   123
                  Top             =   15
                  Width           =   15
               End
            End
            Begin VB.TextBox TxtBillComment 
               Alignment       =   1  'Right Justify
               Height          =   1035
               Left            =   15
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   106
               Top             =   1245
               Width           =   12585
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   495
               Index           =   4
               Left            =   15
               TabIndex        =   113
               TabStop         =   0   'False
               Top             =   15
               Width           =   12585
               _cx             =   22199
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
               Begin VB.CheckBox XPChkTAX 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì»… «·„»Ì⁄« "
                  Height          =   225
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   115
                  Top             =   120
                  Width           =   30
               End
               Begin VB.TextBox XPTxtTaxValue 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   90
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   114
                  Top             =   75
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   255
                  Index           =   51
                  Left            =   60
                  RightToLeft     =   -1  'True
                  TabIndex        =   136
                  Top             =   90
                  Width           =   30
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
                  Height          =   255
                  Index           =   45
                  Left            =   90
                  RightToLeft     =   -1  'True
                  TabIndex        =   130
                  Top             =   90
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   165
                  Index           =   4
                  Left            =   105
                  RightToLeft     =   -1  'True
                  TabIndex        =   116
                  Top             =   135
                  Width           =   15
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈÷«›… √Ì… „·«ÕŸ«  ⁄·Ï «·›« Ê—…"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   1035
               Index           =   44
               Left            =   15
               RightToLeft     =   -1  'True
               TabIndex        =   129
               Top             =   1245
               Visible         =   0   'False
               Width           =   12585
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4215
            Index           =   7
            Left            =   45
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   45
            Width           =   12615
            _cx             =   22251
            _cy             =   7435
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
            GridRows        =   3
            GridCols        =   4
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmOutProductionOrder.frx":1DB7
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   690
               Index           =   2
               Left            =   30
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   30
               Width           =   12555
               _cx             =   22146
               _cy             =   1217
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
               Begin VB.ComboBox CboItemCase 
                  Height          =   315
                  Left            =   5010
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   16
                  Top             =   315
                  Width           =   1110
               End
               Begin VB.TextBox TxtQuantity 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   360
                  Left            =   2265
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   18
                  Top             =   315
                  Width           =   960
               End
               Begin VB.TextBox TxtSerial 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  Height          =   360
                  Left            =   3690
                  MaxLength       =   20
                  RightToLeft     =   -1  'True
                  TabIndex        =   17
                  Top             =   315
                  Width           =   1335
               End
               Begin VB.TextBox TxtPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   360
                  Left            =   1140
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   19
                  Top             =   315
                  Width           =   1110
               End
               Begin MSDataListLib.DataCombo DCboItemsName 
                  Height          =   315
                  Left            =   6180
                  TabIndex        =   15
                  Top             =   315
                  Width           =   4125
                  _ExtentX        =   7276
                  _ExtentY        =   556
                  _Version        =   393216
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
               Begin MSDataListLib.DataCombo DCboItemsCode 
                  Height          =   315
                  Left            =   10335
                  TabIndex        =   14
                  Top             =   315
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdAdd 
                  Height          =   360
                  Left            =   600
                  TabIndex        =   20
                  Top             =   315
                  Width           =   375
                  _ExtentX        =   661
                  _ExtentY        =   635
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
                  ButtonImage     =   "FrmOutProductionOrder.frx":1E29
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
               Begin ImpulseButton.ISButton CmdSearch 
                  Height          =   285
                  Left            =   3270
                  TabIndex        =   56
                  Top             =   330
                  Width           =   435
                  _ExtentX        =   767
                  _ExtentY        =   503
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "..."
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
                  ButtonImage     =   "FrmOutProductionOrder.frx":21C3
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·’‰›"
                  Height          =   285
                  Index           =   31
                  Left            =   10695
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   45
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈”„ «·’‰›"
                  Height          =   285
                  Index           =   30
                  Left            =   7905
                  RightToLeft     =   -1  'True
                  TabIndex        =   44
                  Top             =   15
                  Width           =   1320
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·… «·’‰›"
                  Height          =   285
                  Index           =   29
                  Left            =   5145
                  RightToLeft     =   -1  'True
                  TabIndex        =   43
                  Top             =   15
                  Width           =   1245
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ì—Ì«·"
                  Height          =   285
                  Index           =   28
                  Left            =   3900
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   15
                  Width           =   810
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂ„Ì…"
                  Height          =   285
                  Index           =   27
                  Left            =   2550
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   45
                  Width           =   720
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· ﬂ·›…"
                  Height          =   285
                  Index           =   26
                  Left            =   1500
                  RightToLeft     =   -1  'True
                  TabIndex        =   40
                  Top             =   15
                  Width           =   630
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid FG 
               Height          =   3075
               Left            =   30
               TabIndex        =   13
               Top             =   735
               Width           =   12555
               _cx             =   22146
               _cy             =   5424
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
               Rows            =   2
               Cols            =   22
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmOutProductionOrder.frx":255D
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
            Begin MSComctlLib.Toolbar TBar 
               Height          =   630
               Left            =   495
               TabIndex        =   54
               Top             =   3825
               Width           =   11625
               _ExtentX        =   20505
               _ExtentY        =   1111
               ButtonWidth     =   609
               ButtonHeight    =   1005
               Appearance      =   1
               _Version        =   393216
            End
            Begin VB.Label LblItemsCount 
               Alignment       =   2  'Center
               BackColor       =   &H00404040&
               ForeColor       =   &H0000FFFF&
               Height          =   360
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   3825
               Width           =   450
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4215
            Index           =   5
            Left            =   13350
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   45
            Width           =   12615
            _cx             =   22251
            _cy             =   7435
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
            BackColor       =   255
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
            GridRows        =   3
            GridCols        =   4
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmOutProductionOrder.frx":28DA
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   2040
               Index           =   10
               Left            =   0
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   270
               Width           =   5940
               _cx             =   10478
               _cy             =   3598
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
               GridRows        =   4
               GridCols        =   4
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"FrmOutProductionOrder.frx":294A
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   60
                  Index           =   14
                  Left            =   15
                  TabIndex        =   98
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   20100
                  _cx             =   35454
                  _cy             =   106
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
                  Begin VB.CheckBox XPChkPayType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‘Ìﬂ« "
                     Height          =   270
                     Index           =   2
                     Left            =   8295
                     RightToLeft     =   -1  'True
                     TabIndex        =   99
                     Top             =   60
                     Width           =   1080
                  End
                  Begin ImpulseButton.ISButton CmdCheque 
                     Height          =   270
                     Left            =   2535
                     TabIndex        =   109
                     Top             =   60
                     Width           =   1500
                     _ExtentX        =   2646
                     _ExtentY        =   476
                     ButtonStyle     =   1
                     Caption         =   " ”ÃÌ· «·‘Ìﬂ« "
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
                     DrawFocusRectangle=   0   'False
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   270
                     Index           =   19
                     Left            =   6540
                     RightToLeft     =   -1  'True
                     TabIndex        =   111
                     Top             =   60
                     Width           =   600
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄œœ «·‘Ìﬂ« "
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
                     Height          =   270
                     Index           =   17
                     Left            =   7290
                     RightToLeft     =   -1  'True
                     TabIndex        =   110
                     Top             =   60
                     Width           =   930
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "≈Ã„«·Ï ﬁÌ„… «·‘Ìﬂ« "
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
                     Height          =   270
                     Index           =   16
                     Left            =   4920
                     RightToLeft     =   -1  'True
                     TabIndex        =   101
                     Top             =   60
                     Width           =   1590
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   270
                     Index           =   18
                     Left            =   4050
                     RightToLeft     =   -1  'True
                     TabIndex        =   100
                     Top             =   60
                     Width           =   855
                  End
               End
               Begin VSFlex8UCtl.VSFlexGrid FgCheques 
                  Height          =   3210
                  Left            =   3555
                  TabIndex        =   59
                  Top             =   90
                  Width           =   16560
                  _cx             =   29210
                  _cy             =   5662
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
                  Rows            =   50
                  Cols            =   8
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmOutProductionOrder.frx":29BE
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   2040
               Index           =   6
               Left            =   0
               TabIndex        =   66
               TabStop         =   0   'False
               Top             =   270
               Width           =   5940
               _cx             =   10478
               _cy             =   3598
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
               GridRows        =   3
               GridCols        =   4
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"FrmOutProductionOrder.frx":2AF2
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VSFlex8UCtl.VSFlexGrid FgInstallments 
                  Height          =   6750
                  Left            =   3735
                  TabIndex        =   67
                  Top             =   105
                  Width           =   16380
                  _cx             =   28892
                  _cy             =   11906
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
                  Rows            =   50
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmOutProductionOrder.frx":2B5E
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
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   60
                  Index           =   13
                  Left            =   15
                  TabIndex        =   68
                  TabStop         =   0   'False
                  Top             =   6795
                  Width           =   20100
                  _cx             =   35454
                  _cy             =   106
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
                     Caption         =   "«·ﬁÌ„… «·„»œ∆Ì…"
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
                     Height          =   225
                     Index           =   37
                     Left            =   255
                     RightToLeft     =   -1  'True
                     TabIndex        =   108
                     Top             =   60
                     Visible         =   0   'False
                     Width           =   990
                  End
                  Begin VB.Label LblStartValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   225
                     Left            =   30
                     RightToLeft     =   -1  'True
                     TabIndex        =   107
                     Top             =   60
                     Width           =   210
                  End
                  Begin VB.Label LblInstallSeprator 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     ForeColor       =   &H00FF0000&
                     Height          =   225
                     Left            =   1920
                     RightToLeft     =   -1  'True
                     TabIndex        =   104
                     Top             =   60
                     Width           =   225
                  End
                  Begin VB.Label LblPrecenValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   225
                     Left            =   7305
                     RightToLeft     =   -1  'True
                     TabIndex        =   103
                     Top             =   60
                     Width           =   270
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰”»… «·›«∆œ…"
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
                     Height          =   225
                     Index           =   35
                     Left            =   7590
                     RightToLeft     =   -1  'True
                     TabIndex        =   102
                     Top             =   60
                     Visible         =   0   'False
                     Width           =   420
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰Ê⁄ «·›«∆œ…"
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
                     Height          =   225
                     Index           =   34
                     Left            =   8670
                     RightToLeft     =   -1  'True
                     TabIndex        =   78
                     Top             =   60
                     Visible         =   0   'False
                     Width           =   720
                  End
                  Begin VB.Label LblPrecenType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   225
                     Left            =   8025
                     RightToLeft     =   -1  'True
                     TabIndex        =   77
                     Top             =   60
                     Width           =   630
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„»·€ «·ﬂ·Ï"
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
                     Height          =   225
                     Index           =   36
                     Left            =   6435
                     RightToLeft     =   -1  'True
                     TabIndex        =   76
                     Top             =   60
                     Visible         =   0   'False
                     Width           =   855
                  End
                  Begin VB.Label LblInstallTotal 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   225
                     Left            =   5850
                     RightToLeft     =   -1  'True
                     TabIndex        =   75
                     Top             =   60
                     Width           =   555
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄œœ «·√ﬁ”«ÿ"
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
                     Height          =   225
                     Index           =   38
                     Left            =   4935
                     RightToLeft     =   -1  'True
                     TabIndex        =   74
                     Top             =   60
                     Visible         =   0   'False
                     Width           =   900
                  End
                  Begin VB.Label LblInstallCount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   225
                     Left            =   4590
                     RightToLeft     =   -1  'True
                     TabIndex        =   73
                     Top             =   60
                     Width           =   330
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«Ê· ﬁ”ÿ"
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
                     Height          =   225
                     Index           =   40
                     Left            =   3900
                     RightToLeft     =   -1  'True
                     TabIndex        =   72
                     Top             =   60
                     Visible         =   0   'False
                     Width           =   660
                  End
                  Begin VB.Label LblFirstInstallDate 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   225
                     Left            =   3135
                     RightToLeft     =   -1  'True
                     TabIndex        =   71
                     Top             =   60
                     Width           =   750
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "› —… «· ﬁ”Ìÿ"
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
                     Height          =   225
                     Index           =   42
                     Left            =   2175
                     RightToLeft     =   -1  'True
                     TabIndex        =   70
                     Top             =   60
                     Visible         =   0   'False
                     Width           =   930
                  End
                  Begin VB.Label LblInstallmentType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     ForeColor       =   &H00FF0000&
                     Height          =   225
                     Left            =   1275
                     RightToLeft     =   -1  'True
                     TabIndex        =   69
                     Top             =   60
                     Width           =   630
                  End
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   75
                  Index           =   12
                  Left            =   15
                  TabIndex        =   79
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   20100
                  _cx             =   35454
                  _cy             =   132
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
                  Begin VB.CheckBox ChkInstall 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ”Ìÿ"
                     Height          =   345
                     Left            =   1155
                     RightToLeft     =   -1  'True
                     TabIndex        =   83
                     Top             =   15
                     Width           =   1080
                  End
                  Begin VB.TextBox XPTxtSerial 
                     Alignment       =   1  'Right Justify
                     Height          =   345
                     Index           =   1
                     Left            =   4995
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   82
                     Top             =   30
                     Width           =   990
                  End
                  Begin VB.TextBox XPTxtValue 
                     Alignment       =   1  'Right Justify
                     Height          =   345
                     Index           =   1
                     Left            =   6840
                     MaxLength       =   10
                     RightToLeft     =   -1  'True
                     TabIndex        =   81
                     Top             =   30
                     Width           =   840
                  End
                  Begin VB.CheckBox XPChkPayType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "¬Ã· "
                     Height          =   315
                     Index           =   1
                     Left            =   8430
                     RightToLeft     =   -1  'True
                     TabIndex        =   80
                     Top             =   30
                     Width           =   960
                  End
                  Begin ImpulseButton.ISButton CmdINSTALLMENT 
                     Height          =   420
                     Left            =   180
                     TabIndex        =   84
                     Top             =   -15
                     Width           =   1185
                     _ExtentX        =   2090
                     _ExtentY        =   741
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   "Õ”«» «·√ﬁ”«ÿ"
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
                     ButtonImage     =   "FrmOutProductionOrder.frx":2C2F
                     ColorButton     =   14871017
                     ColorHighlight  =   16777215
                     ColorHoverText  =   16711680
                     ColorShadow     =   4210752
                     ColorOutline    =   0
                     DrawFocusRectangle=   0   'False
                     ColorToggledHoverText=   16711680
                     ColorTextShadow =   4210752
                  End
                  Begin MSComCtl2.DTPicker DtpDelayDate 
                     Height          =   330
                     Left            =   2355
                     TabIndex        =   85
                     Top             =   30
                     Width           =   1290
                     _ExtentX        =   2275
                     _ExtentY        =   582
                     _Version        =   393216
                     Format          =   278724609
                     CurrentDate     =   38784
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
                     Height          =   285
                     Index           =   21
                     Left            =   3705
                     RightToLeft     =   -1  'True
                     TabIndex        =   88
                     Top             =   75
                     Width           =   1110
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁÌ„…"
                     Height          =   330
                     Index           =   15
                     Left            =   7695
                     RightToLeft     =   -1  'True
                     TabIndex        =   87
                     Top             =   75
                     Width           =   525
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„”·”·"
                     Height          =   315
                     Index           =   14
                     Left            =   6150
                     RightToLeft     =   -1  'True
                     TabIndex        =   86
                     Top             =   75
                     Width           =   525
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   270
               Index           =   11
               Left            =   0
               TabIndex        =   89
               TabStop         =   0   'False
               Top             =   0
               Width           =   5940
               _cx             =   10478
               _cy             =   476
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
               Begin VB.TextBox XPTxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Index           =   0
                  Left            =   7680
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   92
                  Top             =   60
                  Width           =   855
               End
               Begin VB.TextBox XPTxtSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Index           =   0
                  Left            =   5760
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   91
                  Top             =   60
                  Width           =   1035
               End
               Begin VB.CheckBox XPChkPayType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰ﬁœ«"
                  Height          =   345
                  Index           =   0
                  Left            =   9180
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   90
                  Width           =   1080
               End
               Begin MSDataListLib.DataCombo DcboBox 
                  Height          =   315
                  Left            =   2820
                  TabIndex        =   93
                  Top             =   105
                  Width           =   1935
                  _ExtentX        =   3413
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
                  Height          =   345
                  Index           =   20
                  Left            =   270
                  RightToLeft     =   -1  'True
                  TabIndex        =   97
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1095
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   345
                  Index           =   13
                  Left            =   8775
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   90
                  Width           =   450
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·”·"
                  Height          =   345
                  Index           =   12
                  Left            =   6795
                  RightToLeft     =   -1  'True
                  TabIndex        =   95
                  Top             =   90
                  Width           =   615
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·Œ“‰…"
                  Height          =   345
                  Index           =   11
                  Left            =   4710
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   90
                  Width           =   870
               End
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   615
         Index           =   9
         Left            =   15
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   15
         Width           =   12705
         _cx             =   22410
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
         Caption         =   "”‰œ ’—› „Ê«œ Œ«„ ··«‰ «Ã"
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
         Begin VB.TextBox oldtxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10680
            RightToLeft     =   -1  'True
            TabIndex        =   179
            Top             =   360
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   285
            Left            =   5415
            RightToLeft     =   -1  'True
            TabIndex        =   135
            Top             =   0
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   285
            Left            =   4905
            RightToLeft     =   -1  'True
            TabIndex        =   134
            Top             =   0
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   285
            Left            =   4335
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   0
            Visible         =   0   'False
            Width           =   510
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   2025
            TabIndex        =   36
            Top             =   30
            Width           =   810
            _ExtentX        =   1429
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
            ButtonImage     =   "FrmOutProductionOrder.frx":2FC9
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
            Left            =   1005
            TabIndex        =   37
            Top             =   30
            Width           =   975
            _ExtentX        =   1720
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
            ButtonImage     =   "FrmOutProductionOrder.frx":3363
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
            Left            =   2865
            TabIndex        =   38
            Top             =   30
            Width           =   900
            _ExtentX        =   1588
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
            ButtonImage     =   "FrmOutProductionOrder.frx":36FD
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
            Left            =   60
            TabIndex        =   39
            Top             =   30
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
            ButtonImage     =   "FrmOutProductionOrder.frx":3A97
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin ImpulseButton.ISButton CmdNotes 
            Height          =   345
            Left            =   7125
            TabIndex        =   60
            Top             =   120
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   3
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
            ButtonImage     =   "FrmOutProductionOrder.frx":3E31
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdRetruns 
            Height          =   345
            Left            =   5655
            TabIndex        =   61
            Top             =   120
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   3
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
            ButtonImage     =   "FrmOutProductionOrder.frx":41CB
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdInfo 
            Height          =   480
            Left            =   6600
            TabIndex        =   151
            Top             =   0
            Visible         =   0   'False
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   847
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
            ButtonImage     =   "FrmOutProductionOrder.frx":4765
            ButtonImageHover=   "FrmOutProductionOrder.frx":543F
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
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
            Index           =   56
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   180
            Top             =   0
            Width           =   4275
         End
         Begin VB.Label LblShortcutKeys 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "ÃœÌœ F12 Or Enter ,  ⁄œÌ· F11 , Õ›Ÿ F10 ,  —«Ã⁄ F9 ,Õ–› F8 ,»ÕÀ F7 "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   390
            Width           =   7485
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   435
         Index           =   3
         Left            =   15
         TabIndex        =   152
         TabStop         =   0   'False
         Top             =   7155
         Width           =   12705
         _cx             =   22410
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
         Begin VB.CommandButton sameCmd 
            Caption         =   "‰”Œ… „„«À·Â"
            Height          =   375
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   188
            Top             =   0
            Width           =   840
         End
         Begin VB.TextBox XPTxtSum 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Height          =   375
            Left            =   7155
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   153
            TabStop         =   0   'False
            Top             =   30
            Visible         =   0   'False
            Width           =   270
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   2370
            TabIndex        =   154
            Top             =   30
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·ﬂ„ÌÂ"
            Height          =   375
            Index           =   63
            Left            =   12015
            TabIndex        =   177
            Top             =   60
            Width           =   615
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
            Height          =   345
            Left            =   10890
            TabIndex        =   176
            Top             =   0
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·≈Ã„«·Ï"
            Height          =   285
            Index           =   3
            Left            =   10215
            RightToLeft     =   -1  'True
            TabIndex        =   165
            Top             =   75
            Width           =   645
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "/"
            Height          =   285
            Index           =   0
            Left            =   570
            RightToLeft     =   -1  'True
            TabIndex        =   164
            Top             =   75
            Width           =   225
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”Ã·"
            Height          =   285
            Index           =   2
            Left            =   1290
            RightToLeft     =   -1  'True
            TabIndex        =   163
            Top             =   75
            Width           =   825
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   285
            Left            =   930
            RightToLeft     =   -1  'True
            TabIndex        =   162
            Top             =   75
            Width           =   315
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   285
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   161
            Top             =   75
            Width           =   405
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„” Œœ„"
            Height          =   330
            Index           =   1
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   160
            Top             =   75
            Width           =   660
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
            Left            =   6225
            RightToLeft     =   -1  'True
            TabIndex        =   159
            Top             =   0
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«›Ì"
            Height          =   285
            Index           =   49
            Left            =   7230
            RightToLeft     =   -1  'True
            TabIndex        =   158
            Top             =   75
            Visible         =   0   'False
            Width           =   690
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
            Left            =   8490
            RightToLeft     =   -1  'True
            TabIndex        =   157
            Top             =   30
            Width           =   1575
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œ’Ê„« "
            Height          =   285
            Index           =   50
            Left            =   7800
            RightToLeft     =   -1  'True
            TabIndex        =   156
            Top             =   315
            Visible         =   0   'False
            Width           =   600
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
            Left            =   6615
            RightToLeft     =   -1  'True
            TabIndex        =   155
            Top             =   270
            Visible         =   0   'False
            Width           =   990
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   540
         Index           =   1
         Left            =   15
         TabIndex        =   166
         TabStop         =   0   'False
         Top             =   7605
         Width           =   12705
         _cx             =   22410
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
            Height          =   375
            Index           =   0
            Left            =   11355
            TabIndex        =   167
            Top             =   90
            Width           =   1290
            _ExtentX        =   2275
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
            ColorToggledText=   -2147483631
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   1
            Left            =   9900
            TabIndex        =   168
            Top             =   120
            Width           =   1350
            _ExtentX        =   2381
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
            Left            =   8490
            TabIndex        =   169
            Top             =   90
            Width           =   1320
            _ExtentX        =   2328
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
            Height          =   375
            Index           =   3
            Left            =   7185
            TabIndex        =   170
            Top             =   90
            Width           =   1215
            _ExtentX        =   2143
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
            Height          =   375
            Index           =   4
            Left            =   5580
            TabIndex        =   171
            Top             =   90
            Width           =   1455
            _ExtentX        =   2566
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
            Left            =   4245
            TabIndex        =   172
            Top             =   90
            Width           =   1230
            _ExtentX        =   2170
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   6
            Left            =   30
            TabIndex        =   173
            TabStop         =   0   'False
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
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
            Index           =   7
            Left            =   2820
            TabIndex        =   174
            Top             =   90
            Width           =   1290
            _ExtentX        =   2275
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
            Height          =   375
            Left            =   1395
            TabIndex        =   175
            Top             =   90
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3720
      TabIndex        =   149
      Top             =   960
      Width           =   1050
   End
End
Attribute VB_Name = "FrmOutProductionOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim NewGrid As New ClsGrid
Dim SaleReport As ClsSaleReport
Dim cSearchDcbo(4)   As clsDCboSearch
Dim Dcombos As ClsDataCombos

Public BolPrint As Boolean

Public WithEvents m_Menu1 As Menu
Attribute m_Menu1.VB_VarHelpID = -1
Dim WithEvents m_MenuRefesh As Menu
Attribute m_MenuRefesh.VB_VarHelpID = -1
Dim WithEvents m_MenuCusBalance As Menu
Attribute m_MenuCusBalance.VB_VarHelpID = -1
Dim WithEvents m_MenuViewList As Menu
Attribute m_MenuViewList.VB_VarHelpID = -1
Dim WithEvents m_MenuViewNotes As Menu
Attribute m_MenuViewNotes.VB_VarHelpID = -1
Dim WithEvents m_MenuScreenPremission As Menu
Attribute m_MenuScreenPremission.VB_VarHelpID = -1
Dim WithEvents StrCashCustomerPhone As TextBox
Attribute StrCashCustomerPhone.VB_VarHelpID = -1
Dim WithEvents StrCashCustomerMobile As TextBox
Attribute StrCashCustomerMobile.VB_VarHelpID = -1
Dim WithEvents StrCashCustomerAddress As TextBox
Attribute StrCashCustomerAddress.VB_VarHelpID = -1
Dim WithEvents m_FrmSearch As Form
Attribute m_FrmSearch.VB_VarHelpID = -1
Dim general_noteid As Long
Dim Account_Code_dynamic As String
Dim mIsFinishSave As Boolean
Dim IsSaveWithOutMsg As Boolean
Dim mIsStart As Boolean

Private Sub txtPassword_Change()
    If Trim(txtPassword) = "Alex2025" Then
        cmdReSave.Visible = True
        txtFromDateReSave.Visible = True
        txtToDateReSave.Visible = True
        chkIsBranch.Visible = True
        txtFromDateReSave.value = Date
        txtToDateReSave.value = Date
        chkIgnorDetails.Visible = True
        chkIgnorDetails.value = 1
    Else
        cmdReSave.Visible = False
        txtFromDateReSave.Visible = False
        txtToDateReSave.Visible = False
        chkIsBranch.Visible = False
        chkIgnorDetails.Visible = False
        
    End If

End Sub

Private Sub cmdReSave_Click()
    Dim s         As String
    Dim rsDummy   As ADODB.Recordset
    Dim mBranchID As Integer
    mBranchID = 0
    If chkIsBranch.value = vbChecked Then
        mBranchID = val(dcBranch.BoundText)
        
    End If
 
    Set rsDummy = New ADODB.Recordset
    s = " SELECT * FROM Transactions WHERE Transaction_Type = 27"
    s = s & "   and ( Transaction_Date >= " & SQLDate(txtFromDateReSave.value, True) & " and "
    s = s & "   Transaction_Date <=   " & SQLDate(txtToDateReSave.value, True) & " )"
    If mBranchID <> 0 Then
        s = s & "  and BranchID =   " & mBranchID
    End If
        
    s = s & "  and Transaction_ID in "
    s = s & "  (SELECT        TT.Transaction_ID"
    s = s & "   FROM            dbo.Transactions TT INNER JOIN"
    s = s & "                 dbo.Transaction_Details ON TT.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
    s = s & " Where (TT.Transaction_Type = 27) )"
    'Transaction_Details.Price > 3000) "
     
    s = s & " ORDER BY  Transaction_Date Desc "
    ', BranchId, Transaction_ID"
    rs.Close
     
    rs.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    '    If rs.RecordCount > 0 Then
    '    rs.MoveLast
    '    End If
    XPBtnMove_Click (2)
    Dim i As Double
    For i = 1 To rs.RecordCount
        '        On Error GoTo NextRow
     
        Cmd_Click (1)
        NewGrid.updateProfit
        DoEvents
        DoEvents
        DoEvents
        SaveData True, True
         
        XPBtnMove_Click (0)
        'rs.MoveNext
        
    Next i
 
    MsgBox " „ «·Õ›Ÿ"
End Sub
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
        



   FG.TextMatrix(Num, FG.ColIndex("ColorID")) = 1
    FG.TextMatrix(Num, FG.ColIndex("itemsize")) = 1
    FG.TextMatrix(Num, FG.ColIndex("ClassId")) = 1
    
   
         
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
            XPTxtValue(1).text = ""
        Else
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            XPChkPayType(0).value = Unchecked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            XPTxtValue(0).text = ""
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub ChkInstall_Click()

    If ChkInstall.value = vbChecked Then
        Me.CmdINSTALLMENT.Enabled = True
    Else
        Me.CmdINSTALLMENT.Enabled = False

        With Me.FgInstallments
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            LblPrecenType.Caption = ""
            LblPrecenValue.Caption = ""
            LblInstallTotal.Caption = ""
            LblInstallCount.Caption = ""
            LblFirstInstallDate.Caption = ""
            LblInstallmentType.Caption = ""
        End With

    End If

End Sub

Private Sub ChkTaxAdd_Click()

    If ChkTaxAdd.value = Checked Then
        TxtTaxAddValue.Enabled = True
        lbl(39).Enabled = True
        lbl(46).Enabled = True
    Else
        TxtTaxAddValue.text = ""
        TxtTaxAddValue.Enabled = False
        lbl(39).Enabled = False
        lbl(46).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChkTaxSerivce_Click()
    On Error GoTo ErrTrap

    If ChkTaxSerivce.value = Checked Then
        TxtTaxServiceValue.Enabled = True
        lbl(43).Enabled = True
        lbl(47).Enabled = True
    Else
        TxtTaxServiceValue.text = ""
        TxtTaxServiceValue.Enabled = False
        lbl(43).Enabled = False
        lbl(47).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChkTaxStamp_Click()

    If ChkTaxStamp.value = Checked Then
        TxtTaxStampValue.Enabled = True
        lbl(41).Enabled = True
        lbl(48).Enabled = True
    Else
        TxtTaxStampValue.text = ""
        TxtTaxStampValue.Enabled = False
        lbl(41).Enabled = False
        lbl(48).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Cmd_Click(index As Integer)
    Dim AskOption As Boolean
    Dim intDef As Integer
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTest As ADODB.Recordset
    Dim RsOptions As ADODB.Recordset
    BolPrint = True
    On Error GoTo ErrTrap

    Select Case index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
        
            If SystemOptions.SysRegisterState = DemoRun Then
                Set RsTest = New ADODB.Recordset
                StrSQL = "Select Count(Transaction_ID) AS CountX From Transactions"
                RsTest.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (RsTest.BOF Or RsTest.EOF) Then
                    If RsTest("CountX").value >= 50 Then
                        Msg = "≈‰ Â  ‰”Œ… ⁄—÷ «·»—‰«„Ã ... »—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï"
                        Msg = Msg & CHR(13) & " "
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        Exit Sub
                    End If
                End If
            End If
        
            clear_all Me
            ClearNotes
            TxtModFlg.text = "N"
      
            SetDefaults
            NewGrid.GridDefaultValue 1
            Me.DCboUserName.BoundText = user_id
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultClient", 2)
            DBCboClientName.BoundText = intDef
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSaleStore", 1)
            DCboStoreName.BoundText = intDef
        
            Set RsOptions = New ADODB.Recordset
            RsOptions.Open "tbloptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable

            If Not (RsOptions.BOF Or RsOptions.EOF) Then
                Me.DcboBox.BoundText = IIf(IsNull(RsOptions("SalesBoxID").value), "", RsOptions("SalesBoxID").value)
            End If

            XPTab301.CurrTab = 0
            '------------------
            Me.XPDtbBill.SetFocus
            '--------------------
            Me.dcBranch.BoundText = Current_branch
      FG.rows = FG.FixedRows
            FG.rows = 2
      
            If Voucher_coding(val(my_branch), XPDtbBill.value, 18, 240) = "" And val(my_branch) <> 0 Then
                TxtNoteSerial1.locked = False
            Else
                TxtNoteSerial1.locked = True
 
            End If
 
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If


            If Not IsSaveWithOutMsg Then
                    If Trim(TxtNoteSerial1) = "" Then Exit Sub
                    
                    If ChekClodePeriod(XPDtbBill.value) = True And cmdReSave.Visible = False Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                        Else
                            MsgBox "Please Change Date Becouse This is Period is Closed"
                        End If
                        Exit Sub
                    End If
        End If

            'If AvailableDeal = True Then
            '«·√ﬁ”«ÿ «·„”Ã·… ⁄·Ï «·›« Ê—…
            
            If Text1.text <> "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "Â–« «·«–‰ ‰« Ã ⁄‰ ›« Ê—… ”«»ﬁ… Ê·« Ì„‰‰  ⁄œÌ·…  ›« Ê—… —ﬁ„  " & Space$(5) & Txtnots2.text
                Else
                    Msg = "This Voucher Created From Sales Invoice And Cant Modify"
                End If

                MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            If XPTxtValue(1).Tag <> "" Then
                StrSQL = "select * From InstallMent where NoteID=" & XPTxtValue(1).Tag
                Set RsTest = New ADODB.Recordset
                RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (RsTest.EOF Or RsTest.BOF) Then
                    Msg = "·ﬁœ  „  ﬁ”Ìÿ «·ﬁÌ„ «·¬Ã·… ⁄·Ï Â–Â «·›« Ê—…" & CHR(13)
                    Msg = Msg + " ⁄œÌ· «·›« Ê—… ”ÌƒœÌ ≈·Ï Õ–› Â–Â «·√ﬁ”«ÿ" & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì  ⁄œÌ· Â–Â «·›« Ê—…ø"

                    If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbNo Then
                        Exit Sub
                    End If
                End If
            End If

            '«·√ﬁ”«ÿ «·„”œœ… ⁄·Ï «·›« Ê—…
            If XPTxtValue(1).Tag <> "" Then
                StrSQL = "select * From ReceiptQestForBill where NoteID=" & XPTxtValue(1).Tag
                Set RsTest = New ADODB.Recordset
                RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (RsTest.EOF Or RsTest.BOF) Then
                    Msg = "·ﬁœ  „  Õ’Ì· »⁄÷ «·√ﬁ”«ÿ «·„”Ã·… ⁄·Ï Â–Â «·›« Ê—…" & CHR(13)
                    Msg = Msg + "Ê·« Ì„ﬂ‰  ⁄œÌ· »Ì«‰« Â«" & CHR(13)
                    Msg = Msg + "≈–« ﬂ‰   —€» ›Ì  ⁄œÌ· »Ì«‰«  Â–Â «·›« Ê—…" & CHR(13)
                    Msg = Msg + "ÌÃ» Õ–› ⁄„·Ì«  «· Õ’Ì· «·Œ«’… »Â«"
                    MsgBox Msg, vbOKOnly + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If

            '⁄„·Ì«  «·’Ì«‰… «·„— »ÿ… »«·›« Ê—…
            StrSQL = "select * From MaintenanceJuncTransaction where Transaction_ID=" & Trim(XPTxtBillID.text)
            Set RsTest = New ADODB.Recordset
            RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsTest.EOF Or RsTest.BOF) Then
                Msg = "·ﬁœ  „ ≈Ã—«¡ »⁄÷ ⁄„·Ì«  «·’Ì«‰… ⁄·Ï Â–Â «·›« Ê—… Ê·« Ì„ﬂ‰  ⁄œÌ·Â«"
                Msg = Msg + "≈–« ﬂ‰   —€» ›Ì  ⁄œÌ· »Ì«‰«  Â–Â «·›« Ê—…" & CHR(13)
                Msg = Msg + "ÌÃ» Õ–› ⁄„·Ì«  «·’Ì«‰… «·Œ«’… »Â«"
                MsgBox Msg, vbOKOnly + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            TxtModFlg.text = "E"
            CuurentLogdata
            Me.DCboUserName.BoundText = user_id
            'End If
   
            '-------------------------------

        Case 2
     
            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Departement"
                Else
                    Msg = "Õœœ «·›—⁄ «Ê·«  "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                dcBranch.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = val(Me.dcBranch.BoundText)

            If Text1.text <> "" Then
                Msg = "·ﬁœ  „  ÕÊÌ· Â–« «·«–‰ «·Ï ›« Ê—… „»Ì⁄«    .."
                Msg = Msg & CHR(13) & "Ê·«Ì„ﬂ‰  ÕÊÌ·… „—… «Œ—Ï  ..!!"
                MsgBox Msg, vbOKOnly, App.Title
                Exit Sub
                Else:
     
                '                        If Me.TxtModFlg.text = "N" Then
             
                '             End If
     
                Account_Code_dynamic = get_account_code_branch(37, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  „’«—Ì› «·«‰ «Ã „Ê«œ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
            
                    End If
                End If
                
                SaveData
     
            End If

        Case 3
     
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            If Text1.text <> "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "Â–« «·«–‰ ‰« Ã ⁄‰ ›« Ê—… ”«»ﬁ… Ê·« Ì„‰‰  ⁄œÌ·…  ›« Ê—… —ﬁ„  " & Space$(5) & Txtnots2.text
                Else
                    Msg = "This Voucher Created From Sales Invoice And Cant Modify"
                End If

                MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            Del_TransAction

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            If m_FrmSearch Is Nothing Then
                Set m_FrmSearch = New FrmBuySearch
                m_FrmSearch.DealingForm = RowMaterialIssue

                If SystemOptions.UserInterface = ArabicInterface Then
                    m_FrmSearch.Caption = "«·»ÕÀ ⁄‰ ⁄„·Ì… ’—› „Ê«œ Œ«„"
                Else
                    m_FrmSearch.Caption = "Search Production Issue Voucher"
                End If

                Set m_FrmSearch.RetrunFrm = Me
                m_FrmSearch.show vbModal ', mdifrmmain
            Else
                Msg = "Â‰«ﬂ ‘«‘… »ÕÀ  "
                Msg = Msg & CHR(13) & "Ÿ«Â—… «„«„ﬂ ›⁄·«...·«Ì„ﬂ‰ ⁄—÷ «ﬂÀ— „‰ ‘«‘… »ÕÀ  "
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                m_FrmSearch.ZOrder 0
                'm_FrmSearch.SetFocus
            End If

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If Me.XPTxtBillID.text = "" Then
                Msg = "·« ÊÃ·« ÌÊÃœ   ”‰œ ·ÿ»«⁄ …"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowMe", False)

            If AskOption = False Then

                FrmSallReportOptions.show vbModal

                If FrmSallReportOptions.UserCanceled = True Then
                    Unload FrmSallReportOptions
                    Exit Sub
                End If
'
                Unload FrmSallReportOptions
            End If


            PrintReport
    
        Case 6
            Unload Me

        Case 10
            ShowGL_cc TxtNoteSerial.text, , 200, val(Me.TXTNoteID.text)
    End Select

    Exit Sub

ErrTrap:

End Sub

Private Sub CmdCash_Click(index As Integer)

    Select Case index

        Case 0

        Case 1
    End Select

End Sub

Private Sub cmdCommand1_Click()
End Sub

Private Sub CmdConvert_Click()
    Dim RowNum As Integer
    Dim Frm As Form
    Dim Msg As String
    On Error GoTo ErrTrap

    If Text1.text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Â–« «·«–‰ ‰« Ã ⁄‰ ›« Ê—… ”«»ﬁ… Ê·« Ì„‰‰  ÕÊÌ·Â " & Space$(5) & Text1.text
        Else
            Msg = "This Voucher Created From Sales Invoice And Cant Convert Again"
        End If

        MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
    Set Frm = New frmsalebill

    With Frm

        .Convert
        '    .XPTxtBillID.Text = XPTxtBillID.Text
        .XPDtbBill.value = XPDtbBill.value
        .DBCboClientName.BoundText = DBCboClientName.BoundText
        .DCboStoreName.BoundText = DCboStoreName.BoundText
        .Text1.text = TxtTransSerial.text
        .XPCboDiscountType.ListIndex = Me.XPCboDiscountType.ListIndex
        .CboPayMentType.ListIndex = 0 ' Me.CboPaymentType.ListIndex
        .XPTxtDiscountVal.text = XPTxtDiscountVal.text
    
        For RowNum = 1 To FG.rows - 1

            If .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("Code")) <> "" Then
                .FG.rows = .FG.rows + 1
            End If

            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("Name")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Name")))
            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("Code")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Code")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Code")))
            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("ItemCase")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")))
            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("HaveSerial")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")))
            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("Count")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Count")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Count")))
            ' .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Price")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Price")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Price")))
            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("Price")) = GetItemPrice(.FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("Code")), 1)
      
            .FG.TextMatrix(.FG.rows - 1, .FG.ColIndex("DiscountType")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")))
            Dim StrSQL As String
            Dim RsUnit As New ADODB.Recordset
            StrSQL = "SELECT TOP 100 PERCENT dbo.TblItemsUnits.UnitID, dbo.TblUnites.UnitName, dbo.Transactions.Transaction_Serial,dbo.Transactions.Transaction_Type FROM dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN dbo.TblUnites INNER JOIN dbo.TblItemsUnits ON dbo.TblUnites.UnitID = dbo.TblItemsUnits.UnitID ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID AND dbo.Transaction_Details.Item_ID = dbo.TblItemsUnits.ItemID WHERE (dbo.Transactions.Transaction_Serial = '" & TxtTransSerial & "') AND (dbo.Transactions.Transaction_Type = 27) AND (dbo.TblItemsUnits.ItemID = " & FG.TextMatrix(RowNum, FG.ColIndex("Code")) & ") ORDER BY dbo.TblItemsUnits.SecOrder"
            Set RsUnit = New ADODB.Recordset
            RsUnit.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
            .FG.cell(flexcpData, RowNum, .FG.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitID")), "", (RsUnit("UnitID").value))
            .FG.TextMatrix(RowNum, .FG.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitName")), "", (RsUnit("UnitName").value))

            '        FG.Cell(flexcpData, I, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").Value))
            '        FG.TextMatrix(I, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").Value))
            '           StrSQL = "SELECT dbo.Transactions.Transaction_Type, dbo.Transaction_Details.UnitId, dbo.TblUnites.UnitName, dbo.Transactions.Transaction_Serial FROM dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID WHERE (dbo.Transactions.Transaction_Type = 19) AND (dbo.Transactions.Transaction_Serial = '" & TxtTransSerial & "')"
            '        .FG.Cell(flexcpData, .FG.Rows - 1, FG.ColIndex("UnitID")) = 1 'FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) ' IIf(IsNull(RsUnit("UnitID")), "", (RsUnit("UnitID").Value))
            '        .FG.TextMatrix(.FG.Rows - 1, FG.ColIndex("UnitID")) = "Ã—«„" 'FG.TextMatrix(RowNum, FG.ColIndex("UnitID")) ' IIf(IsNull(RsUnit("UnitName")), "", (RsUnit("UnitName").Value))

        Next RowNum

        .Cala
    End With

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault

End Sub

Private Sub CmdConvert1_Click()

End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub CmdInfo_Click()
    Dim xPoint As POINTAPI
    
    mdifrmmain.MnuInvInsertTemp.Visible = True
    
    'mdifrmmain.MnuInvSales_Mnu4.Enabled = Me.CmdNotes.Visible
    

    'ClientToScreen Me.CmdInfo.hwnd, xPoint
    'Me.PopupMenu MDIFrmMain.MnuInvSales, , (xPoint.X * Screen.TwipsPerPixelX), (xPoint.Y * Screen.TwipsPerPixelY)
    'Me.PopupMenu MDIFrmMain.MnuInvSales, vbPopupMenuRightAlign + vbPopupMenuRightButton, (xPoint.X * Screen.TwipsPerPixelX), (xPoint.Y * Screen.TwipsPerPixelY)

End Sub

Private Sub CmdINSTALLMENT_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim i As Integer

    If XPTxtValue(1).text = "" Then
        Msg = "ÌÃ»  ÕœÌœ «·ﬁÌ„… «·¬Ã·… ﬁ»·  ”ÃÌ· «·√ﬁ”«ÿ"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

        If XPTxtValue(1).Enabled = True Then
            XPTxtValue(1).SetFocus
        End If

        Exit Sub
    End If

    Load FrmInstallMent
    Set FrmInstallMent.Frm = Me

    With FrmInstallMent

        If Me.TxtModFlg.text = "R" Then
            .Tag = "R"
            .Retrive val(XPTxtValue(1).Tag)
        Else
            .Tag = "N"
            .Txt(1).text = XPTxtValue(1).text
            .LblNoteID.Caption = XPTxtSerial(1).text
            .CboPrecenType.ListIndex = val(Me.LblPrecenType.Tag)
            .Txt(3).text = val(LblPrecenValue.Caption)
            .Txt(5).text = val(LblInstallCount.Caption)

            If IsDate(Me.LblFirstInstallDate.Caption) Then
                .Dtp_First.value = Me.LblFirstInstallDate.Caption
            End If

            .Txt(7).text = val(LblInstallSeprator.Caption)

            If val(LblInstallmentType.Tag) = 0 Then
                .OptInt(0).value = True
            ElseIf val(LblInstallmentType.Tag) = 1 Then
                .OptInt(1).value = True
            ElseIf val(LblInstallmentType.Tag) = 2 Then
                .OptInt(2).value = True
            End If

            With .FG
                .rows = Me.FgInstallments.rows

                For i = 1 To Me.FgInstallments.rows - 1
                    .TextMatrix(i, .ColIndex("Serial")) = i
                    .TextMatrix(i, .ColIndex("Value")) = Me.FgInstallments.TextMatrix(i, Me.FgInstallments.ColIndex("Value"))
                    .TextMatrix(i, .ColIndex("Due_Date")) = Me.FgInstallments.TextMatrix(i, Me.FgInstallments.ColIndex("Due_Date"))
                Next i

                .AutoSize 0, .Cols - 1, False
            End With

        End If

        .show vbModal
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdInvProfit_Click()

    If SystemOptions.SysMainStockCostMethod = LastPurPriceType Or SystemOptions.SysMainStockCostMethod = ModernWeightAverage Then
        NewGrid.ShowInvProfDialog
    End If

    'If Me.TxtModFlg.Text = "R" Then
    '
    'Else
    '    NewGrid.ShowInvProfDialog
    'End If
End Sub

Private Sub CmdNotes_Click()
    ShowRelatedNotes val(Me.XPTxtBillID.text), 1
End Sub

Private Sub CmdNotes_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    Dim StrTemp As String

    If val(Me.CmdNotes.Tag) = 0 Then
        Me.CmdNotes.ToolTipText = ""
    Else
        StrTemp = " ÊÃœ ⁄·Ï Â–Â «·Õ—ﬂ… ⁄„·Ì«  „«·Ì… „ﬁœ«—Â« : " & val(Me.CmdNotes.Tag)
        Me.CmdNotes.ToolTipText = StrTemp
    End If

End Sub

Private Sub CmdRetruns_Click()
    ShowRelatedTransactions val(Me.XPTxtBillID.text), 1
End Sub

Private Sub CmdRetruns_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    Dim StrTemp As String

    If val(Me.CmdRetruns.Tag) = 0 Then
        Me.CmdRetruns.ToolTipText = ""
    Else
        StrTemp = " ÊÃœ ⁄·Ï Â–Â «·Õ—ﬂ… Õ—ﬂ«   Ã«—Ì… √Œ—Ï ·Â« ⁄·«ﬁ… »Â« ≈Ã„«·ÌÂ«: " & val(Me.CmdRetruns.Tag)
        Me.CmdRetruns.ToolTipText = StrTemp
    End If

End Sub

Private Sub CmdSearch_Click()
    'Dim LngItemID As Long
    'Dim LngStoreID As Long
    'LngItemID = Val(Me.DCboItemsName.BoundText)
    'LngStoreID = Val(Me.DCboStoreName.BoundText)
    'If LngItemID = 0 Or LngStoreID = 0 Then
    '    Exit Sub
    'End If
    'Load FrmSerialList
    'FrmSerialList.RetrunType = 1
    'Set FrmSerialList.m_TextBox = Me.TxtSerial
    'FrmSerialList.GetData LngItemID, LngStoreID
    'FrmSerialList.Show vbModal
End Sub

Private Sub Command1_Click()
    Dim MYWAER As String
    Dim Msg As String
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
    Dim i As Long
    Dim StrCurrentItemName As String
    Dim DblNotesTotal As Double

    Dim IntLineNO As Integer
    Dim StrAccountCode As String

    'Dim RsTranse As ADODB.Recordset
    Msg = "”Ê› Ì „ «‰‘«¡ ›« Ê—… »Ì⁄ »—ﬁ„ «–‰ «·’—›  .."
    Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

    If MsgBox(Msg, vbYesNo, App.Title) = vbYes Then
   
        rs.Close
        rs.Open "select * from Transactions where Transaction_Serial = " & TxtTransSerial.text & " and Transaction_type = 27"
         
        If Text1.text <> "" Then
            Msg = "·ﬁœ  „  ÕÊÌ· Â–« «·«–‰ «·Ï ›« Ê—… „»Ì⁄«    .."
            Msg = Msg & CHR(13) & "Ê·«Ì„ﬂ‰  ÕÊÌ·… „—… «Œ—Ï  ..!!"
            MsgBox Msg, vbOKOnly, App.Title
            Exit Sub
        End If

        rs!nots = TxtTransSerial.text
         
        rs.update
        '      MYWAER = " And Transaction_Type = 19"
        ''  "select * From ReplacedItems where ReturnID=" & XPTxtBillID.text
        ''                StrSQL = StrSQL + " and ItemID=" & RsDetails("Item_ID")
        Cn.Execute "INSERT INTO  Transactions (Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,CusID,StoreID,UserID,Emp_ID)SELECT Transaction_ID +1,Transaction_Serial,Transaction_Date,Transaction_Type = 21,CusID,StoreID,UserID,Emp_ID From Transactions Where Transaction_ID =" & XPTxtBillID.text + " And Transaction_Type = 27"
        '
        Cn.Execute "INSERT INTO  dbo.Transaction_Details(Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,UnitId,ShowQty,QtyBySmalltUnit)SELECT Transaction_ID+1,Item_ID,ItemCase,ItemSerial , Quantity, Price, ColorID, UnitId, ShowQty, QtyBySmalltUnit From dbo.Transaction_Details Where Transaction_ID = " & XPTxtBillID.text
        '
        '
        '      MYinvnum = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type= 2"))
        '
        '
        ''        End If
        '     Cn.Execute " update Transactions Set Transaction_Type = 2  , Transaction_Serial = '" & MYinvnum & "'  Where Transaction_Serial = " & TxtTransSerial.text
        '...............................................

        Set RsNotes = New ADODB.Recordset
        StrSQL = "Select * From Notes Where Transaction_ID=" & val(rs("Transaction_ID").value)
        RsNotes.Open StrSQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

        If (RsNotes.EOF Or RsNotes.BOF) Then
            If Me.XPChkPayType(0).value = Checked Then

                RsNotes.AddNew
                RsNotes("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))

                If Me.TxtModFlg.text = "N" Then
                    RsNotes("NoteSerial").value = CStr(new_id("Notes", "NoteSerial", "", True))
                    XPTxtSerial(0).text = RsNotes("NoteSerial").value
                ElseIf Trim(XPTxtSerial(0).text) <> "" Then
                    RsNotes("NoteSerial").value = Trim(XPTxtSerial(0).text)
                Else
                    RsNotes("NoteSerial").value = CStr(new_id("Notes", "NoteSerial", "", True))
                    XPTxtSerial(0).text = RsNotes("NoteSerial").value
                End If

                RsNotes("Transaction_ID").value = val(XPTxtBillID.text)
                RsNotes("NoteDate").value = XPDtbBill.value
                RsNotes("NoteType").value = 0
                RsNotes("Note_Value").value = IIf(XPTxtValue(0).text = "", Null, val(XPTxtValue(0).text))
                RsNotes("Member_ID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
                RsNotes("BankID").value = Null
                RsNotes("BoxID").value = IIf(DcboBox.BoundText = "", Null, val(DcboBox.BoundText))
                RsNotes("CUSID").value = Null

                RsNotes.update
            End If

            '«·ﬁÌ„ «·¬Ã·…
            If Me.XPChkPayType(1).value = Checked Then
                RsNotes.AddNew
                RsNotes("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
                XPTxtValue(1).Tag = IIf(IsNull(RsNotes("NoteID").value), "", (RsNotes("NoteID").value))
                note_id = RsNotes("NoteID").value
                RsNotes("NoteDate").value = XPDtbBill.value

                If Me.TxtModFlg.text = "N" Then
                    RsNotes("NoteSerial").value = CStr(new_id("Notes", "NoteSerial", "", True))
                    XPTxtSerial(1).text = RsNotes("NoteSerial").value
                ElseIf Trim(XPTxtSerial(1).text) <> "" Then
                    RsNotes("NoteSerial").value = Trim(XPTxtSerial(1).text)
                Else
                    RsNotes("NoteSerial").value = CStr(new_id("Notes", "NoteSerial", "", True))
                    XPTxtSerial(1).text = RsNotes("NoteSerial").value
                End If

                RsNotes("Transaction_ID").value = val(XPTxtBillID.text)
                RsNotes("NoteType").value = 1
                RsNotes("Note_Value").value = IIf(XPTxtValue(1).text = "", Null, val(XPTxtValue(1).text))
                RsNotes("Member_ID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
                RsNotes("BankID").value = Null
                RsNotes("CUSID").value = Null
                RsNotes("DueDate").value = DtpDelayDate.value
                RsNotes.update
            End If

            If Me.XPChkPayType(2).value = Checked Then

                With Me.FgCheques

                    For i = .FixedRows To .rows - 1
                        RsNotes.AddNew
                        RsNotes("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
                        RsNotes("NoteDate").value = XPDtbBill.value
                        RsNotes("Transaction_ID").value = val(XPTxtBillID.text)
                        RsNotes("NoteType").value = 2
                        RsNotes("Note_Value").value = val(.TextMatrix(i, .ColIndex("CheckValue")))
                        RsNotes("BankID").value = val(.TextMatrix(i, .ColIndex("BankID")))
                        RsNotes("ChqueNum").value = Trim$(.TextMatrix(i, .ColIndex("CheckNumber")))
                        RsNotes("DueDate").value = CDate(Trim$(.TextMatrix(i, .ColIndex("DueDate"))))
                        RsNotes("Member_ID").value = val(Me.DBCboClientName.BoundText)
                        RsNotes("CUSID").value = val(Me.DBCboClientName.BoundText)
                        RsNotes.update
                    Next i

                End With

            End If

            Else: Exit Sub
        End If
    End If

End Sub

Private Sub Command2_Click()
End Sub

Private Sub DBCboClientName_MouseUp(Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    Y As Single)

    If Button = vbRightButton Then
        mdifrmmain.MnuCusTools.Tag = Me.DBCboClientName.BoundText
        Me.PopupMenu mdifrmmain.MnuCusTools
    End If

End Sub

Private Sub DCboItemsCode_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 13
        FrmItemSearch.show vbModal
    End If

End Sub

Private Sub DCboStoreName_Change()
 

 TxtStoreID.text = getStoreCoding(val(DCboStoreName.BoundText))
 
    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then

     If CheckStoreCoding(val(dcBranch.BoundText), 18) = True Then
     TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""

     End If
     
    End If

End Sub

Private Sub DCboStoreName_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetStores Me.DCboStoreName
    End If

End Sub

Private Sub Dcbranch_Change()

    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
        Dcombos.GetDocTypebyid Me.DCDocTypes, 27, val(Me.dcBranch.BoundText)
        TxtNoteSerial1.text = ""
        
        TxtNoteSerial.text = ""
 

    End If

End Sub

Private Sub Dcbranch_Click(Area As Integer)
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
End Sub

Private Sub dcBranch_KeyUp(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetBranches dcBranch
    End If

End Sub

Private Sub Ele_DblClick(index As Integer)
    On Error GoTo ErrTrap

    If index = 9 Then
        If Me.WindowState = vbNormal Then
            Me.WindowState = vbMaximized
        Else
            Me.WindowState = vbNormal
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub FG_AfterEdit(ByVal row As Long, _
                         ByVal Col As Long)

    If Me.TxtModFlg <> "E" Then Exit Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    If Col = FG.ColIndex("Code") Or Col = FG.ColIndex("Name") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, row, FG.ColIndex("Name")), , , , , , , , , , val(Me.TxtNoteSerial), (Me.TxtNoteSerial1), 240
    ElseIf Col = FG.ColIndex("UnitID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, row, FG.ColIndex("Name")), FG.cell(flexcpTextDisplay, row, FG.ColIndex("UnitID")), , , , , , , , , val(Me.TxtNoteSerial), val(Me.TxtNoteSerial1), 240
    ElseIf Col = FG.ColIndex("Count") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, row, FG.ColIndex("Name")), , (FG.TextMatrix(row, FG.ColIndex("Count"))), , , , , , , , val(Me.TxtNoteSerial), val(Me.TxtNoteSerial1), 240
    ElseIf Col = FG.ColIndex("Price") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, row, FG.ColIndex("Name")), , , (FG.TextMatrix(row, FG.ColIndex("Price"))), , , , , , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 240
    ElseIf Col = FG.ColIndex("ColorID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, row, FG.ColIndex("Name")), , , , , FG.cell(flexcpTextDisplay, row, FG.ColIndex("ColorID")), , , , , val(Me.TxtNoteSerial), val(Me.TxtNoteSerial1), 240
    ElseIf Col = FG.ColIndex("ItemSize") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, row, FG.ColIndex("Name")), , , , , , FG.cell(flexcpTextDisplay, row, FG.ColIndex("ItemSize")), , , , val(Me.TxtNoteSerial), val(Me.TxtNoteSerial1), 240
    ElseIf Col = FG.ColIndex("ClassId") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, row, FG.ColIndex("Name")), , , , , , , FG.cell(flexcpTextDisplay, row, FG.ColIndex("ClassId")), , , val(Me.TxtNoteSerial), val(Me.TxtNoteSerial1), 240
    ElseIf Col = FG.ColIndex("DiscountType") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, row, FG.ColIndex("Name")), , , , , , , , FG.cell(flexcpTextDisplay, row, FG.ColIndex("DiscountType")), , val(Me.TxtNoteSerial), val(Me.TxtNoteSerial1), 240
    ElseIf Col = FG.ColIndex("DiscountVal") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, row, FG.ColIndex("Name")), , , , , , , , , FG.TextMatrix(row, FG.ColIndex("DiscountVal")), val(Me.TxtNoteSerial), val(Me.TxtNoteSerial1), 240

    End If

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
End Sub

Private Sub Fg_DblClick()
    'FrmItemsDetails.Show
End Sub

Private Sub Form_Activate()
    Set m_Menu1 = mdifrmmain.MnuInvInsertTemp
    Set m_MenuRefesh = mdifrmmain.MnuInvSales_Refresh
    Set m_MenuCusBalance = mdifrmmain.MnuInvSales_Mnu1
    Set m_MenuViewList = mdifrmmain.MnuInvViewList
    'Set m_MenuViewNotes = mdifrmmain.MnuInvSales_Mnu4
    Set m_MenuScreenPremission = mdifrmmain.MnuInvSales_Mnu7

    If TxtTransSerial.Enabled = True Then
        '    TxtTransSerial.SetFocus
    End If

End Sub

Private Sub lbl_MouseMove(index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          X As Single, _
                          Y As Single)

    If val(lbl(index).Caption) <> 0 Then
        lbl(index).ToolTipText = WriteNo(lbl(index).Caption, 0, True)
    End If

End Sub

Private Sub LblInstallCount_MouseMove(Button As Integer, _
                                      Shift As Integer, _
                                      X As Single, _
                                      Y As Single)
    LblInstallCount.ToolTipText = WriteNo(LblInstallCount.Caption, 0, True)
End Sub

Private Sub LblInstallTotal_MouseMove(Button As Integer, _
                                      Shift As Integer, _
                                      X As Single, _
                                      Y As Single)
    LblInstallTotal.ToolTipText = WriteNo(LblInstallTotal.Caption, 0, True)
End Sub

Private Sub LblInvProfit_Change()
    CalculateInvPrecent
End Sub

Private Sub LblPrecenValue_MouseMove(Button As Integer, _
                                     Shift As Integer, _
                                     X As Single, _
                                     Y As Single)
    LblPrecenValue.ToolTipText = WriteNo(LblPrecenValue.Caption, 0, True)
End Sub

Private Sub LblTotal_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    LblTotal.ToolTipText = WriteNo(LblTotal.Caption, 0, True)
End Sub

Private Sub m_FrmSearch_Unload(Cancel As Integer)
    Set m_FrmSearch = Nothing
End Sub

Private Sub m_Menu1_Click()
    On Error GoTo ErrTrap

    With FrmBuySearch
        .DealingForm = InsertTemplateToInvoice
        .Caption = "«·⁄—Ê÷ «·Ã«Â“…"
        .FG.TextMatrix(0, .FG.ColIndex("Transaction_ID")) = "ﬂÊœ «·⁄—÷"
        .FG.TextMatrix(0, .FG.ColIndex("BillDate")) = "«”„ «·⁄—÷"
        .FG.TextMatrix(0, .FG.ColIndex("ClientNmae")) = " «—ÌŒ «·⁄—÷"
        .FG.TextMatrix(0, .FG.ColIndex("StorName")) = "ﬁÌ„… «·⁄—÷"
        .XPChkSearchType.Visible = False
        .TxtVal.Visible = True
        .XPLbl(2).Visible = True
        .XPLbl(1).Visible = False
        .XPLbl(0).Visible = False
        .XPLbl(3).Visible = True
        .XPLbl(4).Visible = True
        .show vbModal
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub m_MenuCusBalance_Click()
    Dim cReport As ClsCustemerReport
    Dim LngCusID As Long

    With Me.FG

        If Me.DBCboClientName.BoundText = "" Then Exit Sub
        LngCusID = val(Me.DBCboClientName.BoundText)
        OpenScreen PopUpShowCustomerBalanceScreen, LngCusID, 0
    End With

End Sub

Private Sub m_MenuRefesh_Click()
    Dim Msg As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
        Msg = " ÕœÌÀ «·»Ì«‰«  €Ì— „ «Õ ≈·« «‰  ﬂÊ‰ «·‘«‘… ›Ï Õ«·… «·⁄—÷ ›ﬁÿ..!"
        'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        'Exit Sub
    End If

    LoadCombosData
    NewGrid.FillGrid
    rs.Requery
    Exit Sub
ErrTrap:
End Sub

Private Sub m_MenuScreenPremission_Click()
    ShowScreenPermission Me.Name
End Sub

Private Sub m_MenuViewList_Click()
    Dim FrmView As FrmViewList
    Dim FG As VSFlex8UCtl.VSFlexGrid
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim StrComboList As String
    Dim GrdBack As ClsBackGroundPic
    Dim cProgress As ClsProgress
    Dim BolFrmLoaded As Boolean
    Set FrmView = New FrmViewList
    Set FG = FrmView.vsfGroup1.VSFlexGrid

    With FG
        .Cols = 10
        .RowHeightMin = 320
        .TextMatrix(0, 0) = "—ﬁ„ «·»—‰«„Ã"
        .TextMatrix(0, 1) = "—ﬁ„ «·›« Ê—…"
        .TextMatrix(0, 2) = " «—ÌŒ «·›« Ê—…"
        .ColDataType(2) = flexDTDate
        .TextMatrix(0, 3) = "«”„ «·⁄„Ì·"
        .TextMatrix(0, 4) = "ÿ—Ìﬁ… «·œ›⁄"
        StrComboList = "#0;‰ﬁœÏ|#1;√Ã·"
        .ColComboList(4) = StrComboList
    
        .TextMatrix(0, 5) = "«”„ «·„Œ“‰"
        .TextMatrix(0, 6) = "«”„ «·„ÊŸ›"
    
        .TextMatrix(0, 7) = "‰Ê⁄ «·Œ’„"
        .TextMatrix(0, 8) = "ﬁÌ„… «·Œ’„"
        .TextMatrix(0, 9) = "≈Ã„«·Ï «·›« Ê—…"

        ',
        'QryTransactionsTotal.TransSum
        'QryTransactionsTotal.TransNet,
        If SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = "SELECT QryTransactionsTotal.Transaction_ID, QryTransactionsTotal.Transaction_Serial," & "QryTransactionsTotal.Transaction_Date,dbo.TblCustemers.CusName, QryTransactionsTotal.PaymentType, " & "dbo.TblStore.StoreName,dbo.TblEmployee.Emp_Name ,QryTransactionsTotal.Trans_DiscountType," & "QryTransactionsTotal.Trans_Discount,QryTransactionsTotal.TotalAfterTax"
            StrSQL = StrSQL + " FROM dbo.QryTransactionsTotal() QryTransactionsTotal LEFT OUTER JOIN"
            StrSQL = StrSQL + " dbo.TblStore ON QryTransactionsTotal.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
            StrSQL = StrSQL + " dbo.TblEmployee ON QryTransactionsTotal.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
            StrSQL = StrSQL + " dbo.TblCustemers ON QryTransactionsTotal.CusID = dbo.TblCustemers.CusID"
            StrSQL = StrSQL + " WHERE QryTransactionsTotal.Transaction_Type=2 "
            StrSQL = StrSQL + " Order  By QryTransactionsTotal.Transaction_ID"
        ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "SELECT QryTransactionsTotal.Transaction_ID , QryTransactionsTotal.Transaction_Serial," & "QryTransactionsTotal.Transaction_Date,TblCustemers.CusName, QryTransactionsTotal.PaymentType," & "TblStore.StoreName,TblEmployee.Emp_Name ,QryTransactionsTotal.Trans_DiscountType," & "QryTransactionsTotal.Trans_Discount,QryTransactionsTotal.TotalAfterTax "
            StrSQL = StrSQL + "FROM (TblEmployee RIGHT JOIN (TblCustemers RIGHT JOIN QryTransactionsTotal " & "ON TblCustemers.CusID = QryTransactionsTotal.CusID) ON TblEmployee.Emp_ID = QryTransactionsTotal.Emp_ID) " & "LEFT JOIN TblStore ON QryTransactionsTotal.StoreID = TblStore.StoreID "
            StrSQL = StrSQL + " WHERE QryTransactionsTotal.Transaction_Type=2 "
            StrSQL = StrSQL + " Order  By QryTransactionsTotal.Transaction_ID"
        End If

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenKeyset, adLockReadOnly, adAsyncExecute + adAsyncFetch
        Set cProgress = New ClsProgress
        BolFrmLoaded = True
        cProgress.ProgressType = Waiting
        cProgress.StartProgress

        Do While rs.State = adStateExecuting
            DoEvents
        Loop

        If BolFrmLoaded = True Then
            cProgress.StopProgess
            Set cProgress = Nothing
        End If

        Set .DataSource = rs
        .TextMatrix(0, 0) = "—ﬁ„ «·»—‰«„Ã"
        .TextMatrix(0, 1) = "—ﬁ„ «·›« Ê—…"
        .TextMatrix(0, 2) = " «—ÌŒ «·›« Ê—…"
        .ColDataType(2) = flexDTDate
        .TextMatrix(0, 3) = "«”„ «·⁄„Ì·"
        .TextMatrix(0, 4) = "ÿ—Ìﬁ… «·œ›⁄"
        StrComboList = "#0;‰ﬁœÏ|#1;√Ã·"
        .ColComboList(4) = StrComboList
        .TextMatrix(0, 5) = "«”„ «·„Œ“‰"
        .TextMatrix(0, 6) = "«”„ «·„ÊŸ›"
    
        .TextMatrix(0, 7) = "‰Ê⁄ «·Œ’„"
        .TextMatrix(0, 8) = "ﬁÌ„… «·Œ’„"
        .TextMatrix(0, 9) = "≈Ã„«·Ï «·›« Ê—…"
        .ColKey(9) = "TotalAfterTax"
        'Rs.Close
        'Set Rs = Nothing
    End With

    Set GrdBack = New ClsBackGroundPic
    FrmView.vsfGroup1.VSFlexGrid.WallPaper = GrdBack.Picture
    FrmView.vsfGroup1.SetRTL = True
    FrmView.vsfGroup1.TotalOnColKey = "TotalAfterTax"
    FrmView.vsfGroup1.update
    FrmView.show

End Sub

Private Sub m_MenuViewNotes_Click()
    CmdNotes_Click
End Sub

Private Sub sameCmd_Click()
Me.TxtModFlg.text = "N"
Me.XPDtbBill.value = Date
TxtNoteSerial1.text = ""
TxtNoteSerial.text = ""
End Sub

Public Function add_item_to_parts_grid1(Optional ItemID As Long, _
                                       Optional itemcode As String, _
                                      Optional ItemName As String, _
                                       Optional cost As Double, _
                                       Optional Qty As Double, _
                                       Optional productQty As Double, Optional UnitID As Integer, Optional ByVal mOrderNo As Long = 0)
    Dim Msg As String
    Dim LngFindRow As Long
    Dim LngNewRow As Long
    Dim StrSQL As String
  
  If val(txtMixID.text) = 0 Then Exit Function
    LngNewRow = ModFgLib.SetFgForNewRow(FG, FG.ColIndex("Code"))

  '  StrSQL = "SELECT TblItemsUnits.JunckID, TblItemsUnits.ItemID, TblItemsUnits.UnitID," & "TblUnites.UnitName, TblItemsUnits.UnitFactor, TblItemsUnits.SecOrder,TblItemsUnits.DefaultUnit," & "TblItemsUnits.UnitSalesPrice,TblItemsUnits.UnitPurPrice"
  '  StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits ON TblUnites.UnitID =" & "TblItemsUnits.UnitID "
  '  StrSQL = StrSQL + " Where  TblUnites.UnitID=" & val(unitid)
    
   StrSQL = "SELECT   TblDefComItem.ID,  dbo.TblDefComItemDet.ItemID, dbo.TblDefComItemDet.UnitID, dbo.TblDefComItemDet.IDDefCIT, dbo.TblDefComItemDet.Qty, dbo.TblItems.ItemID, "
  StrSQL = StrSQL + "  dbo.TblItems.itemcode , dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblUnites.Unitname, dbo.TblUnites.UnitNamee"
 StrSQL = StrSQL + " FROM         dbo.TblDefComItemDet INNER JOIN"
  
StrSQL = StrSQL + " TblDefComItem   On TblDefComItemDet.IDDefCIT = TblDefComItem.ID"
 StrSQL = StrSQL + " INNER JOIN"
 StrSQL = StrSQL + " dbo.TblItems ON dbo.TblDefComItemDet.ItemID = dbo.TblItems.ItemID INNER JOIN"
 StrSQL = StrSQL + " dbo.TblUnites ON dbo.TblDefComItemDet.UnitID = dbo.TblUnites.UnitID"
 'StrSQL = StrSQL + " WHERE     (dbo.TblDefComItemDet.IDDefCIT = " & val(txtMixID.Text) & ")"
 If mOrderNo <> 0 Then
    StrSQL = StrSQL + "  Where (TblDefComItemDet.IDDefCIT = " & val(txtMixID.text) & ") and IsNull(IsDeleted,0) = 0"
Else
    StrSQL = StrSQL + "  Where (TblDefComItem.ID = " & val(txtMixID.text) & " ) and IsNull(IsDeleted,0) = 0"
End If


productQty = val(LblTotalQty.Caption)
Dim i As Integer
    Dim rs As New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
       FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
    
    Dim UnitName As String
Dim item_cost As Double
    If Not (rs.BOF Or rs.EOF) Then
    FG.rows = rs.RecordCount + 1
    For i = 1 To rs.RecordCount
    LngNewRow = i
        UnitID = IIf(IsNull(rs("UnitID").value), 0, rs("UnitID").value)
        UnitName = IIf(IsNull(rs("UnitName").value), "", rs("UnitName").value)
        ItemID = IIf(IsNull(rs("ItemID").value), 0, rs("ItemID").value)
        itemcode = IIf(IsNull(rs("itemcode").value), "", rs("itemcode").value)
        ItemName = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
        Qty = IIf(IsNull(rs("Qty").value), 0, rs("Qty").value)
    cost = ModItemCostPrice.GetCostItemPrice(rs("ItemID").value, 0, , , SystemOptions.SysMainStockCostMethod, , , , , rs("UnitID").value)
    
       With Me.FG
       ' .TextMatrix(LngNewRow, .ColIndex("Item_ID")) = ItemID
        .TextMatrix(LngNewRow, .ColIndex("order_no")) = val(rs("ID") & "")
        .TextMatrix(LngNewRow, .ColIndex("code")) = ItemID
        .TextMatrix(LngNewRow, .ColIndex("Name")) = ItemName
        .TextMatrix(LngNewRow, .ColIndex("count")) = Qty
        .TextMatrix(LngNewRow, .ColIndex("UnitId")) = UnitID
        
          FG.TextMatrix(LngNewRow, FG.ColIndex("ItemCase")) = 1 ' IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
                FG.TextMatrix(LngNewRow, FG.ColIndex("DiscountType")) = 0 ' IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
                FG.TextMatrix(LngNewRow, FG.ColIndex("DiscountVal")) = 0 '  IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
                FG.TextMatrix(LngNewRow, FG.ColIndex("ColorID")) = 1 'IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
                FG.TextMatrix(LngNewRow, FG.ColIndex("ItemSize")) = 1 'IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
                FG.TextMatrix(LngNewRow, FG.ColIndex("ClassID")) = 1 'IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
                FG.TextMatrix(LngNewRow, FG.ColIndex("ItemType")) = 0 '  IIf(IsNull(RsDetails("ItemType")), 0, (RsDetails("ItemType").value))
              '    Fg.TextMatrix(currentrow, Fg.ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(LngItemID2, 0, "", , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, , val(Fg.TextMatrix(currentrow, Fg.ColIndex("UnitID"))))
  
                '   If RsDetails("HaveSerial") = True Then
                '       FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
                '   End If
        
                FG.cell(flexcpData, LngNewRow, FG.ColIndex("UnitID")) = UnitID
                FG.TextMatrix(LngNewRow, FG.ColIndex("UnitID")) = UnitName
'                Fg.TextMatrix(currentrow, Fg.ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(LngItemID2, 0, "", , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(XPTxtBillID.Text), val(Fg.Cell(flexcpData, currentrow, Fg.ColIndex("UnitID"))))
               FG.TextMatrix(LngNewRow, FG.ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(ItemID, 0, "", , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(XPTxtBillID.text), val(FG.TextMatrix(LngNewRow, FG.ColIndex("UnitID"))), val(Me.DCboStoreName.BoundText))
        
        .TextMatrix(LngNewRow, .ColIndex("ItemCostPrice")) = cost
       ' .TextMatrix(LngNewRow, .ColIndex("cost")) = cost
        '.TextMatrix(LngNewRow, .ColIndex("Price")) = rs!Price & ""
        .TextMatrix(LngNewRow, .ColIndex("Valu")) = cost * Qty
        '.TextMatrix(LngNewRow, .ColIndex("TotalQty")) = productQty * Qty
       ' .TextMatrix(LngNewRow, .ColIndex("Total")) = productQty * cost * Qty
    
        .AutoSize 0, .Cols - 1, False
   
        If .rows > 1 Then
            Me.LblTotalAll.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Valu"), .rows - 1, .ColIndex("Valu"))
            Me.LblTotalQty.Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("count"), .rows - 1, .ColIndex("count"))
        Else
          Me.LblTotalAll.Caption = 0
           Me.LblTotalQty.Caption = 0
        End If

    LblTotal = LblTotalAll
    End With
    
    rs.MoveNext
    Next i
    End If



End Function



Private Sub TxtFillData_Change()

    If TxtFillData.text = "F" Then
        NewGrid.Calculate 1, , , True
    End If

End Sub

Private Sub txtMIxCode_Change()
'If txtMIxCode.text = "" Then FG1.Rows = 1
 Me.txtMixID = ""
Me.txtMixID = GetMixIdFormCode(txtMIxCode)
    If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" And Me.TxtModFlg <> "E" Then
        add_item_to_parts_grid1
    End If
End Sub

Private Sub TxtTransSerial_KeyDown(KeyCode As Integer, _
                                   Shift As Integer)
    Dim StrSearch As String
    Dim VarBookMark As Variant
    Dim Msg As String

    If Me.TxtModFlg.text = "R" Then
        If KeyCode = vbKeyReturn Then
            If Trim$(TxtTransSerial.text) <> "" Then
                StrSearch = Trim$(TxtTransSerial.text)

                If Not (rs.BOF Or rs.EOF) Then
                    If rs.EditMode = adEditNone Then
                        VarBookMark = rs.Bookmark
                        rs.Find "Transaction_Serial='" & StrSearch & "'", , adSearchForward, adBookmarkFirst

                        If Not (rs.BOF Or rs.EOF) Then
                            Me.Retrive rs("Transaction_ID").value
                        Else
                            rs.Bookmark = VarBookMark
                            Msg = "Â–Â «·›« Ê—… €Ì— „ÊÃÊœ…...!!!"
                            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        End If
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Sub TxtTransSerial_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtTransSerial.text, 1)
End Sub

Private Sub TxtWorkOrderNO_Change()

    If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" And Me.TxtModFlg <> "E" Then
        RetriveOrder (Me.TxtWorkOrderNO.text)
    End If

End Sub

Public Sub RetriveOrder(Optional order_no As String = "")
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 1
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh

    Dim RsMainData  As New ADODB.Recordset
    Dim StrSQLMain As String
    Dim i As Integer
    Dim LngItemID As Long
    Dim LngItemID2 As Long
    Dim lngShowQty As Long
    Dim currentrow As Integer
    currentrow = 0
    StrSQLMain = " SELECT     dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ShowQty,Transactions.Transaction_ID"
    StrSQLMain = StrSQLMain & " FROM         dbo.Transactions INNER JOIN"
    StrSQLMain = StrSQLMain & " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
    StrSQLMain = StrSQLMain & "  WHERE     (dbo.Transactions.Transaction_Type = 26) AND (dbo.Transactions.Transaction_Serial = N'" & order_no & "')"
    If SystemOptions.AllowProductOrderOne = True Then
    StrSQLMain = StrSQLMain & " and dbo.Transactions.FlgProductOrder is null"
    End If
    RsMainData.Open StrSQLMain, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsMainData.RecordCount < 1 Then
 
        Exit Sub
    Else
        txtProductionOrderID.text = RsMainData!Transaction_ID & ""
    End If

    For i = 1 To RsMainData.RecordCount
        LngItemID = IIf(IsNull(RsMainData("Item_ID")), 0, (RsMainData("Item_ID").value))
        lngShowQty = IIf(IsNull(RsMainData("ShowQty")), 0, (RsMainData("ShowQty").value))
 
        StrSQL = "SELECT     TOP 100 PERCENT dbo.TblItemsParts.Unitid, dbo.TblItemsParts.PartItemPrice, dbo.TblItemsParts.PartItemQty, dbo.TblItemsParts.PartItemID, "
        StrSQL = StrSQL + " dbo.TblItemsParts.ItemID , dbo.TblItemsParts.TableID, dbo.TblUnites.unitname, dbo.TblUnites.UnitNamee"
        StrSQL = StrSQL + " FROM         dbo.TblItemsParts INNER JOIN"
        StrSQL = StrSQL + "  dbo.TblUnites ON dbo.TblItemsParts.Unitid = dbo.TblUnites.UnitID"
        StrSQL = StrSQL + " Where (dbo.TblItemsParts.ItemID = " & LngItemID & ")"
        StrSQL = StrSQL + " ORDER BY dbo.TblItemsParts.TableID"

        RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        'XPTxtSum.text = ""
        If Not (RsDetails.EOF Or RsDetails.BOF) Then

            For Num = 1 To RsDetails.RecordCount
                currentrow = currentrow + 1
                FG.rows = FG.rows + 1
                LngItemID2 = IIf(IsNull(RsDetails("partItemID")), 0, (RsDetails("partItemID").value))
                
                FG.TextMatrix(currentrow, FG.ColIndex("ItemId2")) = LngItemID
                FG.TextMatrix(currentrow, FG.ColIndex("ItemName2")) = GetItemName(LngItemID)
                
                FG.TextMatrix(currentrow, FG.ColIndex("Code")) = LngItemID2
                FG.TextMatrix(currentrow, FG.ColIndex("Name")) = LngItemID2
                FG.TextMatrix(currentrow, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("partitemqty")), 0, (RsDetails("partitemqty").value)) * lngShowQty
                'FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
                'FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("ShowPrice")), 0, (RsDetails("ShowPrice").value)) ' GET_COST_PRICE_FOR_PRODUCT_ITEM(Val(FG.TextMatrix(Num, FG.ColIndex("Code"))))
      
                '  FG.TextMatrix(Num, FG.ColIndex("Expenses")) = IIf(IsNull(RsDetails("Lineexpenses")), "", (RsDetails("Lineexpenses").value))
         
                FG.TextMatrix(currentrow, FG.ColIndex("ItemCase")) = 1 ' IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
                FG.TextMatrix(currentrow, FG.ColIndex("DiscountType")) = 0 ' IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
                FG.TextMatrix(currentrow, FG.ColIndex("DiscountVal")) = 0 '  IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
                FG.TextMatrix(currentrow, FG.ColIndex("ColorID")) = 1 'IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
                FG.TextMatrix(currentrow, FG.ColIndex("ItemSize")) = 1 'IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
                FG.TextMatrix(currentrow, FG.ColIndex("ClassID")) = 1 'IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
                FG.TextMatrix(currentrow, FG.ColIndex("ItemType")) = 0 '  IIf(IsNull(RsDetails("ItemType")), 0, (RsDetails("ItemType").value))
              '    Fg.TextMatrix(currentrow, Fg.ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(LngItemID2, 0, "", , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, , val(Fg.TextMatrix(currentrow, Fg.ColIndex("UnitID"))))
  
                '   If RsDetails("HaveSerial") = True Then
                '       FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
                '   End If
        
                FG.cell(flexcpData, currentrow, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
                FG.TextMatrix(currentrow, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
'                Fg.TextMatrix(currentrow, Fg.ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(LngItemID2, 0, "", , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(XPTxtBillID.Text), val(Fg.Cell(flexcpData, currentrow, Fg.ColIndex("UnitID"))))
               FG.TextMatrix(currentrow, FG.ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(LngItemID2, 0, "", , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(XPTxtBillID.text), val(FG.TextMatrix(currentrow, FG.ColIndex("UnitID"))), val(Me.DCboStoreName.BoundText))
                 
                RsDetails.MoveNext
         
                '    Debug.Print Num
                '    If FG.Rows > 10 Then
                '        If Num = 8 Then FG.Refresh
                '    End If
            Next Num

        End If

        RsDetails.Close
        RsMainData.MoveNext
    Next i

    TxtFillData.text = "F"
    Screen.MousePointer = vbDefault
    ' XPDtbBill_Change

    'XPTxtCurrent.Caption = rs.AbsolutePosition
    'XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub
Private Function GetItemName(ByVal mItemNo As Long) As String
Dim rsDummy As New ADODB.Recordset
Dim s As String
s = "Select ItemName from TblItems Where ItemId = " & mItemNo
rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
If Not rsDummy.EOF Then
    GetItemName = rsDummy!ItemName & ""
End If

End Function
Private Sub TxtWorkOrderNO_KeyUp(KeyCode As Integer, _
                                 Shift As Integer)

    If KeyCode = vbKeyF3 Then
       Order_no_search2.show
        Order_no_search2.RetrunType = 1
 
    End If

End Sub

Private Sub XPBtnMove_Click(index As Integer)
'    On Error GoTo ErrTrap

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

'
Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" And Not (Me.ActiveControl Is TxtTransSerial) Then
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
            'XPBtnAdd_Click
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            'XPBtnRemove_Click
        End If
    End If

    If KeyCode = vbKeyDelete Then
        If Me.ActiveControl Is FG Then
            If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
                'XPBtnRemove_Click
            End If
        End If
    End If

    If KeyCode = vbKeyF5 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            XPBtnNewClients_Click
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeySpace Then
            If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
                'XPFillData_Click
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

    If Shift = VBRUN.ShiftConstants.vbShiftMask Then

        'vbKeyX
        If KeyCode = vbKeyEscape Then
            Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    Dim StrSQL As String
    Dim Num As Integer
    Dim StrList As String
    Dim BGround As New ClsBackGroundPic
    Dim ShowTax As Boolean

   ' On Error GoTo ErrTrap

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    ScreenNameArabic = " ”‰œ ’—› „Ê«œ Œ«„ ··«‰ «Ã "
    ScreenNameEnglish = " Production Issue Voucher "
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 240
 
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    'Set m_menu1.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Excute").Picture

    Dim My_SQL As String
    'My_SQL = "  select branch_id,branch_name from TblBranchesData   "
    'fill_combo DcBranch, My_SQL

    If SystemOptions.usertype <> UserAdminAll Then
 
        Me.dcBranch.Enabled = True
    End If

    Set NewGrid.Grid = FG

    ShowTax = GetSetting(StrAppRegPath, "SallBill", "HaveTaxOnSalles", False)
    Ele(4).Visible = ShowTax
    'NewGrid.GridTrans = InvoiceTransaction
    NewGrid.GridTrans = InventoryOut
    Set NewGrid.DtpBillDate = Me.XPDtbBill
    Set NewGrid.TxtInvID = Me.XPTxtBillID
    Set NewGrid.TxtModFlag = TxtModFlg
    Set NewGrid.txtTotal = XPTxtSum
    Set NewGrid.CboDiscount_Type = XPCboDiscountType
    Set NewGrid.TxtDiscount_Val = XPTxtDiscountVal
    Set NewGrid.TxtValueCash = XPTxtValue(0)
    Set NewGrid.TxtValueDelay = XPTxtValue(1)
    Set NewGrid.TxtValuechque = XPTxtValue(2)
    '--------------------------------------
    Set NewGrid.TxtTaxValue = Me.XPTxtTaxValue
    Set NewGrid.TxtAddTax = Me.TxtTaxAddValue
    Set NewGrid.TxtStampTax = Me.TxtTaxStampValue
    Set NewGrid.TxtServiceTax = Me.TxtTaxServiceValue
    Set NewGrid.LblTotalQty = Me.LblTotalQty
    '------------------------------------------------
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.StoreName = Me.DCboStoreName
    Set NewGrid.DtpBillDate = Me.XPDtbBill
    Set NewGrid.CmdAddSerialLIst = Me.CmdSearch
    'Set NewGrid.CboDiscountType = CboDiscountType
    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    Set NewGrid.DCboItemName = DCboItemsName
    Set NewGrid.DCboItemCode = DCboItemsCode
    Set NewGrid.CboItemCase = CboItemCase
    Set NewGrid.CmdAddData = CmdAdd
    Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.TxtQuantity = TxtQuantity
    Set NewGrid.TxtPrice = TxtPrice
    Set NewGrid.LblInvProfit = Me.LblInvProfit
    Set NewGrid.LblItemsCount = Me.LblItemsCount
    Set NewGrid.GrdTBar = Me.TBar
    Set NewGrid.LblTotalAll = Me.LblTotalAll
    Set NewGrid.LblDiscountsTotal = Me.LblDiscountsTotal
    Set NewGrid.LblTaxSalesValue = Me.lbl(51)
    Set NewGrid.LblTaxAddValue = Me.lbl(52)
    Set NewGrid.LblTaxStampValue = Me.lbl(53)
    Set NewGrid.LblTaxServiceValue = Me.lbl(54)

    NewGrid.FillGrid
    FG.WallPaper = BGround.Picture
    AddTip
    XPTab301.CurrTab = 0
    XPDtbBill.value = Date

    If SystemOptions.UserInterface = ArabicInterface Then

        With XPCboDiscountType
            .Clear
            .AddItem "·«ÌÊÃœ Œ’„"
            .AddItem "Œ’„ »ﬁÌ„…"
            .AddItem "Œ’„ »‰”»…"
        End With

        With CboPayMentType
            .Clear
            .AddItem "‰ﬁœ«"
            .AddItem "¬Ã·"
        End With

        With Me.CboSaleType
            .Clear
            .AddItem "ﬁÿ«⁄Ì"
            .AddItem " Ã«—Ï"
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then

        With XPCboDiscountType
            .Clear
            .AddItem "No Discount"
            .AddItem "Value Discount"
            .AddItem "Precetage Discount"
        End With

        With CboPayMentType
            .Clear
            .AddItem "Cash"
            .AddItem "Due"
        End With

        With Me.CboSaleType
            .Clear
            .AddItem "Retail"
            .AddItem "WholeSale"
        End With

    End If

    '--------------------------------
    Set Dcombos = New ClsDataCombos
    LoadCombosData

    '--------------------------------
    If SystemOptions.UserInvoiceShowProfit = 0 Then
        Me.Ele(8).Visible = False
    Else
        ' Me.Ele(8).Visible = True
    End If

    SetDtpickerDate Me.XPDtbBill
    '----------------------------
    SetDtpickerDate Me.DtpDelayDate
    '≈⁄œ«œ Ã—œ «·√ﬁ”«ÿ
    ChkInstall.value = Unchecked
    ChkInstall.Enabled = False

    With Me.FgInstallments
        .rows = .FixedRows
        Set .WallPaper = BGround.Picture
        .RowHeightMin = 300
        .AutoSize 0, .Cols - 1, False
    End With

    With Me.FgCheques
        .rows = .FixedRows
        Set .WallPaper = BGround.Picture
        .RowHeightMin = 300
        .AutoSize 0, .Cols - 1, False
    End With

    Me.XPChkTAX.value = vbUnchecked
    XPChkTAX_Click
    Me.ChkTaxAdd.value = vbUnchecked
    ChkTaxAdd_Click
    Me.ChkTaxStamp.value = vbUnchecked
    ChkTaxStamp_Click
    Me.ChkTaxSerivce.value = vbUnchecked
    ChkTaxSerivce_Click
    '---------------------------
    Resize_Form Me, TransactionSize
    '----------------------------
    'DB_CreateField "Transactions", "TransactionComment", adVarWChar, adColNullable, 255, , " ”ÃÌ· „·«ÕŸ«  ⁄·Ï «·›« Ê—…", False, True
    '----------------------------
    
    
    

    StrSQL = "SELECT  * FROM Transactions WHERE Transaction_Type= 27"
StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"
    If SystemOptions.usertype <> UserAdminAll Then
        'StrSQL = StrSQL & " AND   BranchId=" & branch_id
    End If

    StrSQL = StrSQL & "  Order by Transaction_ID "

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveLast
    End If

    Retrive
    Me.TxtModFlg.text = "R"

    InvType = 27
If SystemOptions.HideCost = True Then
LblTotalAll.Visible = False
LblTotal.Visible = False

TxtPrice.Visible = False
       FG.ColHidden(FG.ColIndex("Price")) = True
       FG.ColHidden(FG.ColIndex("Valu")) = True


End If
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub
 
Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·«–‰   " & TxtNoteSerial1.text & CHR(13) & "  «—ÌŒ «·«–‰ " & XPDtbBill.value & CHR(13) & " «·›—⁄   " & dcBranch.text & CHR(13) & "—ﬁ„ «„— «·«‰ «Ã  " & TxtWorkOrderNO & CHR(13) & " «·„Œ“‰  " & DCboStoreName.text
                     
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Vchr. No.  " & TxtNoteSerial1.text & CHR(13) & "Date " & XPDtbBill.value & CHR(13) & " Branch   " & dcBranch.text & CHR(13) & " To  Order No " & TxtWorkOrderNO & CHR(13) & " Inventory  " & DCboStoreName.text
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 240, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, , , Me.TxtNoteSerial, Me.TxtNoteSerial1
    Else
        AddToLogFile CInt(user_id), 240, Date, Time, LogTextA, LogTexte, Me.Name, "D", , , Me.TxtNoteSerial, Me.TxtNoteSerial1
    End If
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    Dim i As Integer
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, , 240

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    Set Dcombos = Nothing

    For i = LBound(cSearchDcbo) To UBound(cSearchDcbo)
        Set cSearchDcbo(i) = Nothing
    Next i

    Set rs = Nothing
    Set TTP = Nothing
    NewGrid.Class_Terminate
    Set NewGrid = Nothing
    Set SaleReport = Nothing

    Set m_Menu1 = Nothing
    Set m_MenuRefesh = Nothing

    If Not m_FrmSearch Is Nothing Then
        Unload m_FrmSearch
        Set m_FrmSearch = Nothing
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap
    Dim RsTest As ADODB.Recordset
    Dim StrSQL As String

    Select Case Me.TxtModFlg.text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = " ”‰œ ’—› „Ê«œ Œ«„ ··«‰ «Ã"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.Caption = "Production Issue Voucer"
            End If

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
            XPBtnNewClients.Enabled = False
        
            XPCboDiscountType.locked = True
            Me.XPDtbBill.Enabled = False
            Me.DBCboClientName.locked = True
            Me.DCboStoreName.locked = True
        
            Me.XPTxtDiscountVal.locked = True
            XPChkPayType(0).Enabled = False
            XPChkPayType(1).Enabled = False
            XPChkPayType(2).Enabled = False
            XPTxtValue(0).Enabled = False
            XPTxtSerial(0).Enabled = False
            XPTxtValue(1).Enabled = False
            XPTxtSerial(1).Enabled = False
        
            FG.Editable = flexEDNone
            XPChkTAX.Enabled = False

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

            If Not m_Menu1 Is Nothing Then
                m_Menu1.Enabled = False
            End If

            CmdINSTALLMENT.Enabled = False
            CmdCheque.Enabled = False

            '⁄—÷ «·√ﬁ”«ÿ «·„”Ã·… ⁄·Ï «·›« Ê—…
            If XPTxtValue(1).Tag <> "" Then
                StrSQL = "select * From InstallMent where NoteID=" & XPTxtValue(1).Tag
                Set RsTest = New ADODB.Recordset
                RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (RsTest.EOF Or RsTest.BOF) Then
                    CmdINSTALLMENT.Enabled = True
                    CmdINSTALLMENT.Caption = "⁄—÷ «·√ﬁ”«ÿ «·„”Ã·…"
                Else
                    CmdINSTALLMENT.Enabled = False
                    CmdINSTALLMENT.Caption = " ﬁ”Ìÿ «·ﬁÌ„… «·¬Ã·…"
                End If
            End If

            Ele(2).Enabled = False
            DcboEmp.Enabled = False
            XPChkTAX.Enabled = False
            ChkTaxAdd.Enabled = False
            ChkTaxSerivce.Enabled = False
            ChkTaxStamp.Enabled = False

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "”‰œ ’—› „Ê«œ Œ«„ ··«‰ «Ã ( ÃœÌœ )"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.Caption = "  Production Issue Voucher(New)"
            End If

            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
            DcboEmp.Enabled = True

            '  Me.XPBtnMove(0).Enabled = False
            '  Me.XPBtnMove(1).Enabled = False
            '  Me.XPBtnMove(2).Enabled = False
            '  Me.XPBtnMove(3).Enabled = False
            XPBtnNewClients.Enabled = True
            FG.Enabled = True
            XPCboDiscountType.locked = False
            Me.XPDtbBill.Enabled = True
            XPDtbBill.value = Date
            Me.DBCboClientName.locked = False
            CboPayMentType.locked = False
            Me.DCboStoreName.locked = False
            Me.XPTxtDiscountVal.locked = False
        
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            XPChkPayType(0).value = Unchecked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            FG.Editable = flexEDKbdMouse
            XPChkTAX.Enabled = True
            XPTxtTaxValue.text = ""
            XPChkTAX.value = Unchecked
            XPCboDiscountType.ListIndex = 0
            CboPayMentType.ListIndex = 0
            '        XPFillData.Enabled = True
            DtpDelayDate.Enabled = True
            m_Menu1.Enabled = True
            DtpDelayDate.value = Date
       
            CmdINSTALLMENT.Enabled = False
            CmdCheque.Enabled = False
            Ele(2).Enabled = True
            CboItemCase.ListIndex = 0
        
            Me.LblInvProfit.Caption = "0.0"
            Me.LblInvProfit.ForeColor = vbBlack
        
            DcboEmp.Enabled = True
            XPChkTAX.Enabled = True
            ChkTaxAdd.Enabled = True
            ChkTaxSerivce.Enabled = True
            ChkTaxStamp.Enabled = True
        
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "”‰œ ’—› „Ê«œ Œ«„ ··«‰ «Ã  (   ⁄œÌ· )"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.Caption = "  Production Issue Voucher( Edit )"
            End If

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
            XPCboDiscountType.locked = False
            Me.XPDtbBill.Enabled = True
            Me.DBCboClientName.locked = False
            Me.DCboStoreName.locked = False
            Me.XPTxtDiscountVal.locked = False
            CboPayMentType.locked = False
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
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
            XPBtnNewClients.Enabled = True
            XPChkTAX.Enabled = True

            If Not m_Menu1 Is Nothing Then
                m_Menu1.Enabled = False
            End If

            If XPChkPayType(1).value = vbChecked Then
                If XPTxtValue(1).text <> "" Then
                    CmdINSTALLMENT.Enabled = True
                    CmdINSTALLMENT.Caption = " ﬁ”Ìÿ «·ﬁÌ„… «·¬Ã·…"
                Else
                    CmdINSTALLMENT.Enabled = False
                End If
            End If

            If Me.XPChkPayType(2).value = vbChecked Then
                CmdCheque.Enabled = True
            Else
                CmdCheque.Enabled = False
            End If

            DBCboClientName_Change
            Ele(2).Enabled = True
        
            DcboEmp.Enabled = True
            XPChkTAX.Enabled = True
            ChkTaxAdd.Enabled = True
            ChkTaxSerivce.Enabled = True
            ChkTaxStamp.Enabled = True
        
    End Select

    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTest  As ADODB.Recordset
    Dim RsReplace As ADODB.Recordset
    Dim LngPartID As Long
    Dim RsPartDetails As ADODB.Recordset
    Dim i As Long

  '  On Error GoTo ErrTrap
    '---------------------------------------------
    'Here We Reset all Setting
    Me.CmdNotes.Visible = False
    Me.CmdNotes.Tag = ""
    Me.CmdRetruns.Visible = False
    Me.CmdRetruns.Tag = ""

    ChkTaxAdd.value = vbUnchecked
    Me.TxtTaxAddValue.text = ""
    ChkTaxStamp.value = vbUnchecked
    Me.TxtTaxStampValue.text = ""
    ChkTaxStamp.value = vbUnchecked
    Me.TxtTaxStampValue.text = ""
    ChkTaxSerivce.value = vbUnchecked
    Me.TxtTaxServiceValue.text = ""

    '---------------------------------------------
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
    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", (rs("NoteSerial").value))
    Me.TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", (rs("NoteSerial1").value))
    Me.oldtxtNoteSerial1.text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)

    lbl(56).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

    Me.TXTNoteID.text = IIf(IsNull(rs("NoteID").value), "", (rs("NoteID").value))
    txtProductionOrderID.text = IIf(IsNull(rs("ProductionOrderID").value), "", (rs("ProductionOrderID").value))
    
    XPTxtBillID.text = IIf(IsNull(rs("Transaction_ID").value), "", val(rs("Transaction_ID").value))
    TxtTransSerial.text = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
    XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", (rs("Transaction_Date").value))
    txtremark.text = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)

    DCDocTypes.BoundText = IIf(IsNull(rs("Doctype").value), "", rs("Doctype").value)

    XPCboDiscountType.ListIndex = IIf(IsNull(rs("Trans_DiscountType").value), -1, val(rs("Trans_DiscountType").value))
    CboPayMentType.ListIndex = IIf(IsNull(rs("PaymentType").value), 0, rs("PaymentType").value)
    XPTxtDiscountVal.text = IIf(IsNull(rs("Trans_Discount").value), "", (rs("Trans_Discount").value))
    Me.DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    FG.Clear flexClearScrollable, flexClearEverything
    Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
    Me.DcboEmp.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
    XPTxtTaxValue.text = IIf(IsNull(rs("TaxValue").value), "", (rs("TaxValue").value))
    XPChkTAX.value = IIf(rs("TaxFound") = True, Checked, Unchecked)
    Text1.text = IIf(IsNull(rs("nots").value), "", (rs("nots").value))
    Txtnots2.text = IIf(IsNull(rs("nots2").value), "", (rs("nots2").value))
    TxtWorkOrderNO.text = IIf(IsNull(rs("WorkOrderNO").value), "", (rs("WorkOrderNO").value))
    txtMIxCode.text = IIf(IsNull(rs("MIxCode").value), "", (rs("MIxCode").value))
    txtMixID.text = IIf(IsNull(rs("MixID").value), "", (rs("MixID").value))
       

    
    dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", (rs("BranchId").value))

    If IsNull(rs("SaleType").value) Then
        Me.CboSaleType.ListIndex = 0
    Else
        Me.CboSaleType.ListIndex = IIf(rs("SaleType").value = 0, 0, 1)
    End If

    If Not (IsNull(rs("CashCustomerName").value)) Then
        Me.TxtCashCustomerName.text = rs("CashCustomerName").value
    Else
        Me.TxtCashCustomerName.text = ""
    End If

    '÷—»Ì… «·Œ’„ Ê«·≈÷«›…
    If Not IsNull(rs("TaxAddValue").value) Then
        If rs("TaxAddValue").value > 0 Then
            ChkTaxAdd.value = vbChecked
            Me.TxtTaxAddValue.text = rs("TaxAddValue").value
        End If
    End If

    '÷—»Ì… «·œ„€…
    If Not IsNull(rs("TaxStampValue").value) Then
        If rs("TaxStampValue").value > 0 Then
            ChkTaxStamp.value = vbChecked
            Me.TxtTaxStampValue.text = rs("TaxStampValue").value
        End If
    End If

    '÷—»Ì… «·Œœ„…
    If Not IsNull(rs("TaxServiceValue").value) Then
        If rs("TaxServiceValue").value > 0 Then
            ChkTaxSerivce.value = vbChecked
            Me.TxtTaxServiceValue.text = rs("TaxServiceValue").value
        End If
    End If

    TxtBillComment.text = IIf(IsNull(rs("TransactionComment").value), "", (rs("TransactionComment").value))
    FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)

    Set RsDetails = New ADODB.Recordset
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.rows = RsDetails.RecordCount + 1

        For i = 1 To RsDetails.RecordCount
            FG.TextMatrix(i, FG.ColIndex("FoxyNo")) = IIf(IsNull(RsDetails("FoxyNo")), "", (RsDetails("FoxyNo").value))
            FG.TextMatrix(i, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no").value), "", RsDetails("order_no").value)
            FG.TextMatrix(i, FG.ColIndex("OrderArrivalDate")) = IIf(IsNull(RsDetails("OrderArrivalDate").value), "", RsDetails("OrderArrivalDate").value)
   
            FG.cell(flexcpPicture, i, FG.ColIndex("Ser")) = ""
            FG.cell(flexcpData, i, FG.ColIndex("Ser")) = ""
            FG.TextMatrix(i, FG.ColIndex("ItemId2")) = IIf(IsNull(RsDetails("ItemId2")), "", (RsDetails("ItemId2").value))
            FG.TextMatrix(i, FG.ColIndex("ItemName2")) = GetItemName(val(RsDetails!ItemID2 & ""))
            FG.TextMatrix(i, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(i, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim$(RsDetails("Item_ID").value))
            FG.TextMatrix(i, FG.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial")), "", Trim(RsDetails("ItemSerial").value))

            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(i, FG.ColIndex("HaveSerial")) = True

                '«·»ÕÀ ⁄‰ ⁄„·Ì«  «·«” »œ«· «·Œ«’… »«·›« Ê—…
                If (RsDetails("Item_ID")) <> "" And RsDetails("ItemSerial") <> "" Then
                    StrSQL = "select * From ReplacedItems where ReturnID=" & XPTxtBillID.text
                    StrSQL = StrSQL + " and ItemID=" & RsDetails("Item_ID")
                    StrSQL = StrSQL + " and ItemSerial='" & RsDetails("ItemSerial") & "'"
                    Set RsReplace = New ADODB.Recordset
                    RsReplace.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (RsReplace.EOF Or RsReplace.BOF) Then
                        FG.cell(flexcpPicture, i, FG.ColIndex("Ser")) = mdifrmmain.ImgLstTree.ListImages("Request").Picture
                        FG.cell(flexcpData, i, FG.ColIndex("Ser")) = "X"
                    End If
                End If
            End If

            FG.TextMatrix(i, FG.ColIndex("ItemType")) = IIf(IsNull(RsDetails("ItemType").value), "", (RsDetails("ItemType").value))

            If RsDetails("ItemType").value = 1 Then
                FG.cell(flexcpPicture, i, FG.ColIndex("Ser")) = mdifrmmain.ImgLstTree.ListImages("Maintenance").Picture
            
            End If
FG.TextMatrix(i, FG.ColIndex("NProductionOrderNO")) = IIf(IsNull(RsDetails("NProductionOrderNO")), "", Trim(RsDetails("NProductionOrderNO").value))

            FG.TextMatrix(i, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("ShowQty")), "", (RsDetails("ShowQty").value))
            FG.TextMatrix(i, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
            FG.TextMatrix(i, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice")), "", (RsDetails("showPrice").value))
        
            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                FG.TextMatrix(i, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            Else
                FG.TextMatrix(i, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("Transaction_Details.ItemCase")), "", (RsDetails("Transaction_Details.ItemCase").value))
            End If
        
            FG.TextMatrix(i, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(i, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(i, FG.ColIndex("guaranteeTime")) = IIf(IsNull(RsDetails("guaranteeTime")), "", (RsDetails("guaranteeTime").value))
        
            FG.TextMatrix(i, FG.ColIndex("ItemCostPrice")) = IIf(IsNull(RsDetails("CostPrice")), "", (RsDetails("CostPrice").value))
            FG.TextMatrix(i, FG.ColIndex("PofTransID")) = IIf(IsNull(RsDetails("CostTransID")), "", (RsDetails("CostTransID").value))
            FG.TextMatrix(i, FG.ColIndex("ItemProfit")) = IIf(IsNull(RsDetails("ItemProfit")), "", (RsDetails("ItemProfit").value))
            FG.TextMatrix(i, FG.ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks")), "", (RsDetails("Remarks").value))
        
            FG.TextMatrix(i, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(i, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(i, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
                
            If val(FG.TextMatrix(i, FG.ColIndex("ItemProfit"))) = 0 Then
                Me.FG.cell(flexcpBackColor, i, 1, i, FG.Cols - 1) = vbYellow
            ElseIf val(FG.TextMatrix(i, FG.ColIndex("ItemProfit"))) < 0 Then
                Me.FG.cell(flexcpBackColor, i, 1, i, FG.Cols - 1) = vbRed
            Else
                Me.FG.cell(flexcpBackColor, i, 1, i, FG.Cols - 1) = 0
            End If

            FG.cell(flexcpData, i, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(i, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))

            RsDetails.MoveNext
        
            If FG.rows > 10 Then
                If i = 8 Then FG.Refresh
            End If

        Next i

        '----------------------------
        Me.LblInvProfit.Caption = FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("ItemProfit"), FG.rows - 1, FG.ColIndex("ItemProfit"))

        If val(Me.LblInvProfit.Caption) > 0 Then
            Me.LblInvProfit.ForeColor = &H4000&
        ElseIf val(Me.LblInvProfit.Caption) = 0 Then
            Me.LblInvProfit.ForeColor = vbBlack
        ElseIf val(Me.LblInvProfit.Caption) < 0 Then
            Me.LblInvProfit.ForeColor = vbRed
        End If

        '---------------------------
        FG.AutoSize 0, FG.Cols - 1, False
    End If
             mIsFinishSave = True
    XPChkPayType(0).value = Unchecked
    XPChkPayType(1).value = Unchecked
    XPChkPayType(2).value = Unchecked
    XPTxtValue(0).text = ""
    XPTxtValue(1).text = ""
    XPTxtSerial(0).text = ""
    XPTxtSerial(1).text = ""
    XPTxtValue(1).Tag = ""
    DtpDelayDate.value = Date
    '----------------------------------------------------------------------------------------
'    StrSQL = "Select * From Notes Where Transaction_ID=" & val(rs("Transaction_ID").value)
'    RsNotes.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
'     Set RsNotes = New ADODB.Recordset
'    StrSQL = "SELECT Notes.NoteID, Notes.NoteDate, Notes.NoteType, Notes.NoteSerial," & "Notes.Note_Value, Notes.BankID,BanksData.BankName , Notes.ChqueNum, Notes.DueDate "
'    StrSQL = StrSQL + " FROM Notes INNER JOIN BanksData ON Notes.BankID = BanksData.BankID "
'    StrSQL = StrSQL + " Where NoteType=2 AND NOTES.Transaction_ID=" & val(rs("Transaction_ID").value)
'    StrSQL = StrSQL + " Order BY Notes.NoteID"
'    RsNotes.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
  
    TxtFillData.text = "F"
    '-----------------------------------------------------------------------------------------------
    Dim SngRelatedNotesValues As Single
 '   Me.CmdNotes.Visible = ShowRelatedNotes(val(Me.XPTxtBillID.Text), 0, SngRelatedNotesValues)
 '   Me.CmdNotes.Tag = SngRelatedNotesValues

    SngRelatedNotesValues = 0
 '   Me.CmdRetruns.Visible = ShowRelatedTransactions(val(Me.XPTxtBillID.Text), 0, SngRelatedNotesValues)
 '   Me.CmdRetruns.Tag = SngRelatedNotesValues

    '-----------------------------------------------------------------------------------------------
    Screen.MousePointer = vbDefault
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
    Resume
    Screen.MousePointer = vbDefault
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
    Dim RsTest As ADODB.Recordset
    Dim StrSQL As String
    Dim IntRes As Integer
    Dim BegainTrans As Boolean
    On Error GoTo ErrTrap

    If XPTxtBillID.text = "" Then
        clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Sub
    End If

    If AvailableDeal = False Then
        Exit Sub
    End If

    '«·√ﬁ”«ÿ «·„”œœ… ⁄·Ï «·›« Ê—…
    If XPTxtValue(1).Tag <> "" Then
        StrSQL = "select * From ReceiptQestForBill Where NoteID=" & XPTxtValue(1).Tag
        Set RsTest = New ADODB.Recordset
        RsTest.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTest.EOF Or RsTest.BOF) Then
            Msg = "·ﬁœ  „  Õ’Ì· »⁄÷ «·√ﬁ”«ÿ «·„”Ã·… ⁄·Ï Â–Â «·›« Ê—…" & CHR(13)
            Msg = Msg + "Ê·« Ì„ﬂ‰ Õ–› »Ì«‰« Â«" & CHR(13)
            Msg = Msg + "≈–« ﬂ‰   —€» ›Ì Õ–› »Ì«‰«  Â–Â «·›« Ê—…" & CHR(13)
            Msg = Msg + "ÌÃ» Õ–› ⁄„·Ì«  «· Õ’Ì· «·Œ«’… »Â«"
            MsgBox Msg, vbOKOnly + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If
    End If

    '⁄„·Ì«  «·’Ì«‰… «·„— »ÿ… »«·›« Ê—…
    StrSQL = "select * From MaintenanceJuncTransaction Where Transaction_ID=" & Trim(XPTxtBillID.text)
    Set RsTest = New ADODB.Recordset
    RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTest.EOF Or RsTest.BOF) Then
        Msg = "·ﬁœ  „ ≈Ã—«¡ »⁄÷ ⁄„·Ì«  «·’Ì«‰… ⁄·Ï Â–Â «·›« Ê—… Ê·« Ì„ﬂ‰ Õ–›Â«"
        Msg = Msg + "≈–« ﬂ‰   —€» ›Ì Õ–› »Ì«‰«  Â–Â «·›« Ê—…" & CHR(13)
        Msg = Msg + "ÌÃ» Õ–› ⁄„·Ì«  «·’Ì«‰… «·Œ«’… »Â«"
        MsgBox Msg, vbOKOnly + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If Me.CboPayMentType.ListIndex = 0 Then

        '›« Ê—… ‰ﬁœÌ…
        If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtValue(0).text), XPDtbBill.value, False) = False Then
            Msg = "·‰ Ì„ﬂ‰ «·”„«Õ »Õ–› Â–« «·⁄„·Ì…..!!!"
            Msg = Msg & CHR(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï Õ”«»«  «·Œ“‰…"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If
    End If

    Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
    Msg = Msg + (TxtTransSerial.text) & CHR(13)
    Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
    IntRes = MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)

    If IntRes = vbYes Then
        If Not rs.RecordCount < 1 Then
            Cn.BeginTrans
            BegainTrans = True
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS  " & "Where DOUBLE_ENTREY_VOUCHERS.Transaction_ID=" & rs("Transaction_ID").value
            Cn.Execute StrSQL, , adExecuteNoRecords
            If TxtWorkOrderNO.text <> "" Then
                 Cn.Execute " Update Transactions set FlgProductOrder=null   WHERE     (dbo.Transactions.Transaction_Type = 26) AND (dbo.Transactions.Transaction_Serial = N'" & TxtWorkOrderNO.text & "')"
             End If
            StrSQL = "delete From Notes where noteid=" & val(TXTNoteID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            CuurentLogdata ("D")
            rs.delete
            Cn.CommitTrans
            BegainTrans = False
            Msg = " „  ⁄„·Ì… «·Õ–› "
            MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            rs.MoveFirst

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
    Dim BolRtl As Boolean

    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… »Ì⁄ ÃœÌœ…" & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F12 OR Enter", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…" & Wrap & "„›« ÌÕ «·«Œ ’«— F6", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  ⁄„·Ì… «·»Ì⁄" & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F11", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  ⁄„·Ì… «·»Ì⁄ «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F10", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·»Ì⁄" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F9", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  ⁄„·Ì… »Ì⁄" & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F8", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄„·Ì… »Ì⁄" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F7", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— Ctrl + X", BolRtl
        End With
    
        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnNewClients, "≈÷«›… ⁄„Ì· ÃœÌœ ..." & Wrap & "· ”ÃÌ· »Ì«‰«  ⁄„Ì· ÃœÌœ" & Wrap & " «÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F5", BolRtl
        End With
    
        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, BolRtl
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        BolRtl = False

        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "New..." & Wrap & "Click here to add new Bill Invoice" & Wrap & "" & Wrap & "Shortcut (Enter Or F12)", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "Print..." & Wrap & "Print this Bill Invoice" & Wrap & "" & Wrap & "Shortcut (F6)", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), "Edit..." & Wrap & "Edit this Bill Invoice Record" & Wrap & "  " & Wrap & "Shortcut (F11)", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Save..." & Wrap & "Save the New Bill Invoice Or Save the edit" & Wrap & "in the current Bill Invoice" & Wrap & "" & Wrap & "Shortcut (F10)", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), "Undo..." & Wrap & "Undo in the New Bill Invoice" & Wrap & "Or Undo in the Editing" & Wrap & "" & Wrap & "Shortcut (F9)", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Delete..." & Wrap & "Delete this current Bill Invoice" & Wrap & "" & Wrap & "Shortcut (F8)", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "Search..." & Wrap & "Click here to display the search" & Wrap & "Screen" & Wrap & "Shortcut (F7)", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Exit..." & Wrap & "Close this Window", BolRtl
        End With
    
        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnNewClients, "Add New Customer...." & Wrap & "To add New Customer Click here..." & Wrap & "Shortcut (F5)", BolRtl
        End With
    
        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "First..." & Wrap & "Move to first Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "Previous..." & Wrap & "Move to Previous Record" & Wrap & " , BolRTL"
        End With

        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "Next..." & Wrap & "Move to Next Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "Last..." & Wrap & "Move to Last Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "Help..." & Wrap & "to View Help Files" & Wrap & "click Here" & Wrap & "Shortcut(F1)" & Wrap, BolRtl
        End With

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData(Optional ByVal IsSaveWithOutMsg As Boolean = False, _
                     Optional ByVal fromResave As Boolean = False)
    Dim Msg            As String
    Dim RowNum         As Integer
    Dim RSTransDetails As ADODB.Recordset
    Dim RsNotes        As ADODB.Recordset
    Dim RsTemp         As New ADODB.Recordset
    Dim RsTest         As New ADODB.Recordset
    Dim RsRepeat       As ADODB.Recordset
    Dim RsDetalis      As ADODB.Recordset
    Dim StrSQL         As String
    Dim StrSqlDel      As String
    Dim note_id        As Long
    Dim TransBegine    As Boolean
    Dim BolTemp        As Boolean
    Dim LnItemID       As Long
    Dim i              As Integer
    Dim DblNotesTotal  As Double
    Dim SngTemp        As Variant
    '****************************
    '· Ã«Â· Õ›Ÿ «· ›«’Ì· „⁄ «⁄«œÂ Ÿ»ÿ «·Õ—ﬂ« 
    Dim mSaveDetails   As Boolean
    mSaveDetails = (fromResave And chkIgnorDetails.value = 1) Or Not fromResave
    '***********************
    On Error GoTo ErrTrap

    Me.FG.FinishEditing True

    DoEvents

    Screen.MousePointer = vbArrowHourglass
    If IsSaveWithOutMsg Then GoTo SaveDirect
    If Trim(Me.TxtTransSerial.text) = "" Then
        Msg = "ÌÃ» ≈œŒ«· —ﬁ„ «·”‰œ...!!"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtTransSerial.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    
    Else

        If Me.TxtModFlg.text = "N" Then
            BolTemp = UniqueTransSerial(Trim(Me.TxtTransSerial.text), 27, , val(Me.dcBranch.BoundText))
        ElseIf Me.TxtModFlg.text = "E" Then
            BolTemp = UniqueTransSerial(Trim(Me.TxtTransSerial.text), 27, val(Me.XPTxtBillID.text), val(Me.dcBranch.BoundText))
        End If

    End If
    
    If DCboStoreName.text = "" Then
        Msg = "ÌÃ»  ÕœÌœ «·„Œ“‰"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        DCboStoreName.SetFocus
        Sendkeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
 
    '----------------------------------------------
    If val(Me.XPTxtValue(1).text) > 0 Then
        If ChkInstall.value = vbChecked Then
            If val(Me.LblInstallTotal.Caption) = 0 Then
                Msg = "ÌÃ» Õ”«» «·√ﬁ”«ÿ ﬁ»· ⁄„·Ì… «·Õ›Ÿ..!!!"
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Me.XPTab301.CurrTab = 1
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            If val(Me.LblInstallTotal.Caption) <> val(Me.XPTxtValue(1).text) Then
                Me.XPTxtValue(1).text = val(Me.LblInstallTotal.Caption)
            End If
        End If
    End If

    '-----------------------------------------
    If XPChkPayType(2).value = vbChecked Then
        If val(Me.lbl(18).Caption) = 0 Then
            Msg = "ÌÃ» ≈œŒ«· «·‘Ìﬂ«  ﬁ»· ⁄„·Ì… «·Õ›Ÿ..!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Me.XPTab301.CurrTab = 1
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If

    If XPChkTAX.value = Checked Then
        If XPTxtTaxValue.text = "" Then
            Msg = "ÌÃ» «œŒ«· ﬁÌ„… ÷—Ì»… «·„»Ì⁄« "
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPTxtTaxValue.SetFocus
            FG.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If

    If XPCboDiscountType.ListIndex = 1 Or XPCboDiscountType.ListIndex = 2 Then
        If XPTxtDiscountVal.text = "" Then
            Msg = "≈–« ﬂ«‰ Â‰«ﬂ Œ’„ ⁄·Ï «·›« Ê—… " & CHR(13)
            Msg = Msg + "ÌÃ»  ÕœÌœ ﬁÌ„… Â–« «·Œ’„ " & CHR(13)
            Msg = Msg + "√Ê √Œ Ì«— ·« ÌÊÃœ Œ’„ "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPCboDiscountType.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If

    '--------------------------------
    '«·ﬂ‘› ⁄·Ï „œÌÊ‰Ì… «·⁄„Ì· '
    '    If val(Me.DBCboClientName.BoundText) <> 1 Or val(Me.DBCboClientName.BoundText <> 2) Then
    '        If Me.CboPaymentType.ListIndex = 1 Then
    '            If val(Me.XPTxtValue(1).text) > 0 Then
    '                If CheckCusCredit(val(Me.DBCboClientName.BoundText), val(Me.XPTxtValue(1).text), 0) = False Then
    '                    Screen.MousePointer = vbDefault
    '                    Exit Sub
    '                End If
    '            End If
    '        End If
    '    End If

    '--------------------------------
    Me.XPTab301.CurrTab = 0

    If NewGrid.CheckDataEntered = False Then
        Exit Sub
    End If

    '-------------------------------
    If NewGrid.Calculate(1, True, False, True) = False Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '-------------------------------
    If Me.XPChkPayType(0).value = vbChecked Then
        DblNotesTotal = val(Me.XPTxtValue(0).text)
    End If

    If Me.XPChkPayType(1).value = vbChecked Then
        DblNotesTotal = DblNotesTotal + val(Me.XPTxtValue(1).text)
    End If

    If Me.XPChkPayType(2).value = vbChecked Then
        DblNotesTotal = DblNotesTotal + val(Me.lbl(18).Caption)
    End If

    '---------------------------------
    Dim RsNotesGeneral As ADODB.Recordset
    Dim Vchr_result    As String
    Dim notes_result   As String
    my_branch = val(Me.dcBranch.BoundText)
             
    If TxtNoteSerial1.text = "" Then
        Vchr_result = Voucher_coding(val(my_branch), XPDtbBill.value, 18, 240, , 27, , val(DCboStoreName.BoundText))

        If Vchr_result = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ’—› „Œ“‰Ì ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
        Else
                       
            If Vchr_result = "" Then
                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
            Else
                '         txtNoteSerial1.text = Vchr_result
            End If
        End If
    End If
                    
    If TxtNoteSerial.text = "" Then
        notes_result = Notes_coding(val(my_branch), XPDtbBill.value)

        If notes_result = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
        Else
                       
            If notes_result = "" Then
                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
            Else
                '         TxtNoteSerial.text = notes_result
            End If
        End If
    End If
SaveDirect:
    Set RsNotesGeneral = New ADODB.Recordset
    'RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
    RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
    If Me.TxtModFlg.text = "N" Then
        XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
        Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
    Else
        '
        StrSqlDel = "delete From Notes where noteid=" & val(TXTNoteID.text)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        
        general_noteid = val(TXTNoteID.text)
    End If

    RsNotesGeneral.AddNew
    RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
    general_noteid = RsNotesGeneral("NoteID").value
    TXTNoteID.text = general_noteid
    ' RsNotesGeneral("Transaction_ID").value = Val(XPTxtBillID.text)
    RsNotesGeneral("NoteDate").value = XPDtbBill.value
    RsNotesGeneral("NoteType").value = 240 ' «–‰ «÷«›…
    RsNotesGeneral("Note_Value").value = val(LblTotal.Caption)

    If TxtNoteSerial.text = "" Then
        TxtNoteSerial = Notes_coding(val(my_branch), XPDtbBill.value)
    End If
        
    RsNotesGeneral("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))

    If TxtNoteSerial1.text = "" Then
        TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbBill.value, 18, 240, , 27, , val(DCboStoreName.BoundText))
    End If
          
    '    RsNotesGeneral("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.Text) = "", Null, Trim(Me.TxtNoteSerial1.Text))
    RsNotesGeneral("REMARK").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))

    RsNotesGeneral("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '
        
    RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
    RsNotesGeneral("numbering_type1").value = sand_numbering_type(18) '  «–‰ ’—›
    RsNotesGeneral("sanad_year").value = year(XPDtbBill.value)
    RsNotesGeneral("sanad_month").value = Month(XPDtbBill.value)
    'RsNotes("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
    RsNotesGeneral("branch_no").value = val(Me.dcBranch.BoundText)
    RsNotesGeneral.update
        
    Set RSTransDetails = New ADODB.Recordset
    '  RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    '    Set RsNotes = New ADODB.Recordset
    '  RsNotes.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    '     StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
    '   RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
    StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
    RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If SystemOptions.SysRegisterState <> Registered And SystemOptions.SysRegisterState <> DevelopVersion Then
        If rs.RecordCount > 50 Then
            'Exit Sub
        End If
    End If

    Screen.MousePointer = vbArrowHourglass
    Cn.BeginTrans
    TransBegine = True

    If Me.TxtModFlg.text = "N" Then
        
        rs.AddNew
        XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
        rs("Transaction_ID").value = val(XPTxtBillID.text)
    ElseIf Me.TxtModFlg.text = "E" Then
        If mSaveDetails Then
            StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
        End If
        StrSqlDel = "delete From Notes where Transaction_ID=" & val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        'StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & val(Me.XPTxtBillID.Text)
        'Cn.Execute StrSQL, , adExecuteNoRecords
    End If

    rs("Doctype").value = IIf(Me.DCDocTypes.BoundText = "", Null, val(DCDocTypes.BoundText))

    rs("remark").value = IIf(Trim(Me.txtremark.text) = "", Null, Trim(Me.txtremark.text))
    rs("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
    rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
    rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
    rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '
    rs("ProductionOrderID").value = val(Me.txtProductionOrderID.text) '

    rs("NoteId").value = val(TXTNoteID.text)
    rs("Transaction_Serial").value = IIf(Trim(Me.TxtTransSerial.text) = "", "", Trim(Me.TxtTransSerial.text))
    rs("Transaction_Date").value = XPDtbBill.value
    rs("Transaction_Type").value = 27
    rs("UserID").value = user_id

    rs("Nots").value = Me.Text2.text
    rs("nots2").value = Txtnots2.text
    rs("WorkOrderNO").value = TxtWorkOrderNO.text
    
    rs("MixID").value = val(txtMixID.text)
    rs("MIxCode").value = txtMIxCode.text
    
    Dim rs2 As New ADODB.Recordset
    '           rs2.Close
    '  rs2.Open "select * from Transactions where Transaction_Serial = " & TxtTransSerial.Text & " and Transaction_type = 21", Cn, adOpenDynamic, adLockOptimistic

    '  If Not rs2.EOF Then
    '      rs2("Nots2").value = Me.Text2.Text & ""
    '      rs2.update
    '      rs2.Close
    '  End If
    
    If XPCboDiscountType.ListIndex = -1 Then
        rs("Trans_DiscountType").value = 0
    Else
        rs("Trans_DiscountType").value = val(XPCboDiscountType.ListIndex)
    End If

    rs("Trans_Discount").value = IIf(XPTxtDiscountVal.text = "", Null, val(XPTxtDiscountVal.text))
    rs("CusID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
    rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, val(DCboStoreName.BoundText))

    If CboPayMentType.ListIndex = -1 Then
        rs("PaymentType").value = 0
    Else
        rs("PaymentType").value = val(CboPayMentType.ListIndex)
    End If

    rs("TaxFound").value = IIf(XPChkTAX.value = Checked, True, False)
    rs("TaxValue").value = IIf(XPTxtTaxValue.text = "", Null, val(XPTxtTaxValue.text))
    rs("Emp_ID").value = IIf(DcboEmp.BoundText = "", Null, DcboEmp.BoundText)

    If Me.CboSaleType.ListIndex = 0 Or Me.CboSaleType.ListIndex = -1 Then
        rs("SaleType").value = 0
    Else
        rs("SaleType").value = 1
    End If

    If Trim$(Me.TxtCashCustomerName.text) <> "" Then
        rs("CashCustomerName").value = Trim$(Me.TxtCashCustomerName.text)
    Else
        rs("CashCustomerName").value = Null
    End If

    rs("TransactionComment").value = IIf(Trim$(TxtBillComment.text) = "", Null, Trim$(TxtBillComment.text))

    '÷—»Ì… Œ’„ Ê≈÷«›…
    If ChkTaxAdd.value = vbChecked And val(Me.TxtTaxAddValue.text) > 0 Then
        rs("TaxAddValue").value = val(Me.TxtTaxAddValue.text)
    Else
        rs("TaxAddValue").value = 0
    End If

    '÷—»Ì… œ„€…
    If ChkTaxStamp.value = vbChecked And val(Me.TxtTaxStampValue.text) > 0 Then
        rs("TaxStampValue").value = val(Me.TxtTaxStampValue.text)
    Else
        rs("TaxStampValue").value = 0
    End If

    '÷—»Ì… Œœ„…
    If ChkTaxSerivce.value = vbChecked And val(Me.TxtTaxServiceValue.text) > 0 Then
        rs("TaxServiceValue").value = val(Me.TxtTaxServiceValue.text)
    Else
        rs("TaxServiceValue").value = 0
    End If

    rs.update

    CuurentLogdata
    If mSaveDetails Then
        For RowNum = 1 To FG.rows - 1

            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then

                'Check Repeat Serial
                If FG.TextMatrix(RowNum, FG.ColIndex("Serial")) <> "" Then
                    StrSQL = "select * From Transaction_Details where ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                    StrSQL = StrSQL + " and Transaction_ID =" & XPTxtBillID.text
                    Set RsTemp = New ADODB.Recordset
                    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (RsTemp.EOF Or RsTemp.BOF) Then
                        Msg = "«·”Ì—Ì«· «·Œ«’ »«·’‰›" & CHR(13)
                        Msg = Msg + FG.cell(flexcpTextDisplay, RowNum, FG.ColIndex("name")) & CHR(13)
                        Msg = Msg + " „ √œŒ«·Â ·ﬁÿ⁄… √Œ—Ï ›Ì Â–Â «·›« Ê—…"
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        RsTemp.Close
                        XPTab301.CurrTab = 0
                        FG.row = RowNum
                        FG.Col = FG.ColIndex("name")
                        FG.ShowCell RowNum, FG.ColIndex("name")
                        FG.SetFocus
                
                        TransBegine = False
                        Cn.RollbackTrans
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If

                    RsTemp.Close
                End If

                If IsEmpty(Me.FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))) Then
                    If val(Me.FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))) = 0 Then
                        Msg = " ÌÃ»  ÕœÌœ ÊÕœ… «·ﬂ„Ì… «·Œ«’… »«·’‰›" & CHR(13)
                        Msg = Msg + FG.cell(flexcpTextDisplay, RowNum, FG.ColIndex("Name")) & CHR(13)
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        XPTab301.CurrTab = 0
                        FG.row = RowNum
                        FG.Col = FG.ColIndex("UnitID")
                        FG.ShowCell RowNum, FG.ColIndex("UnitID")
                        FG.SetFocus
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                End If

                RSTransDetails.AddNew
                RSTransDetails("order_no").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("order_no")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("order_no")))
                RSTransDetails("OrderArrivalDate").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")))
                RSTransDetails("FoxyNo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")))
                RSTransDetails("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
                RSTransDetails("Transaction_ID").value = val(XPTxtBillID.text)
                RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
                RSTransDetails("ItemID2").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemID2")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemID2"))))

                'RSTransDetails("Quantity").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
                '            RSTransDetails("ItemName").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Name")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Name"))))
        
                RSTransDetails("ItemDiscountType").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType"))))
                RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
                RSTransDetails("ItemDiscount").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal"))))
                RSTransDetails("guaranteeTime").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime"))))
                RSTransDetails("CostPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice"))))
                RSTransDetails("CostTransID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("PofTransID")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("PofTransID"))))
                RSTransDetails("ItemProfit").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemProfit")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemProfit"))))
        
                RSTransDetails("UnitID").value = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
                RSTransDetails("ShowQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
                Dim cnt As Double
                cnt = FG.TextMatrix(RowNum, FG.ColIndex("Count"))

                RSTransDetails("NProductionOrderNO").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("NProductionOrderNO")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("NProductionOrderNO"))))

                RSTransDetails("Remarks").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Remarks")) = ""), Null, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("Remarks"))))
                
                RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
                RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
        
                RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
   
                RSTransDetails("showPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
      
                '«·ÊÕœ« 
           
                Dim RsUnitData   As ADODB.Recordset
                Dim LngCurItemID As Long
                Dim LngUnitID    As Long
                Dim DblQty       As Double
        
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
            
                End If

                RSTransDetails("Price").value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
       
                SngTemp = SngTemp + (val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice"))) * RSTransDetails("quantity").value)

                Dim OldQty  As Double
                Dim OldCost As Double
                Dim NewQty  As Double
                Dim NewCost As Double
               
              '  getItemCostData XPDtbBill.value, RSTransDetails("Item_ID").value, val(DCboStoreName.BoundText), val(Me.XPTxtBillID.text), OldQty, OldCost, NewQty, NewCost, , LngUnitID
       
              '  RSTransDetails("OldQty").value = NewQty
              '  RSTransDetails("OldCost").value = NewCost
       
               ' RSTransDetails("NewQty").value = RSTransDetails("OldQty").value - RSTransDetails("Quantity").value
              '  RSTransDetails("NewCost").value = RSTransDetails("OldCost").value ' ((RSTransDetails("OldQty").value * RSTransDetails("OldCost").value) + (RSTransDetails("Quantity").value * RSTransDetails("Price").value)) / (RSTransDetails("Quantity").value + RSTransDetails("OldQty").value)
       
                RSTransDetails.update
                '-------------
        
            End If

        Next RowNum
    End If
    'If Me.XPChkPayType(0).Value = Checked Then
    '    RsNotes.AddNew
    '    RsNotes("NoteID").Value = CStr(new_id("Notes", "NoteID", "", True))
    '    If Me.TxtModFlg.text = "N" Then
    '        RsNotes("NoteSerial").Value = CStr(new_id("Notes", "NoteSerial", "", True))
    '        XPTxtSerial(0).text = RsNotes("NoteSerial").Value
    '    ElseIf Trim(XPTxtSerial(0).text) <> "" Then
    '        RsNotes("NoteSerial").Value = Trim(XPTxtSerial(0).text)
    '    Else
    '        RsNotes("NoteSerial").Value = CStr(new_id("Notes", "NoteSerial", "", True))
    '        XPTxtSerial(0).text = RsNotes("NoteSerial").Value
    '    End If
    '    RsNotes("Transaction_ID").Value = Val(XPTxtBillID.text)
    '    RsNotes("NoteDate").Value = XPDtbBill.Value
    '    RsNotes("NoteType").Value = 0
    '    RsNotes("Note_Value").Value = IIf(XPTxtValue(0).text = "", Null, Val(XPTxtValue(0).text))
    '    RsNotes("Member_ID").Value = IIf(DBCboClientName.BoundText = "", Null, Val(DBCboClientName.BoundText))
    '    RsNotes("BankID").Value = Null
    '    RsNotes("BoxID").Value = IIf(DcboBox.BoundText = "", Null, Val(DcboBox.BoundText))
    '    RsNotes("CUSID").Value = Null
    '
    '    RsNotes.update
    'End If
    ''«·ﬁÌ„ «·¬Ã·…
    'If Me.XPChkPayType(1).Value = Checked Then
    '    RsNotes.AddNew
    '    RsNotes("NoteID").Value = CStr(new_id("Notes", "NoteID", "", True))
    '    XPTxtValue(1).Tag = IIf(IsNull(RsNotes("NoteID").Value), "", (RsNotes("NoteID").Value))
    '    Note_ID = RsNotes("NoteID").Value
    '    RsNotes("NoteDate").Value = XPDtbBill.Value
    '    If Me.TxtModFlg.text = "N" Then
    '        RsNotes("NoteSerial").Value = CStr(new_id("Notes", "NoteSerial", "", True))
    '        XPTxtSerial(1).text = RsNotes("NoteSerial").Value
    '    ElseIf Trim(XPTxtSerial(1).text) <> "" Then
    '        RsNotes("NoteSerial").Value = Trim(XPTxtSerial(1).text)
    '    Else
    '        RsNotes("NoteSerial").Value = CStr(new_id("Notes", "NoteSerial", "", True))
    '        XPTxtSerial(1).text = RsNotes("NoteSerial").Value
    '    End If
    '    RsNotes("Transaction_ID").Value = Val(XPTxtBillID.text)
    '    RsNotes("NoteType").Value = 1
    '    RsNotes("Note_Value").Value = IIf(XPTxtValue(1).text = "", Null, Val(XPTxtValue(1).text))
    '    RsNotes("Member_ID").Value = IIf(DBCboClientName.BoundText = "", Null, Val(DBCboClientName.BoundText))
    '    RsNotes("BankID").Value = Null
    '    RsNotes("CUSID").Value = Null
    '    RsNotes("DueDate").Value = DtpDelayDate.Value
    '    RsNotes.update
    'End If
    'If Me.XPChkPayType(2).Value = Checked Then
    '    With Me.FgCheques
    '        For I = .FixedRows To .Rows - 1
    '            RsNotes.AddNew
    '                RsNotes("NoteID").Value = CStr(new_id("Notes", "NoteID", "", True))
    '                RsNotes("NoteDate").Value = XPDtbBill.Value
    '                RsNotes("Transaction_ID").Value = Val(XPTxtBillID.text)
    '                RsNotes("NoteType").Value = 2
    '                RsNotes("Note_Value").Value = Val(.TextMatrix(I, .ColIndex("CheckValue")))
    '                RsNotes("BankID").Value = Val(.TextMatrix(I, .ColIndex("BankID")))
    '                RsNotes("ChqueNum").Value = Trim$(.TextMatrix(I, .ColIndex("CheckNumber")))
    '                RsNotes("DueDate").Value = CDate(Trim$(.TextMatrix(I, .ColIndex("DueDate"))))
    '                RsNotes("Member_ID").Value = Val(Me.DBCboClientName.BoundText)
    '                RsNotes("CUSID").Value = Val(Me.DBCboClientName.BoundText)
    '            RsNotes.update
    '        Next I
    '    End With
    'End If
    ''Õ›Ÿ «·√›”«ÿ
    'If Me.XPChkPayType(1).Value = Checked Then
    '    If ChkInstall.Value = vbChecked Then
    '        'Save installment Data
    '        Set RsTemp = New ADODB.Recordset
    '        RsTemp.Open "InstallMent", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    '        RsTemp.AddNew
    '            RsTemp("PartID").Value = CStr(new_id("InstallMent", "PartID", "", True))
    '            RsTemp("NoteID").Value = Note_ID
    '            RsTemp("BasicAmmount").Value = IIf(XPTxtValue(1).text = "", 0, Val(XPTxtValue(1).text))
    '            RsTemp("InterestType").Value = Val(Me.LblPrecenType.Tag)
    '            RsTemp("InterestVal").Value = Val(LblPrecenValue.Caption)
    '            RsTemp("Total").Value = Val(LblInstallTotal.Caption)
    '            RsTemp("InstallCount").Value = Val(LblInstallCount.Caption)
    '            RsTemp("FirstInstallDate").Value = CDate(Me.LblFirstInstallDate.Caption)
    '            If Val(LblInstallmentType.Tag) = 0 Then
    '                RsTemp("InstallmentType").Value = 0
    '            ElseIf Val(LblInstallmentType.Tag) = 1 Then
    '                RsTemp("InstallmentType").Value = 1
    '            ElseIf Val(LblInstallmentType.Tag) = 2 Then
    '                RsTemp("InstallmentType").Value = 2
    '            End If
    '            RsTemp("InstallSeprator").Value = Val(Me.LblInstallSeprator.Caption)
    '            RsTemp("StartValue").Value = IIf(Val(Me.LblStartValue.Caption) = 0, Null, Val(Me.LblStartValue.Caption))
    '            RsTemp("CustID").Value = IIf(DBCboClientName.BoundText = "", Null, Val(DBCboClientName.BoundText))
    '            RsTemp("Type").Value = 0
    '        RsTemp.update
    '        'save installment Details
    '        Set RsDetalis = New ADODB.Recordset
    '        RsDetalis.Open "InstallMentDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    '        With Me.FgInstallments
    '            For RowNum = 1 To .Rows - 1
    '                RsDetalis.AddNew
    '                    RsDetalis("QestID").Value = CStr(new_id("InstallMentDetails", "QestID", "", True))
    '                    RsDetalis("PartID").Value = RsTemp("PartID").Value
    '                    RsDetalis("QeqtNum").Value = IIf(.TextMatrix(RowNum, .ColIndex("Serial")) = "", "", .TextMatrix(RowNum, .ColIndex("Serial")))
    '                    RsDetalis("Value").Value = IIf(.TextMatrix(RowNum, .ColIndex("Value")) = "", "", Val(.TextMatrix(RowNum, .ColIndex("Value"))))
    '                    RsDetalis("DueDate").Value = IIf(.TextMatrix(RowNum, .ColIndex("Due_Date")) = "", "", .TextMatrix(RowNum, .ColIndex("Due_Date")))
    '                    RsDetalis("Receipt").Value = False
    '
    '                RsDetalis.update
    '            Next RowNum
    '        End With
    '    End If
    'End If

    Dim LngDevID           As Long
    Dim LngDevNO           As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes         As String
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '----------------
    Dim Account_Code_dynamic As String
    'SngTemp = NewGrid.GetItemsCostTotal * RSTransDetails("quantity").value / Cnt
    SngTemp = NewGrid.GetItemsTotal(ItemsGoodType) 'ﬁÌœ

    If SngTemp > 0 Then
        '1 work with branch
        '2 work with inventory
        '3 work with groups

        If detect_inventory_work_type = 1 Then
            Account_Code_dynamic = get_account_code_branch(37, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»     „’«—Ì› «‰ «Ã „Ê«œ  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    GoTo ErrTrap
         
                End If
            End If

            StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄«  1

            'StrTempAccountCode = "a3a2" ' ﬂ·›… «·„»Ì⁄« 
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "√–‰ ’—›  —ﬁ„ " & Me.TxtNoteSerial1.text
            Else
                StrTempDes = "Issue Voucher No. " & Me.TxtNoteSerial1.text
            End If

            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
    
            '«·„Œ“Ê‰ ›Ì «·›—⁄
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
                StrTempDes = "√–‰ ’—› „Ê«œ ··«‰ «Ã  —ﬁ„ " & Me.TxtNoteSerial1.text
            Else
                StrTempDes = "Production Issue Voucher No. " & Me.TxtNoteSerial1.text
            End If
    
            If TxtWorkOrderNO.text <> "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = StrTempDes & "»‰«¡ ·√„— «‰ «Ã —ﬁ„  " & TxtWorkOrderNO.text
                Else
                    StrTempDes = StrTempDes & "To PO No " & TxtWorkOrderNO.text
                End If
            End If
    
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
    
        ElseIf detect_inventory_work_type = 2 Then
            Account_Code_dynamic = get_account_code_branch(37, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» „’«—Ì› «‰ «Ã „Ê«œ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    GoTo ErrTrap
         
                End If
            End If

            StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄«  1

            Dim usedaccount    As Integer
            Dim UseCustomerAcc As Integer
            If val(DCDocTypes.BoundText) > 0 Then
        
                getDocAccounts val(DCDocTypes.BoundText), StrTempAccountCode, , , , , usedaccount, , , , , UseCustomerAcc
        
                If StrTempAccountCode = "" And usedaccount = 1 Then
                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·„œÌ‰ ··”‰œ", vbCritical
                    GoTo ErrTrap
                ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
                
                ElseIf usedaccount = 0 And UseCustomerAcc = 0 Then
                
                    StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄«  1
                ElseIf usedaccount = 0 And UseCustomerAcc = 1 Then
                 
                    StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
                       
                End If

            Else
                StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄«  1
            End If

            'StrTempAccountCode = "a3a2" ' ﬂ·›… «·„»Ì⁄« 
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "√–‰ ’—›  „Ê«œ ··«‰ «Ã —ﬁ„ " & Me.TxtNoteSerial1.text
            Else
                StrTempDes = "Production Issue Voucher No. " & Me.TxtNoteSerial1.text
            End If
    
            If TxtWorkOrderNO.text <> "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = StrTempDes & "»‰«¡ ·√„— «‰ «Ã —ﬁ„  " & TxtWorkOrderNO.text
                Else
                    StrTempDes = StrTempDes & "To PO No " & TxtWorkOrderNO.text
                End If
            End If
    
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

            '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
    
            Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")

            If Account_Code_dynamic = "" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                GoTo ErrTrap
            End If
    
            StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰

            ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "√–‰ ’—›  —ﬁ„ " & Me.TxtNoteSerial1.text
            Else
                StrTempDes = "Issue Voucher No. " & Me.TxtNoteSerial1.text
            End If
    
            If TxtWorkOrderNO.text <> "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = StrTempDes & "»‰«¡ ·√„— «‰ «Ã —ﬁ„  " & TxtWorkOrderNO.text
                Else
                    StrTempDes = StrTempDes & "To PO No " & TxtWorkOrderNO.text
                End If
            End If
    
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value   As Single

            With FG

                For i = 1 To FG.rows - 1

                    If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                        ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                        groupAccount = get_item_group_account_in_branch(FG.TextMatrix(i, FG.ColIndex("Code")), val(my_branch), 1)

                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»   ﬂ·›… ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "√–‰ ’—›  —ﬁ„ " & Me.TxtNoteSerial1.text
                        Else
                            StrTempDes = "Issue Voucher No. " & Me.TxtNoteSerial1.text
                        End If
    
                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

                Next i

            End With

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

                        line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "√–‰ ’—›  —ﬁ„ " & Me.TxtNoteSerial1.text
                        Else
                            StrTempDes = "Issue Voucher No. " & Me.TxtNoteSerial1.text
                        End If

                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

                Next i

            End With

        End If

        '----------------
        'LngDevID = LngDevID + 1
        'LngDevNO = 0
    End If

    'If Me.XPChkPayType(0).value = vbChecked Then
    '    StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", Val(Me.DcboBox.BoundText))
    '    StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtTransSerial.text
    '    LngDevNO = LngDevNO + 1
    '    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Val(Me.XPTxtValue(0).text), _
    '        0, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
    '        GoTo ErrTrap
    '    End If
    'End If
    'If Me.XPChkPayType(1).Value = vbChecked Then
    '    StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", Val(Me.DBCboClientName.BoundText))
    '    StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtTransSerial.text
    '    LngDevNO = LngDevNO + 1
    '    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Val(Me.LblTotalAll.Caption), _
    '        0, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
    '        GoTo ErrTrap
    '    End If
    'End If
    If Me.XPChkPayType(2).value = vbChecked Then
        '   StrTempAccountCode = "a1a2a4" '«Ê—«ﬁ ﬁ»÷
        '   StrTempDes = "⁄œœ " & Me.lbl(19).Caption & "  ‘Ìﬂ«  " & Chr(13)
        '   StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtTransSerial.text
        '   LngDevNO = LngDevNO + 1
        '   If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Val(Me.lbl(18).Caption), _
        '       0, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        '       GoTo ErrTrap
        '   End If
    End If

    If val(Me.LblDiscountsTotal.Caption) > 0 Then
        '
        '        Account_Code_dynamic = get_account_code_branch(12, my_branch)
        '        If Account_Code_dynamic = "NO branch" Then
        '        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        '        GoTo ErrTrap
        '        Else
        '        If Account_Code_dynamic = "NO account" Then
        '           MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··Œ’„ «·„”„ÊÕ »Â ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
        '        GoTo ErrTrap
         
        '        End If
        '        End If
        '    StrTempAccountCode = Account_Code_dynamic '«·Œ’„ «·„”„ÊÕ »Â 12
        '    'StrTempAccountCode = "a3a5" '«·Œ’„ «·„”„ÊÕ »Â
        '    StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtTransSerial.text
        '    LngDevNO = LngDevNO + 1
        '    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Val(Me.LblDiscountsTotal.Caption), _
        '        0, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        '        GoTo ErrTrap
        '    End If
    End If

    If Me.ChkTaxAdd.value = vbChecked Then
        '   StrTempAccountCode = "a2a5a4" '÷—»Ì… √—»«Õ  Ã«—Ì… (Œ’„ Ê≈÷«›…
        '   StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtTransSerial.text
        '   SngTemp = Val(Me.lbl(52).Caption)
        '   LngDevNO = LngDevNO + 1
        '   If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
        '       0, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        '       GoTo ErrTrap
        '   End If
    End If

    If Me.ChkTaxStamp.value = vbChecked Then
        '   StrTempAccountCode = "a3a9" 'œ„€«  ÕﬂÊ„Ì…
        '   StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtTransSerial.text
        '   SngTemp = Val(Me.lbl(53).Caption)
        '   LngDevNO = LngDevNO + 1
        '   If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
        '       0, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        '       GoTo ErrTrap
        '   End If
    End If

    '«·œ«∆‰
    'SngTemp = NewGrid.GetItemsTotal(ItemsGoodType)
    'If SngTemp > 0 Then
    '
    '        Account_Code_dynamic = get_account_code_branch(2, my_branch)
    '        If Account_Code_dynamic = "NO branch" Then
    '        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
    '        GoTo ErrTrap
    '        Else
    '        If Account_Code_dynamic = "NO account" Then
    '           MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„»Ì⁄«  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
    '        GoTo ErrTrap
    '
    '        End If
    '        End If
    '    StrTempAccountCode = Account_Code_dynamic '«·„»Ì⁄« 2
    ' '   StrTempAccountCode = "a4a1" '«·„»Ì⁄« 
    '    StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtTransSerial.text
    '    LngDevNO = LngDevNO + 1
    '    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
    '        1, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
    '        GoTo ErrTrap
    '    End If
    'End If
    'SngTemp = NewGrid.GetItemsTotal(ItemsServiceType)
    'If SngTemp > 0 Then
    '        Account_Code_dynamic = get_account_code_branch(23, my_branch)
    '        If Account_Code_dynamic = "NO branch" Then
    '        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
    '        GoTo ErrTrap
    '        Else
    '        If Account_Code_dynamic = "NO account" Then
    '           MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «Ì—«œ«  «·Œœ„«   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
    '        GoTo ErrTrap
    '
    '        End If
    '        End If
    '    StrTempAccountCode = Account_Code_dynamic '≈Ì—«œ«  «·Œœ„« 23
    '  '  StrTempAccountCode = "a4a7" '≈Ì—«œ«  «·Œœ„« 
    '
    '    StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtTransSerial.text
    '    LngDevNO = LngDevNO + 1
    '    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
    '        1, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
    '        GoTo ErrTrap
    '    End If
    'End If
    '
    If XPChkTAX.value = vbChecked Then
        'StrTempAccountCode = "a1a3a5" '÷—»Ì… „»Ì⁄«  „œÌ‰…
        'SngTemp = Val(Me.lbl(51).Caption)
        'StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtTransSerial.text
        'LngDevNO = LngDevNO + 1
        'If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
        '    1, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        '    GoTo ErrTrap
        'End If
    End If

    If ChkTaxSerivce.value = vbChecked Then
        'StrTempAccountCode = "a4a9" '÷—»Ì… Œœ„… „»Ì⁄« 
        'SngTemp = Val(Me.lbl(54).Caption)
        'StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtTransSerial.text
        'LngDevNO = LngDevNO + 1
        'If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
        '    1, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        '    GoTo ErrTrap
        'End If
    End If
    If TxtWorkOrderNO.text <> "" Then
        Cn.Execute " Update Transactions set FlgProductOrder=1,Product_Issue_voucher_Serial = N'" & Trim(TxtWorkOrderNO) & "',nots2 =" & val(XPTxtBillID) & "    WHERE     (dbo.Transactions.Transaction_Type = 26) AND (dbo.Transactions.Transaction_Serial = N'" & TxtWorkOrderNO.text & "')"
    End If
    Cn.CommitTrans
    TransBegine = False
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    If IsSaveWithOutMsg Then Exit Sub
    Select Case Me.TxtModFlg.text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
                Msg = " Data Was Saved do you want Another Entry" & CHR(13)
    
            End If
    
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton1, App.Title) = vbYes Then
                Cmd_Click (0)
                Screen.MousePointer = vbDefault
            Else
                TxtModFlg.text = "R"
            End If
   
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Else
                Msg = " changes Was Saved " & CHR(13)
    
            End If

            lbl(56).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
            TxtModFlg.text = "R"
    End Select

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:

    If TransBegine = True Then
        TransBegine = False
        Cn.RollbackTrans
    End If

    'Resume
    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
    End If

    If Not RsNotes Is Nothing Then
        If RsNotes.EditMode <> adEditNone Then
            RsNotes.CancelUpdate
        End If
    End If

    If Not RSTransDetails Is Nothing Then
        If RSTransDetails.EditMode <> adEditNone Then
            RSTransDetails.CancelUpdate
        End If
    End If

    Screen.MousePointer = vbDefault

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Msg = Msg & CHR(13) & Err.Description
        Msg = Msg & CHR(13) & Err.Number
        Msg = Msg & CHR(13) & Err.Source
        Msg = Msg & CHR(13) & Err.LastDllError
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Msg = Msg & CHR(13) & Err.Description
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & Err.Source
    Msg = Msg & CHR(13) & Err.LastDllError
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub XPBtnNewClients_Click()
    On Error GoTo ErrTrap

    'With FrmAddNewCustemer
    '    .DealingForm = InvoiceTransaction
    '    FrmAddNewCustemer.AddType = 1
    '    .Caption = "≈÷«›… ⁄„Ì· ÃœÌœ"
    '    .lbl(1).Caption = "ﬂÊœ «·⁄„Ì·"
    '    .lbl(0).Caption = "«”„ «·⁄„Ì·"
    '    Set .DcboCustomers = DBCboClientName
    '    .show vbModal
    '    cSearchDcbo(0).Refresh
    'End With

    Exit Sub
ErrTrap:
End Sub

Private Sub XPCboDiscountType_Change()
    XPCboDiscountType_Click
End Sub

Private Sub XPCboDiscountType_Click()
    On Error GoTo ErrTrap

    If XPCboDiscountType.ListIndex = 0 Or XPCboDiscountType.ListIndex = 3 Or XPCboDiscountType.ListIndex = -1 Then
    
        XPTxtDiscountVal.Enabled = False
        XPTxtDiscountVal.text = ""
    Else
    
        XPTxtDiscountVal.Enabled = True
        XPTxtDiscountVal.text = ""
    End If

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If FG.TextMatrix(1, FG.ColIndex("Code")) <> "" Then
            NewGrid.Calculate 1, , , True
        End If
    End If

    Me.lbl(55).Visible = (Me.XPCboDiscountType.ListIndex = 2)

    Me.lbl(21).Visible = (Me.XPCboDiscountType.ListIndex = 2)

    If XPCboDiscountType.ListIndex = 0 Then
        ' lbl(8).Visible = False
        ' XPTxtDiscountVal.Visible = False
        ' lbl(8).Visible = False
    Else
        ' lbl(8).Visible = True
        ' XPTxtDiscountVal.Visible = True
        ' lbl(8).Visible = True
    End If

    Exit Sub

ErrTrap:
End Sub

Private Sub XPChkPayType_Click(index As Integer)
    On Error GoTo ErrTrap

    Select Case index

        Case 0

            If XPChkPayType(0).value = Checked Then
                If Me.TxtModFlg.text = "N" Then
                    XPTxtValue(0).text = ""
                    XPTxtSerial(0).text = ""
                End If

                If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
                    XPTxtValue(0).Enabled = True
                    '                XPTxtSerial(0).Enabled = True
                    XPTxtValue(0).locked = False
                    '                XPTxtSerial(0).Locked = False
                End If

            Else
                XPTxtValue(0).Enabled = False
                XPTxtValue(0).text = ""
                '            XPTxtSerial(0).Enabled = False
            End If

        Case 1

            If XPChkPayType(1).value = Checked Then
                If Me.TxtModFlg.text = "N" Then
                    XPTxtValue(1).text = ""
                    XPTxtSerial(1).text = ""
                    DtpDelayDate.value = Date
                    XPTxtSerial(1).text = CStr(new_id("Notes", "NoteSerial", "", True))
                End If

                If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
                    XPTxtValue(1).Enabled = True
                    XPTxtValue(1).locked = False
                    DtpDelayDate.Enabled = True
                Else
                    DtpDelayDate.Enabled = False
                End If

                Me.ChkInstall.Enabled = True
            Else
                XPTxtValue(1).Enabled = False
                XPTxtSerial(1).Enabled = False
                XPTxtValue(1).text = ""
                Me.ChkInstall.Enabled = False
            End If

        Case 2

            If XPChkPayType(2).value = Checked And Me.TxtModFlg.text <> "R" Then
                Me.CmdCheque.Enabled = True
            Else
                Me.CmdCheque.Enabled = False
                Me.lbl(18).Caption = 0
                Me.lbl(19).Caption = 0
                Me.FgCheques.rows = Me.FgCheques.FixedRows
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub XPChkTAX_Click()
    On Error GoTo ErrTrap

    If XPChkTAX.value = Checked Then
        XPTxtTaxValue.Enabled = True
        lbl(4).Enabled = True
        lbl(45).Enabled = True
    Else
        XPTxtTaxValue.text = ""
        XPTxtTaxValue.Enabled = False
        lbl(4).Enabled = False
        lbl(45).Enabled = False
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
End Sub

Private Sub XPTxtDiscountVal_Change()
    Dim Msg As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        NewGrid.Calculate 1, , , True
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub PrintReport(Optional PrinterTarget As Boolean = False)
    On Error GoTo ErrTrap
    Dim BuyReport As ClsBuyReport
    Dim Msg As String
    
    
    If Not XPTxtBillID.text Then
        Set BuyReport = New ClsBuyReport
        
         Msg = "”Ì „ ÿ»«⁄… «· ﬁ—Ì— «Ã„«·Ì  " & CHR(13)
        Msg = Msg + "«÷€ÿ ‰⁄„ ··„Ê«›ﬁ… «Ê ·« ··ÿ»«⁄…  ›’Ì·Ì"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbNo Then
        
            BuyReport.ShowBuyData XPTxtBillID.text, 33, True, LblTotal.Caption
        Else
            BuyReport.ShowBuyData XPTxtBillID.text, 3, True, LblTotal.Caption
        End If
        
     
    End If

    Exit Sub
ErrTrap:
 
End Sub

Private Sub XPTxtDiscountVal_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtDiscountVal.text, 0)
End Sub

Private Sub XPTxtSum_Change()
    On Error GoTo ErrTrap
Exit Sub
    If CboPayMentType.ListIndex = 0 Then
        XPChkPayType(0).value = Checked
        XPTxtValue(0).text = XPTxtSum.text
    End If

    Me.LblTotal.Caption = XPTxtSum.text
    CalculateInvPrecent
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

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub

Public Sub Convert()
    Cmd_Click (0)
End Sub

Public Sub Cala()
    NewGrid.Calculate 1, , , True
End Sub

Private Sub DBCboClientName_Change()
    Dim Msg As String
    Dim RsTemp  As ADODB.Recordset
    Dim StrSQL As String

    On Error GoTo ErrTrap

    If val(DBCboClientName.BoundText) <> 0 Then
        If (DBCboClientName.BoundText = 1 Or DBCboClientName.BoundText = 2) And Me.TxtModFlg.text <> "R" Then
            CboPayMentType.locked = True
            '        CboPayMentType.ListIndex = 0
            Me.TxtCashCustomerName.Enabled = True
            Me.CmdCash(0).Enabled = True
            Me.CmdCash(1).Enabled = True
        Else
            CboPayMentType.locked = False
            Me.TxtCashCustomerName.Enabled = False
            Me.CmdCash(0).Enabled = False
            Me.CmdCash(1).Enabled = False
        End If

        If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
            StrSQL = "Select * From TblCustemers Where CusID=" & val(DBCboClientName.BoundText)
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsTemp.BOF Or RsTemp.EOF) Then
                If Not (IsNull(RsTemp("SaleType").value)) Then
                    If RsTemp("SaleType").value = 0 Then
                        Me.CboSaleType.ListIndex = 0
                    ElseIf RsTemp("SaleType").value = 1 Then
                        Me.CboSaleType.ListIndex = 1
                    End If

                Else
                    Me.CboSaleType.ListIndex = -1
                End If

                If Not (IsNull(RsTemp("Trans_DiscountType").value)) Then
                    If RsTemp("Trans_DiscountType").value = 0 Then
                        Me.XPCboDiscountType.ListIndex = 0
                        Me.XPTxtDiscountVal.text = 0
                    ElseIf RsTemp("Trans_DiscountType").value = 1 Then
                        Me.XPCboDiscountType.ListIndex = 1
                        Me.XPTxtDiscountVal.text = IIf(IsNull(RsTemp("Trans_Discount").value), "", RsTemp("Trans_Discount").value)
                    ElseIf RsTemp("Trans_DiscountType").value = 2 Then
                        Me.XPCboDiscountType.ListIndex = 2
                        Me.XPTxtDiscountVal.text = IIf(IsNull(RsTemp("Trans_Discount").value), "", RsTemp("Trans_Discount").value)
                    End If

                Else
                    Me.XPCboDiscountType.ListIndex = 0
                    Me.XPTxtDiscountVal.text = 0
                End If

            Else
                Me.CboSaleType.ListIndex = -1
                Me.XPCboDiscountType.ListIndex = 0
                Me.XPTxtDiscountVal.text = 0
            End If

            RsTemp.Close
            Set RsTemp = Nothing
        End If
    End If

    Exit Sub
ErrTrap:
    Msg = Err.Description & CHR(13) & ""
    Msg = Msg & Err.Source & CHR(13) & ""
    Msg = Msg & Me.Name & " DBCboClientName_Change:" & CHR(13) & ""
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub DBCboClientName_Click(Area As Integer)
    DBCboClientName_Change
End Sub

Private Sub XPTxtValue_Change(index As Integer)
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If XPTxtValue(1).text <> "" Then
            If val(Me.XPTxtValue(1).text) > 0 Then
                ChkInstall.Enabled = True
            End If
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Public Sub ReplacementData()
    Dim Msg As String
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsReplace As ADODB.Recordset

    If Me.TxtModFlg.text <> "R" Then Exit Sub

    '«·»ÕÀ ⁄‰ ⁄„·Ì«  «·«” »œ«· «·Œ«’… »«·›« Ê—…
    If FG.TextMatrix(FG.row, FG.ColIndex("Code")) <> "" And FG.TextMatrix(FG.row, FG.ColIndex("Serial")) <> "" Then
        StrSQL = "select * From ReplacedItems where ReturnID=" & XPTxtBillID.text
        StrSQL = StrSQL + " and ItemID=" & FG.TextMatrix(FG.row, FG.ColIndex("Code"))
        StrSQL = StrSQL + " and ItemSerial='" & FG.TextMatrix(FG.row, FG.ColIndex("Serial")) & "'"
        Set RsReplace = New ADODB.Recordset
        RsReplace.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsReplace.EOF Or RsReplace.BOF) Then
            Msg = "·ﬁœ  „ «” »œ«· «·ﬁÿ⁄… : " & FG.cell(flexcpTextDisplay, FG.row, FG.ColIndex("Name")) & CHR(13)
            Msg = Msg + "–«  «·”Ì—Ì«· : " & FG.TextMatrix(FG.row, FG.ColIndex("Serial")) & CHR(13)
            Msg = Msg + " »«·ﬁÿ⁄… –«  «·”Ì—Ì«· : " & RsReplace("newSerial").value & CHR(13)
            Msg = Msg + "›Ì ⁄„·Ì… ’Ì«‰…"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, "ﬁÿ⁄…  „ «” »œ«·Â«"
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Function AvailableDeal() As Boolean
    On Error GoTo ErrTrap
    Dim RowNum As Integer
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    Dim RsSalle As ADODB.Recordset
    Dim LngItemID As Long

    For RowNum = 1 To FG.rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            StrSQL = "select * From QryDelPurchase where Transaction_Date >=" & SQLDate(XPDtbBill.value, True) & ""
            StrSQL = StrSQL + " and Item_ID=" & FG.TextMatrix(RowNum, FG.ColIndex("Code"))
            StrSQL = StrSQL + " and Transaction_Type=9"

            If FG.cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then
                If FG.TextMatrix(RowNum, FG.ColIndex("Serial")) <> "" Then
                    StrSQL = StrSQL + " and ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                End If
            End If

            Set RsSalle = New ADODB.Recordset
            RsSalle.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsSalle.EOF Or RsSalle.BOF) Then
                If FG.cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then

                    '                StrSql = "select * From QryGardComplete where ItemID=" & FG.TextMatrix(RowNum, FG.ColIndex("Code"))
                    '                StrSql = StrSql + " AND ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                    '                StrSql = StrSql + " AND StoreID=" & DCboStoreName.BoundText
                    '                Set RsTemp = New ADODB.Recordset
                    '                RsTemp.Open StrSql, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    '                If RsTemp.EOF Or RsTemp.BOF Then
                    With FrmAlarm
                        .DealingForm = InvoiceTransaction
                        .show vbModal
                    End With

                    AvailableDeal = False
                    Exit Function
                    '                End If
                    RsTemp.Close
                Else
                    LngItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                    Set RsTemp = New ADODB.Recordset
                    Set RsTemp = GetItemQuantityStock(LngItemID, Me.DCboStoreName.BoundText, Me.XPDtbBill.value, val(Me.XPTxtBillID.text))

                    If Not (RsTemp.EOF Or RsTemp.BOF) Then
                        If val(RsTemp("QTY").value) < val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))) Then

                            With FrmAlarm
                                .DealingForm = InvoiceTransaction
                                .show vbModal
                            End With

                            AvailableDeal = False
                            Exit Function
                        End If
                    End If

                    RsTemp.Close
                End If
            End If

            RsSalle.Close
        End If

    Next RowNum

    AvailableDeal = True
    Exit Function
ErrTrap:
End Function

Private Sub SetDefaults()
    Dim StrTemp As String
    Dim RsTemp As ADODB.Recordset

    Me.CboSaleType.ListIndex = 0

    If SystemOptions.SysPurDateTakeType = InvDateFromLocalCompuer Then
        XPDtbBill.value = Date
    ElseIf SystemOptions.SysPurDateTakeType = InvDateFromServerComputer Then
        StrTemp = "select Getdate() as ServerDate"
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrTemp, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            If Not IsNull(RsTemp("ServerDate").value) Then
                XPDtbBill.value = Format(RsTemp("ServerDate").value, "yyyy/M/d")
            End If

            'XPDtbBill.Value = IIf(IsNull(RsTemp("ServerDate").Value), Date, (RsTemp("ServerDate").Value))
        End If

        RsTemp.Close
        Set RsTemp = Nothing
    End If

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveLast

        If SystemOptions.SysPurDateTakeType = InvDateFromLastInvDate Then
            XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), Date, (rs("Transaction_Date").value))
        End If

        Me.DcboEmp.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)

        If Not IsNull(rs("Transaction_Serial").value) Then
            StrTemp = rs("Transaction_Serial").value
            StrTemp = val(StrTemp) + 1
            TxtTransSerial.text = StrTemp
        Else
            TxtTransSerial.text = 1
        End If

    Else
        TxtTransSerial.text = 1
    End If

    'Me.CboPayMentType.ListIndex = 1
    CboPayMentType.ListIndex = 1

End Sub

Private Sub CalculateInvPrecent()
Exit Sub
    Dim DblInvTotal As Double
    Dim DblInvProfit As Double
    Dim DblRes As Double

    DblInvProfit = val(Me.LblInvProfit.Caption)
    DblInvTotal = val(Me.XPTxtSum.text)

    If DblInvProfit = 0 Or DblInvTotal = 0 Then
        DblRes = 0
    Else
        DblRes = 100 * (DblInvProfit / DblInvTotal)
    End If

    Me.lblInvPrecent.Caption = "%" & CStr(Int(DblRes)) 'Format(DblRes, SystemOptions.SysDefCurrencyForamt)
End Sub

Private Sub LoadCombosData()
    Dcombos.GetEmployees Me.DcboEmp
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName
    Dcombos.GetStores Me.DCboStoreName
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetDocTypebyid Me.DCDocTypes, 27, val(Me.dcBranch.BoundText)

    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DBCboClientName
    cSearchDcbo(0).SetBuddyText Me.TxtCusID

    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DCboStoreName
   ' cSearchDcbo(1).SetBuddyText Me.TxtStoreID

    Set cSearchDcbo(3) = New clsDCboSearch
    Set cSearchDcbo(3).Client = Me.DcboEmp
    cSearchDcbo(3).SetBuddyText Me.TxtEmployeeID
End Sub

Private Sub ChangeLang()
    CmdConvert.Caption = "Convert to bill"
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Label4.Caption = "Remarks"
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    Cmd(10).Caption = "Print JL"
    Frame3.Caption = "JL NO"
    Label5.Caption = "Doc. Type"
    Me.LblShortcutKeys.Caption = "(New F12 OR Enter) ,(Edit F11),(Save F10),(Undo F9),(Delete F8),(Search F7)"
    Me.Caption = "Production Issue Voucher"
    Ele(9).Caption = Me.Caption
    lbl(5).Caption = "Invoice ID"
    lbl(6).Caption = "Invoice Date"
    lbl(7).Caption = "Customer Name"
    lbl(24).Caption = "Store "
    lbl(25).Caption = "Employee "
    lbl(9).Caption = "Payment Type"
    lbl(10).Caption = "Discount Type"
    Label3.Caption = "Branch"
    Label2.Caption = "Based On"
    lbl(63).Caption = "Total Qty"
    lbl(10).Caption = "Discount Type"

    lbl(8).Caption = "Discount Value"
    lbl(22).Caption = "Profit Value"
    lbl(23).Caption = "Profit Perce"

    lbl(3).Caption = " Total:"
    lbl(50).Caption = "Disc"
    lbl(49).Caption = " Net:"

    lbl(1).Caption = " By:"
    lbl(2).Caption = "Rec. Count:"

    lbl(31).Caption = "Item Code"
    lbl(30).Caption = "Item Name"
    lbl(29).Caption = " Case"
    lbl(28).Caption = " Serial"
    lbl(27).Caption = "QTY"
    lbl(26).Caption = "Price"
    lbl(32).Caption = "Production Order NO"

    lbl(33).Caption = "Cash Customer"
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    Me.CmdHelp.Caption = "Help"
    Me.XPTab301.TabCaption(0) = "Items"
    
    Me.XPTab301.TabCaption(1) = "Notes"
    lbl(20).Caption = "Payment Method"
    XPChkPayType(0).Caption = "Cahs"
    XPChkPayType(1).Caption = "Credit"
    XPChkPayType(2).Caption = "Cheque"
    lbl(13).Caption = "Value"
    lbl(15).Caption = "Value"
    lbl(16).Caption = "Value"
    lbl(12).Caption = "Serial"
    lbl(14).Caption = "Serial"
    lbl(11).Caption = "Box"
    lbl(21).Caption = "Due Date"
    
    lbl(18).Caption = "Check NO."
    lbl(17).Caption = "Bank"
    lbl(19).Caption = "Due Date"
    CmdINSTALLMENT.Caption = "INSTALLMENT"
    Me.XPTab301.TabCaption(2) = "Comment On Invoice"
    Me.Ele(15).Caption = "Write any Comments about this Invoice"
    
    With FgInstallments
        .TextMatrix(0, .ColIndex("QestID")) = "ID"
        .TextMatrix(0, .ColIndex("value")) = "value"
        .TextMatrix(0, .ColIndex("Due_Date")) = "Due_Date"
 
    End With

    With FgCheques
 
        .TextMatrix(0, .ColIndex("CheckValue")) = "Value"
        .TextMatrix(0, .ColIndex("CheckNumber")) = "Cheque Number"
        .TextMatrix(0, .ColIndex("BankName")) = "Bank Name"
        .TextMatrix(0, .ColIndex("DueDate")) = "Due Date"
        .TextMatrix(0, .ColIndex("ReleaseDate")) = "Release Date"
 
    End With

    CmdINSTALLMENT.Caption = "Calc"
    ChkInstall.Caption = "Install."
End Sub

Private Sub XPTxtValue_KeyPress(index As Integer, _
                                KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtValue(index).text, 0)
End Sub

Private Function CheckCashCustomer() As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    If Trim$(Me.TxtCashCustomerName.text) = "" Then
        CheckCashCustomer = True
    Else
        StrSQL = "Select * From Transactions Where CashCustomerName='" & Trim$(Me.TxtCashCustomerName.text) & "'"
    
    End If

End Function

Private Sub XPTxtValue_MouseMove(index As Integer, _
                                 Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)

    If val(Me.XPTxtValue(index).text) <> 0 Then
        Me.XPTxtValue(index).ToolTipText = WriteNo(Me.XPTxtValue(index).text, 1, True)
    Else
        Me.XPTxtValue(index).ToolTipText = ""
    End If

End Sub

Private Sub SumChecks()

    With Me.FgCheques

        If .rows > 1 Then
            Me.lbl(19).Caption = .Aggregate(flexSTCount, .FixedRows, .ColIndex("CheckNumber"), .rows - 1, .ColIndex("CheckNumber"))
            Me.lbl(18).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CheckValue"), .rows - 1, .ColIndex("CheckValue"))
        Else
            Me.lbl(19).Caption = 0
            Me.lbl(18).Caption = 0
        End If

    End With

End Sub

Private Sub ClearNotes()

    LblPrecenType.Caption = 0
    LblPrecenValue.Caption = 0
    LblInstallTotal.Caption = 0
    LblInstallCount.Caption = 0
    LblFirstInstallDate.Caption = ""
    LblInstallSeprator.Caption = ""
    LblInstallmentType.Caption = ""
    LblStartValue.Caption = ""
    lbl(19).Caption = ""
    lbl(18).Caption = ""
End Sub
