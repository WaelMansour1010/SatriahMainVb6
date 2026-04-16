VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmOut 
   Caption         =   "”‰œ ’—›"
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12690
   HelpContextID   =   160
   Icon            =   "FrmOut.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   12690
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   9375
      Left            =   0
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   0
      Width           =   12690
      _cx             =   22384
      _cy             =   16536
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
      _GridInfo       =   $"FrmOut.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1785
         Index           =   0
         Left            =   15
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   645
         Width           =   12660
         _cx             =   22331
         _cy             =   3149
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
         Begin VB.TextBox TxtItemCodeB12 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3360
            TabIndex        =   245
            Top             =   300
            Width           =   1440
         End
         Begin VB.TextBox txtOrderID 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   199
            Top             =   0
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.TextBox txtTradingContractID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1830
            RightToLeft     =   -1  'True
            TabIndex        =   197
            Top             =   0
            Width           =   945
         End
         Begin VB.ComboBox DcbType 
            Height          =   315
            Left            =   4335
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   194
            Top             =   585
            Width           =   1305
         End
         Begin VB.TextBox TxtOldOpOrderID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   177
            Top             =   1344
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton Command10 
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄—÷"
            Height          =   285
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   175
            Top             =   285
            Width           =   525
         End
         Begin VB.TextBox txtManualNO 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FFFF&
            Height          =   315
            Left            =   8730
            RightToLeft     =   -1  'True
            TabIndex        =   173
            Top             =   0
            Width           =   945
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   2  'Center
            BackColor       =   &H0000FFFF&
            Height          =   300
            Left            =   10635
            TabIndex        =   163
            Top             =   630
            Width           =   750
         End
         Begin VB.TextBox txtEmpCode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FFFF&
            Height          =   270
            Left            =   7395
            RightToLeft     =   -1  'True
            TabIndex        =   160
            Top             =   1488
            Width           =   570
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FFFF&
            Height          =   285
            Left            =   7395
            RightToLeft     =   -1  'True
            TabIndex        =   159
            Top             =   0
            Width           =   570
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FFFF&
            Height          =   270
            Left            =   7395
            RightToLeft     =   -1  'True
            TabIndex        =   155
            Top             =   1176
            Width           =   570
         End
         Begin MSDataListLib.DataCombo DcboEmpDepartments 
            Height          =   315
            Left            =   4335
            TabIndex        =   156
            Top             =   1170
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   65535
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCEquipments 
            Height          =   315
            Left            =   4335
            TabIndex        =   154
            Top             =   840
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   65535
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FFFF&
            Height          =   264
            Left            =   7395
            RightToLeft     =   -1  'True
            TabIndex        =   153
            Top             =   840
            Width           =   570
         End
         Begin VB.ComboBox CBoBasedON 
            BackColor       =   &H0000FFFF&
            Height          =   315
            ItemData        =   "FrmOut.frx":03F0
            Left            =   6045
            List            =   "FrmOut.frx":03F2
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   150
            Top             =   270
            Width           =   1905
         End
         Begin VB.ComboBox DCOPrType 
            Height          =   315
            Left            =   6045
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   149
            Top             =   585
            Width           =   1905
         End
         Begin VB.TextBox TXT_order_no 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4830
            RightToLeft     =   -1  'True
            TabIndex        =   148
            Top             =   330
            Width           =   1230
         End
         Begin VB.TextBox TxtBillComment 
            Alignment       =   1  'Right Justify
            Height          =   396
            Left            =   60
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   147
            Top             =   765
            Width           =   2865
         End
         Begin VB.TextBox TxtTicketNO 
            Alignment       =   1  'Right Justify
            Height          =   240
            Left            =   1230
            RightToLeft     =   -1  'True
            TabIndex        =   143
            Top             =   1608
            Visible         =   0   'False
            Width           =   1770
         End
         Begin VB.TextBox TxtWorkOrderNO 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2625
            RightToLeft     =   -1  'True
            TabIndex        =   138
            Top             =   1608
            Visible         =   0   'False
            Width           =   1785
         End
         Begin VB.TextBox TXTNoteID 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   133
            Top             =   2136
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   10230
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   132
            Top             =   0
            Width           =   1140
         End
         Begin ALLButtonS.ALLButton CmdConvert 
            Height          =   285
            Left            =   0
            TabIndex        =   131
            Top             =   -390
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   503
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
            MICON           =   "FrmOut.frx":03F4
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
            Height          =   225
            Left            =   2010
            RightToLeft     =   -1  'True
            TabIndex        =   129
            Top             =   2244
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   240
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Top             =   2040
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.TextBox TxtCashCustomerName 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FFFF&
            Height          =   300
            Left            =   8820
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   948
            Width           =   2565
         End
         Begin VB.TextBox TxtEmployeeID 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   5145
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   2040
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.TextBox TxtStoreID 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FFFF&
            Height          =   315
            Left            =   10635
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   1308
            Width           =   750
         End
         Begin VB.TextBox TxtCusID 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   11070
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   720
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.ComboBox CboSaleType 
            Height          =   288
            Left            =   120
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1956
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   10905
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   -255
            Visible         =   0   'False
            Width           =   2580
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   570
            Index           =   8
            Left            =   4815
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   2715
            Visible         =   0   'False
            Width           =   3360
            _cx             =   5927
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
               Height          =   330
               Left            =   90
               TabIndex        =   48
               Top             =   135
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   582
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
               ButtonImage     =   "FrmOut.frx":0410
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰”»… «·—»Õ"
               ForeColor       =   &H00C00000&
               Height          =   195
               Index           =   23
               Left            =   2550
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·—»Õ"
               ForeColor       =   &H00C00000&
               Height          =   225
               Index           =   22
               Left            =   2550
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   120
               Width           =   1335
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
               Height          =   210
               Left            =   1125
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   330
               Width           =   1635
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
               Height          =   210
               Left            =   1125
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   105
               Width           =   1635
            End
         End
         Begin VB.ComboBox XPCboDiscountType 
            Height          =   288
            Left            =   495
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1908
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.TextBox XPTxtDiscountVal 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   225
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   1728
            Visible         =   0   'False
            Width           =   1230
         End
         Begin VB.ComboBox CboPayMentType 
            Height          =   288
            Left            =   120
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1764
            Visible         =   0   'False
            Width           =   3135
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   285
            Left            =   8820
            TabIndex        =   3
            Top             =   630
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   65535
            ListField       =   "6"
            BoundColumn     =   ""
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   8820
            TabIndex        =   6
            Top             =   1290
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   65535
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   285
            Left            =   8820
            TabIndex        =   1
            Top             =   345
            Width           =   2550
            _ExtentX        =   4498
            _ExtentY        =   503
            _Version        =   393216
            Format          =   224460801
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton XPBtnNewClients 
            Height          =   288
            Left            =   4968
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   900
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   503
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
            ButtonImage     =   "FrmOut.frx":07AA
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   312
            Left            =   2340
            TabIndex        =   8
            Top             =   2028
            Visible         =   0   'False
            Width           =   2376
            _ExtentX        =   4180
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton CmdCash 
            Height          =   225
            Index           =   0
            Left            =   6540
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   930
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   397
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
            ButtonImage     =   "FrmOut.frx":0B44
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdCash 
            Height          =   225
            Index           =   1
            Left            =   6105
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   930
            Visible         =   0   'False
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   397
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
            ButtonImage     =   "FrmOut.frx":0EDE
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   5715
            TabIndex        =   134
            Top             =   0
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   65535
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCDocTypes 
            Height          =   315
            Left            =   3360
            TabIndex        =   141
            Top             =   0
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcCostCenter 
            Bindings        =   "FrmOut.frx":1278
            Height          =   315
            Left            =   30
            TabIndex        =   144
            Top             =   405
            Width           =   2070
            _ExtentX        =   3651
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
            Left            =   4335
            TabIndex        =   161
            Top             =   1470
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   65535
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbProject 
            Height          =   312
            Left            =   60
            TabIndex        =   191
            Top             =   1188
            Width           =   2868
            _ExtentX        =   5054
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcCustmer 
            Bindings        =   "FrmOut.frx":128D
            Height          =   315
            Left            =   30
            TabIndex        =   198
            Top             =   0
            Width           =   1800
            _ExtentX        =   3175
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
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "« ›«ﬁÌ…"
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   1
            Left            =   2745
            TabIndex        =   196
            Top             =   30
            Width           =   555
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‰Ê⁄"
            Height          =   285
            Index           =   67
            Left            =   5265
            RightToLeft     =   -1  'True
            TabIndex        =   193
            Top             =   600
            Width           =   765
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„‘—Ê⁄"
            Height          =   216
            Index           =   66
            Left            =   3072
            RightToLeft     =   -1  'True
            TabIndex        =   192
            Top             =   1308
            Width           =   768
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÌœÊÌ"
            Height          =   195
            Index           =   53
            Left            =   9690
            RightToLeft     =   -1  'True
            TabIndex        =   174
            Top             =   0
            Width           =   420
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„ÊŸ›"
            Height          =   240
            Index           =   64
            Left            =   7770
            RightToLeft     =   -1  'True
            TabIndex        =   162
            Top             =   1470
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„⁄œÂ/«·”Ì«—…"
            Height          =   225
            Index           =   62
            Left            =   7860
            RightToLeft     =   -1  'True
            TabIndex        =   158
            Top             =   870
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·«œ«—… «·ÿ«·»…"
            Height          =   210
            Index           =   61
            Left            =   7860
            RightToLeft     =   -1  'True
            TabIndex        =   157
            Top             =   1170
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·⁄„·Ì…"
            Height          =   285
            Index           =   60
            Left            =   8010
            RightToLeft     =   -1  'True
            TabIndex        =   152
            Top             =   615
            Width           =   765
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï"
            Height          =   195
            Index           =   59
            Left            =   8070
            RightToLeft     =   -1  'True
            TabIndex        =   151
            Top             =   285
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„·«ÕŸ« "
            Height          =   204
            Left            =   2952
            RightToLeft     =   -1  'True
            TabIndex        =   146
            Top             =   876
            Width           =   768
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„—ﬂ“ «· ﬂ·›… "
            Height          =   210
            Left            =   1770
            RightToLeft     =   -1  'True
            TabIndex        =   145
            Top             =   375
            Width           =   750
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·”‰œ"
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   0
            Left            =   4935
            TabIndex        =   142
            Top             =   0
            Width           =   765
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   7875
            TabIndex        =   135
            Top             =   0
            Width           =   795
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
            Height          =   216
            Index           =   55
            Left            =   108
            RightToLeft     =   -1  'True
            TabIndex        =   130
            Top             =   1656
            Visible         =   0   'False
            Width           =   216
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì· «·‰ﬁœÏ"
            Height          =   255
            Index           =   33
            Left            =   10905
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   1050
            Width           =   1725
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„  «„— «·‘€·  "
            Height          =   192
            Index           =   32
            Left            =   3036
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   1548
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„‰œÊ»"
            Height          =   210
            Index           =   25
            Left            =   5865
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   2070
            Visible         =   0   'False
            Width           =   1470
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Œ’„"
            Height          =   228
            Index           =   10
            Left            =   3420
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   2460
            Visible         =   0   'False
            Width           =   1452
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   276
            Index           =   9
            Left            =   3288
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   2148
            Visible         =   0   'False
            Width           =   1416
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„…"
            Height          =   288
            Index           =   8
            Left            =   1512
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   1728
            Visible         =   0   'False
            Width           =   432
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ «·„Œ“‰"
            Height          =   195
            Index           =   24
            Left            =   11355
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   1305
            Width           =   1290
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì·"
            Height          =   225
            Index           =   7
            Left            =   10860
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   645
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·«–‰"
            Height          =   225
            Index           =   6
            Left            =   10785
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   345
            Width           =   1800
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·«–‰"
            Height          =   210
            Index           =   5
            Left            =   11685
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   75
            Width           =   900
         End
      End
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   5910
         Left            =   15
         TabIndex        =   22
         Top             =   2445
         Width           =   12660
         _cx             =   22331
         _cy             =   10425
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
         Caption         =   "«·√’‰«›|»Ì«‰«  ›« Ê—… «·„»Ì⁄« |»Ì«‰«  «·‘Õ‰|«·„—›ﬁ« |»Ì«‰«  «·‘Õ‰2|«·«⁄ „«œ"
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
         Picture(0)      =   "FrmOut.frx":12A2
         Picture(1)      =   "FrmOut.frx":163C
         Flags(1)        =   2
         Picture(2)      =   "FrmOut.frx":19D6
         Flags(2)        =   2
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   5445
            Left            =   14205
            TabIndex        =   249
            Top             =   45
            Width           =   12570
            _Version        =   786432
            _ExtentX        =   22172
            _ExtentY        =   9604
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Begin VB.TextBox txtShipBatched 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   2115
               RightToLeft     =   -1  'True
               TabIndex        =   299
               Top             =   1980
               Width           =   1590
            End
            Begin VB.TextBox txtShipPlantNo 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   5835
               RightToLeft     =   -1  'True
               TabIndex        =   289
               Top             =   4500
               Width           =   1590
            End
            Begin VB.TextBox txtShipTripNo 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   5835
               RightToLeft     =   -1  'True
               TabIndex        =   287
               Top             =   3960
               Width           =   1590
            End
            Begin VB.TextBox txtShipDayOrder 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   5850
               RightToLeft     =   -1  'True
               TabIndex        =   285
               Top             =   3480
               Width           =   1590
            End
            Begin VB.TextBox txtShipThisLoad 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   5835
               RightToLeft     =   -1  'True
               TabIndex        =   283
               Top             =   2940
               Width           =   1590
            End
            Begin VB.TextBox txtShipTotalDeleveryd 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   5835
               RightToLeft     =   -1  'True
               TabIndex        =   281
               Top             =   2535
               Width           =   1590
            End
            Begin VB.TextBox txtShipIceTemp 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   5835
               RightToLeft     =   -1  'True
               TabIndex        =   279
               Top             =   2040
               Width           =   1590
            End
            Begin VB.TextBox txtShipTruckNo 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   5835
               RightToLeft     =   -1  'True
               TabIndex        =   277
               Top             =   1575
               Width           =   1590
            End
            Begin VB.TextBox txtShipPump 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   5835
               RightToLeft     =   -1  'True
               TabIndex        =   275
               Top             =   1170
               Width           =   1590
            End
            Begin VB.TextBox txtShipPipeLine 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   5835
               RightToLeft     =   -1  'True
               TabIndex        =   273
               Top             =   720
               Width           =   1590
            End
            Begin VB.TextBox txtShipDriverName 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   5835
               RightToLeft     =   -1  'True
               TabIndex        =   271
               Top             =   330
               Width           =   1590
            End
            Begin VB.TextBox txtShipMixDescription 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   9300
               RightToLeft     =   -1  'True
               TabIndex        =   269
               Top             =   4455
               Width           =   1590
            End
            Begin VB.TextBox txtShipStructuralElement 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   9300
               RightToLeft     =   -1  'True
               TabIndex        =   267
               Top             =   3975
               Width           =   1590
            End
            Begin VB.TextBox txtShipProjectName 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   9315
               RightToLeft     =   -1  'True
               TabIndex        =   265
               Top             =   3555
               Width           =   1590
            End
            Begin VB.TextBox txtShipSiteNo 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   9300
               RightToLeft     =   -1  'True
               TabIndex        =   263
               Top             =   3015
               Width           =   1590
            End
            Begin VB.TextBox txtShipArea 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   9300
               RightToLeft     =   -1  'True
               TabIndex        =   261
               Top             =   2580
               Width           =   1590
            End
            Begin VB.TextBox txtShipDistance 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   9300
               RightToLeft     =   -1  'True
               TabIndex        =   259
               Top             =   2130
               Width           =   1590
            End
            Begin VB.TextBox txtShipCustomerName 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   9300
               RightToLeft     =   -1  'True
               TabIndex        =   257
               Top             =   1635
               Width           =   1590
            End
            Begin VB.TextBox txtShipAccountNo 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   9300
               RightToLeft     =   -1  'True
               TabIndex        =   255
               Top             =   1215
               Width           =   1590
            End
            Begin VB.TextBox txtShipEnquieryNo 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   9300
               RightToLeft     =   -1  'True
               TabIndex        =   253
               Top             =   735
               Width           =   1590
            End
            Begin VB.TextBox txtShipOrderNo 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   9300
               RightToLeft     =   -1  'True
               TabIndex        =   250
               Top             =   315
               Width           =   1590
            End
            Begin MSComCtl2.DTPicker txtShipRestunedPlant 
               Height          =   285
               Left            =   2115
               TabIndex        =   291
               Top             =   315
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   503
               _Version        =   393216
               Format          =   220921857
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker txtShipEndDischarge 
               Height          =   285
               Left            =   2115
               TabIndex        =   293
               Top             =   750
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   503
               _Version        =   393216
               Format          =   220921857
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker txtShipStartDisCharge 
               Height          =   285
               Left            =   2115
               TabIndex        =   295
               Top             =   1185
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   503
               _Version        =   393216
               Format          =   220921857
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker txtShipOnSite 
               Height          =   285
               Left            =   2115
               TabIndex        =   297
               Top             =   1560
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   503
               _Version        =   393216
               Format          =   220921857
               CurrentDate     =   38784
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Êﬁ  «· Õ„Ì·"
               Height          =   195
               Index           =   96
               Left            =   4335
               RightToLeft     =   -1  'True
               TabIndex        =   300
               Top             =   1920
               Width           =   1290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Êﬁ  «·Ê’Ê·"
               Height          =   225
               Index           =   94
               Left            =   3825
               RightToLeft     =   -1  'True
               TabIndex        =   298
               Top             =   1455
               Width           =   1800
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»œ¡ «· ›—Ì€"
               Height          =   225
               Index           =   93
               Left            =   3825
               RightToLeft     =   -1  'True
               TabIndex        =   296
               Top             =   1080
               Width           =   1800
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Â«ÌÂ «· ›—Ì€"
               Height          =   225
               Index           =   92
               Left            =   3825
               RightToLeft     =   -1  'True
               TabIndex        =   294
               Top             =   645
               Width           =   1800
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄ÊœÂ ··„’‰⁄"
               Height          =   225
               Index           =   91
               Left            =   3825
               RightToLeft     =   -1  'True
               TabIndex        =   292
               Top             =   330
               Width           =   1800
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·„’‰⁄"
               Height          =   195
               Index           =   90
               Left            =   7605
               RightToLeft     =   -1  'True
               TabIndex        =   290
               Top             =   4530
               Width           =   1290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·—Õ·Â"
               Height          =   195
               Index           =   89
               Left            =   7605
               RightToLeft     =   -1  'True
               TabIndex        =   288
               Top             =   3990
               Width           =   1290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ·» «·ÌÊ„"
               Height          =   195
               Index           =   88
               Left            =   7605
               RightToLeft     =   -1  'True
               TabIndex        =   286
               Top             =   3510
               Width           =   1290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ„Ê·Â «·‘Õ‰Â"
               Height          =   195
               Index           =   87
               Left            =   7605
               RightToLeft     =   -1  'True
               TabIndex        =   284
               Top             =   2970
               Width           =   1290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ã„«·Ï «·„” ·„"
               Height          =   195
               Index           =   86
               Left            =   7605
               RightToLeft     =   -1  'True
               TabIndex        =   282
               Top             =   2580
               Width           =   1290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "œ—ÃÂ «·Õ—«—Â"
               Height          =   195
               Index           =   85
               Left            =   7605
               RightToLeft     =   -1  'True
               TabIndex        =   280
               Top             =   2070
               Width           =   1290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·‘«Õ‰Â"
               Height          =   195
               Index           =   84
               Left            =   7605
               RightToLeft     =   -1  'True
               TabIndex        =   278
               Top             =   1605
               Width           =   1290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„÷ŒÂ"
               Height          =   195
               Index           =   83
               Left            =   7605
               RightToLeft     =   -1  'True
               TabIndex        =   276
               Top             =   1200
               Width           =   1290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„Ê«”Ì— «÷«›ÌÂ"
               Height          =   195
               Index           =   82
               Left            =   7605
               RightToLeft     =   -1  'True
               TabIndex        =   274
               Top             =   750
               Width           =   1290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·”«∆ﬁ"
               Height          =   195
               Index           =   81
               Left            =   7605
               RightToLeft     =   -1  'True
               TabIndex        =   272
               Top             =   360
               Width           =   1290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Ê’› «·Œ·ÿÂ"
               Height          =   195
               Index           =   80
               Left            =   11190
               RightToLeft     =   -1  'True
               TabIndex        =   270
               Top             =   4515
               Width           =   1290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·’»Â"
               Height          =   195
               Index           =   79
               Left            =   11190
               RightToLeft     =   -1  'True
               TabIndex        =   268
               Top             =   4035
               Width           =   1290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„‘—Ê⁄"
               Height          =   195
               Index           =   78
               Left            =   11190
               RightToLeft     =   -1  'True
               TabIndex        =   266
               Top             =   3615
               Width           =   1290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·„Êﬁ⁄"
               Height          =   195
               Index           =   77
               Left            =   11190
               RightToLeft     =   -1  'True
               TabIndex        =   264
               Top             =   3075
               Width           =   1290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„‰ÿﬁÂ"
               Height          =   195
               Index           =   76
               Left            =   11190
               RightToLeft     =   -1  'True
               TabIndex        =   262
               Top             =   2640
               Width           =   1290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„”«›Â"
               Height          =   195
               Index           =   75
               Left            =   11190
               RightToLeft     =   -1  'True
               TabIndex        =   260
               Top             =   2190
               Width           =   1290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·⁄„Ì·"
               Height          =   195
               Index           =   74
               Left            =   11190
               RightToLeft     =   -1  'True
               TabIndex        =   258
               Top             =   1695
               Width           =   1290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·Õ”«»"
               Height          =   195
               Index           =   73
               Left            =   11190
               RightToLeft     =   -1  'True
               TabIndex        =   256
               Top             =   1275
               Width           =   1290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·ÿ·»"
               Height          =   195
               Index           =   72
               Left            =   11190
               RightToLeft     =   -1  'True
               TabIndex        =   254
               Top             =   795
               Width           =   1290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·«„—"
               Height          =   195
               Index           =   70
               Left            =   11190
               RightToLeft     =   -1  'True
               TabIndex        =   251
               Top             =   375
               Width           =   1290
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5445
            Index           =   19
            Left            =   13905
            TabIndex        =   112
            TabStop         =   0   'False
            Top             =   45
            Width           =   12570
            _cx             =   22172
            _cy             =   9604
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic2 
               Height          =   1560
               Left            =   4185
               TabIndex        =   184
               TabStop         =   0   'False
               Top             =   1245
               Width           =   8265
               _cx             =   14579
               _cy             =   2752
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
               Begin VB.TextBox TxtNoteSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   525
                  Left            =   1050
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   187
                  Top             =   165
                  Width           =   2370
               End
               Begin VB.TextBox Txtnots2 
                  Alignment       =   1  'Right Justify
                  Height          =   495
                  Left            =   1050
                  RightToLeft     =   -1  'True
                  TabIndex        =   186
                  Top             =   780
                  Width           =   2355
               End
               Begin ImpulseButton.ISButton Cmd 
                  CausesValidation=   0   'False
                  Height          =   480
                  Index           =   10
                  Left            =   0
                  TabIndex        =   188
                  Top             =   165
                  Width           =   1065
                  _ExtentX        =   1879
                  _ExtentY        =   847
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
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "»‰«¡ ⁄·Ï  ›« Ê—Â —ﬁ„"
                  ForeColor       =   &H00000000&
                  Height          =   420
                  Left            =   3405
                  TabIndex        =   190
                  Top             =   780
                  Width           =   1320
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·ﬁÌœ"
                  ForeColor       =   &H00000000&
                  Height          =   165
                  Left            =   3420
                  TabIndex        =   189
                  Top             =   555
                  Width           =   1320
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "»Ì«‰«  ﬁÌœ «·”‰œ"
                  Height          =   510
                  Index           =   51
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   185
                  Top             =   90
                  Width           =   1515
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic1 
               Height          =   960
               Left            =   4140
               TabIndex        =   178
               TabStop         =   0   'False
               Top             =   90
               Width           =   8325
               _cx             =   14684
               _cy             =   1693
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
               Begin VB.TextBox TxtExtraValue 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   135
                  RightToLeft     =   -1  'True
                  TabIndex        =   179
                  Top             =   330
                  Width           =   1740
               End
               Begin MSDataListLib.DataCombo DCExtraAccount 
                  Height          =   315
                  Left            =   2685
                  TabIndex        =   180
                  Top             =   330
                  Width           =   2865
                  _ExtentX        =   5054
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
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "Õ”«»«  «·«÷«›« "
                  Height          =   210
                  Index           =   48
                  Left            =   4815
                  RightToLeft     =   -1  'True
                  TabIndex        =   183
                  Top             =   90
                  Width           =   1530
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Õ”«»"
                  Height          =   210
                  Index           =   58
                  Left            =   5610
                  RightToLeft     =   -1  'True
                  TabIndex        =   182
                  Top             =   330
                  Width           =   735
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   150
                  Index           =   57
                  Left            =   1875
                  RightToLeft     =   -1  'True
                  TabIndex        =   181
                  Top             =   330
                  Width           =   675
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5445
            Index           =   15
            Left            =   13605
            TabIndex        =   105
            TabStop         =   0   'False
            Top             =   45
            Width           =   12570
            _cx             =   22172
            _cy             =   9604
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
            _GridInfo       =   $"FrmOut.frx":1D70
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   5415
               Index           =   18
               Left            =   15
               TabIndex        =   115
               TabStop         =   0   'False
               Top             =   15
               Width           =   12540
               _cx             =   22119
               _cy             =   9551
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
               Begin VB.TextBox Text7 
                  Alignment       =   1  'Right Justify
                  Height          =   465
                  Left            =   22935
                  RightToLeft     =   -1  'True
                  TabIndex        =   167
                  Top             =   1050
                  Width           =   1935
               End
               Begin VB.TextBox Text8 
                  Alignment       =   1  'Right Justify
                  Height          =   480
                  Left            =   22935
                  RightToLeft     =   -1  'True
                  TabIndex        =   170
                  Top             =   1740
                  Width           =   1935
               End
               Begin VB.TextBox Text5 
                  Alignment       =   1  'Right Justify
                  Height          =   465
                  Left            =   22935
                  RightToLeft     =   -1  'True
                  TabIndex        =   164
                  Top             =   345
                  Width           =   1935
               End
               Begin VB.TextBox TxtTaxServiceValue 
                  Alignment       =   1  'Right Justify
                  Height          =   0
                  Left            =   180
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   119
                  Top             =   435
                  Width           =   0
               End
               Begin VB.CheckBox ChkTaxSerivce 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì… Œœ„…"
                  Height          =   0
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   116
                  Top             =   675
                  Width           =   0
               End
               Begin MSDataListLib.DataCombo DCboStoreName2 
                  Height          =   315
                  Left            =   11490
                  TabIndex        =   165
                  Top             =   345
                  Width           =   7305
                  _ExtentX        =   12885
                  _ExtentY        =   556
                  _Version        =   393216
                  ListField       =   "7"
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCCar 
                  Height          =   315
                  Left            =   11490
                  TabIndex        =   168
                  Top             =   1050
                  Width           =   7305
                  _ExtentX        =   12885
                  _ExtentY        =   556
                  _Version        =   393216
                  ListField       =   "7"
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCDriver 
                  Height          =   315
                  Left            =   11490
                  TabIndex        =   171
                  Top             =   1740
                  Width           =   7305
                  _ExtentX        =   12885
                  _ExtentY        =   556
                  _Version        =   393216
                  ListField       =   "7"
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker txtManualDate 
                  Height          =   315
                  Left            =   9360
                  TabIndex        =   241
                  Top             =   150
                  Width           =   1560
                  _ExtentX        =   2752
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   193986561
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker txtRegDate 
                  Height          =   315
                  Left            =   6000
                  TabIndex        =   243
                  Top             =   150
                  Width           =   1560
                  _ExtentX        =   2752
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   193986561
                  CurrentDate     =   38784
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· ”ÃÌ·"
                  Height          =   300
                  Index           =   69
                  Left            =   6495
                  TabIndex        =   244
                  Top             =   240
                  Width           =   2445
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· «—ÌŒ «·ÌœÊÏ"
                  Height          =   300
                  Index           =   68
                  Left            =   9855
                  TabIndex        =   242
                  Top             =   120
                  Width           =   2445
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ì"
                  Height          =   270
                  Index           =   45
                  Left            =   18855
                  RightToLeft     =   -1  'True
                  TabIndex        =   176
                  Top             =   345
                  Width           =   915
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”«∆ﬁ"
                  Height          =   240
                  Index           =   41
                  Left            =   18855
                  RightToLeft     =   -1  'True
                  TabIndex        =   172
                  Top             =   1695
                  Width           =   915
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„⁄œÂ/«·”Ì«—…"
                  Height          =   255
                  Index           =   4
                  Left            =   18855
                  RightToLeft     =   -1  'True
                  TabIndex        =   169
                  Top             =   1080
                  Width           =   915
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ï «·„Œ“‰"
                  Height          =   255
                  Index           =   65
                  Left            =   24510
                  RightToLeft     =   -1  'True
                  TabIndex        =   166
                  Top             =   375
                  Width           =   1890
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   1545
                  Index           =   54
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   127
                  Top             =   435
                  Width           =   60
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
                  Height          =   1545
                  Index           =   47
                  Left            =   180
                  RightToLeft     =   -1  'True
                  TabIndex        =   123
                  Top             =   435
                  Width           =   60
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   2025
                  Index           =   43
                  Left            =   180
                  RightToLeft     =   -1  'True
                  TabIndex        =   120
                  Top             =   435
                  Width           =   60
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   105
               Index           =   16
               Left            =   15
               TabIndex        =   113
               TabStop         =   0   'False
               Top             =   585
               Width           =   12540
               _cx             =   22119
               _cy             =   185
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
                  Left            =   180
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   118
                  Top             =   15
                  Width           =   0
               End
               Begin VB.CheckBox ChkTaxAdd 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì»… Œ’„ Ê≈÷«›… (√—»«Õ  Ã«—Ì…)"
                  Height          =   90
                  Left            =   180
                  RightToLeft     =   -1  'True
                  TabIndex        =   114
                  Top             =   0
                  Width           =   60
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   45
                  Index           =   52
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   126
                  Top             =   15
                  Width           =   60
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
                  Left            =   180
                  RightToLeft     =   -1  'True
                  TabIndex        =   122
                  Top             =   15
                  Width           =   60
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   60
                  Index           =   39
                  Left            =   180
                  RightToLeft     =   -1  'True
                  TabIndex        =   117
                  Top             =   15
                  Width           =   60
               End
            End
            Begin VB.TextBox TxtBillComment1 
               Alignment       =   1  'Right Justify
               Height          =   3045
               Left            =   15
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   106
               Top             =   2385
               Width           =   12540
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
               Height          =   3045
               Index           =   44
               Left            =   15
               RightToLeft     =   -1  'True
               TabIndex        =   121
               Top             =   2385
               Visible         =   0   'False
               Width           =   12540
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5445
            Index           =   7
            Left            =   45
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   45
            Width           =   12570
            _cx             =   22172
            _cy             =   9604
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
            _GridInfo       =   $"FrmOut.frx":1DDF
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1185
               Index           =   2
               Left            =   30
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   30
               Width           =   12405
               _cx             =   21881
               _cy             =   2090
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
               Begin VB.TextBox TxtItemCodeB1 
                  Alignment       =   1  'Right Justify
                  Height          =   288
                  Left            =   9315
                  TabIndex        =   237
                  Top             =   150
                  Width           =   1728
               End
               Begin VB.TextBox TxtShortName 
                  Height          =   312
                  Left            =   120
                  TabIndex        =   236
                  Top             =   105
                  Width           =   6840
               End
               Begin VB.TextBox TxtItemsIDes 
                  Alignment       =   1  'Right Justify
                  Height          =   348
                  Left            =   11370
                  TabIndex        =   195
                  Top             =   -264
                  Visible         =   0   'False
                  Width           =   1710
               End
               Begin VB.ComboBox CboItemCase 
                  BackColor       =   &H0000FFFF&
                  Height          =   315
                  Left            =   4275
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   16
                  Top             =   795
                  Width           =   1080
               End
               Begin VB.TextBox TxtQuantity 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H0000FFFF&
                  Enabled         =   0   'False
                  Height          =   384
                  Left            =   1605
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   18
                  Top             =   795
                  Width           =   900
               End
               Begin VB.TextBox TxtSerial 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H0000FFFF&
                  Enabled         =   0   'False
                  Height          =   384
                  Left            =   2970
                  MaxLength       =   20
                  RightToLeft     =   -1  'True
                  TabIndex        =   17
                  Top             =   795
                  Width           =   1305
               End
               Begin VB.TextBox TxtPrice 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H0000FFFF&
                  Height          =   384
                  Left            =   435
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   19
                  Top             =   795
                  Width           =   1104
               End
               Begin MSDataListLib.DataCombo DCboItemsName 
                  Height          =   315
                  Left            =   5400
                  TabIndex        =   15
                  Top             =   795
                  Width           =   5145
                  _ExtentX        =   9075
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   65535
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
                  Height          =   285
                  Left            =   10575
                  TabIndex        =   14
                  Top             =   795
                  Width           =   1800
                  _ExtentX        =   3175
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   65535
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdAdd 
                  Height          =   390
                  Left            =   -45
                  TabIndex        =   20
                  Top             =   795
                  Width           =   360
                  _ExtentX        =   635
                  _ExtentY        =   688
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
                  ButtonImage     =   "FrmOut.frx":1E80
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
                  Height          =   315
                  Left            =   2565
                  TabIndex        =   56
                  Top             =   825
                  Width           =   405
                  _ExtentX        =   714
                  _ExtentY        =   556
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
                  ButtonImage     =   "FrmOut.frx":221A
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·»«—ﬂÊœ"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000000&
                  Height          =   270
                  Index           =   95
                  Left            =   10905
                  TabIndex        =   239
                  Top             =   105
                  Width           =   1410
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·»ÕÀ «·”—Ì⁄"
                  Height          =   330
                  Index           =   97
                  Left            =   7410
                  TabIndex        =   238
                  Top             =   105
                  Width           =   1260
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·’‰›"
                  Height          =   270
                  Index           =   31
                  Left            =   10920
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   555
                  Width           =   1110
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈”„ «·’‰›"
                  Height          =   270
                  Index           =   30
                  Left            =   7800
                  RightToLeft     =   -1  'True
                  TabIndex        =   44
                  Top             =   525
                  Width           =   1320
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·… «·’‰›"
                  Height          =   270
                  Index           =   29
                  Left            =   4425
                  RightToLeft     =   -1  'True
                  TabIndex        =   43
                  Top             =   525
                  Width           =   1245
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ì—Ì«·"
                  Height          =   270
                  Index           =   28
                  Left            =   3180
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   525
                  Width           =   825
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂ„Ì…"
                  Height          =   270
                  Index           =   27
                  Left            =   1875
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   555
                  Width           =   690
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· ﬂ·›…"
                  Height          =   270
                  Index           =   26
                  Left            =   855
                  RightToLeft     =   -1  'True
                  TabIndex        =   40
                  Top             =   525
                  Width           =   630
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid FG 
               Height          =   3330
               Left            =   30
               TabIndex        =   13
               Top             =   1230
               Width           =   12420
               _cx             =   21907
               _cy             =   5874
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
               Cols            =   26
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmOut.frx":25B4
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
               Left            =   510
               TabIndex        =   54
               Top             =   4575
               Width           =   11460
               _ExtentX        =   20214
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
               Height          =   375
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   4575
               Width           =   465
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5445
            Index           =   5
            Left            =   13305
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   45
            Width           =   12570
            _cx             =   22172
            _cy             =   9604
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
            _GridInfo       =   $"FrmOut.frx":29EA
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   2550
               Index           =   10
               Left            =   0
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   2895
               Width           =   12570
               _cx             =   22172
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
               _GridInfo       =   $"FrmOut.frx":2A5A
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   2265
                  Index           =   14
                  Left            =   15
                  TabIndex        =   98
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   2160
                  _cx             =   3810
                  _cy             =   3995
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
                     Height          =   13470
                     Index           =   2
                     Left            =   945
                     RightToLeft     =   -1  'True
                     TabIndex        =   99
                     Top             =   2985
                     Width           =   135
                  End
                  Begin ImpulseButton.ISButton CmdCheque 
                     Height          =   13470
                     Left            =   300
                     TabIndex        =   109
                     Top             =   2985
                     Width           =   165
                     _ExtentX        =   291
                     _ExtentY        =   23760
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
                     Height          =   13470
                     Index           =   19
                     Left            =   750
                     RightToLeft     =   -1  'True
                     TabIndex        =   111
                     Top             =   2985
                     Width           =   75
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
                     Height          =   13470
                     Index           =   17
                     Left            =   840
                     RightToLeft     =   -1  'True
                     TabIndex        =   110
                     Top             =   2985
                     Width           =   105
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
                     Height          =   13470
                     Index           =   16
                     Left            =   555
                     RightToLeft     =   -1  'True
                     TabIndex        =   101
                     Top             =   2985
                     Width           =   195
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   13470
                     Index           =   18
                     Left            =   465
                     RightToLeft     =   -1  'True
                     TabIndex        =   100
                     Top             =   2985
                     Width           =   90
                  End
               End
               Begin VSFlex8UCtl.VSFlexGrid FgCheques 
                  Height          =   2265
                  Left            =   15
                  TabIndex        =   59
                  Top             =   15
                  Width           =   2160
                  _cx             =   3810
                  _cy             =   3995
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
                  FormatString    =   $"FrmOut.frx":2ACE
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
               Height          =   2745
               Index           =   6
               Left            =   0
               TabIndex        =   66
               TabStop         =   0   'False
               Top             =   150
               Width           =   12570
               _cx             =   22172
               _cy             =   4842
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
               _GridInfo       =   $"FrmOut.frx":2C02
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VSFlex8UCtl.VSFlexGrid FgInstallments 
                  Height          =   2430
                  Left            =   15
                  TabIndex        =   67
                  Top             =   15
                  Width           =   2265
                  _cx             =   3995
                  _cy             =   4286
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
                  FormatString    =   $"FrmOut.frx":2C70
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
                  Height          =   30
                  Index           =   13
                  Left            =   15
                  TabIndex        =   68
                  TabStop         =   0   'False
                  Top             =   2460
                  Width           =   2265
                  _cx             =   3995
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
                     Left            =   45
                     RightToLeft     =   -1  'True
                     TabIndex        =   108
                     Top             =   60
                     Visible         =   0   'False
                     Width           =   195
                  End
                  Begin VB.Label LblStartValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   225
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   107
                     Top             =   60
                     Width           =   45
                  End
                  Begin VB.Label LblInstallSeprator 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     ForeColor       =   &H00FF0000&
                     Height          =   225
                     Left            =   360
                     RightToLeft     =   -1  'True
                     TabIndex        =   104
                     Top             =   60
                     Width           =   45
                  End
                  Begin VB.Label LblPrecenValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   225
                     Left            =   1365
                     RightToLeft     =   -1  'True
                     TabIndex        =   103
                     Top             =   60
                     Width           =   45
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
                     Left            =   1425
                     RightToLeft     =   -1  'True
                     TabIndex        =   102
                     Top             =   60
                     Visible         =   0   'False
                     Width           =   75
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
                     Left            =   1620
                     RightToLeft     =   -1  'True
                     TabIndex        =   78
                     Top             =   60
                     Visible         =   0   'False
                     Width           =   135
                  End
                  Begin VB.Label LblPrecenType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   225
                     Left            =   1500
                     RightToLeft     =   -1  'True
                     TabIndex        =   77
                     Top             =   60
                     Width           =   120
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
                     Left            =   1200
                     RightToLeft     =   -1  'True
                     TabIndex        =   76
                     Top             =   60
                     Visible         =   0   'False
                     Width           =   165
                  End
                  Begin VB.Label LblInstallTotal 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   225
                     Left            =   1095
                     RightToLeft     =   -1  'True
                     TabIndex        =   75
                     Top             =   60
                     Width           =   105
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
                     Left            =   930
                     RightToLeft     =   -1  'True
                     TabIndex        =   74
                     Top             =   60
                     Visible         =   0   'False
                     Width           =   165
                  End
                  Begin VB.Label LblInstallCount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   225
                     Left            =   855
                     RightToLeft     =   -1  'True
                     TabIndex        =   73
                     Top             =   60
                     Width           =   60
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
                     Left            =   735
                     RightToLeft     =   -1  'True
                     TabIndex        =   72
                     Top             =   60
                     Visible         =   0   'False
                     Width           =   120
                  End
                  Begin VB.Label LblFirstInstallDate 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   225
                     Left            =   585
                     RightToLeft     =   -1  'True
                     TabIndex        =   71
                     Top             =   60
                     Width           =   135
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
                     Left            =   405
                     RightToLeft     =   -1  'True
                     TabIndex        =   70
                     Top             =   60
                     Visible         =   0   'False
                     Width           =   180
                  End
                  Begin VB.Label LblInstallmentType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     ForeColor       =   &H00FF0000&
                     Height          =   225
                     Left            =   240
                     RightToLeft     =   -1  'True
                     TabIndex        =   69
                     Top             =   60
                     Width           =   120
                  End
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   2430
                  Index           =   12
                  Left            =   15
                  TabIndex        =   79
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   2265
                  _cx             =   3995
                  _cy             =   4286
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
                     Height          =   18675
                     Left            =   135
                     RightToLeft     =   -1  'True
                     TabIndex        =   83
                     Top             =   810
                     Width           =   120
                  End
                  Begin VB.TextBox XPTxtSerial 
                     Alignment       =   1  'Right Justify
                     Height          =   18675
                     Index           =   1
                     Left            =   570
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   82
                     Top             =   1620
                     Width           =   120
                  End
                  Begin VB.TextBox XPTxtValue 
                     Alignment       =   1  'Right Justify
                     Height          =   18675
                     Index           =   1
                     Left            =   780
                     MaxLength       =   10
                     RightToLeft     =   -1  'True
                     TabIndex        =   81
                     Top             =   1620
                     Width           =   105
                  End
                  Begin VB.CheckBox XPChkPayType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "¬Ã· "
                     Height          =   17055
                     Index           =   1
                     Left            =   960
                     RightToLeft     =   -1  'True
                     TabIndex        =   80
                     Top             =   1620
                     Width           =   120
                  End
                  Begin ImpulseButton.ISButton CmdINSTALLMENT 
                     Height          =   22725
                     Left            =   15
                     TabIndex        =   84
                     Top             =   -810
                     Width           =   135
                     _ExtentX        =   238
                     _ExtentY        =   40084
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
                     ButtonImage     =   "FrmOut.frx":2D41
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
                     Height          =   17865
                     Left            =   270
                     TabIndex        =   85
                     Top             =   1620
                     Width           =   150
                     _ExtentX        =   265
                     _ExtentY        =   31512
                     _Version        =   393216
                     Format          =   193724417
                     CurrentDate     =   38784
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
                     Height          =   15420
                     Index           =   21
                     Left            =   420
                     RightToLeft     =   -1  'True
                     TabIndex        =   88
                     Top             =   4065
                     Width           =   135
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁÌ„…"
                     Height          =   17850
                     Index           =   15
                     Left            =   885
                     RightToLeft     =   -1  'True
                     TabIndex        =   87
                     Top             =   4065
                     Width           =   60
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„”·”·"
                     Height          =   17040
                     Index           =   14
                     Left            =   705
                     RightToLeft     =   -1  'True
                     TabIndex        =   86
                     Top             =   4065
                     Width           =   60
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   150
               Index           =   11
               Left            =   0
               TabIndex        =   89
               TabStop         =   0   'False
               Top             =   0
               Width           =   12570
               _cx             =   22172
               _cy             =   265
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
                  Left            =   1905
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   92
                  Top             =   60
                  Width           =   210
               End
               Begin VB.TextBox XPTxtSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Index           =   0
                  Left            =   1425
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   91
                  Top             =   60
                  Width           =   270
               End
               Begin VB.CheckBox XPChkPayType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰ﬁœ«"
                  Height          =   345
                  Index           =   0
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   90
                  Width           =   270
               End
               Begin MSDataListLib.DataCombo DcboBox 
                  Height          =   315
                  Left            =   705
                  TabIndex        =   93
                  Top             =   105
                  Width           =   480
                  _ExtentX        =   847
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
                  Left            =   60
                  RightToLeft     =   -1  'True
                  TabIndex        =   97
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   285
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   345
                  Index           =   13
                  Left            =   2175
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   90
                  Width           =   120
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·”·"
                  Height          =   345
                  Index           =   12
                  Left            =   1695
                  RightToLeft     =   -1  'True
                  TabIndex        =   95
                  Top             =   90
                  Width           =   150
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·Œ“‰…"
                  Height          =   345
                  Index           =   11
                  Left            =   1170
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   90
                  Width           =   225
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   5445
            Left            =   14505
            TabIndex        =   301
            TabStop         =   0   'False
            Top             =   45
            Width           =   12570
            _cx             =   22172
            _cy             =   9604
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
               Height          =   2460
               Left            =   180
               TabIndex        =   302
               Tag             =   "1"
               Top             =   240
               Width           =   12270
               _cx             =   21643
               _cy             =   4339
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
               FormatString    =   $"FrmOut.frx":30DB
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
            Begin ImpulseButton.ISButton Accredit 
               Height          =   390
               Left            =   90
               TabIndex        =   305
               Top             =   5010
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   688
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
            Begin VB.Label Label1100 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
               Height          =   255
               Left            =   13080
               RightToLeft     =   -1  'True
               TabIndex        =   304
               Top             =   4560
               Width           =   3375
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
               Height          =   255
               Left            =   6780
               RightToLeft     =   -1  'True
               TabIndex        =   303
               Top             =   2790
               Width           =   3375
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
         Width           =   12660
         _cx             =   22331
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
         Caption         =   "”‰œ ’—› /«· ”·Ì„"
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
         Begin VB.CheckBox chkIsProplemOnly 
            BackColor       =   &H00FFFFFF&
            Caption         =   " „ «· ”·Ì„"
            Height          =   225
            Left            =   6810
            TabIndex        =   306
            Top             =   360
            Width           =   960
         End
         Begin VB.CheckBox chkIgnorDetails 
            Alignment       =   1  'Right Justify
            Caption         =   " Ã«Â· «· ›«’Ì·"
            Height          =   270
            Left            =   7065
            RightToLeft     =   -1  'True
            TabIndex        =   248
            Top             =   360
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.CheckBox chkWithoutCost 
            Caption         =   "»œÊ‰ Õ”«»   ﬂ·›…"
            Height          =   225
            Left            =   2760
            TabIndex        =   247
            Top             =   360
            Visible         =   0   'False
            Width           =   1725
         End
         Begin VB.CheckBox chkStore 
            Caption         =   "»«·„Œ“‰"
            Height          =   225
            Left            =   4500
            TabIndex        =   246
            Top             =   360
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.CheckBox chkIsPosOnly 
            Caption         =   "»‰ﬁ«ÿ «·»Ì⁄ "
            Height          =   225
            Left            =   4500
            TabIndex        =   240
            Top             =   0
            Value           =   1  'Checked
            Width           =   1125
         End
         Begin VB.CheckBox withoutJL 
            Caption         =   "»œÊ‰ ﬁÌÊœ"
            Height          =   225
            Left            =   4500
            TabIndex        =   235
            Top             =   180
            Width           =   1125
         End
         Begin VB.CheckBox chkDone 
            BackColor       =   &H00FFFFFF&
            Caption         =   " „ «· ”·Ì„"
            Height          =   225
            Left            =   5820
            TabIndex        =   234
            Top             =   360
            Width           =   960
         End
         Begin VB.CommandButton cmdReSave 
            Caption         =   "÷»ÿ «·Õ—ﬂ« "
            Height          =   285
            Left            =   9075
            TabIndex        =   202
            Top             =   0
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.TextBox txtPassword 
            Height          =   315
            IMEMode         =   3  'DISABLE
            Left            =   3720
            PasswordChar    =   "*"
            TabIndex        =   201
            Top             =   60
            Width           =   750
         End
         Begin VB.CheckBox chkIsBranch 
            Caption         =   "»«·›—⁄"
            Height          =   225
            Left            =   5745
            TabIndex        =   200
            Top             =   60
            Width           =   735
         End
         Begin VB.TextBox oldtxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1785
            RightToLeft     =   -1  'True
            TabIndex        =   139
            Top             =   0
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   285
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   125
            Top             =   30
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   285
            Left            =   1890
            RightToLeft     =   -1  'True
            TabIndex        =   124
            Top             =   30
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   285
            Left            =   1335
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   30
            Visible         =   0   'False
            Width           =   510
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   2010
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
            ButtonImage     =   "FrmOut.frx":321E
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
            Width           =   960
            _ExtentX        =   1693
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
            ButtonImage     =   "FrmOut.frx":35B8
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
            Left            =   2835
            TabIndex        =   38
            Top             =   30
            Width           =   915
            _ExtentX        =   1614
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
            ButtonImage     =   "FrmOut.frx":3952
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
            ButtonImage     =   "FrmOut.frx":3CEC
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
            Left            =   8025
            TabIndex        =   60
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
            ButtonImage     =   "FrmOut.frx":4086
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdRetruns 
            Height          =   345
            Left            =   8985
            TabIndex        =   61
            Top             =   120
            Width           =   870
            _ExtentX        =   1535
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
            ButtonImage     =   "FrmOut.frx":4420
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdInfo 
            Height          =   480
            Left            =   10275
            TabIndex        =   137
            Top             =   0
            Visible         =   0   'False
            Width           =   795
            _ExtentX        =   1402
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
            ButtonImage     =   "FrmOut.frx":49BA
            ButtonImageHover=   "FrmOut.frx":5694
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
         End
         Begin MSComCtl2.DTPicker txtFromDateReSave 
            Height          =   315
            Left            =   7665
            TabIndex        =   203
            Top             =   30
            Visible         =   0   'False
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            Format          =   193134593
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker txtToDateReSave 
            Height          =   315
            Left            =   6465
            TabIndex        =   204
            Top             =   30
            Visible         =   0   'False
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            Format          =   193134593
            CurrentDate     =   38784
         End
         Begin VB.Image ImgFavorites 
            Height          =   390
            Left            =   8670
            Picture         =   "FrmOut.frx":636E
            Stretch         =   -1  'True
            Top             =   0
            Width           =   420
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
            Left            =   75
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   870
            Width           =   7485
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
            Left            =   3825
            RightToLeft     =   -1  'True
            TabIndex        =   140
            Top             =   120
            Width           =   5250
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   435
         Index           =   3
         Left            =   15
         TabIndex        =   205
         TabStop         =   0   'False
         Top             =   8370
         Width           =   12660
         _cx             =   22331
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
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Height          =   375
            Left            =   11325
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   206
            TabStop         =   0   'False
            Top             =   30
            Visible         =   0   'False
            Width           =   450
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   2610
            TabIndex        =   207
            Top             =   30
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lblTotalSalesPrice 
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
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   233
            Top             =   0
            Visible         =   0   'False
            Width           =   1380
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
            Left            =   8940
            RightToLeft     =   -1  'True
            TabIndex        =   220
            Top             =   30
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œ’Ê„« "
            Height          =   285
            Index           =   50
            Left            =   9975
            RightToLeft     =   -1  'True
            TabIndex        =   219
            Top             =   75
            Width           =   585
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
            Left            =   10590
            RightToLeft     =   -1  'True
            TabIndex        =   218
            Top             =   30
            Width           =   1380
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«›Ì"
            Height          =   285
            Index           =   49
            Left            =   8430
            RightToLeft     =   -1  'True
            TabIndex        =   217
            Top             =   75
            Width           =   495
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
            Left            =   7005
            RightToLeft     =   -1  'True
            TabIndex        =   216
            Top             =   30
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„” Œœ„"
            Height          =   330
            Index           =   1
            Left            =   3990
            RightToLeft     =   -1  'True
            TabIndex        =   215
            Top             =   75
            Width           =   855
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   285
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   214
            Top             =   75
            Width           =   645
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   285
            Left            =   1110
            RightToLeft     =   -1  'True
            TabIndex        =   213
            Top             =   75
            Width           =   495
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”Ã·"
            Height          =   285
            Index           =   2
            Left            =   1860
            RightToLeft     =   -1  'True
            TabIndex        =   212
            Top             =   75
            Width           =   735
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "/"
            Height          =   285
            Index           =   0
            Left            =   705
            RightToLeft     =   -1  'True
            TabIndex        =   211
            Top             =   75
            Width           =   375
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·≈Ã„«·Ï"
            Height          =   285
            Index           =   3
            Left            =   12000
            RightToLeft     =   -1  'True
            TabIndex        =   210
            Top             =   75
            Width           =   630
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
            Left            =   4890
            TabIndex        =   209
            Top             =   0
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·ﬂ„ÌÂ"
            Height          =   375
            Index           =   63
            Left            =   5925
            TabIndex        =   208
            Top             =   60
            Visible         =   0   'False
            Width           =   945
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   540
         Index           =   1
         Left            =   15
         TabIndex        =   221
         TabStop         =   0   'False
         Top             =   8820
         Width           =   12660
         _cx             =   22331
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
         Begin VB.CommandButton sameCmd 
            Caption         =   "‰”Œ… „„«À·Â"
            Height          =   375
            Left            =   1725
            RightToLeft     =   -1  'True
            TabIndex        =   232
            Top             =   90
            Width           =   855
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   360
            Index           =   0
            Left            =   11280
            TabIndex        =   222
            Top             =   90
            Width           =   1305
            _ExtentX        =   2302
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
            Height          =   360
            Index           =   1
            Left            =   9855
            TabIndex        =   223
            Top             =   90
            Width           =   1320
            _ExtentX        =   2328
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
            Height          =   360
            Index           =   2
            Left            =   8445
            TabIndex        =   224
            Top             =   90
            Width           =   1320
            _ExtentX        =   2328
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
            Height          =   360
            Index           =   3
            Left            =   7140
            TabIndex        =   225
            Top             =   90
            Width           =   1215
            _ExtentX        =   2143
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
            Height          =   360
            Index           =   4
            Left            =   5565
            TabIndex        =   226
            Top             =   90
            Width           =   1425
            _ExtentX        =   2514
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
            Height          =   360
            Index           =   5
            Left            =   4230
            TabIndex        =   227
            Top             =   90
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   635
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
            Height          =   360
            Index           =   6
            Left            =   30
            TabIndex        =   228
            TabStop         =   0   'False
            Top             =   90
            Width           =   900
            _ExtentX        =   1588
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
            Height          =   360
            Index           =   7
            Left            =   2820
            TabIndex        =   229
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   635
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
            Height          =   360
            Left            =   2490
            TabIndex        =   230
            Top             =   90
            Visible         =   0   'False
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   635
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   360
            Index           =   9
            Left            =   930
            TabIndex        =   231
            TabStop         =   0   'False
            Top             =   90
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   635
            ButtonStyle     =   1
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
      Caption         =   "—ﬁ„ «·«„—"
      Height          =   195
      Index           =   71
      Left            =   10470
      RightToLeft     =   -1  'True
      TabIndex        =   252
      Top             =   3960
      Width           =   1290
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3720
      TabIndex        =   136
      Top             =   960
      Width           =   1050
   End
End
Attribute VB_Name = "FrmOut"
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
Dim OtherInformation As New ClsGLOther
Public BolPrint As Boolean
Public invoiceSerach As Boolean
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
Dim Line1 As Double
Dim Line2 As Double
Dim DebitAccount As String
Dim CreditAccount As String
Dim MintDone As Integer
 Dim mClicked As Boolean
Dim s As String
Dim mIsFinishSave As Boolean
Dim IsSaveWithOutMsg As Boolean
Dim mIsStart As Boolean
Sub SerchItems(Optional str As String)
 
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
sql = " select  ItemID,barCodeNO   from  dbo.TblItems where 1=1"
If SystemOptions.UserInterface = ArabicInterface Then
SQL1 = " select  ItemID,ItemName   from  dbo.TblItems where 1=1"
Else
SQL1 = " select  ItemID,ItemNamee   from  dbo.TblItems where 1=1"
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
  
                        If str = "" Then
                                 sql = " select  ItemID,barCodeNO   from  dbo.TblItems where 1=1"
                                 If SystemOptions.UserInterface = ArabicInterface Then
                                 SQL1 = " select  ItemID,ItemName   from  dbo.TblItems where 1=1"
                                     SQL1 = SQL1 + " Order BY ItemName "
                                 Else
                                 SQL1 = " select  ItemID,ItemNamee   from  dbo.TblItems where 1=1"
                                     SQL1 = SQL1 + " Order BY ItemNameE "
                                 End If
                                 
                                     fill_combo DCboItemsCode, sql
                                         fill_combo DCboItemsName, SQL1
                End If
                
       Exit Sub
       
If str <> "" Then
'Dim Sql As String
'Dim StrWhere As String
'  Dim astrSplit2tems2() As String
'  Dim j As Integer
'  Dim nElements As Integer
StrWhere = ""
If SystemOptions.UserInterface = ArabicInterface Then
sql = " select  ItemID,ItemName   from  dbo.TblItems where 1=1"
Else
sql = " select  ItemID,ItemNamee   from  dbo.TblItems where 1=1"
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

Private Sub Accredit_Click()
    Dim BeginTrans As Boolean
If val(XPTxtBillID.text) = 0 Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "«Õ›Ÿ «·”‰œ «Ê·«", vbCritical
     Else
     MsgBox "Save Doc First", vbCritical
     End If
      
      Exit Sub
      End If
 
 
    SendTopost Me.Name, "Transactions", "Transaction_ID", 0, val(dcBranch.BoundText), val(XPTxtBillID.text), TxtNoteSerial1.text
  rs.Resync
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
    Accredit.Caption = "Sent To Approval "
End If
    Retrive (val(Me.XPTxtBillID.text))
End Sub

Private Sub DCboItemsCode_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
DCboItemsName.SetFocus
End If
End Sub

Private Sub DCboItemsName_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
TxtQuantity.SetFocus
End If
End Sub

Private Sub Text17_Change()

End Sub

Private Sub TxtItemCodeB12_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
 '   On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    Dim LngItemID As Long
    Dim LngUnitID As Long
  Dim ColorID As Integer
   Dim sizeid As Integer
    Dim ClassId As Integer
    Dim ParrtNoCode As String
    Dim ItemDetailedCode As String

ItemDetailedCode = Replace(TxtItemCodeB12, "*", "")
Dim mArr As Variant

mArr = Split(ItemDetailedCode, " ")
ItemDetailedCode = mArr(0)
    StrSQL = " SELECT     * from TblDefComItem where Id  = " & val(ItemDetailedCode)
     
     
       
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        TXT_order_no = Trim(RsTemp!NoteSerial13 & "")
        
    End If
'       StrSQL = " SELECT     * from TblDefComItem where Id  = " & val(ItemDetailedCode)
'        StrSQL = " Select tblItems.barcodeno, LineID,TblDefComItemData.ItemID,Qty from TblDefComItemData Inner join tblItems On TblDefComItemData.ItemId = tblItems.ItemId Where LineID = " & val(mArr(1)) & " and IDDefCIT = " & val(mArr(0))
'        Set RsTemp = New ADODB.Recordset
'        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'        If Not (RsTemp.EOF Or RsTemp.BOF) Then
'            DCboItemsName.BoundText = val(RsTemp!ItemID & "")
'            TxtItemCodeB1 = Trim(RsTemp!barCodeNO & "")
'            TxtItemCodeB1.SetFocus
'            'TxtItemCodeB1_KeyDown KeyCode, Shift
'
'            TxtQuantity = val(RsTemp!Qty & "")
'            cmdAdd_Click
'        End If
     
    
 End If
  '
End Sub

Private Sub TxtPrice_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
CmdAdd.SetFocus
End If
End Sub

Private Sub TxtQuantity_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
'TxtPrice.SetFocus
End If
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


Private Sub cmdReSave_Click()

    XPBtnMove_Click (2) 'ÌÃ„⁄ «·”‰œ« 
    DoEvents
    
    If rs Is Nothing Then Exit Sub
    If rs.EOF Or rs.BOF Then Exit Sub
    
    IsSaveWithOutMsg = True
    
    '«»œ√ „‰ ¬Œ— ”‰œ (√‰  »«·›⁄· ⁄«„· MoveLast œ«Œ· XPBtnMove_Click(2))
    Do While Not rs.BOF And Not rs.EOF
        
        '«œŒ· Ê÷⁄ «· ⁄œÌ· ·Ê “— «· ⁄œÌ· »Ì⁄„· œÂ
        Cmd_Click (1)
        DoEvents
        
        If chkWithoutCost.value = vbUnchecked Then
            'NewGrid.DtpBillDate_Change   '≈⁄«œ… Õ”«» «· ﬂ·›…
        End If
        If TxtNoteSerial1.text = "42506054" Then
                IsSaveWithOutMsg = IsSaveWithOutMsg
        End If
         NewGrid.DtpBillDate_Change   '≈⁄«œ… Õ”«» «· ﬂ·›…
        DoEvents
        Cmd_Click (2) 'Save
        DoEvents
        
        rs.MovePrevious
        If rs.BOF Then Exit Do
        
        '„Â„: —Ã¯⁄ ‘«‘… «·⁄—÷ ··”Ã· «·Õ«·Ì ·Ê ⁄‰œﬂ Retrive »Ì ⁄„· »⁄œ «·Õ—ﬂ…
        XPBtnMove_Click (0)
        DoEvents
        
    Loop
    
    IsSaveWithOutMsg = False
    MsgBox " „ «·Õ›Ÿ"

End Sub


Private Sub chkDone_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
         
If SystemOptions.UserInterface = ArabicInterface Then
        X = MsgBox("·« Ì„ﬂ‰  ⁄œÌ· Õ«·Â Â–« «·”‰œ „—Â  «Œ—Ì    √ﬂÌœ", vbCritical + vbYesNo)
 Else
        X = MsgBox("Confirm lock y/n", vbCritical + vbYesNo)
 End If
 
 If X = vbYes Then
  Cn.Execute "update Transactions set  chkDone=1 where  Transaction_ID=" & val(XPTxtBillID)
              
              If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "      „  ⁄œÌ· Õ«·Â Â–« «·”‰œ »‰Ã«Õ", vbInformation
                Else
                          MsgBox "  Done", vbInformation
            End If

 
 End If
 chkDone.Enabled = False
 chkDone.value = vbChecked
rs.Resync adAffectCurrent
 
 
End Sub

Private Sub cmdReSave_ClickOld()
  
    Dim s         As String
    Dim rsDummy   As ADODB.Recordset
    Dim mBranchID As Integer
    
    XPBtnMove_Click (2)
    DoEvents
 
    Dim i As Double
    For i = 1 To rs.RecordCount
        IsSaveWithOutMsg = True
        Cmd_Click (1)
        DoEvents
        If chkWithoutCost.value = vbUnchecked Then
             NewGrid.DtpBillDate_Change
        End If
   
        DoEvents
        Cmd_Click (2)
         
        XPBtnMove_Click (0)
        DoEvents
    Next i
  
    IsSaveWithOutMsg = False
    MsgBox " „ «·Õ›Ÿ"

End Sub
 
Private Sub txtPassword_Change()
If Trim(txtPassword) = "Alex2025" Then
    cmdReSave.Visible = True
    txtFromDateReSave.Visible = True
    txtToDateReSave.Visible = True
    chkIsBranch.Visible = True
    withoutJL.Visible = True
    chkStore.Visible = True
    chkWithoutCost.Visible = True
    chkIgnorDetails.Visible = True
        chkIgnorDetails.value = 1
Else
    withoutJL.Visible = False
    cmdReSave.Visible = False
    txtFromDateReSave.Visible = False
    txtToDateReSave.Visible = False
   chkIsBranch.Visible = False
   chkStore.Visible = False
    chkWithoutCost.Visible = False
    chkIgnorDetails.Visible = False
       
End If
txtFromDateReSave.value = Date
txtToDateReSave.value = Date
End Sub
 
Public Function generalSearch(StrSQL As String)
rs.Close
Set rs = Nothing


    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
          If Not (rs.BOF Or rs.EOF) Then
                rs.MoveLast
            End If
   Me.TxtModFlg.text = "R"
            Retrive
          
            Me.TxtModFlg.text = "R"
End Function
 
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


Public Sub RetriveOrder(Optional order_no As String = "", _
                        Optional Transaction_Type As Integer = 0)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    Dim StoreId2 As Double
    Dim issuedQty As Double
    Dim rsDummy As New ADODB.Recordset
    Dim mCostPrice As Double
    Dim s As String
   On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh

   
        StrSQL = "Select * from transactions where  Transaction_Type=" & Transaction_Type & " and noteserial1='" & order_no & "'"
 

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount < 1 Then
 
        Exit Sub
    Else
        DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
       ' Me.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
        TxtBillComment.text = IIf(IsNull(rs("TransactionComment").value), "", (rs("TransactionComment").value))
        If Transaction_Type = 21 And SystemOptions.MultyStore = True Then
        
        Else
            Me.DCboStoreName.BoundText = IIf(IsNull(rs("storeid").value), "", rs("storeid").value)
        End If

        Me.dcBranch.BoundText = IIf(IsNull(rs("Branchid").value), "", rs("Branchid").value)
        TxtOldOpOrderID.text = IIf(IsNull(rs("OldOpOrderID").value), "", rs("OldOpOrderID").value)
        TxtCashCustomerName.text = IIf(IsNull(rs("CashCustomerName").value), "", rs("CashCustomerName").value)
        DBCboClientName.BoundText = IIf(IsNull(rs("Cusid").value), "", rs("Cusid").value)
        
             
        If Trim(rs!TransactionComment & "") <> "" Then
            TxtBillComment.text = IIf(IsNull(rs("TransactionComment").value), "", (rs("TransactionComment").value))
        Else
            TxtBillComment.text = IIf(IsNull(rs("Remark").value), "", (rs("Remark").value))
        End If
        
        'txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If
    txtOrderID = rs!Transaction_ID & ""


    Screen.MousePointer = vbArrowHourglass
    'if SystemOptions.InsertItemManualOut Then Exit Sub
    
    
    
If CBoBasedON.ListIndex = 13 Then
    StrSQL = "SELECT TblItems.HaveSerial, dbo.[GetBalanceQtyPO5] (Transaction_Details.Item_ID,N'" & Trim(order_no) & "'," & val(Me.XPTxtBillID) & ",  19,13) as Showqty6 , * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
End If
    'StrSQL = StrSQL + " where Transaction_ID=" & Transaction_ID
 'If str = "" And Transaction_ID <> 0 Then
   '  StrSQL = StrSQL + " where Transaction_ID=" & Transaction_ID
 'ElseIf str <> "" Then
 '   StrSQL = StrSQL + " where Transaction_ID in (" & str & ")"
'Else
'Exit Sub
' End If
 If CBoBasedON.ListIndex = 13 Then
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)
    StrSQL = StrSQL + " And dbo.[GetBalanceQtyPO5] (Transaction_Details.Item_ID,N'" & Trim(order_no) & "'," & val(Me.XPTxtBillID) & ",19,13) <> 0"
Else

    
    
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)
End If
If Transaction_Type = 21 And SystemOptions.MultyStore = True Then
    If val(Me.DCboStoreName.BoundText) <> 0 And Trim(Me.DCboStoreName.BoundText) <> "" Then
        StrSQL = StrSQL + " and  Transaction_Details.StoreID2=" & val(Me.DCboStoreName.BoundText)
    End If
End If

StrSQL = StrSQL & " order by Transaction_Details.id "

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            StoreId2 = val(DCboStoreName.BoundText)
            
            
    
            
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))
            
 If Transaction_Type = 38 Then
 If SystemOptions.poWithatotalQty = True Then
 FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), 0, (RsDetails("showqty").value))
 Else
FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("ItemBalance")), 0, (RsDetails("ItemBalance").value)) - IIf(IsNull(RsDetails("showqty")), 0, (RsDetails("showqty").value))
End If

End If

If val(RsDetails("OldID") & "") = 0 Then
    FG.TextMatrix(Num, FG.ColIndex("OldID")) = IIf(IsNull(RsDetails("ID")), 0, (RsDetails("ID").value))
Else
    FG.TextMatrix(Num, FG.ColIndex("OldID")) = IIf(IsNull(RsDetails("OldID")), 0, (RsDetails("OldID").value))
End If
FG.TextMatrix(Num, FG.ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks")), "", (RsDetails("Remarks").value))
issuedQty = GetIssuedQty(TXT_order_no, val(Me.XPTxtBillID), StoreId2, val(FG.TextMatrix(Num, FG.ColIndex("Code"))), val(FG.TextMatrix(Num, FG.ColIndex("OldID"))))

 If Transaction_Type = 21 Then
 
 FG.TextMatrix(Num, FG.ColIndex("TotalInvoiceQty")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))
 FG.TextMatrix(Num, FG.ColIndex("ISSUEDQTY")) = issuedQty
  FG.TextMatrix(Num, FG.ColIndex("Count")) = val(FG.TextMatrix(Num, FG.ColIndex("TotalInvoiceQty"))) - val(FG.TextMatrix(Num, FG.ColIndex("ISSUEDQTY")))

End If


issuedQty = GetIssuedQty2(TXT_order_no, val(Me.XPTxtBillID), StoreId2, val(FG.TextMatrix(Num, FG.ColIndex("Code"))), val(FG.TextMatrix(Num, FG.ColIndex("OldID"))))

 If Transaction_Type = 29 Then
 
 FG.TextMatrix(Num, FG.ColIndex("TotalInvoiceQty")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))
 FG.TextMatrix(Num, FG.ColIndex("ISSUEDQTY")) = issuedQty
 FG.TextMatrix(Num, FG.ColIndex("StillQty")) = issuedQty
  FG.TextMatrix(Num, FG.ColIndex("Count")) = val(FG.TextMatrix(Num, FG.ColIndex("TotalInvoiceQty"))) - val(FG.TextMatrix(Num, FG.ColIndex("ISSUEDQTY")))

End If

            'FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
         '   If Transaction_Type = 0 Then
                'FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("ShowPrice")), 0, (RsDetails("ShowPrice").value)) ' GET_COST_PRICE_FOR_PRODUCT_ITEM(Val(FG.TextMatrix(Num, FG.ColIndex("Code"))))
         '   End If
      
       
            '  FG.TextMatrix(Num, FG.ColIndex("Expenses")) = IIf(IsNull(RsDetails("Lineexpenses")), "", (RsDetails("Lineexpenses").value))
         
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
          FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = 1 ' IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
           FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = 0  '(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemType")) = IIf(IsNull(RsDetails("ItemType")), 0, (RsDetails("ItemType").value))
         
         
            FG.TextMatrix(Num, FG.ColIndex("length")) = IIf(IsNull(RsDetails("length")), "", (RsDetails("length").value))
             FG.TextMatrix(Num, FG.ColIndex("Height")) = IIf(IsNull(RsDetails("Height")), "", (RsDetails("Height").value))
             FG.TextMatrix(Num, FG.ColIndex("Width")) = IIf(IsNull(RsDetails("Width")), "", (RsDetails("Width").value))

            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            
            If Transaction_Type = 42 Then
                s = "SELECT T2.* "
                s = s & " from  Transactions AS t "
                s = s & " Inner Join Transaction_Details T2 On T2.Transaction_ID = t.Transaction_ID"
                s = s & " WHERE t.Transaction_Type = 26 and t.OrderID =  " & val(txtOrderID)
                s = s & " and  T2.Item_ID = " & val(RsDetails("Item_ID").value & "")
                s = s & " and T2.UnitId= " & val(RsDetails("UnitId").value & "")
                Set rsDummy = New ADODB.Recordset
                
                rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If rsDummy.EOF Then
                    mCostPrice = 0
                Else
                    mCostPrice = val(rsDummy!ShowPrice & "")
                End If
                           
            End If

            If mCostPrice <> 0 Then
                FG.TextMatrix(Num, FG.ColIndex("Price")) = mCostPrice
            Else
                FG.TextMatrix(Num, FG.ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(Num, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Me.XPTxtBillID), val(FG.cell(flexcpData, Num, FG.ColIndex("UnitID"))), val(Me.DCboStoreName.BoundText))
            End If
            
            If CBoBasedON.ListIndex = 13 Then
         '   FG.TextMatrix(Num, FG.ColIndex("showqty")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))
            End If
            FG.TextMatrix(Num, FG.ColIndex("SalesPrice")) = GetItemPrice(val(FG.TextMatrix(Num, FG.ColIndex("Code"))), 0, val(FG.cell(flexcpData, Num, FG.ColIndex("UnitID"))))
            FG.TextMatrix(Num, FG.ColIndex("TotalSalesPrice")) = val(FG.TextMatrix(Num, FG.ColIndex("SalesPrice"))) * val(FG.TextMatrix(Num, FG.ColIndex("Count")))
        
            RsDetails.MoveNext
            Debug.Print Num

            If FG.rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num

    End If

    TxtFillData.text = "F"
    Screen.MousePointer = vbDefault
'    XPTxtCurrent.Caption = rs.AbsolutePosition
'    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub


Public Sub RetriveOrderDef(Optional order_no As String = "", _
                        Optional Transaction_Type As Integer = 0)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    Dim StoreId2 As Double
    Dim issuedQty As Double
    Dim rsDummy As New ADODB.Recordset
    Dim mCostPrice As Double
    Dim s As String
   On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
    txtOrderID = order_no
      
      
             StrSQL = "SELECT    TblDefComItem.order_no,TblDefComItem.OrderID, TblDefComItem.CusID,TblDefComItem.storeid,"
    StrSQL = StrSQL + " TblDefComItem.Branchid, dbo.TblDefComItemData.*,"
    StrSQL = StrSQL + " TblItems.*,TblUnites.UnitName,TblUnites.UnitNamee "
    StrSQL = StrSQL + " From TblDefComItemData"
    StrSQL = StrSQL + " Left Outer join TblDefComItem On TblDefComItem.ID = TblDefComItemData.IDDefCIT"
    StrSQL = StrSQL + " Left Outer Join TblItems On TblItems.ItemID = TblDefComItemData.ItemID"
    StrSQL = StrSQL + " Left Outer Join TblUnites On TblUnites.UnitID = TblDefComItemData.UnitId"
    StrSQL = StrSQL + " Where TblDefComItem.Id = " & val(txtOrderID)
    
    StrSQL = StrSQL & " order by TblDefComItemData.id "

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    

    If RsDetails.RecordCount < 1 Then
 
        Exit Sub
    Else
        DBCboClientName.BoundText = IIf(IsNull(RsDetails("CusID").value), "", RsDetails("CusID").value)
       ' Me.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
        
If Transaction_Type = 21 And SystemOptions.MultyStore = True Then

Else
Me.DCboStoreName.BoundText = IIf(IsNull(RsDetails("storeid").value), "", RsDetails("storeid").value)
End If

        Me.dcBranch.BoundText = IIf(IsNull(RsDetails("Branchid").value), "", RsDetails("Branchid").value)
       ' TxtOldOpOrderID.Text = IIf(IsNull(rs("OldOpOrderID").value), "", rs("OldOpOrderID").value)
       ' TxtCashCustomerName.Text = IIf(IsNull(rs("CashCustomerName").value), "", rs("CashCustomerName").value)
        
        
        'txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
    End If

    If RsDetails.EOF Or RsDetails.BOF Then
        Exit Sub
    End If
  '  txtOrderID = rs!Transaction_ID & ""


    Screen.MousePointer = vbArrowHourglass
  


    XPTxtSum.text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemID")), "", (RsDetails("ItemID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemID")), "", Trim(RsDetails("ItemID").value))
            StoreId2 = val(DCboStoreName.BoundText)
            
                        
            FG.TextMatrix(Num, FG.ColIndex("Height")) = IIf(IsNull(RsDetails("hight")), "", (RsDetails("hight").value))
            FG.TextMatrix(Num, FG.ColIndex("length")) = IIf(IsNull(RsDetails("length")), "", (RsDetails("length").value))

            
            FG.TextMatrix(Num, FG.ColIndex("Width")) = IIf(IsNull(RsDetails("widtj")), "", (RsDetails("widtj").value))
            

            
    
            
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Qty")), "", (RsDetails("Qty").value))
            
 
        issuedQty = GetIssuedQty(TXT_order_no, val(Me.XPTxtBillID), StoreId2, val(FG.TextMatrix(Num, FG.ColIndex("Code"))))


            'FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
         '   If Transaction_Type = 0 Then
                'FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("ShowPrice")), 0, (RsDetails("ShowPrice").value)) ' GET_COST_PRICE_FOR_PRODUCT_ITEM(Val(FG.TextMatrix(Num, FG.ColIndex("Code"))))
         '   End If
      
       
            '  FG.TextMatrix(Num, FG.ColIndex("Expenses")) = IIf(IsNull(RsDetails("Lineexpenses")), "", (RsDetails("Lineexpenses").value))
         
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = 0 'IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("TotalDisc")), "", (RsDetails("TotalDisc").value))
            'FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            'FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            'FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemType")) = IIf(IsNull(RsDetails("ItemType")), 0, (RsDetails("ItemType").value))
         
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
             
            If val(RsDetails!OrderID & "") <> 0 And SystemOptions.CostByProduction Then
                s = "SELECT T2.* "
                s = s & " from  Transactions AS t "
                s = s & " Inner Join Transaction_Details T2 On T2.Transaction_ID = t.Transaction_ID"
                s = s & " WHERE t.Transaction_Type = 26 and t.OrderID =  " & val(RsDetails!OrderID & "")
                s = s & " and  T2.Item_ID = " & val(RsDetails("ItemID").value & "")
                s = s & " and T2.UnitId= " & val(RsDetails("UnitId").value & "")
                Set rsDummy = New ADODB.Recordset
                
                rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If rsDummy.EOF Then
                    mCostPrice = 0
                Else
                    mCostPrice = val(rsDummy!ShowPrice & "")
                End If
                           
            End If

            If mCostPrice <> 0 Then
                FG.TextMatrix(Num, FG.ColIndex("Price")) = mCostPrice
            Else
                FG.TextMatrix(Num, FG.ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(Num, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Me.XPTxtBillID), val(FG.cell(flexcpData, Num, FG.ColIndex("UnitID"))), val(Me.DCboStoreName.BoundText))
            End If
            FG.TextMatrix(Num, FG.ColIndex("SalesPrice")) = GetItemPrice(val(FG.TextMatrix(Num, FG.ColIndex("Code"))), 0, val(FG.cell(flexcpData, Num, FG.ColIndex("UnitID"))))
            FG.TextMatrix(Num, FG.ColIndex("TotalSalesPrice")) = val(FG.TextMatrix(Num, FG.ColIndex("SalesPrice"))) * val(FG.TextMatrix(Num, FG.ColIndex("Count")))
        
            RsDetails.MoveNext
            Debug.Print Num

         

        Next Num

    End If

    TxtFillData.text = "F"
    Screen.MousePointer = vbDefault
'    XPTxtCurrent.Caption = rs.AbsolutePosition
'    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdAdd_Click()
 If SystemOptions.InsertItemManualOut Then
        Dim StrSQL  As String, Msg As String
        Dim RsDetails As New ADODB.Recordset
           StrSQL = "Select * from transactions where  Transaction_Type=" & 21 & " and noteserial1='" & TXT_order_no & "'"
 
    Dim rsDummy As New ADODB.Recordset
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rsDummy.RecordCount < 1 Then
 
        Exit Sub
    End If
   
       ' Me.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
        
     
        
        StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
        StrSQL = StrSQL + " where Transaction_ID=" & val(rsDummy("Transaction_ID").value)
        StrSQL = StrSQL + " And Item_Id  = " & val(DCboItemsName.BoundText)
'        If Transaction_Type = 21 And SystemOptions.MultyStore = True Then
'            StrSQL = StrSQL + " and  Transaction_Details.StoreID2=" & val(Me.DCboStoreName.BoundText)
'        End If
        
        StrSQL = StrSQL & " order by Transaction_Details.id "

        RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If RsDetails.EOF Then
            Msg = "Â–« «·’‰› €Ì— „ÊÃÊœ ›Ï «·›« Ê—…"
            Msg = Msg & Chr(13) & ""
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCboItemsName.BoundText = ""
        Else
            TxtQuantity.text = RsDetails!ShowQty & ""
        
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command10_Click()

    Dim Transaction_ID As String
    Dim Transaction_Type As Integer
    Dim Transaction_Type2 As Integer

    Transaction_Type2 = 0
    
    If CBoBasedON.ListIndex = 7 Then
        If TXT_order_no.text <> "" Then
            Transaction_Type = 38
            Transaction_ID = get_transactionData("noteserial1", TXT_order_no.text, "Transaction_ID", Transaction_Type, Transaction_Type2)
            'Unload FrmPO6
            'FrmPO6.Show
         
             FrmPO6.Retrive val(Transaction_ID)
            
        End If
    ElseIf CBoBasedON.ListIndex = 2 Then
        If TXT_order_no.text <> "" Then
            Transaction_Type = 21
            Transaction_ID = get_transactionData("noteserial1", TXT_order_no.text, "Transaction_ID", Transaction_Type, Transaction_Type2)
            'Unload FrmPO6
            'FrmPO6.Show
                frmsalebill.show
                frmsalebill.XPBtnMove_Click (2)
                
            frmsalebill.Retrive val(Transaction_ID)
        End If
        
    ElseIf CBoBasedON.ListIndex = 10 Then
        If TXT_order_no.text <> "" Then
            Transaction_Type = 9
            Transaction_ID = get_transactionData("noteserial1", TXT_order_no.text, "Transaction_ID", Transaction_Type, Transaction_Type2)
            'Unload FrmPO6
            'FrmPO6.Show
            FrmReturnSalling.Retrive val(Transaction_ID)
        End If
    ElseIf CBoBasedON.ListIndex = 8 Then
        If TXT_order_no.text <> "" Then
            If MintDone = 1 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "√„— ‘€· «·’Ì«‰… „‰ ÂÌ Ê·« Ì„ﬂ‰ «‰‘«¡ ”‰œ ’—›/ ”·„ »‰«¡ ⁄·ÌÂ"
                Else
                    MsgBox "this maintenance order is done and can't creat Issue Voucher based on it "
                End If
            ElseIf MintDone = -1 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·« ÌÊÃœ «„— ‘€· «·’Ì«‰… »Â–« «·—ﬁ„"
                Else
                    MsgBox "there's no maintenance order associated with this No."
                End If
            End If
        End If
        ElseIf CBoBasedON.ListIndex = 9 Then
        If TXT_order_no.text <> "" Then
            If MintDone = 1 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "√„— «’·«Õ «·’Ì«‰… „‰ ÂÌ Ê·« Ì„ﬂ‰ «‰‘«¡ ”‰œ ’—›/ ”·„ »‰«¡ ⁄·ÌÂ"
                Else
                    MsgBox "his maintenance order is done and can't creat Issue Voucher based on it "
                End If
            ElseIf MintDone = -1 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·« ÌÊÃœ «„— «’·«Õ «·’Ì«‰… »Â–« «·—ﬁ„"
                Else
                    MsgBox "There's no maintenance order associated with this No."
                End If
                TXT_order_no.text = ""
            End If
        End If
        
    End If
    
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
       If val(DcboEmpName.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    txtEmpCode.text = EmpCode
    
End Sub

Private Sub DcboEmpName_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 19
        Set FrmEmployeeSearch.RetrunFrm = Me
        FrmEmployeeSearch.show
    End If
    
End Sub

Private Sub DCboItemsName_Validate(Cancel As Boolean)
If SystemOptions.InsertItemManualOut And DCboItemsName.BoundText <> "" Then
        Dim StrSQL  As String, Msg As String
        Dim RsDetails As New ADODB.Recordset
           StrSQL = "Select * from transactions where  Transaction_Type=" & 21 & " and noteserial1='" & TXT_order_no & "'"
 
    Dim rsDummy As New ADODB.Recordset
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rsDummy.RecordCount < 1 Then
 
        Exit Sub
    End If
   
       ' Me.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
        
     
        
        StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
        StrSQL = StrSQL + " where Transaction_ID=" & val(rsDummy("Transaction_ID").value)
        StrSQL = StrSQL + " And Item_Id  = " & val(DCboItemsName.BoundText)
'        If Transaction_Type = 21 And SystemOptions.MultyStore = True Then
'            StrSQL = StrSQL + " and  Transaction_Details.StoreID2=" & val(Me.DCboStoreName.BoundText)
'        End If
        
        StrSQL = StrSQL & " order by Transaction_Details.id "

        RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If RsDetails.EOF Then
            Msg = "Â–« «·’‰› €Ì— „ÊÃÊœ ›Ï «·›« Ê—…"
            Msg = Msg & Chr(13) & ""
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCboItemsName.BoundText = ""
        Else
            TxtQuantity.text = RsDetails!ShowQty & ""
        
        End If
    End If
    Screen.MousePointer = vbDefault

End Sub

Private Sub DCboItemsName_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF3 Then


        Load FrmItemSearch
        FrmItemSearch.RetrunType = 6
        FrmItemSearch.show vbModal
    End If
End Sub


Private Sub DCboStoreName_Change()
 TxtStoreID.text = getStoreCoding(val(DCboStoreName.BoundText))
     If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then

     If CheckStoreCoding(val(dcBranch.BoundText), 10) = True Then
     TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""

     End If
     
    End If
    
End Sub

Function CuurentLogdata(Optional Currentmode As String)
      LogTextA = "    ‘«‘… " & ScreenNameArabic & Chr(13) & " —ﬁ„ «·”‰œ   " & TxtNoteSerial1.text & Chr(13) & " —ﬁ„ ÌœÊÌ     " & txtManualNO.text & Chr(13) & " «· «—ÌŒ " & XPDtbBill.value & Chr(13) & " «·Œ“Ì‰… " & DcboBox.text & Chr(13) & " «·„Œ“‰  " & DCboStoreName.text & Chr(13) & "  «·⁄„Ì· / «·„Ê—œ   " & DBCboClientName.text & Chr(13) & "‰Ê⁄ «·”‰œ " & DCDocTypes & Chr(13) & "»‰«¡ ⁄·Ï " & CBoBasedON & "»—ﬁ„   " & TXT_order_no & Chr(13) & "⁄„Ì· ‰ﬁœÌ" & TxtCashCustomerName & Chr(13) & "  «·„⁄œ… " & DCEquipments & Chr(13) & "    «·«œ«—… " & DcboEmpDepartments & Chr(13) & " „—ﬂ“ «· ﬂ·›… " & DcCostCenter & Chr(13) & "«·„ÊŸ›" & DcboEmpName & Chr(13) & "«·«œ«—… " & DcboEmpDepartments & Chr(13) & " „·«ÕŸ«  " & TxtBillComment & Chr(13) & "—ﬁ„ «·ﬁÌœ " & TxtNoteSerial & Chr(13) & "«Ã„«·Ì «·”‰œ   " & LblTotalAll.Caption
                     
    LogTexte = "" '    Screen  " & ScreenNameEnglish & Chr(13) & " Bill No " & TxtNoteSerial1.text & Chr(13) & "Supplier Bill No " & txtManualNO.text & Chr(13) & " Date " & XPDtbBill.value & Chr(13) & " Box " & DcboBox.text & Chr(13) & " Store  " & DCboStoreName.text & Chr(13) & " Supplier/Cuxtomer" & DBCboClientName.text & Chr(13) & "Doc Type" & DCDocTypes & Chr(13) & "Based On" & CBoBasedON & "No :   " & TXT_order_no & Chr(13) & "Payment Type" & CboPayMentType & Chr(13) & "Discount Type  " & XPCboDiscountType & Chr(13) & " Discount Vaalue   " & XPTxtDiscountVal & Chr(13) & " Shipment Arival Date" & DTArrivalDate & Chr(13) & "Due Date " & DtpDelayDate & Chr(13) & " Currency " & Dccurrency & Chr(13) & " GE NO" & TxtNoteSerial
                       
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 180, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", , val(TxtNoteSerial), TxtNoteSerial1
    Else
        AddToLogFile CInt(user_id), 180, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", , val(TxtNoteSerial), TxtNoteSerial1
    End If
    
End Function

Private Sub Dcbranch_GotFocus()
Dcbranch_Click (0)
End Sub

Private Sub DcbType_Change()
LoadCar
End Sub

Private Sub DcbType_Click()
DcbType_Change
End Sub

Private Sub DcCustmer_Click(Area As Integer)
    Dim My_SQL As String

    My_SQL = "  select CusID,CusName,TT.ID from TblCustemers  "
    
    My_SQL = My_SQL & " INNER  JOIN Tbl_TradingContract TT ON TblCustemers.CusID =TT.TContract_CustID "
    My_SQL = My_SQL & " Where TT.TContract_CustID =  " & val(DcCustmer.BoundText)
    My_SQL = My_SQL & " And IsNull(IsCanceld,0) <> 1"
    My_SQL = My_SQL & " order by CusName "
    Dim rsDummy As New ADODB.Recordset
    rsDummy.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Not rsDummy.EOF Then
        txtTradingContractID = rsDummy!ID & ""
    End If
End Sub

Private Sub DCEquipments_Change()
If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
TxtBillComment.text = DCEquipments.text
End If

End Sub

Private Sub DCEquipments_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
         Load FrmCasrShearches
        FrmCasrShearches.SendForm = "FrmOut"
        FrmCasrShearches.show vbModal
    End If
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String
Dim LngRow As Long
With FG
Select Case .ColKey(Col)
     Case "GroupMint"
              StrAccountCode = .ComboData
              LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("GroupIDMint"), False, True)
              .TextMatrix(Row, .ColIndex("GroupIDMint")) = StrAccountCode
              .TextMatrix(Row, .ColIndex("MintName")) = ""
              .TextMatrix(Row, .ColIndex("MintID")) = 0
      Case "MintName"
              StrAccountCode = .ComboData
              LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("MintID"), False, True)
             .TextMatrix(Row, .ColIndex("MintID")) = StrAccountCode
        Case "Code", "Name"
          
         If SystemOptions.InsertItemManualOut And CBoBasedON.ListIndex = 2 Then
               Dim StrSQL  As String, Msg As String
               Dim RsDetails As New ADODB.Recordset
               StrSQL = "Select * from transactions where  Transaction_Type=" & 21 & " and noteserial1='" & TXT_order_no & "'"
    
               Dim rsDummy As New ADODB.Recordset
               Set rsDummy = New ADODB.Recordset
               rsDummy.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

               If rsDummy.RecordCount < 1 Then
            
                   Exit Sub
               End If
   
       ' Me.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
        
     
        
        StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
        StrSQL = StrSQL + " where Transaction_ID=" & val(rsDummy("Transaction_ID").value)
        StrSQL = StrSQL + " And Item_Id  = " & val(FG.TextMatrix(Row, FG.ColIndex("Name")))
'        If Transaction_Type = 21 And SystemOptions.MultyStore = True Then
'            StrSQL = StrSQL + " and  Transaction_Details.StoreID2=" & val(Me.DCboStoreName.BoundText)
'        End If
        
        StrSQL = StrSQL & " order by Transaction_Details.id "

        RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If RsDetails.EOF Then
            Msg = "Â–« «·’‰› €Ì— „ÊÃÊœ ›Ï «·›« Ê—…"
            Msg = Msg & Chr(13) & ""
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
           ' DCboItemsName.BoundText = ""
           FG.TextMatrix(Row, FG.ColIndex("Name")) = ""
        Else
            FG.TextMatrix(Row, FG.ColIndex("count")) = RsDetails!ShowQty & ""
        End If
    End If
    Screen.MousePointer = vbDefault
End Select
End With

If Me.TxtModFlg <> "E" Then Exit Sub

'val (TxtNoteSerial)
    If Col = FG.ColIndex("Code") Or Col = FG.ColIndex("Name") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , , val(TxtNoteSerial), Me.TxtNoteSerial1, 180
    ElseIf Col = FG.ColIndex("UnitID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("UnitID")), , , , , , , , , val(TxtNoteSerial), Me.TxtNoteSerial1, 180
    ElseIf Col = FG.ColIndex("Count") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , (FG.TextMatrix(Row, FG.ColIndex("Count"))), , , , , , , , val(TxtNoteSerial), Me.TxtNoteSerial1, 180
    ElseIf Col = FG.ColIndex("Price") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , (FG.TextMatrix(Row, FG.ColIndex("Price"))), , , , , , , val(TxtNoteSerial), Me.TxtNoteSerial1, 180
    ElseIf Col = FG.ColIndex("ColorID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("ColorID")), , , , , val(TxtNoteSerial), Me.TxtNoteSerial1, 180
    ElseIf Col = FG.ColIndex("ItemSize") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("ItemSize")), , , , val(TxtNoteSerial), Me.TxtNoteSerial1, 180
    ElseIf Col = FG.ColIndex("ClassId") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("ClassId")), , , val(TxtNoteSerial), Me.TxtNoteSerial1, 180
    ElseIf Col = FG.ColIndex("DiscountType") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , FG.cell(flexcpTextDisplay, Row, FG.ColIndex("DiscountType")), , val(TxtNoteSerial), Me.TxtNoteSerial1, 180
    ElseIf Col = FG.ColIndex("DiscountVal") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , FG.TextMatrix(Row, FG.ColIndex("DiscountVal")), val(TxtNoteSerial), Me.TxtNoteSerial1, 180
    
    End If
    

End Sub

Private Sub FG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With FG
Select Case .ColKey(Col)
Case "GroupIDMint", "MintID"
Cancel = True
End Select
End With
End Sub

Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim StrSQL As String
Dim StrComboList As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

    If CBoBasedON.ListIndex = 7 Then
         If Not SystemOptions.CanChangeOut Then
            Cancel = True
             Exit Sub
         End If
     End If

With FG
If CBoBasedON.ListIndex = 8 Then
  Select Case .ColKey(Col)
       Case "GroupMint"
                StrSQL = "select * from TblMaintenanceType  where MainType=1 and id in(SELECT     GroupID FROM         dbo.tblordermaintenancetypes"
                StrSQL = StrSQL & "  WHERE     (ORderID = " & TXT_order_no.text & ") AND (TypeTrans = 0))"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "name", "id")
                Else
                    StrComboList = .BuildComboList(rs, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
    Case "MintName"
                StrSQL = "select * from TblMaintenanceType  "
                StrSQL = StrSQL & " where ( MainType =0 or MainType is null)and id in(SELECT     maintenanceid FROM         dbo.tblordermaintenancetypes"
                StrSQL = StrSQL & "  WHERE     (ORderID = " & TXT_order_no.text & ") AND (TypeTrans = 0))"
                If val(.TextMatrix(Row, .ColIndex("GroupIDMint"))) <> 0 Then
               StrSQL = StrSQL & "  and  FollowID=" & val(.TextMatrix(Row, .ColIndex("GroupIDMint"))) & "   "
               End If
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "name", "id")
                Else
                    StrComboList = .BuildComboList(rs, "namee", "id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
     End Select
     End If
  End With
End Sub

Private Sub sameCmd_Click()
Me.TxtModFlg.text = "N"
Me.XPDtbBill.value = Date
TxtNoteSerial1.text = ""
TxtNoteSerial.text = ""
End Sub
Private Sub Txt_order_no_Change()
    Dim StrSQL As String
    Dim rs2 As ADODB.Recordset
    DcbType.Visible = False
    lbl(67).Visible = False
    Dim Transaction_Type As Integer
    If CBoBasedON.ListIndex = 1 Then
        Transaction_Type = 6
        
    ElseIf CBoBasedON.ListIndex = 11 Then
        Transaction_Type = 22
         
         
    ElseIf CBoBasedON.ListIndex = 2 Then
        Transaction_Type = 21
    ElseIf CBoBasedON.ListIndex = 3 Then
        Transaction_Type = 5
    ElseIf CBoBasedON.ListIndex = 4 Then
        Transaction_Type = 41
    ElseIf CBoBasedON.ListIndex = 5 Then
        Transaction_Type = 42
    ElseIf CBoBasedON.ListIndex = 5 Then
    ElseIf CBoBasedON.ListIndex = 7 Then
        Transaction_Type = 38
    ElseIf CBoBasedON.ListIndex = 13 Then
        Transaction_Type = 29
    ElseIf CBoBasedON.ListIndex = 14 Then
        Transaction_Type = 26
    ElseIf CBoBasedON.ListIndex = 15 Then
        Transaction_Type = 28
    ElseIf CBoBasedON.ListIndex = 8 Then
    If TXT_order_no.text = "" Then Exit Sub
        DcbType.Visible = True
        lbl(67).Visible = True
        Set rs2 = New ADODB.Recordset
        StrSQL = "select * from TblOrderMaint where ID = " & TXT_order_no.text & " "
        'ended
        StrSQL = "SELECT LeaderID,TblOrderMaint.ended,"
        StrSQL = StrSQL & " Te.Emp_Name,"
        StrSQL = StrSQL & "te.DepartmentID,"
        StrSQL = StrSQL & "TblEmpDepartments.DepartmentName,TblOrderMaint.EquepID,FixedAssets.Name FixedAssetName"
        StrSQL = StrSQL & " From TblOrderMaint"
        StrSQL = StrSQL & " LEFT OUTER JOIN TblEmployee AS te"
        StrSQL = StrSQL & "     ON  TblOrderMaint.LeaderID = te.Emp_ID"
        StrSQL = StrSQL & " LEFT OUTER JOIN TblEmpDepartments"
        StrSQL = StrSQL & " ON  TblEmpDepartments.DeparmentID = te.DepartmentID"
        StrSQL = StrSQL & " LEFT OUTER JOIN FixedAssets ON FixedAssets.id = TblOrderMaint.EquepID"
        
        StrSQL = StrSQL & " where TblOrderMaint.ID = " & TXT_order_no.text & " "
        rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If rs2.RecordCount > 0 Then
                MintDone = val(rs2("ended").value & "")
         '       If MintDone = 1 Then MintDone = -1: Txt_order_no = "": MsgBox "Â–« «·«„— „€·ﬁ": Exit Sub
                DCEquipments.BoundText = val(rs2!EquepID & "")
                DcboEmpDepartments.BoundText = val(rs2!DepartmentID & "")
                DcboEmpName.BoundText = val(rs2!LeaderID & "")
                
            Else
                MintDone = -1
            End If
        'GetOrderMaint
        Exit Sub
       ElseIf CBoBasedON.ListIndex = 9 Then
       Dim orderStatus As Integer
     
      MintDone = 0
        Set rs2 = New ADODB.Recordset
        StrSQL = "select * from TblCardAuthorizationReform where WorkOrder = " & val(TXT_order_no.text) & " "
        rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If rs2.RecordCount > 0 Then
            orderStatus = IIf(IsNull(rs2("OrderStatus").value), 0, rs2("OrderStatus").value)
            TxtCashCustomerName.text = IIf(IsNull(rs2("ClientName").value), "", rs2("ClientName").value)
            If orderStatus = 2 Or orderStatus = 4 Or orderStatus = 5 Then
                MintDone = 1
            End If
            Else
            TxtCashCustomerName.text = ""
                MintDone = -1
            End If
        Exit Sub
    ElseIf CBoBasedON.ListIndex = 12 Then
        RetriveOrderDef Me.TXT_order_no, 0
        Exit Sub
    Else
         Exit Sub
    End If

   'Transaction_ID = get_transactionData("order_no", Txt_order_no.text, "Transaction_ID", Transaction_Type)
    If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
        RetriveOrder Me.TXT_order_no, Transaction_Type
    End If
        If CBoBasedON.ListIndex = 7 Then
                     If Not SystemOptions.CanChangeOut Then
                        FG.Enabled = False
                        Ele(2).Enabled = False
                         Exit Sub
                    
                     End If
        Else
            FG.Enabled = True
            Ele(2).Enabled = True
        End If
      
End Sub
Sub GetOrderMaintdet()
If 1 = 1 Then
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     PartID"
sql = sql & " From dbo.tblordermaintenancetypes"
sql = sql & "  where ORderID =" & val(TXT_order_no.text) & " and TypeTrans=2"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
DCEquipments.BoundText = IIf(IsNull(rs2("PartID").value), "", rs2("PartID").value)
Else
DCEquipments.BoundText = 0
End If
End If
End Sub

Sub GetOrderMaint()
If 1 = 1 Then
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     EquepID"
sql = sql & " From dbo.TblOrderMaint"
sql = sql & "  where ID =" & val(TXT_order_no.text) & ""
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
DCEquipments.BoundText = IIf(IsNull(rs2("EquepID").value), "", rs2("EquepID").value)
Else
DCEquipments.BoundText = 0
End If
End If
End Sub
Private Sub CBoBasedON_Change()
 DcbType.Visible = False
 lbl(67).Visible = False
If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then

    If TXT_order_no.text <> "" Then
        TXT_order_no.text = ""
        TxtOldOpOrderID.text = ""
    End If
    LoadCar
    
    
  End If
End Sub

Private Sub CBoBasedON_Click()
CBoBasedON_Change
End Sub

Private Sub txt_ORDER_NO_KeyUp(KeyCode As Integer, Shift As Integer)


    Dim transactiontype As Integer
    Dim transactionName As String

    If KeyCode = vbKeyF3 Then
       If CBoBasedON.ListIndex = 1 Then
            transactiontype = 6
            If SystemOptions.UserInterface = ArabicInterface Then
                transactionName = "«·»ÕÀ ⁄‰ «„— »Ì⁄"
            Else
                transactionName = "Search  Sales Order"
            End If

            Order_no_search.show
             Order_no_search.RetrunType = 11
             Order_no_search.Label1(2).Caption = transactionName
             Order_no_search.lblSpecificsearch = transactiontype
'
            If val(Me.DBCboClientName.BoundText) <> 2 Then

                 Order_no_search.DBCboClientName.BoundText = Me.DBCboClientName.BoundText
   
            End If
        ElseIf CBoBasedON.ListIndex = 13 Then
            transactiontype = 29
            If SystemOptions.UserInterface = ArabicInterface Then
                transactionName = "«·»ÕÀ ⁄‰ «„— ‘—«¡"
            Else
                transactionName = "Search  Sales Order"
            End If

            Order_no_search.show
             Order_no_search.RetrunType = 89
             Order_no_search.Label1(2).Caption = transactionName
             Order_no_search.lblSpecificsearch = transactiontype
'
            If val(Me.DBCboClientName.BoundText) <> 2 Then

                 Order_no_search.DBCboClientName.BoundText = Me.DBCboClientName.BoundText
   
            End If
            
        ElseIf CBoBasedON.ListIndex = 14 Then
            transactiontype = 26
            If SystemOptions.UserInterface = ArabicInterface Then
                transactionName = "«·»ÕÀ ⁄‰ «„— «‰ «Ã"
            Else
                transactionName = "Search  Sales Order"
            End If

            Order_no_search.show
             Order_no_search.RetrunType = 99
             Order_no_search.Label1(2).Caption = transactionName
             Order_no_search.lblSpecificsearch = transactiontype
'
            If val(Me.DBCboClientName.BoundText) <> 2 Then

                 Order_no_search.DBCboClientName.BoundText = Me.DBCboClientName.BoundText
   
            End If
            
            
        ElseIf CBoBasedON.ListIndex = 15 Then
            transactiontype = 28
            If SystemOptions.UserInterface = ArabicInterface Then
                transactionName = "Production Recive Voucher"
            Else
                transactionName = "Production Recive Voucher"
            End If

            Order_no_search.show
             Order_no_search.RetrunType = 98
             Order_no_search.Label1(2).Caption = transactionName
             Order_no_search.lblSpecificsearch = transactiontype
'
            If val(Me.DBCboClientName.BoundText) <> 2 Then

                 Order_no_search.DBCboClientName.BoundText = Me.DBCboClientName.BoundText
   
            End If
        ElseIf CBoBasedON.ListIndex = 2 Then
            'If SystemOptions.UserInterface = ArabicInterface Then
                'transactionName = "«·»ÕÀ ⁄‰ ›Ê« Ì— »Ì⁄"
            'Else
                'transactionName = "Search  Sales Invoices"
            'End If
            'transactionType = 21
      
            FrmBuySearch.DealingForm = GridTransType.InvoiceTransaction
            FrmBuySearch.Index = 12
            FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ›« Ê—… „»Ì⁄«    "
            FrmBuySearch.show vbModal
        ElseIf CBoBasedON.ListIndex = 3 Then
            'If SystemOptions.UserInterface = ArabicInterface Then
                'transactionName = "«·»ÕÀ ⁄‰   „—œÊœ«  „‘ —Ì« "
            'Else
                'transactionName = "Search  Return Purchase"
            'End If
            'transactionType = 5
        ElseIf CBoBasedON.ListIndex = 7 Then
            If KeyCode = vbKeyF3 Then
                FrmBuySearch.Index = 7
                FrmBuySearch.DealingForm = GridTransType.internalorder
                FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ÿ·»«   œ«Œ·Ì…"
                FrmBuySearch.show vbModal
            End If
        ElseIf CBoBasedON.ListIndex = 8 Then


              FrmSearchOrderMainten.lbltypr = 1
              Load FrmSearchOrderMainten
              FrmSearchOrderMainten.show
    
        Else
            'transactiontype = 0
            Exit Sub
        End If
    End If
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
 
Private Sub Cmd_Click(Index As Integer)
    Dim AskOption As Boolean
    Dim intDef    As Integer
    Dim Msg       As String
    Dim StrSQL    As String
    Dim RsTest    As ADODB.Recordset
    Dim RsOptions As ADODB.Recordset
    BolPrint = True
    On Error GoTo ErrTrap

    Select Case Index

        Case 9
            ShowAttachments TxtNoteSerial1, 19, "0991201403"
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
                        Msg = Msg & Chr(13) & ""
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        Exit Sub
                    End If
                End If
            End If
        
            clear_all Me
              GRID2.Clear flexClearScrollable, flexClearEverything
            GRID2.rows = 1
            DcCostCenter.text = ""
            ClearNotes
            TxtModFlg.text = "N"
            '       XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            SetDefaults
            NewGrid.GridDefaultValue 1
            Me.DCboUserName.BoundText = user_id
            intDef = val(GetSetting(StrAppRegPath, "DefaultOptions", "DefaultClient", 2))
            DBCboClientName.BoundText = intDef
            '       intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSaleStore", 1)
            '       DCboStoreName.BoundText = intDef

            Dim dstore       As Integer
            Dim dBox         As Integer
            Dim usertype     As Integer
            Dim EmpID        As Integer
            Dim userbranchid As Integer
            'GetBranchData branch_id, dstore, dBox
                 
            Dim CUSTID       As Integer

            GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID, , CUSTID
            DBCboClientName.BoundText = CUSTID

            If usertype <> 0 Then 'admin
                dcBranch.Enabled = False
 
                DCboStoreName.Enabled = True
                '     TxtStoreID.Enabled = False
                Me.DCboStoreName.BoundText = dstore
            Else
                dcBranch.Enabled = True
 
                DCboStoreName.Enabled = True
 
                Me.dcBranch.BoundText = ""
                Me.DCboStoreName.BoundText = ""
                '     TxtStoreID.Enabled = True
            End If
        
            Set RsOptions = New ADODB.Recordset
            RsOptions.Open "tbloptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable

            If Not (RsOptions.BOF Or RsOptions.EOF) Then
                Me.DcboBox.BoundText = IIf(IsNull(RsOptions("SalesBoxID").value), "", RsOptions("SalesBoxID").value)
            End If

            XPTab301.CurrTab = 0
            '------------------
            Me.XPDtbBill.SetFocus
            '--------------------
         
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

            If Voucher_coding(val(Me.dcBranch.BoundText), XPDtbBill.value, 10, 180, , 19) = "" Then
                TxtNoteSerial1.locked = False
            Else
                TxtNoteSerial1.locked = True
 
            End If

            DCOPrType.ListIndex = 0
            CBoBasedON.ListIndex = 0
           
            FG.rows = FG.FixedRows
            FG.rows = 2
            DCExtraAccount.BoundText = get_account_code_branch(213, val(Me.dcBranch.BoundText))
        Case 1

          

            If IsSaveWithOutMsg Then GoTo SaveDirect
            
              If chkDone.value = vbChecked Then
                MsgBox "·« Ì„ﬂ‰  ⁄œÌ· Õ«·Â Â–« «·”‰œ ·«‰Â „”·„ »«·›⁄·"
                Exit Sub

            End If
            If CBoBasedON.ListIndex = 7 Then
                If Not SystemOptions.CanChangeOut Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·« Ì„ﬂ‰ «· ⁄œÌ· ·⁄œ„ ÊÃÊœ ’·«ÕÌ… «· ⁄œÌ· ⁄·Ì ”‰œ «·’—› »ı‰«¡« ⁄·Ì «·ÿ·» «·œ«Œ·Ì"
                    Else
                        MsgBox "The amendment can not be modified because there is no validity of the amendment on the exchange certificate based on the internal request"
                    End If

                    Exit Sub
                End If
            End If

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

            'If AvailableDeal = True Then
            '«·√ﬁ”«ÿ «·„”Ã·… ⁄·Ï «·›« Ê—…
            
            If Text1.text <> "" And txtPassword <> "Alex2025" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "Â–« «·«–‰ «·Ì Ê·« Ì„ﬂ‰  ⁄œÌ·Â       " & Space$(5) & Txtnots2.text
                Else
                    Msg = "This Voucher Created Automatically And Cant Modify"
                End If

                MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            If XPTxtValue(1).Tag <> "" Then
                StrSQL = "select * From InstallMent where NoteID=" & XPTxtValue(1).Tag
                Set RsTest = New ADODB.Recordset
                RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (RsTest.EOF Or RsTest.BOF) Then
                    Msg = "·ﬁœ  „  ﬁ”Ìÿ «·ﬁÌ„ «·¬Ã·… ⁄·Ï Â–Â «·›« Ê—…" & Chr(13)
                    Msg = Msg + " ⁄œÌ· «·›« Ê—… ”ÌƒœÌ ≈·Ï Õ–› Â–Â «·√ﬁ”«ÿ" & Chr(13)
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
                    Msg = "·ﬁœ  „  Õ’Ì· »⁄÷ «·√ﬁ”«ÿ «·„”Ã·… ⁄·Ï Â–Â «·›« Ê—…" & Chr(13)
                    Msg = Msg + "Ê·« Ì„ﬂ‰  ⁄œÌ· »Ì«‰« Â«" & Chr(13)
                    Msg = Msg + "≈–« ﬂ‰   —€» ›Ì  ⁄œÌ· »Ì«‰«  Â–Â «·›« Ê—…" & Chr(13)
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
                Msg = Msg + "≈–« ﬂ‰   —€» ›Ì  ⁄œÌ· »Ì«‰«  Â–Â «·›« Ê—…" & Chr(13)
                Msg = Msg + "ÌÃ» Õ–› ⁄„·Ì«  «·’Ì«‰… «·Œ«’… »Â«"
                MsgBox Msg, vbOKOnly + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

SaveDirect:
            TxtModFlg.text = "E"

            If Trim(txtPassword) <> "Alex2025" Then
                Me.DCboUserName.BoundText = user_id
            End If

            'End If
        Case 2
        
            If IsSaveWithOutMsg Then GoTo SaveDirect2
            Cmd(2).Enabled = False

            If ChekClodePeriod(XPDtbBill.value) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                Else
                    MsgBox "Please Change Date Becouse This is Period is Closed"
                End If

                Cmd(2).Enabled = True
                Exit Sub
            End If
     
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
                Cmd(2).Enabled = True
                Exit Sub
            End If

            my_branch = val(Me.dcBranch.BoundText)

            If Text1.text <> "" And txtPassword <> "Alex2025" Then
                Msg = "·ﬁœ  „  ÕÊÌ· Â–« «·«–‰ «·Ï ›« Ê—… „»Ì⁄«    .."
                Msg = Msg & Chr(13) & "Ê·«Ì„ﬂ‰  ÕÊÌ·… „—… «Œ—Ï  ..!!"
                MsgBox Msg, vbOKOnly, App.Title
                Cmd(2).Enabled = True
                Exit Sub
                Else:
     
                '         If Me.TxtModFlg.text = "N" Then
             
                ' End If
                '    If CheckFilegrid() = False Then
                '        Cmd(2).Enabled = True
                '        Cmd(2).Enabled = True
                '        Exit Sub
                '     End If
SaveDirect2:
                SaveData
     
            End If

        Case 3
     
            Undo

        Case 4

            If chkDone.value = vbChecked Then
                MsgBox "·« Ì„ﬂ‰  ⁄œÌ· Õ«·Â Â–« «·”‰œ ·«‰Â „”·„ »«·›⁄·"
                Exit Sub

            End If

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
                m_FrmSearch.DealingForm = InventoryOut
                m_FrmSearch.Caption = "«·»ÕÀ ⁄‰ ”‰œ«  «·’—› "
                Set m_FrmSearch.RetrunFrm = Me
                m_FrmSearch.show vbModal
            Else
                Msg = "Â‰«ﬂ ‘«‘… »ÕÀ Œ«’… »‘«‘…      »”‰œ«  «·’—›"
                Msg = Msg & Chr(13) & "Ÿ«Â—… «„«„ﬂ ›⁄·«...·«Ì„ﬂ‰ ⁄—÷ «ﬂÀ— „‰ ‘«‘… »ÕÀ ·ﬂ· ‘«‘… ”‰œ« "
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                m_FrmSearch.ZOrder 0
                'm_FrmSearch.SetFocus
            End If

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
                
            AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowMe", False)

            If AskOption = False Then
             
                '    FrmPrintOptions.show vbModal
            
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

Private Sub CmdCash_Click(Index As Integer)

    Select Case Index

        Case 0

        Case 1
    End Select

End Sub
Function ChekOrderQty(Transaction_ID As Double, Optional Item_ID As Double, Optional OldID As Double) As Double
    Dim StrSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    StrSQL = "   SELECT     dbo.Transaction_Details.ItemSerial, dbo.TblItems.HaveSerial, *, dbo.Transaction_Details.ShowQty - ISNULL(dbo.GetQtyIsue('" & TXT_order_no.text & "', " & val(XPTxtBillID.text) & ","
    StrSQL = StrSQL & "                   dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ClassId, dbo.Transaction_Details.ItemSize,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.UnitId), 0) AS Qty"
    StrSQL = StrSQL & "    FROM         dbo.TblItems INNER JOIN"
    StrSQL = StrSQL & "                  dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID"
    StrSQL = StrSQL & "     WHERE     (dbo.Transaction_Details.Transaction_ID = " & Transaction_ID & ") AND (dbo.Transaction_Details.ShowQty - ISNULL(dbo.GetQtyIsue('" & TXT_order_no.text & "', " & val(XPTxtBillID.text) & ","
    StrSQL = StrSQL & "                  dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ClassId, dbo.Transaction_Details.ItemSize,"
    StrSQL = StrSQL & "                  dbo.Transaction_Details.UnitId), 0) > 0)and dbo.Transaction_Details.Item_ID=" & Item_ID & " and dbo.Transaction_Details.ID=" & OldID & ""
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If RsDetails.RecordCount > 0 Then
ChekOrderQty = IIf(IsNull(RsDetails("Qty").value), 0, RsDetails("Qty").value)
Else
ChekOrderQty = 0
End If

End Function
Function CheckFilegrid() As Boolean

If IsSaveWithOutMsg Then CheckFilegrid = True: Exit Function
Dim i As Integer
Dim j As Integer
Dim Item_ID As Double
Dim OldID As Double
Dim SumQty As Double
Dim total As Double
Dim Msg As String
Dim Transaction_ID2 As Double
Dim Transaction_ID1 As Long
 If CBoBasedON.ListIndex = 2 And SystemOptions.IssueVoucherWorkWithRemain = True Then
    GetTransIDFromNoteSerial1 Me.TXT_order_no.text, Transaction_ID1, , 21
    Transaction_ID2 = Transaction_ID1
With FG
CheckFilegrid = True
For j = .FixedRows To .rows - 1
SumQty = 0
Item_ID = val(.TextMatrix(j, .ColIndex("Code")))
OldID = val(.TextMatrix(j, .ColIndex("OldID")))
For i = .FixedRows To .rows - 1
If Item_ID = val(.TextMatrix(i, .ColIndex("Code"))) And OldID = val(.TextMatrix(i, .ColIndex("OldID"))) Then
SumQty = SumQty + val(.TextMatrix(i, .ColIndex("Count")))
End If
Next i
total = ChekOrderQty(Transaction_ID2, Item_ID, OldID)
If Round(total - SumQty, 2) < 0 Then
If total > 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
Msg = .cell(flexcpTextDisplay, j, .ColIndex("Name")) & "  ·«Ì„ﬂ‰ «œŒ«· ﬂ„Ì… «ﬂ»— „‰ «·ﬂ„Ì… «·«’·Ì… ··’‰› "
Msg = Msg & Chr(13)
Msg = Msg & (total) & " " & "«·ﬂ„Ì… «·„ »ﬁÌ…"
Else
End If
Else
If SystemOptions.UserInterface = ArabicInterface Then
Msg = .cell(flexcpTextDisplay, j, .ColIndex("Name")) & "  ·«ÌÊÃœ  ﬂ„Ì… „‰  «·’‰›  "
Msg = Msg & Chr(13)
Else
End If
End If
MsgBox Msg
GoTo l
Else
CheckFilegrid = True
End If
Next j
Exit Function
End With

l: CheckFilegrid = False
Else
 CheckFilegrid = True
End If
End Function

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
            StrSQL = "SELECT TOP 100 PERCENT dbo.TblItemsUnits.UnitID, dbo.TblUnites.UnitName, dbo.Transactions.Transaction_Serial,dbo.Transactions.Transaction_Type FROM dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN dbo.TblUnites INNER JOIN dbo.TblItemsUnits ON dbo.TblUnites.UnitID = dbo.TblItemsUnits.UnitID ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID AND dbo.Transaction_Details.Item_ID = dbo.TblItemsUnits.ItemID WHERE (dbo.Transactions.Transaction_Serial = '" & TxtTransSerial & "') AND (dbo.Transactions.Transaction_Type = 19) AND (dbo.TblItemsUnits.ItemID = " & FG.TextMatrix(RowNum, FG.ColIndex("Code")) & ") ORDER BY dbo.TblItemsUnits.SecOrder"
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



Private Sub CmdHelp_Click()
'    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
'    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
    On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments TxtNoteSerial1, "0703201703"

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
    Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

    If MsgBox(Msg, vbYesNo, App.Title) = vbYes Then
   
        rs.Close
        rs.Open "select * from Transactions where Transaction_Serial = " & TxtTransSerial.text & " and Transaction_type = 19"
         
        If Text1.text <> "" Then
            Msg = "·ﬁœ  „  ÕÊÌ· Â–« «·«–‰ «·Ï ›« Ê—… „»Ì⁄«    .."
            Msg = Msg & Chr(13) & "Ê·«Ì„ﬂ‰  ÕÊÌ·… „—… «Œ—Ï  ..!!"
            MsgBox Msg, vbOKOnly, App.Title
            Exit Sub
        End If

        rs!nots = TxtTransSerial.text
         
        rs.update
        '      MYWAER = " And Transaction_Type = 19"
        ''  "select * From ReplacedItems where ReturnID=" & XPTxtBillID.text
        ''                StrSQL = StrSQL + " and ItemID=" & RsDetails("Item_ID")
        Cn.Execute "INSERT INTO  Transactions (Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,CusID,StoreID,UserID,Emp_ID)SELECT Transaction_ID +1,Transaction_Serial,Transaction_Date,Transaction_Type = 21,CusID,StoreID,UserID,Emp_ID From Transactions Where Transaction_ID =" & XPTxtBillID.text + " And Transaction_Type = 19"
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



Private Sub DBCboClientName_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName
    End If
    
    
        If KeyCode = vbKeyF3 Then
            FrmCustemerSearch.SearchType = 5551
        FrmCustemerSearch.show vbModal
  
    End If
 

 



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
        FrmItemSearch.RetrunType = 6
        FrmItemSearch.show vbModal
    End If

End Sub
 
Private Sub DCboStoreName_Click(Area As Integer)
 TxtStoreID.text = getStoreCoding(val(DCboStoreName.BoundText))
 
    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then

     If CheckStoreCoding(val(dcBranch.BoundText), 10) = True Then
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
        TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
    End If


    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
       Dim Dcombos As New ClsDataCombos
        Dcombos.GetDocTypebyid Me.DCDocTypes, 19, val(Me.dcBranch.BoundText)
    End If

    If dcBranch.BoundText = "" Then TxtNoteSerial1.locked = True: Exit Sub

    If Voucher_coding(val(Me.dcBranch.BoundText), XPDtbBill.value, 10, 180, , 19) = "" Then
        TxtNoteSerial1.locked = False
    Else
        TxtNoteSerial1.locked = True
 
    End If
 
End Sub

Private Sub Dcbranch_Click(Area As Integer)
    Dcbranch_Change

    If Voucher_coding(val(val(Me.dcBranch.BoundText)), XPDtbBill.value, 10, 180, , 19) = "" Then Exit Sub
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

Private Sub DcCostCenter_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyF3 Then
        CostCenterSearch.show
        CostCenterSearch.RetrunType = 9
    End If

    If KeyCode = vbKeyF5 Then
        Dim StrSQL As String
        StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
        fill_combo Me.DcCostCenter, StrSQL
    End If
        
End Sub

Private Sub DCDocTypes_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetDocTypebyid Me.DCDocTypes, 19, val(Me.dcBranch.BoundText)

    End If

End Sub

Private Sub DCExtraAccount_KeyUp(KeyCode As Integer, _
                                 Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Unload Account_search
        Account_search.show
        Account_search.case_id = 191
            
    End If
            
    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        If SystemOptions.UserInterface = ArabicInterface Then
            Dcombos.GetAccountingCodes DCExtraAccount, True
        Else
                 
            Dcombos.GetAccountingCodesENg DCExtraAccount, True
                
        End If

    End If
        
End Sub

Private Sub Ele_DblClick(Index As Integer)
    On Error GoTo ErrTrap

    If Index = 9 Then
        If Me.WindowState = vbNormal Then
            Me.WindowState = vbMaximized
        Else
            Me.WindowState = vbNormal
        End If
    End If

    Exit Sub
ErrTrap:
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

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
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

Private Sub TXT_order_no_Validate(Cancel As Boolean)
Dim rs2 As New ADODB.Recordset
Dim StrSQL As String
If CBoBasedON.ListIndex = 8 Then
    If TXT_order_no.text = "" Then Exit Sub
        DcbType.Visible = True
        lbl(67).Visible = True
        Set rs2 = New ADODB.Recordset
        StrSQL = "select * from TblOrderMaint where ID = " & TXT_order_no.text & " "
        'ended
        StrSQL = "SELECT LeaderID,TblOrderMaint.ended,"
        StrSQL = StrSQL & " Te.Emp_Name,"
        StrSQL = StrSQL & "te.DepartmentID,"
        StrSQL = StrSQL & "TblEmpDepartments.DepartmentName,TblOrderMaint.EquepID,FixedAssets.Name FixedAssetName"
        StrSQL = StrSQL & " From TblOrderMaint"
        StrSQL = StrSQL & " LEFT OUTER JOIN TblEmployee AS te"
        StrSQL = StrSQL & "     ON  TblOrderMaint.LeaderID = te.Emp_ID"
        StrSQL = StrSQL & " LEFT OUTER JOIN TblEmpDepartments"
        StrSQL = StrSQL & " ON  TblEmpDepartments.DeparmentID = te.DepartmentID"
        StrSQL = StrSQL & " LEFT OUTER JOIN FixedAssets ON FixedAssets.id = TblOrderMaint.EquepID"
        
        StrSQL = StrSQL & " where TblOrderMaint.ID = " & TXT_order_no.text & " "
        rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If rs2.RecordCount > 0 Then
            MintDone = val(rs2("ended").value & "")
            If MintDone = 1 Then MintDone = -1: TXT_order_no = "": MsgBox "Â–« «·«„— „€·ﬁ": Exit Sub
        End If
    End If
End Sub

Private Sub txtEmpCode_KeyPress(KeyAscii As Integer)
   Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode txtEmpCode.text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
    
    
End Sub

Private Sub TxtExtraValue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtExtraValue.text, 0)
End Sub

Private Sub TxtFillData_Change()

    If TxtFillData.text = "F" Then
        NewGrid.Calculate 1, , , True
    End If

End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
  Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.text, 1
        DBCboClientName.BoundText = CUSTID
    End If



End Sub

Private Sub TxtSearchCode_KeyUp(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.show vbModal
        FrmCustemerSearch.SearchType = 1122014
    End If
End Sub

Private Sub TxtStoreID_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim StoreID As Integer

    If KeyCode = vbKeyReturn Then
    StoreID = getStoreInformatin(TxtStoreID)
        DCboStoreName.BoundText = StoreID
    End If
End Sub

Private Sub TxtTicketNO_KeyUp(KeyCode As Integer, _
                              Shift As Integer)

    If KeyCode = vbKeyF3 Then
        TxtTicketNO.text = ""
        Load FrmMaintanenceSearch1
        FrmMaintanenceSearch1.SearchType = 2
        Set FrmMaintanenceSearch1.ExtraRetrunObject = Me.TxtTicketNO
        FrmMaintanenceSearch1.show vbModal

    End If

End Sub

Private Sub txtTradingContractID_GotFocus()
mClicked = True
End Sub

Private Sub txtTradingContractID_KeyPress(KeyAscii As Integer)
Dim TContractCustID As Double
'    Dim My_SQL As String
'
'    My_SQL = "  select CusID,CusName,TT.ID,IsNull(IsCanceld,0) as IsCanceld from TblCustemers  "
'
'    My_SQL = My_SQL & " INNER  JOIN Tbl_TradingContract TT ON TblCustemers.CusID =TT.TContract_CustID "
'    My_SQL = My_SQL & " Where "
'    My_SQL = My_SQL & " TT.Id = " & val(txtTradingContractID)
'
'    'My_SQL = My_SQL & " And IsNull(IsCanceld,0) <> 1"
'    Dim msg As String
'    Dim rsDummy As New ADODB.Recordset
'    rsDummy.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If rsDummy.EOF Then
'        txtTradingContractID = ""
'
'         msg = "—ﬁ„ «·« ›«ﬁÌ… €Ì— „ÊÃÊœ"
'        MsgBox msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'    ElseIf rsDummy!IsCanceld Then
'        txtTradingContractID = ""
'        msg = "·« Ì„ﬂ‰ «Œ Ì«— « ›«ﬁÌ… „·€«…"
'
'        MsgBox msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'    End If


If txtTradingContractID.text = "" Then Exit Sub

If KeyAscii = vbKeyReturn Then
Get_TradingContractinfo txtTradingContractID.text, TContractCustID, 0
If TContractCustID = 0 Then
    txtTradingContractID.text = ""
End If
DcCustmer.BoundText = TContractCustID
End If
End Sub

Private Sub txtTradingContractID_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
    Dim TContractCustID As Double
        
     FrmProjectSearch.C1Tab1 = 4
     FrmProjectSearch.Label11.Caption = 5
     FrmProjectSearch.Caption = "»ÕÀ «·« ›«ﬁÌ«  "
     FrmProjectSearch.show vbModal
     
     Get_TradingContractinfo val(txtTradingContractID), TContractCustID, 0
    
    DcCustmer.BoundText = TContractCustID
    End If
    mClicked = False
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

Public Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap

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
     StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=19 "
     
            If cmdReSave.Visible = False Then
                StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"
            End If
            
   If cmdReSave.Visible = True Then
                StrSQL = StrSQL & "   and ( Transaction_Date >= " & SQLDate(txtFromDateReSave.value, True) & " and "
                StrSQL = StrSQL & "   Transaction_Date <=   " & SQLDate(txtToDateReSave.value, True) & " )"
    End If
    
                If chkIsBranch.value = vbChecked <> 0 Then
        StrSQL = StrSQL & "  and BranchID =   " & val(Me.dcBranch.BoundText)
        
          End If
          
                      If chkStore.value = vbChecked <> 0 Then
        StrSQL = StrSQL & "  and storeid =   " & val(Me.DCboStoreName.BoundText)
        
          End If
          
     If withoutJL.value = vbChecked Then
          StrSQL = StrSQL & "  and Transaction_ID in "
          StrSQL = StrSQL & "  ( Select Transaction_ID from Transactions where Transaction_Type=19 and NoteId not In (SELECT IsNull(notes_id,0) FROM DOUBLE_ENTREY_VOUCHERS where Credit_Or_Debit = 0))"
     End If
     If cmdReSave.Visible = True Then
        If chkIsPosOnly.value = vbChecked Then
            StrSQL = StrSQL & "  and Transaction_ID in "
            StrSQL = StrSQL & "  ( Select nots from Transactions where Transaction_Type=21 and POSBillType = 1 )"
        End If
       
       If chkWithoutCost.value = vbChecked Then
        
       StrSQL = StrSQL & "  and Transaction_ID in "
        StrSQL = StrSQL & "  (SELECT        TT.Transaction_ID"
        StrSQL = StrSQL & "   FROM            dbo.Transactions TT INNER JOIN"
        StrSQL = StrSQL & "                 dbo.Transaction_Details ON TT.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
        StrSQL = StrSQL & " Where (TT.Transaction_Type = 19) And GardQty <> price)"
        'Transaction_Details.Price > 3000) "
        End If
    
    
    If chkIsProplemOnly.value = vbChecked Then
    
            StrSQL = StrSQL & " AND EXISTS ("
            StrSQL = StrSQL & "   SELECT 1"
            StrSQL = StrSQL & "   FROM Transactions t21"
            StrSQL = StrSQL & "   INNER JOIN Transaction_Details td19"
            StrSQL = StrSQL & "       ON td19.Transaction_ID = Transactions.Transaction_ID"
            StrSQL = StrSQL & "   INNER JOIN Transaction_Details td21"
            StrSQL = StrSQL & "       ON td21.Transaction_ID = t21.Transaction_ID"
            StrSQL = StrSQL & "      AND td21.Item_ID = td19.Item_ID"
            StrSQL = StrSQL & "      AND ISNULL(td21.UnitId,0) = ISNULL(td19.UnitId,0)"
            StrSQL = StrSQL & "   WHERE t21.Transaction_Type = 21"
            
            ' --- «·—»ÿ: Nots („„ﬂ‰ ÌﬂÊ‰ ›ÌÂ „”«›« ) ---
            StrSQL = StrSQL & "     AND ISNUMERIC(LTRIM(RTRIM(ISNULL(Transactions.Nots,'')))) = 1"
            StrSQL = StrSQL & "     AND CAST(LTRIM(RTRIM(ISNULL(Transactions.Nots,''))) AS float) = t21.Transaction_ID"
            
            ' --- ›· — «· «—ÌŒ (··›« Ê—… 21) ‰›” › —… «·’—› ---
            StrSQL = StrSQL & "     AND t21.Transaction_Date >= " & SQLDate(txtFromDateReSave.value, True)
            StrSQL = StrSQL & "     AND t21.Transaction_Date <= " & SQLDate(txtToDateReSave.value, True)
            
            ' --- ‘—ÿ «·„‘ﬂ·…:  ﬂ·›… «·’—› > ”⁄— «·»Ì⁄ ---
            StrSQL = StrSQL & "     AND ISNULL(td19.Price,0) > ISNULL(td21.Price,0)"
            
            StrSQL = StrSQL & " )"
            
End If
       ' StrSQL = StrSQL & "  and Transactions.Transaction_ID =188072"

    ' StrSQL = StrSQL & "  and BillBasedOn =0 "
 End If
     If SystemOptions.usertype <> UserAdminAll And cmdReSave.Visible = False Then
 
          If SystemOptions.FixedCustomer = 1 Then
            StrSQL = StrSQL & " and  UserID = " & user_id
             End If
  
  
        Me.dcBranch.Enabled = True
      
      
    End If
    If cmdReSave.Visible = True Then
            StrSQL = StrSQL & " Order by Transaction_Date Desc"
    Else
            If SystemOptions.SortInvoiceByEntry Then
                StrSQL = StrSQL & " Order by Transaction_ID"
            Else
                StrSQL = StrSQL & " Order by noteserial1"
            End If
   End If
            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveLast
            End If
              

   Me.TxtModFlg.text = "R"

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
        If Not mClicked Then
            Cmd_Click (5)
        End If
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
    Dim StrSQL  As String
    Dim Num     As Integer
    Dim StrList As String
    Dim BGround As New ClsBackGroundPic
    Dim ShowTax As Boolean

    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    If SystemOptions.UserInterface = ArabicInterface Then

        With DcbType
            .Clear
            .AddItem "„⁄œ…"
            .AddItem "„·Õﬁ"
        End With

    Else

        With DcbType
            .Clear
            .AddItem "Equipment"
            .AddItem "Part"
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
    'Set m_menu1.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Excute").Picture

    Dim My_SQL As String
    'My_SQL = "  select branch_id,branch_name from TblBranchesData   "
    'fill_combo DcBranch, My_SQL
    ScreenNameArabic = "  ”‰œ ’—›   "
    ScreenNameEnglish = " Issue Voucher  "
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 180

    StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
    fill_combo Me.DcCostCenter, StrSQL

    If SystemOptions.usertype <> UserAdminAll Then
 
        Me.dcBranch.Enabled = True
    End If

    Set NewGrid.Grid = FG
 If SystemOptions.IsHiddenTransportInv Then
        lbl(53).Caption = "—ﬁ„ ÿ·» «—«„ﬂÊ"
    End If
    ShowTax = GetSetting(StrAppRegPath, "SallBill", "HaveTaxOnSalles", False)
    '    Ele(4).Visible = ShowTax
    'NewGrid.GridTrans = InvoiceTransaction
    
    NewGrid.GridTrans = InventoryOut

    Set NewGrid.TxtInvID = Me.XPTxtBillID
    Set NewGrid.TxtModFlag = TxtModFlg
    Set NewGrid.txtTotal = XPTxtSum
    Set NewGrid.lblTotalSalesPrice = lblTotalSalesPrice
    
    Set NewGrid.CboDiscount_Type = XPCboDiscountType
    Set NewGrid.TxtDiscount_Val = XPTxtDiscountVal
    Set NewGrid.TxtValueCash = XPTxtValue(0)
    Set NewGrid.TxtValueDelay = XPTxtValue(1)
    Set NewGrid.TxtValuechque = XPTxtValue(2)
    'Set NewGrid.
    '--------------------------------------
    'Set NewGrid.TxtTaxValue = Me.XPTxtTaxValue
    Set NewGrid.TxtAddTax = Me.TxtTaxAddValue
    'Set NewGrid.TxtStampTax = Me.TxtTaxStampValue
    Set NewGrid.TxtServiceTax = Me.TxtTaxServiceValue
    Set NewGrid.LblTotalQty = Me.LblTotalQty
    Set NewGrid.TxtItemCodeB1 = TxtItemCodeB1
    '------------------------------------------------
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.StoreName = Me.DCboStoreName
    Set NewGrid.DtpBillDate = Me.XPDtbBill
    Set NewGrid.CmdAddSerialLIst = Me.CmdSearch
    'Set NewGrid.CboDiscountType = CboDiscountType
    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    Set NewGrid.DCboItemName = DCboItemsName
    Set NewGrid.DCboItemCode = DCboItemsCode
    Set NewGrid.StoreName = DCboStoreName
    Set NewGrid.DtpBillDate = Me.XPDtbBill
      
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
    
    FG.ColHidden(Me.FG.ColIndex("MintID")) = True
    FG.ColHidden(Me.FG.ColIndex("GroupIDMint")) = True
    FG.WallPaper = BGround.Picture
    AddTip
    XPTab301.CurrTab = 0
    XPDtbBill.value = Date
    '********ship***********
    txtShipOrderNo = ""
    txtShipEnquieryNo = ""
    txtShipAccountNo = ""
    txtShipCustomerName = ""
    txtShipDistance = ""
    txtShipArea = ""
    txtShipSiteNo = ""
    txtShipProjectName = ""
    txtShipStructuralElement = ""
    txtShipMixDescription = ""
    txtShipDriverName = ""
    txtShipPipeLine = ""
    txtShipPump = ""
    txtShipTruckNo = ""
    txtShipIceTemp = ""
    txtShipTotalDeleveryd = "0"
    txtShipThisLoad = ""
    txtShipDayOrder = ""
    txtShipTripNo = ""
    txtShipPlantNo = ""
    txtShipBatched = ""

    SetDtpickerDate txtShipRestunedPlant
    SetDtpickerDate txtShipEndDischarge
    SetDtpickerDate txtShipStartDisCharge
    SetDtpickerDate txtShipOnSite
 
    '********************
    txtManualDate.value = Date
    txtRegDate.value = Date
 
    If SystemOptions.UserInterface = ArabicInterface Then

        With Me.DCOPrType
            .Clear
            .AddItem "»·«"
            .AddItem "„Ê«œ Œ«„ "
            .AddItem "„Â„« "
            .AddItem "ﬁÿ⁄ €Ì«— "
            .AddItem "«‰ «Ã  «„ "
            .AddItem "Âœ«Ì« Ê⁄Ì‰« "
            .AddItem "⁄Âœ…/ «’Ê· À«» …  "
        End With
    
        With CBoBasedON
            .Clear
            .AddItem "»·«"
            .AddItem "  «„— »Ì⁄"
            .AddItem " ›« Ê—…  „»Ì⁄«  "
            .AddItem "„—Êœ«  „‘ —Ì« "
            .AddItem "ÿ·» ⁄—÷ ”⁄—"
            .AddItem "⁄—÷ ”⁄— "
            .AddItem "”‰œ ‘Õ‰ "
            .AddItem "ÿ·» œ«Œ·Ì "
            .AddItem "√„— ‘€·  "
            .AddItem "«„— «’·«Õ-Ê—‘ "
            .AddItem "„—œÊœ«  „»Ì⁄«  "
            .AddItem "›« Ê—… „‘ —Ì« "
            .AddItem "”‰œ  Ã„Ì⁄"
            .AddItem "«„— ‘—«¡"
            .AddItem "«„— «‰ «Ã"
            .AddItem "«” ·«„ «‰ «Ã  «„"
        End With

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

        With Me.DCOPrType
            .Clear
            .AddItem "NA"
            .AddItem "Raw Material "
            .AddItem "Errands"
            .AddItem "Spare Parts"
            .AddItem "Finish Product "
            .AddItem "Gifts & Sambles"
            .AddItem "Fixed Assets  "
        End With
  
        With CBoBasedON
            .Clear
            .AddItem "NA"
            .AddItem " Sales Order"
            .AddItem " Sales Invoice "
            .AddItem "Return Purchase "
            .AddItem "Sales Qut. Order"
            .AddItem "Sales Qut."
            .AddItem "Shipment"
            .AddItem "Internal Request"
            .AddItem "Work Order"
            .AddItem "Repair Order-Cars"
            .AddItem "Return Sales invoice"
            .AddItem "Purchase invoice"
            .AddItem "Assemply"
            .AddItem "Purchase order"
            .AddItem "Production order"
            .AddItem "Production Recive Voucher"
            
       
        End With
    
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
    Dcombos.GetCustomersSuppliers 1, Me.DcCustmer
    
    '  Dim My_SQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select CusID,CusName from TblCustemers  "
        My_SQL = My_SQL & " INNER  JOIN Tbl_TradingContract TT ON TblCustemers.CusID =TT.TContract_CustID "
        My_SQL = My_SQL & " Where IsNull(Id,0) <>0"
        My_SQL = My_SQL & " And IsNull(IsCanceld,0) <> 1"
        My_SQL = My_SQL & " order by CusName "
        
    Else
        My_SQL = "  select CusID,CusNamee from TblCustemers  "
        My_SQL = My_SQL & " INNER  JOIN Tbl_TradingContract TT ON TblCustemers.CusID =TT.TContract_CustID "
        My_SQL = My_SQL & " Where IsNull(Id,0) <>0"
        My_SQL = My_SQL & " And IsNull(IsCanceld,0) <> 1"
        My_SQL = My_SQL & " order by CusName "
       
    End If

    fill_combo DcCustmer, My_SQL
    
    '--------------------------------
    If SystemOptions.UserInvoiceShowProfit = 0 Then
        Me.Ele(8).Visible = False
    Else
        ' Me.Ele(8).Visible = True
    End If

    SetDtpickerDate Me.XPDtbBill
    '----------------------------
    SetDtpickerDate Me.DtpDelayDate
    SetDtpickerDate Me.txtManualDate
    SetDtpickerDate Me.txtRegDate
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

    'Me.XPChkTAX.value = vbUnchecked
    'XPChkTAX_Click
    Me.ChkTaxAdd.value = vbUnchecked
    ChkTaxAdd_Click
    'Me.ChkTaxStamp.value = vbUnchecked
    'ChkTaxStamp_Click
    Me.ChkTaxSerivce.value = vbUnchecked
    ChkTaxSerivce_Click
    '---------------------------
    Resize_Form Me, TransactionSize
    '----------------------------
    'DB_CreateField "Transactions", "TransactionComment", adVarWChar, adColNullable, 255, , " ”ÃÌ· „·«ÕŸ«  ⁄·Ï «·›« Ê—…", False, True
    '----------------------------

    StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type= -19"
    StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"
 
    If SystemOptions.usertype <> UserAdminAll Then
        '      StrSQL = StrSQL & " AND   BranchId=" & branch_id
    End If

    StrSQL = StrSQL & "  Order by NoteSerial1 "

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveLast
    End If

    Retrive
    ' Me.TxtModFlg.Text = "R"

    InvType = 19

    txtPassword_Change

    If SystemOptions.HideCost = True Then
        LblTotalAll.Visible = False
        LblTotal.Visible = False

        TxtPrice.Visible = False
        FG.ColHidden(FG.ColIndex("Price")) = True
        FG.ColHidden(FG.ColIndex("Valu")) = True
    Else
        LblTotalAll.Visible = True
        LblTotal.Visible = True

        TxtPrice.Visible = True
        FG.ColHidden(FG.ColIndex("Price")) = False
        FG.ColHidden(FG.ColIndex("Valu")) = False

    End If

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    chkIsPosOnly.value = vbChecked
    Exit Sub
ErrTrap:
    Debug.Print Err.Description

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    Dim i As Integer
 RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, 180
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
                Me.Caption = "«–‰ «·’—›"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.Caption = "Issue Voucher"
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
            'XPChkTAX.Enabled = False

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
            'XPChkTAX.Enabled = False
            ChkTaxAdd.Enabled = False
            ChkTaxSerivce.Enabled = False
            'ChkTaxStamp.Enabled = False

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "  «–‰ «·’—› ( ÃœÌœ )"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.Caption = "Issue Voucher(New)"
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
            'XPChkTAX.Enabled = True
            'XPTxtTaxValue.text = ""
            'XPChkTAX.value = Unchecked
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
            'XPChkTAX.Enabled = True
            ChkTaxAdd.Enabled = True
            ChkTaxSerivce.Enabled = True
            'ChkTaxStamp.Enabled = True
        
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«–‰ «·’—›(   ⁄œÌ· )"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.Caption = "Issue Voucher( Edit )"
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
            'XPChkTAX.Enabled = True

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
            'XPChkTAX.Enabled = True
            ChkTaxAdd.Enabled = True
            ChkTaxSerivce.Enabled = True
            'ChkTaxStamp.Enabled = True
        
    End Select

    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails     As New ADODB.Recordset
    Dim StrSQL        As String
    Dim RsNotes       As New ADODB.Recordset
    Dim RsTest        As ADODB.Recordset
    Dim RsReplace     As ADODB.Recordset
    Dim LngPartID     As Long
    Dim RsPartDetails As ADODB.Recordset
    Dim i             As Long

    On Error GoTo ErrTrap
    '---------------------------------------------
    'Here We Reset all Setting
    Me.CmdNotes.Visible = False
    Me.CmdNotes.Tag = ""
    Me.CmdRetruns.Visible = False
    Me.CmdRetruns.Tag = ""

    ChkTaxAdd.value = vbUnchecked
    Me.TxtTaxAddValue.text = ""
    '    ChkTaxStamp.value = vbUnchecked
    'Me.TxtTaxStampValue.text = ""
    '    ChkTaxStamp.value = vbUnchecked
    '    Me.TxtTaxStampValue.text = ""
    '    ChkTaxSerivce.value = vbUnchecked
    '    Me.TxtTaxServiceValue.text = ""

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

    'Me.TxtModFlg.text = "R"
    TxtFillData.text = "T"
    Screen.MousePointer = vbArrowHourglass
    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", (rs("NoteSerial").value))
    Me.TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", (rs("NoteSerial1").value))
    Me.oldtxtNoteSerial1.text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)
    TxtOldOpOrderID.text = IIf(IsNull(rs("OldOpOrderID").value), "", (rs("OldOpOrderID").value))
    txtOrderID.text = IIf(IsNull(rs.Fields("OrderID").value), 0, rs.Fields("OrderID").value)
    lbl(56).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
    DcbProject.BoundText = IIf(IsNull(rs("project_id").value), "", (rs("project_id").value))
    Me.TXTNoteID.text = IIf(IsNull(rs("NoteID").value), "", (rs("NoteID").value))

    If Not IsNull(rs("general_cost_center").value) Then
        Me.DcCostCenter.BoundText = IIf(rs("general_cost_center").value = "", "", rs("general_cost_center").value)
    Else
        Me.DcCostCenter.BoundText = ""
    End If
 
    If IsNull(rs("chkDone").value) Then
        Me.chkDone.value = vbUnchecked
        Me.chkDone.Enabled = True
    Else

        If (rs("chkDone").value) = 0 Then
            Me.chkDone.value = vbUnchecked
            Me.chkDone.Enabled = True
        Else
            Me.chkDone.value = vbChecked
            Me.chkDone.Enabled = False
       
        End If
    End If

    XPTxtBillID.text = IIf(IsNull(rs("Transaction_ID").value), "", val(rs("Transaction_ID").value))
    TxtTicketNO.text = IIf(IsNull(rs("TicketNO").value), "", (rs("TicketNO").value))

    TxtTransSerial.text = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
    XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", (rs("Transaction_Date").value))
    txtManualDate.value = IIf(IsNull(rs("ManualDate").value), Date, (rs("ManualDate").value))
    txtRegDate.value = IIf(IsNull(rs("RegDate").value), Date, (rs("RegDate").value))
    DCDocTypes.BoundText = IIf(IsNull(rs("Doctype").value), "", rs("Doctype").value)
    '***********Ship*************************
    txtShipOrderNo = rs!ShipOrderNo & ""
    txtShipEnquieryNo = rs!ShipEnquieryNo & ""
    txtShipAccountNo = rs!ShipAccountNo & ""
    txtShipCustomerName = rs!ShipCustomerName & ""
    txtShipDistance = rs!ShipDistance & ""
    txtShipArea = rs!ShipArea & ""
    txtShipSiteNo = rs!ShipSiteNo & ""
    txtShipProjectName = rs!ShipProjectName & ""
    txtShipStructuralElement = rs!ShipStructuralElement & ""
    txtShipMixDescription = rs!ShipMixDescription & ""
    txtShipDriverName = rs!ShipDriverName & ""
    txtShipPipeLine = rs!ShipPipeLine & ""
    txtShipPump = rs!ShipPump & ""
    txtShipTruckNo = rs!ShipTruckNo & ""
    txtShipIceTemp = rs!ShipIceTemp & ""
    txtShipTotalDeleveryd = val(rs!ShipTotalDeleveryd & "")
    txtShipThisLoad = rs!ShipThisLoad & ""
    txtShipDayOrder = rs!ShipDayOrder & ""
    txtShipTripNo = rs!ShipTripNo & ""
    txtShipPlantNo = rs!ShipPlantNo & ""
    txtShipBatched = rs!ShipBatched & ""

    txtShipRestunedPlant = IIf(IsNull(rs("ShipRestunedPlant").value), Date, (rs("ShipRestunedPlant").value))
    txtShipEndDischarge = IIf(IsNull(rs("ShipEndDischarge").value), Date, (rs("ShipEndDischarge").value))
    txtShipStartDisCharge = IIf(IsNull(rs("ShipStartDisCharge").value), Date, (rs("ShipStartDisCharge").value))
    txtShipOnSite = IIf(IsNull(rs("ShipOnSite").value), Date, (rs("ShipOnSite").value))
    '************************************
    Me.DCExtraAccount.BoundText = IIf(IsNull(rs("ExtraAccount").value), "", rs("ExtraAccount").value)

    If Me.DCExtraAccount.BoundText = "" Then
        TxtExtraValue.text = 0
    Else
        TxtExtraValue.text = IIf(IsNull(rs("ExtraValue").value), 0, (rs("ExtraValue").value))
    End If

    XPCboDiscountType.ListIndex = IIf(IsNull(rs("Trans_DiscountType").value), -1, val(rs("Trans_DiscountType").value))
    CboPayMentType.ListIndex = IIf(IsNull(rs("PaymentType").value), 0, rs("PaymentType").value)
    XPTxtDiscountVal.text = IIf(IsNull(rs("Trans_Discount").value), "", (rs("Trans_Discount").value))
    Me.DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    FG.Clear flexClearScrollable, flexClearEverything
    Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
    Me.DCboStoreName2.BoundText = IIf(IsNull(rs("StoreID1").value), "", rs("StoreID1").value)
    DCEquipments.BoundText = IIf(IsNull(rs("FixesAssetsID").value), "", rs("FixesAssetsID").value)

    Me.DcbType.ListIndex = IIf(IsNull(rs("Head_Details").value), -1, rs("Head_Details").value)
    Me.DcboEmp.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
    '    XPTxtTaxValue.text = IIf(IsNull(rs("TaxValue").value), "", (rs("TaxValue").value))
    '    XPChkTAX.value = IIf(rs("TaxFound") = True, Checked, Unchecked)
    Text1.text = IIf(IsNull(rs("nots").value), "", (rs("nots").value))
    Txtnots2.text = IIf(IsNull(rs("nots2").value), "", (rs("nots2").value))
    TxtWorkOrderNO.text = IIf(IsNull(rs("WorkOrderNO").value), "", (rs("WorkOrderNO").value))

    dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", (rs("BranchId").value))

    DcboEmpDepartments.BoundText = IIf(IsNull(rs("DepartementID").value), "", rs("DepartementID").value)
    DcboEmpDepartments.BoundText = IIf(IsNull(rs("DepartementID").value), "", rs("DepartementID").value)
    DcboEmpName.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)

    Me.DCDriver.BoundText = IIf(IsNull(rs("DriverId").value), "", rs("DriverId").value)
    Me.DCCar.BoundText = IIf(IsNull(rs("CarId").value), "", rs("CarId").value)

    'CBOOrderType.ListIndex = IIf(IsNull(rs("OrderType").value), 0, rs("OrderType").value)
    txtManualNO.text = IIf(IsNull(rs("ManualNO").value), "", (rs("ManualNO").value))
    Dim TContractCustID As Double
    txtTradingContractID.text = IIf(IsNull(rs("TradingContractID").value), "", (rs("TradingContractID").value))
    Get_TradingContractinfo val(txtTradingContractID.text), TContractCustID, 0

    DcCustmer.BoundText = val(TContractCustID)

    CBoBasedON.ListIndex = IIf(IsNull(rs("BillBasedOn").value), 2, rs("BillBasedOn").value)

    If CBoBasedON.ListIndex = 2 And Text1.text <> "" Then
        TXT_order_no.text = IIf(IsNull(rs("nots2").value), "", rs("nots2").value)
    Else
        TXT_order_no.text = IIf(IsNull(rs("order_no").value), "", rs("order_no").value)
    End If
 
    DCOPrType.ListIndex = IIf(IsNull(rs("OPrType").value), 0, rs("OPrType").value)

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
            'ChkTaxStamp.value = vbChecked
            'Me.TxtTaxStampValue.text = rs("TaxStampValue").value
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
    ' StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = " SELECT      dbo.TblItems.HaveSerial, Transaction_Details.Remarks AS Remarks22, *, TblMaintenanceType_2.name AS MintName, TblMaintenanceType_2.namee AS MintNameE, "
    StrSQL = StrSQL + "                  TblMaintenanceType_1.name AS GroupMint, TblMaintenanceType_1.namee AS GroupMintE"
    StrSQL = StrSQL + "   FROM         dbo.TblItems INNER JOIN"
    StrSQL = StrSQL + "                  dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID INNER JOIN"
    StrSQL = StrSQL + "                  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL + "                  dbo.TblMaintenanceType TblMaintenanceType_2 ON dbo.Transaction_Details.MintID = TblMaintenanceType_2.id LEFT OUTER JOIN"
    StrSQL = StrSQL + "                  dbo.TblMaintenanceType TblMaintenanceType_1 ON dbo.Transaction_Details.GroupIDMint = TblMaintenanceType_1.id"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)

    'StrSQL = StrSQL + " ORDER BY dbo.Transaction_Details.ID"
    StrSQL = StrSQL + " Order By IsNull(Transaction_Details.LineID,Transaction_Details.Id)"

    Set RsDetails = New ADODB.Recordset
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.text = ""
    lblTotalSalesPrice = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.rows = RsDetails.RecordCount + 1

        For i = 1 To RsDetails.RecordCount
            FG.TextMatrix(i, FG.ColIndex("MixNo")) = IIf(IsNull(RsDetails("MixNo").value), "", RsDetails("MixNo").value)
           
            FG.TextMatrix(i, FG.ColIndex("OldID")) = IIf(IsNull(RsDetails("OldID").value), 0, RsDetails("OldID").value)
            FG.TextMatrix(i, FG.ColIndex("QtyFaqtors")) = IIf(IsNull(RsDetails("QtyFaqtors").value), "", RsDetails("QtyFaqtors").value)
            FG.TextMatrix(i, FG.ColIndex("MaxQty")) = IIf(IsNull(RsDetails("MaxQty").value), "", RsDetails("MaxQty").value)
            FG.TextMatrix(i, FG.ColIndex("MaxUnitID")) = IIf(IsNull(RsDetails("MaxUnitID").value), "", RsDetails("MaxUnitID").value)
            FG.TextMatrix(i, FG.ColIndex("GroupIDMint")) = IIf(IsNull(RsDetails("GroupIDMint").value), "", RsDetails("GroupIDMint").value)
            FG.TextMatrix(i, FG.ColIndex("MintID")) = IIf(IsNull(RsDetails("MintID").value), "", RsDetails("MintID").value)
        
            FG.TextMatrix(i, FG.ColIndex("FoxyNo")) = IIf(IsNull(RsDetails("FoxyNo")), "", (RsDetails("FoxyNo").value))
            FG.TextMatrix(i, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no").value), "", RsDetails("order_no").value)
            FG.TextMatrix(i, FG.ColIndex("OrderArrivalDate")) = IIf(IsNull(RsDetails("OrderArrivalDate").value), "", RsDetails("OrderArrivalDate").value)
   
            FG.cell(flexcpPicture, i, FG.ColIndex("Ser")) = ""
            FG.cell(flexcpData, i, FG.ColIndex("Ser")) = ""
            FG.TextMatrix(i, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(i, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim$(RsDetails("Item_ID").value))
            FG.TextMatrix(i, FG.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial")), "", Trim(RsDetails("ItemSerial").value))
            
     '       FG.TextMatrix(i, FG.ColIndex("EmpID4")) = IIf(IsNull(RsDetails("EmpID4")), "", Trim(RsDetails("EmpID4").value))
            
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

            FG.TextMatrix(i, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("ShowQty")), "", (RsDetails("ShowQty").value))
            FG.TextMatrix(i, FG.ColIndex("ItemsDetailsNewidea")) = IIf(IsNull(RsDetails("ItemsDetailsNewidea")), "", (RsDetails("ItemsDetailsNewidea").value))
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
            FG.TextMatrix(i, FG.ColIndex("QtyFaqtors")) = IIf(IsNull(RsDetails("QtyFaqtors")), "", (RsDetails("QtyFaqtors").value))
            FG.TextMatrix(i, FG.ColIndex("ItemCostPrice")) = IIf(IsNull(RsDetails("CostPrice")), "", (RsDetails("CostPrice").value))
            FG.TextMatrix(i, FG.ColIndex("PofTransID")) = IIf(IsNull(RsDetails("CostTransID")), "", (RsDetails("CostTransID").value))
            FG.TextMatrix(i, FG.ColIndex("ItemProfit")) = IIf(IsNull(RsDetails("ItemProfit")), "", (RsDetails("ItemProfit").value))
            FG.TextMatrix(i, FG.ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks22")), "", (RsDetails("Remarks22").value))
 
            FG.TextMatrix(i, FG.ColIndex("Height")) = IIf(IsNull(RsDetails("Height")), "", (RsDetails("Height").value))
            FG.TextMatrix(i, FG.ColIndex("length")) = IIf(IsNull(RsDetails("length")), "", (RsDetails("length").value))
            FG.TextMatrix(i, FG.ColIndex("Width")) = IIf(IsNull(RsDetails("Width")), "", (RsDetails("Width").value))
            FG.TextMatrix(i, FG.ColIndex("OUTR")) = IIf(IsNull(RsDetails("OUTR")), "", (RsDetails("OUTR").value))
            FG.TextMatrix(i, FG.ColIndex("INR")) = IIf(IsNull(RsDetails("INR")), "", (RsDetails("INR").value))
            FG.TextMatrix(i, FG.ColIndex("NoCount")) = IIf(IsNull(RsDetails("NoCount")), "", (RsDetails("NoCount").value))
        
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

            FG.TextMatrix(i, FG.ColIndex("Height")) = IIf(IsNull(RsDetails("Height")), "", (RsDetails("Height").value))
            FG.TextMatrix(i, FG.ColIndex("length")) = IIf(IsNull(RsDetails("length")), "", (RsDetails("length").value))
            FG.TextMatrix(i, FG.ColIndex("Width")) = IIf(IsNull(RsDetails("Width")), "", (RsDetails("Width").value))
            FG.TextMatrix(i, FG.ColIndex("OUTR")) = IIf(IsNull(RsDetails("OUTR")), "", (RsDetails("OUTR").value))
            FG.TextMatrix(i, FG.ColIndex("INR")) = IIf(IsNull(RsDetails("INR")), "", (RsDetails("INR").value))
            FG.TextMatrix(i, FG.ColIndex("NoCount")) = IIf(IsNull(RsDetails("NoCount")), "", (RsDetails("NoCount").value))

            FG.cell(flexcpData, i, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))

            If SystemOptions.UserInterface = ArabicInterface Then
                FG.TextMatrix(i, FG.ColIndex("MintName")) = IIf(IsNull(RsDetails("MintName")), "", (RsDetails("MintName").value))
                FG.TextMatrix(i, FG.ColIndex("GroupMint")) = IIf(IsNull(RsDetails("GroupMint")), "", (RsDetails("GroupMint").value))
                FG.TextMatrix(i, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            Else
                FG.TextMatrix(i, FG.ColIndex("GroupMint")) = IIf(IsNull(RsDetails("GroupMintE")), "", (RsDetails("GroupMintE").value))
                FG.TextMatrix(i, FG.ColIndex("MintName")) = IIf(IsNull(RsDetails("MintNameE")), "", (RsDetails("MintNameE").value))
                FG.TextMatrix(i, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitNamee")), "", (RsDetails("UnitNamee").value))
            End If

            FG.TextMatrix(i, FG.ColIndex("ProductionDate")) = IIf(IsNull(RsDetails("ProductionDate")), "", (RsDetails("ProductionDate").value))
            FG.TextMatrix(i, FG.ColIndex("ExpiryDate")) = IIf(IsNull(RsDetails("ExpiryDate")), "", (RsDetails("ExpiryDate").value))
            FG.TextMatrix(i, FG.ColIndex("LotNO")) = IIf(IsNull(RsDetails("LotNO")), "", (RsDetails("LotNO").value))
    If Me.FG.ColIndex("TotalInvoiceQty") <> -1 Then
            FG.TextMatrix(i, FG.ColIndex("TotalInvoiceQty")) = IIf(IsNull(RsDetails("TotalInvoiceQty")), "", (RsDetails("TotalInvoiceQty").value))
        End If
        If Me.FG.ColIndex("ISSUEDQTY") <> -1 Then
           FG.TextMatrix(i, FG.ColIndex("ISSUEDQTY")) = IIf(IsNull(RsDetails("ISSUEDQTY")), "", (RsDetails("ISSUEDQTY").value))
   End If
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
    StrSQL = "Select * From Notes Where Transaction_ID=" & val(rs("Transaction_ID").value)
    RsNotes.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsNotes.EOF Or RsNotes.BOF) Then

        For i = 1 To RsNotes.RecordCount

            If RsNotes("NoteType").value = 0 Then
                XPChkPayType(0).value = Checked
                XPChkPayType_Click (0)
                XPTxtValue(0).text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtSerial(0).text = IIf(IsNull(RsNotes("NoteSerial").value), "", Trim$(RsNotes("NoteSerial").value))
                Me.DcboBox.BoundText = IIf(IsNull(RsNotes("BoxID").value), "", RsNotes("BoxID").value)
            End If

            If RsNotes("NoteType").value = 1 Then
                XPChkPayType(1).value = Checked
                XPChkPayType_Click (1)
                XPTxtValue(1).text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtValue(1).Tag = IIf(IsNull(RsNotes("NoteID").value), "", (RsNotes("NoteID").value))
                XPTxtSerial(1).text = IIf(IsNull(RsNotes("NoteSerial").value), "", (RsNotes("NoteSerial").value))
                DtpDelayDate.value = IIf(IsNull(RsNotes("DueDate").value), "", (RsNotes("DueDate").value))
            End If

            If RsNotes("NoteType").value = 2 Then
                XPChkPayType(2).value = Checked
                XPChkPayType_Click (2)
            End If

            RsNotes.MoveNext
        Next i

    End If

    Set RsNotes = New ADODB.Recordset
    StrSQL = "SELECT Notes.NoteID, Notes.NoteDate, Notes.NoteType, Notes.NoteSerial," & "Notes.Note_Value, Notes.BankID,BanksData.BankName , Notes.ChqueNum, Notes.DueDate "
    StrSQL = StrSQL + " FROM Notes INNER JOIN BanksData ON Notes.BankID = BanksData.BankID "
    StrSQL = StrSQL + " Where NoteType=2 AND NOTES.Transaction_ID=" & val(rs("Transaction_ID").value)
    StrSQL = StrSQL + " Order BY Notes.NoteID"
    RsNotes.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.FgCheques
        .rows = .FixedRows

        If Not (RsNotes.BOF Or RsNotes.EOF) Then
            .rows = .FixedRows + RsNotes.RecordCount

            For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("CheckValue")) = IIf(IsNull(RsNotes("Note_Value").value), "", RsNotes("Note_Value").value)
                .TextMatrix(i, .ColIndex("CheckNumber")) = IIf(IsNull(RsNotes("ChqueNum").value), "", RsNotes("ChqueNum").value)
                .TextMatrix(i, .ColIndex("BankID")) = IIf(IsNull(RsNotes("BankID").value), "", RsNotes("BankID").value)
                .TextMatrix(i, .ColIndex("BankName")) = IIf(IsNull(RsNotes("BankName").value), "", RsNotes("BankName").value)

                If Not IsNull(RsNotes("DueDate").value) Then
                    .TextMatrix(i, .ColIndex("DueDate")) = DisplayDate(RsNotes("DueDate").value)
                Else
                    .TextMatrix(i, .ColIndex("DueDate")) = ""
                End If

                RsNotes.MoveNext
            Next i

        End If

        .AutoSize 0, .Cols - 1, False
        SumChecks
    End With

    '⁄—÷ «·√ﬁ”«ÿ «·„”Ã·… ⁄·Ï «·›« Ê—…
    If XPTxtValue(1).Tag <> "" Then
        StrSQL = "Select * From InstallMent where NoteID=" & XPTxtValue(1).Tag
        Set RsTest = New ADODB.Recordset
        RsTest.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTest.EOF Or RsTest.BOF) Then
            CmdINSTALLMENT.Enabled = True
            CmdINSTALLMENT.Caption = "⁄—÷ «·√ﬁ”«ÿ «·„”Ã·…"
            LngPartID = RsTest("PartID").value
            Me.LblPrecenType.Tag = RsTest("InterestType").value

            If RsTest("InterestType").value = 0 Then
                LblPrecenType.Caption = "‰”»… „∆ÊÌ…"
            ElseIf RsTest("InterestType").value = 1 Then
                LblPrecenType.Caption = "ﬁÌ„… À«» …"
            ElseIf RsTest("InterestType").value = 2 Then
                LblPrecenType.Caption = "·«ÌÊÃœ"
            End If

            Me.LblPrecenValue.Caption = RsTest("InterestVal").value
            Me.LblInstallTotal.Caption = RsTest("Total").value
            Me.LblInstallCount.Caption = RsTest("InstallCount").value
            Me.LblFirstInstallDate.Caption = DisplayDate(RsTest("FirstInstallDate").value)
            Me.LblInstallmentType.Tag = RsTest("InstallmentType").value

            If RsTest("InstallmentType").value = 0 Then
                LblInstallmentType.Caption = "ÌÊ„"
            ElseIf RsTest("InstallmentType").value = 1 Then
                LblInstallmentType.Caption = "‘Â—"
            ElseIf RsTest("InstallmentType").value = 2 Then
                LblInstallmentType.Caption = "”‰…"
            End If

            Me.LblInstallSeprator.Caption = RsTest("InstallSeprator").value
            Me.LblStartValue.Caption = IIf(IsNull(RsTest("StartValue").value), "", RsTest("StartValue").value)
            Set RsPartDetails = New ADODB.Recordset
            StrSQL = "Select * From InstallMentDetails Where PartID=" & LngPartID
            RsPartDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsPartDetails.BOF Or RsPartDetails.EOF) Then
                RsPartDetails.MoveFirst

                With Me.FgInstallments
                    .rows = .FixedRows + RsPartDetails.RecordCount

                    For i = .FixedRows To .rows - 1
                        .TextMatrix(i, .ColIndex("QestID")) = IIf(IsNull(RsPartDetails("QestID").value), "", RsPartDetails("QestID").value)
                        .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(RsPartDetails("Value").value), "", RsPartDetails("Value").value)

                        If Not IsNull(RsPartDetails("DueDate").value) Then
                            .TextMatrix(i, .ColIndex("Due_Date")) = DisplayDate(RsPartDetails("DueDate").value)
                        Else
                            .TextMatrix(i, .ColIndex("Due_Date")) = ""
                        End If

                        RsPartDetails.MoveNext
                    Next i

                End With

            End If

        Else
            CmdINSTALLMENT.Enabled = False
            CmdINSTALLMENT.Caption = " ﬁ”Ìÿ «·ﬁÌ„… «·¬Ã·…"
        End If
    End If

    TxtFillData.text = "F"
    '-----------------------------------------------------------------------------------------------
    Dim SngRelatedNotesValues As Single
    Me.CmdNotes.Visible = ShowRelatedNotes(val(Me.XPTxtBillID.text), 0, SngRelatedNotesValues)
    Me.CmdNotes.Tag = SngRelatedNotesValues

    SngRelatedNotesValues = 0
    Me.CmdRetruns.Visible = ShowRelatedTransactions(val(Me.XPTxtBillID.text), 0, SngRelatedNotesValues)
    Me.CmdRetruns.Tag = SngRelatedNotesValues
    mIsFinishSave = True
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
    fillapprovData
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
            Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                clear_all Me
                Me.TxtModFlg.text = "R"
                XPBtnMove_Click (1)
            End If

        Case "E"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ⁄œÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

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

    'On Error GoTo ErrTrap
    If XPTxtBillID.text = "" Then
        clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Sub
    End If

    'If AvailableDeal = False Then
    '    Exit Sub
    'End If
    '«·√ﬁ”«ÿ «·„”œœ… ⁄·Ï «·›« Ê—…
 
    '⁄„·Ì«  «·’Ì«‰… «·„— »ÿ… »«·›« Ê—…
 
    Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
    Msg = Msg + (Me.TxtNoteSerial1.text) & Chr(13)
    Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
    IntRes = MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)

    If IntRes = vbYes Then
        If Not rs.RecordCount < 1 Then
            Cn.BeginTrans
            BegainTrans = True
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS  " & "Where DOUBLE_ENTREY_VOUCHERS.Transaction_ID=" & rs("Transaction_ID").value
            Cn.Execute StrSQL, , adExecuteNoRecords
        
            StrSQL = "delete From Notes where  NoteType=180 and   noteid=" & val(TXTNoteID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
         
         
    'Dim StrSQL As String

    StrSQL = "Delete  marakes_taklefa_temp  where general_des=1 AND kedno =" & val(TXTNoteID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords


            StrSQL = "Delete  marakes_taklefa_temp  where   kedno =" & val(TXTNoteID.text)
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

    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·”Ã· "
    Msg = Msg & Chr(13) & Err.Number
    Msg = Msg & Chr(13) & Err.Description
    Msg = Msg & Chr(13) & Err.Source
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title

    If BegainTrans = True Then
        rs.CancelUpdate
        Cn.RollbackTrans
        BegainTrans = False
    End If

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

Public Sub RetriveSerialsx(ItemID As String, _
                          ItemName As String, _
                          seriallist As String, _
                          currentrow As Long, Optional Price As Double)
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
Private Sub AddTip()
    Dim Wrap As String
    Dim BolRtl As Boolean

    On Error GoTo ErrTrap
    Wrap = Chr(13) + Chr(10)
    Set TTP = New clstooltip

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True

        With TTP
            .Create Me.hWnd, "»Ì«‰«    ”‰œ «·’—› «·„Œ“‰Ì ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«      ”‰œ  ’—›  „Œ“‰Ì   ÃœÌœ" & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F12 OR Enter", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«   ”‰œ  ’—›  „Œ“‰Ì", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…" & Wrap & "„›« ÌÕ «·«Œ ’«— F6", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«     ”‰œ  ’—›  „Œ“‰Ì", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«   ”‰œ  ’—›  „Œ“‰Ì " & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F11", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«     ”‰œ  ’—›  „Œ“‰Ì ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  ⁄„·Ì…   ”‰œ  ’—›  „Œ“‰Ì  ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F10", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì…   ”‰œ  ’—›  „Œ“‰Ì " & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F9", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  ⁄„·Ì…  ”‰œ  ’—›  „Œ“‰Ì " & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F8", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«   ’—›  „Œ“‰Ì  ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄„·Ì…  ”‰œ  ’—›  „Œ“‰Ì " & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F7", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«    ’—›  „Œ“‰Ì", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— Ctrl + X", BolRtl
        End With
    
        With TTP
            .Create Me.hWnd, "»Ì«‰«    ”‰œ ’—›  „Œ“‰Ì", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnNewClients, "≈÷«›… ⁄„Ì· ÃœÌœ ..." & Wrap & "· ”ÃÌ· »Ì«‰«  ⁄„Ì· ÃœÌœ" & Wrap & " «÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F5", BolRtl
        End With
    
        With TTP
            .Create Me.hWnd, "»Ì«‰«    ’—›  „Œ“‰Ì  ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«      ’—›  „Œ“‰Ì", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«     ”‰œ  ’—›  „Œ“‰Ì ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«   ”‰œ  ’—›  „Œ“‰Ì   ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«    ”‰œ  ’—›  „Œ“‰Ì   ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, BolRtl
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        BolRtl = False

        With TTP
            .Create Me.hWnd, "Issue Voucher", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "New..." & Wrap & "Click here to add new Issue Voucher" & Wrap & "" & Wrap & "Shortcut (Enter Or F12)", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Issue Voucher", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "Print..." & Wrap & "Print this Issue Voucher" & Wrap & "" & Wrap & "Shortcut (F6)", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Issue Voucher", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), "Edit..." & Wrap & "Edit this Issue Voucher Record" & Wrap & "  " & Wrap & "Shortcut (F11)", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Issue Voucher", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Save..." & Wrap & "Save the New Issue Voucher Or Save the edit" & Wrap & "in the current Issue Voucher" & Wrap & "" & Wrap & "Shortcut (F10)", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Issue Voucher", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), "Undo..." & Wrap & "Undo in the New Issue Voucher" & Wrap & "Or Undo in the Editing" & Wrap & "" & Wrap & "Shortcut (F9)", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Issue Voucher", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Delete..." & Wrap & "Delete this current Issue Voucher" & Wrap & "" & Wrap & "Shortcut (F8)", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Issue Voucher", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "Search..." & Wrap & "Click here to display the search" & Wrap & "Screen" & Wrap & "Shortcut (F7)", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Issue Voucher", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Exit..." & Wrap & "Close this Window", BolRtl
        End With
    
        With TTP
            .Create Me.hWnd, "Issue Voucher", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnNewClients, "Add New Customer...." & Wrap & "To add New Customer Click here..." & Wrap & "Shortcut (F5)", BolRtl
        End With
    
        With TTP
            .Create Me.hWnd, "Issue Voucher", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "First..." & Wrap & "Move to first Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Issue Voucher", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "Previous..." & Wrap & "Move to Previous Record" & Wrap & " , BolRTL"
        End With

        With TTP
            .Create Me.hWnd, "Issue Voucher", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "Next..." & Wrap & "Move to Next Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Issue Voucher", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "Last..." & Wrap & "Move to Last Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Issue Voucher", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "Help..." & Wrap & "to View Help Files" & Wrap & "click Here" & Wrap & "Shortcut(F1)" & Wrap, BolRtl
        End With

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()


    If CheckFilegrid() = False Then
        Cmd(2).Enabled = True
        Cmd(2).Enabled = True
        Exit Sub
    End If
         
    Dim usedaccount    As Integer
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

    On Error GoTo ErrTrap
    
    Me.FG.FinishEditing True
    '****************************
    '· Ã«Â· Õ›Ÿ «· ›«’Ì· „⁄ «⁄«œÂ Ÿ»ÿ «·Õ—ﬂ« 
    Dim mSaveDetails As Boolean
    Dim mRebuildOnly As Boolean
    mRebuildOnly = (IsSaveWithOutMsg = True)
    
    mSaveDetails = (IsSaveWithOutMsg And chkIgnorDetails.value = 1) Or Not IsSaveWithOutMsg

    '***********************
    If IsSaveWithOutMsg Then GoTo SaveDirect

    DoEvents
    
    If MintDone = 1 And CBoBasedON.ListIndex = 8 Then

        For i = 1 To FG.rows - 1

            If val(FG.TextMatrix(i, FG.ColIndex("Code"))) <> 0 And Trim(FG.TextMatrix(i, FG.ColIndex("Remarks"))) = "" Then
                MsgBox "ÌÃ» «œŒ«· «·„·«ÕŸ«  ⁄·Ï «·”ÿ—"
                Cmd(2).Enabled = True
                Exit Sub
            End If

        Next
    
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "√„— ‘€· «·’Ì«‰… „‰ ÂÌ Ê·« Ì„ﬂ‰ «‰‘«¡ ”‰œ ’—›/ ”·Ì„ »‰«¡ ⁄·ÌÂ"
        Else
            MsgBox "this maintenance order is done and can't creat Issue Voucher based on it "
        End If

        Cmd(2).Enabled = True
        Exit Sub
    End If

    If MintDone = 1 And CBoBasedON.ListIndex = 9 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "√„— «’·«Õ «·’Ì«‰… „‰ ÂÌ Ê·« Ì„ﬂ‰ «‰‘«¡ ”‰œ ’—›/ ”·Ì„ »‰«¡ ⁄·ÌÂ"
        Else
            MsgBox "This maintenance order is done and can't creat Issue Voucher based on it "
        End If

        Cmd(2).Enabled = True
        Exit Sub
    End If
    
    Screen.MousePointer = vbArrowHourglass

    If val(CBoBasedON.ListIndex) = 8 Then
        If MintDone = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ «Œ Ì«— «„— ‘€· "
            Else
                MsgBox "Please select type"
            End If

            Cmd(2).Enabled = True
            Exit Sub
        End If

        If val(Me.DcbType.ListIndex) = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ «Œ Ì«— «·‰Ê⁄"
            Else
                MsgBox "Please select type"
            End If

            Me.DcbType.SetFocus
            Cmd(2).Enabled = True
            Exit Sub
        End If

        If val(DCEquipments.BoundText) = 0 Or DCEquipments.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ «Œ Ì«— «·„⁄œÂ"
            Else
                MsgBox "Please select equipment"
            End If

            DCEquipments.SetFocus
            Cmd(2).Enabled = True
            Exit Sub
        End If
    End If
    
    If Trim(Me.TxtTransSerial.text) = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ≈œŒ«· —ﬁ„ «·”‰œ...!!"
        Else
            Msg = "Enter Voucher No..!!"
        End If
    
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtTransSerial.SetFocus
        Screen.MousePointer = vbDefault
        Cmd(2).Enabled = True
        Exit Sub
    Else

        If Me.TxtModFlg.text = "N" Then
            BolTemp = UniqueTransSerial(Trim(Me.TxtTransSerial.text), 19, , val(Me.dcBranch.BoundText))
        ElseIf Me.TxtModFlg.text = "E" Then
            BolTemp = UniqueTransSerial(Trim(Me.TxtTransSerial.text), 19, val(Me.XPTxtBillID.text), val(Me.dcBranch.BoundText))
        End If
    
        'If BolTemp = False Then
        'Msg = "—ﬁ„ «·”‰œ  „”Ã· „”»ﬁ« ›Ï «·»—‰«„Ã.." & Chr(13)
        'Msg = Msg & "Ê·«Ì„ﬂ‰  ﬂ—«— —ﬁ„ «·”‰œ"
        'MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        'TxtTransSerial.SetFocus
        'Screen.MousePointer = vbDefault
        'Exit Sub
        'End If
    End If

    '«· √ﬂœ „‰ ⁄œ„  ﬂ—«— —ﬁ„ «·”‰œ
    If Voucher_coding(val(dcBranch.BoundText), XPDtbBill.value, 10, 180, , 19) = "" Then
        If Me.TxtModFlg.text = "N" Then
            BolTemp = UniqueNoteSerial1(Trim(Me.TxtNoteSerial1.text), 19, , val(dcBranch.BoundText))
        ElseIf Me.TxtModFlg.text = "E" Then
            BolTemp = UniqueNoteSerial1(Trim(Me.TxtNoteSerial1.text), 19, val(Me.XPTxtBillID.text), val(dcBranch.BoundText))
        End If
 
        If BolTemp = False Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "—ﬁ„ «·”‰œ „”Ã· „”»ﬁ« ›Ï «·»—‰«„Ã.." & Chr(13)
                Msg = Msg & "Ê·«Ì„ﬂ‰  ﬂ—«— —ﬁ„ «·”‰œ"
            Else
                Msg = "This Bill No Already Exist" & Chr(13)
        
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            TxtNoteSerial1.SetFocus
            Screen.MousePointer = vbDefault
            Cmd(2).Enabled = True
            Exit Sub
        End If
    End If

    '‰Â«Ì… «· √ﬂœ
    
    If DCboStoreName.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ «·„Œ“‰"
        Else
            Msg = "Specify Store"
        End If
    
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        DCboStoreName.SetFocus
        Sendkeys "{F4}"
        Screen.MousePointer = vbDefault
        Cmd(2).Enabled = True
        Exit Sub
    End If

    If val(TxtExtraValue.text) <> 0 And DCExtraAccount.text = "" Then
        Msg = "ÌÃ» «œŒ«· ﬁÌ„… «·«÷«›«   "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtExtraValue.SetFocus
        Sendkeys "{F4}"
        Screen.MousePointer = vbDefault
        Cmd(2).Enabled = True
        Exit Sub
    End If

    'If Trim(DcboEmp.BoundText) = "" Then
    '    Msg = "ÌÃ»  ÕœÌœ «”„ «·„ÊŸ›..!!!"
    '    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    DcboEmp.SetFocus
    '    SendKeys "{F4}"
    '    Screen.MousePointer = vbDefault
    '    Exit Sub
    'End If
    'If XPDtbBill.value = "" Then
    '    Msg = "ÌÃ»  ÕœÌœ  «—ÌŒ «·»Ì⁄"
    '    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    XPDtbBill.SetFocus
    '    SendKeys "{F4}"
    '    Screen.MousePointer = vbDefault
    '    Exit Sub
    'End If
    'If CboPaymentType.ListIndex = -1 Then
    '    Msg = "ÌÃ»  ÕœÌœ ÿ—Ìﬁ… «·œ›⁄"
    '    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    CboPaymentType.SetFocus
    '    SendKeys "{F4}"
    '    Screen.MousePointer = vbDefault
    '    Exit Sub
    'End If
    'If XPChkPayType(0).value = vbChecked Then
    '    If Me.DcboBox.BoundText = "" Then
    '        MsgBox "ÌÃ»  ÕœÌœ «”„ «·Œ“‰…...!!!", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '        Screen.MousePointer = vbDefault
    '        Exit Sub
    '    End If
    'End If
    '----------------------------------------------

    If val(Me.XPTxtValue(1).text) > 0 Then
        If ChkInstall.value = vbChecked Then
            If val(Me.LblInstallTotal.Caption) = 0 Then
                Msg = "ÌÃ» Õ”«» «·√ﬁ”«ÿ ﬁ»· ⁄„·Ì… «·Õ›Ÿ..!!!"
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Me.XPTab301.CurrTab = 1
                Screen.MousePointer = vbDefault
                Cmd(2).Enabled = True
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
            Cmd(2).Enabled = True
            Exit Sub
        End If
    End If

    ' If XPChkTAX.value = Checked Then
    '     If XPTxtTaxValue.text = "" Then
    '         Msg = "ÌÃ» «œŒ«· ﬁÌ„… ÷—Ì»… «·„»Ì⁄« "
    '         MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    '         XPTxtTaxValue.SetFocus
    '         Fg.SetFocus
    '         Screen.MousePointer = vbDefault
    '         Exit Sub
    '     End If
    ' End If

    If XPCboDiscountType.ListIndex = 1 Or XPCboDiscountType.ListIndex = 2 Then
        If XPTxtDiscountVal.text = "" Then
            Msg = "≈–« ﬂ«‰ Â‰«ﬂ Œ’„ ⁄·Ï «·›« Ê—… " & Chr(13)
            Msg = Msg + "ÌÃ»  ÕœÌœ ﬁÌ„… Â–« «·Œ’„ " & Chr(13)
            Msg = Msg + "√Ê √Œ Ì«— ·« ÌÊÃœ Œ’„ "
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPCboDiscountType.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            Cmd(2).Enabled = True
            Exit Sub
        End If
    End If

    '--------------------------------
    '«·ﬂ‘› ⁄·Ï „œÌÊ‰Ì… «·⁄„Ì· '

    '--------------------------------
    '-------------------------------

    Me.XPTab301.CurrTab = 0

    If NewGrid.CheckDataEntered = False Then
        Cmd(2).Enabled = True
        Exit Sub
    End If

SaveDirect:

    If NewGrid.Calculate(1, , False, True) = False Then
        Cmd(2).Enabled = True
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

    ' mina If DblNotesTotal <> Val(LblTotal.Caption) Then
    '    Msg = "≈Ã„«·Ï «·√Ê—«ﬁ «·„«·Ì… €Ì— „ ”«ÊÏ „⁄ ≈Ã„«·Ï «·›« Ê—…...!!!"
    '    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '    Exit Sub
    'End If

    '---------------------------------
    If TxtNoteSerial.text = "" Then
        If Notes_coding(val(dcBranch.BoundText), XPDtbBill.value) = "error" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ":  Cmd(2).Enabled = True: Exit Sub
                Cmd(2).Enabled = True
            Else
                MsgBox "GE Exceed Coding ": Cmd(2).Enabled = True: Exit Sub
            
            End If

        Else
                       
            If Notes_coding(val(dcBranch.BoundText), XPDtbBill.value) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Cmd(2).Enabled = True: Exit Sub
                Else
                    MsgBox "Can't Create GE Manual No":  Cmd(2).Enabled = True: Exit Sub
                End If

            Else
                TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbBill.value)
            End If
        End If
    End If
 
    '           If TxtNoteSerial1.Text = "" Then
    '                           If Voucher_coding(val(my_branch), XPDtbBill.value, 10, 180, , 19, , val(DCboStoreName.BoundText)) = "error" Then
    '                               MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ’—› „Œ“‰Ì ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
    '                           Else
    '
    '                                       If Voucher_coding(val(my_branch), XPDtbBill.value, 10, 180, , 19, , val(DCboStoreName.BoundText)) = "" Then
    '                                           MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ”‰œ «·’—›  ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
    '                                       Else
    '                                           TxtNoteSerial1.Text = Voucher_coding(val(my_branch), XPDtbBill.value, 10, 180, , 19, , val(DCboStoreName.BoundText))
    '                                       End If
    '                           End If
    '           End If
           
    Dim NoteSerial1str As String

    If TxtNoteSerial1.text = "" Then
        NoteSerial1str = Voucher_coding(val(dcBranch.BoundText), XPDtbBill.value, 10, 180, , 19, , val(DCboStoreName.BoundText))
         
        If NoteSerial1str = "error" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ’—› ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ":  Cmd(2).Enabled = True: Exit Sub
            Else
                MsgBox " Voucher Code Exceed": Cmd(2).Enabled = True: Cmd(2).Enabled = True:   Exit Sub
            End If
                      
        Else
                                   
            If NoteSerial1str = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„  ”‰œ «·’—›  ÌœÊÌ« ﬂ„« Õœœ   ":  Cmd(2).Enabled = True: Exit Sub
                Else
                    MsgBox " Enter Voucher Number Manually": Cmd(2).Enabled = True: Exit Sub
                End If

            Else
                TxtNoteSerial1.text = NoteSerial1str
            End If
        End If
    
    End If

    '„‰ Â‰«
    Dim RsNotesGeneral As ADODB.Recordset
    Set RsNotesGeneral = New ADODB.Recordset
    '  RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
    RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
    If Me.TxtModFlg.text = "N" Then
        Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
    Else
        StrSQL = "Delete  marakes_taklefa_temp  where general_des=1 AND kedno =" & val(TXTNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
   
        StrSqlDel = "delete From Notes where NoteType=180 and  noteid=" & val(TXTNoteID.text)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        
        general_noteid = val(TXTNoteID.text)
    End If

    SngTemp = NewGrid.GetItemsTotal(ItemsGoodType) + val(TxtExtraValue.text) 'ﬁÌœ

    If SngTemp = 0 Then TxtNoteSerial.text = "":   GoTo novalue
    RsNotesGeneral.AddNew
    RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
    
    general_noteid = RsNotesGeneral("NoteID").value
    RsNotesGeneral.update
    TXTNoteID.text = general_noteid
    
    ' RsNotesGeneral("Transaction_ID").value = Val(XPTxtBillID.text)
    RsNotesGeneral("NoteDate").value = XPDtbBill.value
    RsNotesGeneral("NoteType").value = 180 ' «–‰ «÷«›…
    RsNotesGeneral("TradingContractID").value = val(txtTradingContractID)
    
    RsNotesGeneral("Note_Value").value = val(LblTotal.Caption)
    RsNotesGeneral("Note_ValueSales").value = val(lblTotalSalesPrice.Caption)
    RsNotesGeneral("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
    '  RsNotesGeneral("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.Text) = "", Null, Trim(Me.TxtNoteSerial1.Text))
    
    RsNotesGeneral("Remark").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
    '
    
    RsNotesGeneral("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text)
    '  Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
    RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
    RsNotesGeneral("numbering_type1").value = sand_numbering_type(10) '  «–‰ ’—›
    RsNotesGeneral("sanad_year").value = year(XPDtbBill.value)
    RsNotesGeneral("sanad_month").value = Month(XPDtbBill.value)
    'RsNotes("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
    RsNotesGeneral("branch_no").value = val(Me.dcBranch.BoundText)
    RsNotesGeneral.update
 
novalue:
        
    Set RSTransDetails = New ADODB.Recordset
    '   RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    Set RsNotes = New ADODB.Recordset
    'RsNotes.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    '********************
    StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
    RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    '******************************
    
    StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
    RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  
    If SystemOptions.SysRegisterState <> Registered And SystemOptions.SysRegisterState <> DevelopVersion Then
        If rs.RecordCount > 50 Then
            'Exit Sub
        End If
    End If

    CuurentLogdata
    Screen.MousePointer = vbArrowHourglass
    Cn.BeginTrans
    TransBegine = True

    If Me.TxtModFlg.text = "N" Then
        XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
        rs.AddNew
        rs("Transaction_ID").value = val(XPTxtBillID.text)
  
    ElseIf Me.TxtModFlg.text = "E" Then

        If mSaveDetails Then
            StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
        End If

        StrSqlDel = "delete From Notes where Transaction_ID=" & val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & val(Me.XPTxtBillID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
    End If
    If IsSaveWithOutMsg Then GoTo SaveDetailsOnly
    If Me.chkDone.value = vbUnchecked Then
        rs("chkDone").value = 0
    Else
        rs("chkDone").value = 1
    
    End If

    rs("project_id").value = val(Me.DcbProject.BoundText)
    rs("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
    rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
    rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
    rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '
    rs("OrderID").value = val(txtOrderID.text)
    rs("TicketNO").value = IIf(Trim(Me.TxtTicketNO.text) = "", Null, Trim(Me.TxtTicketNO.text))
    rs("Head_Details").value = val(Me.DcbType.ListIndex)
    rs("NoteId").value = val(TXTNoteID.text)
    rs("Transaction_Serial").value = IIf(Trim(Me.TxtTransSerial.text) = "", "", Trim(Me.TxtTransSerial.text))
    rs("Transaction_Date").value = XPDtbBill.value
    rs("ManualDate").value = txtManualDate.value
    rs("RegDate").value = txtRegDate.value
    '******************************
    '***********Ship*************************
    rs!ShipOrderNo = txtShipOrderNo
    rs!ShipEnquieryNo = txtShipEnquieryNo
    rs!ShipAccountNo = txtShipAccountNo
    rs!ShipCustomerName = txtShipCustomerName
    rs!ShipDistance = txtShipDistance
    rs!ShipArea = txtShipArea
    rs!ShipSiteNo = txtShipSiteNo
    rs!ShipProjectName = txtShipProjectName
    rs!ShipStructuralElement = txtShipStructuralElement
    rs!ShipMixDescription = txtShipMixDescription
    rs!ShipDriverName = txtShipDriverName
    rs!ShipPipeLine = txtShipPipeLine
    rs!ShipPump = txtShipPump
    rs!ShipTruckNo = txtShipTruckNo
    rs!ShipIceTemp = txtShipIceTemp
    rs!ShipTotalDeleveryd = val(txtShipTotalDeleveryd)
    rs!ShipThisLoad = txtShipThisLoad
    rs!ShipDayOrder = txtShipDayOrder
    rs!ShipTripNo = txtShipTripNo
    rs!ShipPlantNo = txtShipPlantNo
    rs!ShipBatched = txtShipBatched

    rs!ShipRestunedPlant = txtShipRestunedPlant
    rs!ShipEndDischarge = txtShipEndDischarge
    rs!ShipStartDisCharge = txtShipStartDisCharge
    rs!ShipOnSite = txtShipOnSite
    '************************************
    '******************************
    
    rs("Transaction_Type").value = 19
    rs("UserID").value = user_id
    rs("OldOpOrderID").value = val(TxtOldOpOrderID.text)

    If val(CBoBasedON.ListIndex) = 8 Then
        rs("OpOrderID").value = val(TXT_order_no.text)
    Else
        rs("OpOrderID").value = Null
    End If

    '  DB_CreateField "Transactions", "RepairOrder", adInteger, adColNullable, , , " ???    ", False, True
    If val(CBoBasedON.ListIndex) = 9 Then
        rs("RepairOrder").value = val(TXT_order_no.text)
    Else

        rs("RepairOrder").value = Null
    End If
   
    rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
    rs("CarId").value = IIf(Me.DCCar.BoundText = "", Null, (Me.DCCar.BoundText))
    rs("DriverId").value = IIf(Me.DCDriver.BoundText = "", Null, (Me.DCDriver.BoundText))
    rs("ManualNo").value = IIf(txtManualNO.text = "", Null, val(txtManualNO.text))
    rs("TradingContractID").value = IIf(txtTradingContractID.text = "", Null, val(txtTradingContractID.text))

    rs("Nots").value = Me.Text2.text
    rs("nots2").value = Txtnots2.text
    rs("WorkOrderNO").value = val(TxtWorkOrderNO.text)
    rs("Doctype").value = IIf(Me.DCDocTypes.BoundText = "", Null, val(DCDocTypes.BoundText))
 
    rs("ExtraAccount").value = IIf(DCExtraAccount.BoundText = "", Null, (DCExtraAccount.BoundText))

    If DCExtraAccount.BoundText = "" Then
        rs("ExtraValue").value = 0
        TxtExtraValue.text = 0
    Else
        rs("ExtraValue").value = val(TxtExtraValue.text)
    End If

    Dim rs2 As New ADODB.Recordset
    '           rs2.Close
    rs2.Open "select * from Transactions where Transaction_Serial = '" & TxtTransSerial.text & " 'and Transaction_type = 21", Cn, adOpenDynamic, adLockOptimistic

    If Not rs2.EOF Then
        rs2("Nots2").value = Me.Text2.text & ""
        rs2.update
        rs2.Close
    End If

    If XPCboDiscountType.ListIndex = -1 Then
        rs("Trans_DiscountType").value = 0
    Else
        rs("Trans_DiscountType").value = val(XPCboDiscountType.ListIndex)
    End If

    rs("Emp_ID").value = IIf(DcboEmpName.BoundText = "", Null, DcboEmpName.BoundText)
    rs("Trans_Discount").value = IIf(XPTxtDiscountVal.text = "", Null, val(XPTxtDiscountVal.text))
    rs("CusID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
    rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, val(DCboStoreName.BoundText))
    rs("StoreID1").value = IIf(DCboStoreName2.BoundText = "", Null, val(DCboStoreName2.BoundText))
     
    If CboPayMentType.ListIndex = -1 Then
        rs("PaymentType").value = 0
    Else
        rs("PaymentType").value = val(CboPayMentType.ListIndex)
    End If

    'rs("TaxFound").value = IIf(XPChkTAX.value = Checked, True, False)
    'rs("TaxValue").value = IIf(XPTxtTaxValue.text = "", Null, val(XPTxtTaxValue.text))
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
    'If ChkTaxStamp.value = vbChecked And val(Me.TxtTaxStampValue.text) > 0 Then
    '    rs("TaxStampValue").value = val(Me.TxtTaxStampValue.text)
    'Else
    '    rs("TaxStampValue").value = 0
    'End If

    '÷—»Ì… Œœ„…
    'If ChkTaxSerivce.value = vbChecked And val(Me.TxtTaxServiceValue.text) > 0 Then
    '    rs("TaxServiceValue").value = val(Me.TxtTaxServiceValue.text)
    'Else
    '    rs("TaxServiceValue").value = 0
    'End If

    rs("DepartementID").value = IIf(DcboEmpDepartments.BoundText = "", Null, val(DcboEmpDepartments.BoundText))

    rs("FixesAssetsID").value = IIf(DCEquipments.BoundText = "", Null, val(DCEquipments.BoundText))

    rs("Emp_ID").value = IIf(DcboEmpName.BoundText = "", Null, val(DcboEmpName.BoundText))

    rs("BillBasedOn").value = val(CBoBasedON.ListIndex)

    rs("OPrType").value = val(DCOPrType.ListIndex)
    rs("order_no").value = TXT_order_no.text
   
    rs.update
SaveDetailsOnly:
    Dim rsCC           As ADODB.Recordset
    Dim mItemLimit     As Double
    Dim mItemLimitType As Double
    Dim mPeriodT1      As Double

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
                        Msg = "«·”Ì—Ì«· «·Œ«’ »«·’‰›" & Chr(13)
                        Msg = Msg + FG.cell(flexcpTextDisplay, RowNum, FG.ColIndex("name")) & Chr(13)
                        Msg = Msg + " „ √œŒ«·Â ·ﬁÿ⁄… √Œ—Ï ›Ì Â–Â «·›« Ê—…"
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        RsTemp.Close
                        XPTab301.CurrTab = 0
                        FG.Row = RowNum
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
                        Msg = " ÌÃ»  ÕœÌœ ÊÕœ… «·ﬂ„Ì… «·Œ«’… »«·’‰›" & Chr(13)
                        Msg = Msg + FG.cell(flexcpTextDisplay, RowNum, FG.ColIndex("Name")) & Chr(13)
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        XPTab301.CurrTab = 0
                        FG.Row = RowNum
                        FG.Col = FG.ColIndex("UnitID")
                        FG.ShowCell RowNum, FG.ColIndex("UnitID")
                        FG.SetFocus
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    End If
                End If
            
                s = " SELECT (IsNull(TblItems.ItemLimit,0)) as ItemLimit,TblItems.ItemLimitType,TblItems.PeriodT1,TblItems.ItemLimit "
                s = s & " From TblItems "
                s = s & " Where         "
                s = s & " IsNull(ItemLimit,0) > 0 AND TblItems.ItemID = " & val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
            
                Set rsCC = New ADODB.Recordset
                rsCC.Open s, Cn, adOpenKeyset, adLockReadOnly

                If Not rsCC.EOF Then
                    mItemLimit = val(rsCC!ItemLimit & "")
                    mItemLimitType = val(rsCC!ItemLimitType & "")
                    mPeriodT1 = val(rsCC!PeriodT1 & "")

                    If mItemLimitType = 1 Then
                        mPeriodT1 = mPeriodT1 * 30
                    ElseIf mItemLimitType = 2 Then
                        mPeriodT1 = mPeriodT1 * 360
                    End If

                    mPeriodT1 = mPeriodT1 * -1
                End If

                rsCC.Close
                s = " SELECT SUM(ShowQty )    as CC FROM Transaction_Details td INNER JOIN"
                s = s & " Transactions t ON T.Transaction_ID = td.Transaction_ID"
                s = s & " INNER JOIN TblItems ON td.Item_ID = TblItems.ItemID"
                s = s & " Where t.Transaction_Date BETWEEN   DATEAdd(D," & mPeriodT1 & " ," & SQLDate(XPDtbBill.value, True) & ")   And " & SQLDate(XPDtbBill.value, True) & ""
                'ay(t.Transaction_Date) = " & day(XPDtbBill.value) & "  And Month(t.Transaction_Date) = " & Month(XPDtbBill.value) & " And year(t.Transaction_Date) = " & year(XPDtbBill.value) & ""
            
                s = s & " AND ItemID = " & val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                s = s & " AND (IsNull(td.EqupID,0) = " & val(DCEquipments.BoundText) & " Or T.FixesAssetsID = " & val(DCEquipments.BoundText) & " )"
                's = s & " Group By TblItems.ItemLimitType,TblItems.PeriodT1,TblItems.ItemLimit"
                Set rsCC = New ADODB.Recordset
            
                rsCC.Open s, Cn, adOpenKeyset, adLockReadOnly
            
                '  rsCC.Close
                '  rsCC.Open s, Cn, adOpenForwardOnly, adLockReadOnly
                If (val(rsCC!CC & "") + val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))) > mItemLimit And mItemLimit > 0) And IsSaveWithOutMsg = False Then
                    Msg = " ÌÃ» «‰ ·« Ì ŒÿÏ «·ﬂ„Ì… «·„’—Ê›… Õœ «·ÿ·» ›Ï ﬂ«— … «·’‰›" & Chr(13)
                    Msg = Msg + FG.cell(flexcpTextDisplay, RowNum, FG.ColIndex("Name")) & Chr(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    FG.SetFocus
                    Screen.MousePointer = vbDefault
                    GoTo ErrTrap

                End If

                mPeriodT1 = 0
                RSTransDetails.AddNew
                RSTransDetails("order_no").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("order_no")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("order_no")))
                RSTransDetails("OrderArrivalDate").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")))
                RSTransDetails("FoxyNo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")))
                RSTransDetails("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
                RSTransDetails("Transaction_ID").value = val(XPTxtBillID.text)
                RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
                RSTransDetails("GroupIDMint").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("GroupIDMint")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("GroupIDMint"))))
                RSTransDetails("MintID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("MintID")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("MintID"))))

                If CBoBasedON.ListIndex = 8 Then
                    RSTransDetails("OrderNo").value = val(TXT_order_no.text)
                    RSTransDetails("EqupID").value = val(DCEquipments.BoundText)
                    RSTransDetails("Head_Details").value = val(DcbType.ListIndex)
                End If

                If CBoBasedON.ListIndex = 9 Then
                    RSTransDetails("RepairOrder").value = val(TXT_order_no.text)
                    RSTransDetails("EqupID").value = val(DCEquipments.BoundText)
                    RSTransDetails("Head_Details").value = val(DcbType.ListIndex)
                End If

                'RSTransDetails("Quantity").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
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
                RSTransDetails("LineId").value = RowNum
                RSTransDetails("QtyFaqtors").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("QtyFaqtors")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("QtyFaqtors"))))
                RSTransDetails("ItemDiscountType").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType"))))
                RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
                RSTransDetails("ItemDiscount").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal"))))
                RSTransDetails("guaranteeTime").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime"))))
                RSTransDetails("CostPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice"))))
                RSTransDetails("CostTransID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("PofTransID")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("PofTransID"))))
                RSTransDetails("ItemProfit").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemProfit")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemProfit"))))
            If Me.FG.ColIndex("EmpID4") <> -1 Then
                RSTransDetails("EmpID4").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("EmpID4")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("EmpID4"))))
        End If
                RSTransDetails("OldID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("OldID")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("OldID"))))
             
                RSTransDetails("length").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("length")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("length"))))
                RSTransDetails("Width").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Width")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Width"))))
                RSTransDetails("OUTR").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("OUTR")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("OUTR"))))
                RSTransDetails("INR").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("INR")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("INR"))))
             
                RSTransDetails("NoCount").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("NoCount")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("NoCount"))))
             
                RSTransDetails("Height").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Height")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Height"))))
        
                RSTransDetails("UnitID").value = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
                RSTransDetails("ShowQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
                Dim cnt As Double
                cnt = FG.TextMatrix(RowNum, FG.ColIndex("Count"))

                RSTransDetails("Remarks").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Remarks")) = ""), Null, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("Remarks"))))
                
                RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
                RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
                RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
            
                RSTransDetails("showPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
                RSTransDetails("OperPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
                RSTransDetails("MixNo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("MixNo")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("MixNo"))))
                RSTransDetails("QtyFaqtors").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("QtyFaqtors")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("QtyFaqtors"))))
                RSTransDetails("MaxQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("MaxQty")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("MaxQty"))))
                RSTransDetails("MaxUnitID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("MaxUnitID")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("MaxUnitID"))))
                RSTransDetails("ItemsDetailsNewidea").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("ItemsDetailsNewidea")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("ItemsDetailsNewidea")))

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

                If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                    RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value

                    RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
              
                End If
                  If IsSaveWithOutMsg Then
                        'RSTransDetails("Price").value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
                        
                        FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(LngCurItemID, 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(XPTxtBillID.text), LngUnitID, val(Me.DCboStoreName.BoundText))
                        FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = FG.TextMatrix(RowNum, FG.ColIndex("Price"))
               Else
                 '       RSTransDetails("Price").value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
               End If
               
               FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(LngCurItemID, 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(XPTxtBillID.text), LngUnitID, val(Me.DCboStoreName.BoundText))
                        FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = FG.TextMatrix(RowNum, FG.ColIndex("Price"))
                      RSTransDetails("Price").value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
                        
                SngTemp = SngTemp + (val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice"))) * RSTransDetails("quantity").value)

                RSTransDetails("ProductionDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate"))), "DD/mm/YYYY"))
                RSTransDetails("ExpiryDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate"))), "DD/mm/YYYY"))
                RSTransDetails("LotNO").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("LotNO")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("LotNO")))
                Dim OldQty  As Double
                Dim OldCost As Double
                Dim NewQty  As Double
                Dim NewCost As Double
               
                getItemCostData XPDtbBill.value, RSTransDetails("Item_ID").value, val(DCboStoreName.BoundText), val(Me.XPTxtBillID.text), OldQty, OldCost, NewQty, NewCost, , LngUnitID
                RSTransDetails("OldQty").value = NewQty
                RSTransDetails("OldCost").value = NewCost
       
                RSTransDetails("NewQty").value = RSTransDetails("OldQty").value - RSTransDetails("Quantity").value
                RSTransDetails("NewCost").value = RSTransDetails("OldCost").value ' ((RSTransDetails("OldQty").value * RSTransDetails("OldCost").value) + (RSTransDetails("Quantity").value * RSTransDetails("Price").value)) / (RSTransDetails("Quantity").value + RSTransDetails("OldQty").value)
            If Me.FG.ColIndex("TotalInvoiceQty") <> -1 Then
                RSTransDetails("TotalInvoiceQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("TotalInvoiceQty")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("TotalInvoiceQty"))))
            End If
             If Me.FG.ColIndex("ISSUEDQTY") <> -1 Then
                RSTransDetails("ISSUEDQTY").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ISSUEDQTY")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ISSUEDQTY"))))
            End If
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

s = "Select IsNotCreateEntry from TblStore where isnull(IsNotCreateEntry,0) = 1 and StoreId = " & val(DCboStoreName.BoundText)

Dim rsDummy As New ADODB.Recordset
rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly

If Not rsDummy.EOF Then
        GoTo ⁄ —Ì”
End If


rsDummy.Close


    Dim LngDevID           As Long
    Dim LngDevNO           As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes         As String
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '----------------
    Dim Account_Code_dynamic As String
    'SngTemp = NewGrid.GetItemsCostTotal * RSTransDetails("quantity").value / Cnt
    SngTemp = NewGrid.GetItemsTotal(ItemsGoodType) + val(TxtExtraValue.text) 'ﬁÌœ

    If SngTemp > 0 Then
        '1 work with branch
        '2 work with inventory
        '3 work with groups
        OtherInformation.NextAccount_Code = get_store_Account(DCboStoreName.BoundText, "Account_Code")

        If detect_inventory_work_type = 1 Then
            Account_Code_dynamic = get_account_code_branch(1, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ﬂ·›… «·„»Ì⁄«  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    GoTo ErrTrap
         
                End If
            End If

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

            DebitAccount = StrTempAccountCode
    
            'StrTempAccountCode = "a3a2" ' ﬂ·›… «·„»Ì⁄« 
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "  √–‰ ’—›  —ﬁ„     " & Me.TxtNoteSerial1.text & "  " & TxtBillComment & "  ··„⁄œ… " & DCEquipments.text
            Else
                StrTempDes = "Issue Voucher No.  " & Me.TxtNoteSerial1.text & "  " & TxtBillComment & "  ··„⁄œ… " & DCEquipments.text
            End If

            Line1 = setfoxy_Line
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , Line1, , , , , , val(Me.DCEquipments.BoundText), , , val(Me.dcBranch.BoundText), , , , , , , val(DcboEmpDepartments.BoundText), val(DcboEmpName.BoundText), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
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
        
            If val(DCDocTypes.BoundText) > 0 Then
                getDocAccounts val(DCDocTypes.BoundText), , StrTempAccountCode, , , , , usedaccount

                If StrTempAccountCode = "" And usedaccount = 1 Then
                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·œ«∆‰ ··”‰œ", vbCritical
                    GoTo ErrTrap
                ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
        
                ElseIf usedaccount = 0 Then
        
                    StrTempAccountCode = Account_Code_dynamic '«·„Œ“Ê‰ 0 ›Ì «·›—⁄
                End If

            Else
                StrTempAccountCode = Account_Code_dynamic '«·„Œ“Ê‰ 0 ›Ì «·›—⁄
            End If

            CreditAccount = StrTempAccountCode
    
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "√–‰ ’—›  —ﬁ„ " & Me.TxtNoteSerial1.text & "  " & TxtBillComment & "  ··„⁄œ… " & DCEquipments.text
            Else
                StrTempDes = "Issue Voucher No. " & Me.TxtNoteSerial1.text & "  " & TxtBillComment & "  ··„⁄œ… " & DCEquipments.text
            End If
    
            LngDevNO = LngDevNO + 1
            Line2 = setfoxy_Line

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , Line2, , , , , , val(Me.DCEquipments.BoundText), , , val(Me.dcBranch.BoundText), , , , , , , val(DcboEmpDepartments.BoundText), val(DcboEmpName.BoundText), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If
    
        ElseIf detect_inventory_work_type = 2 Then
            Account_Code_dynamic = get_account_code_branch(1, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ﬂ·›… «·„»Ì⁄«  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    GoTo ErrTrap
         
                End If
            End If

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

            DebitAccount = StrTempAccountCode
            
            Line1 = setfoxy_Line

            'StrTempAccountCode = "a3a2" ' ﬂ·›… «·„»Ì⁄« 
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "√–‰ ’—›  —ﬁ„ " & Me.TxtNoteSerial1.text
            Else
                StrTempDes = "Issue Voucher No. " & Me.TxtNoteSerial1.text
            End If
    
            LngDevNO = LngDevNO + 1
            Dim project_id As Integer
            project_id = IIf(Me.DcbProject.BoundText = "", 0, Me.DcbProject.BoundText)

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , Line1, , , , , , val(Me.DCEquipments.BoundText), , , val(Me.dcBranch.BoundText), , , , , , , val(DcboEmpDepartments.BoundText), val(DcboEmpName.BoundText), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If

            '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
            SngTemp = NewGrid.GetItemsTotal(ItemsGoodType)

            If val(DCDocTypes.BoundText) > 0 Then
                getDocAccounts val(DCDocTypes.BoundText), , StrTempAccountCode, , , , , usedaccount

                If StrTempAccountCode = "" And usedaccount = 1 Then
                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·œ«∆‰ ··”‰œ", vbCritical
                    GoTo ErrTrap
                ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
                    Account_Code_dynamic = StrTempAccountCode
                ElseIf usedaccount = 0 Then
        
                    Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")
                End If

            Else
                Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")
            End If
        
            If Account_Code_dynamic = "" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                GoTo ErrTrap
            End If
    
            StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰
            CreditAccount = StrTempAccountCode
            OtherInformation.NextAccount_Code = DebitAccount

            ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "√–‰ ’—›  —ﬁ„ " & Me.TxtNoteSerial1.text & "  " & TxtBillComment
            Else
                StrTempDes = "Issue Voucher No. " & Me.TxtNoteSerial1.text & "  " & TxtBillComment
            End If

            Line2 = setfoxy_Line
         
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , Line2, , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , val(DcboEmpDepartments.BoundText), val(DcboEmpName.BoundText), , , , , , val(Me.DcbProject.BoundText), , , , , , , , , , , , , OtherInformation) = False Then
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

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , val(Me.DCEquipments.BoundText), , , val(Me.dcBranch.BoundText), , , , , , , val(DcboEmpDepartments.BoundText), val(DcboEmpName.BoundText)) = False Then
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
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "√–‰ ’—›  —ﬁ„ " & Me.TxtNoteSerial1.text & "  " & TxtBillComment
                        Else
                            StrTempDes = "Issue Voucher No. " & Me.TxtNoteSerial1.text & "  " & TxtBillComment
                        End If

                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , val(Me.DCEquipments.BoundText), , , val(Me.dcBranch.BoundText), , , , , , , val(DcboEmpDepartments.BoundText), val(DcboEmpName.BoundText), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
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

    'If Me.ChkTaxStamp.value = vbChecked Then
    '   StrTempAccountCode = "a3a9" 'œ„€«  ÕﬂÊ„Ì…
    '   StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtTransSerial.text
    '   SngTemp = Val(Me.lbl(53).Caption)
    '   LngDevNO = LngDevNO + 1
    '   If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
    '       0, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
    '       GoTo ErrTrap
    '   End If
    'End If

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
    'If XPChkTAX.value = vbChecked Then
    'StrTempAccountCode = "a1a3a5" '÷—»Ì… „»Ì⁄«  „œÌ‰…
    'SngTemp = Val(Me.lbl(51).Caption)
    'StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtTransSerial.text
    'LngDevNO = LngDevNO + 1
    'If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
    '    1, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
    '    GoTo ErrTrap
    'End If
    'End If

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

    'Õ”«» «·«÷«›« 
    If DCExtraAccount.BoundText <> "" And val(TxtExtraValue.text) > 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrTempDes = "√–‰ ’—›  —ﬁ„ " & Me.TxtNoteSerial1.text & "  " & TxtBillComment
        Else
            StrTempDes = "Issue Voucher No. " & Me.TxtNoteSerial1.text & "  " & TxtBillComment
        End If

        LngDevNO = LngDevNO + 1

        If ModAccounts.AddNewDev(LngDevID, LngDevNO, DCExtraAccount.BoundText, val(TxtExtraValue.text), 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , val(Me.DCEquipments.BoundText), , , val(Me.dcBranch.BoundText), , , , , , , val(DcboEmpDepartments.BoundText), val(DcboEmpName.BoundText), , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
            GoTo ErrTrap
        End If
    End If

    '     If Me.DcCostCenter.BoundText <> "" Then
    save_General_cost_center Me.DcCostCenter.BoundText, Me.DcCostCenter.text, "”‰œ ’—› „Ê«œ Œ«„ ", Me.XPDtbBill.value
    
    
⁄ —Ì”:

    SaveItemsData
    save_cost_center
    Cn.CommitTrans
    TransBegine = False
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount

    If IsSaveWithOutMsg Then Exit Sub

    Select Case Me.TxtModFlg.text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
                Msg = " Data Was Saved do you want Another Entry" & Chr(13)
    
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
                '  Msg = " changes Was Saved " & Chr(13)
                MsgBox "changes Was Saved ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            End If
    
            lbl(56).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
    
            TxtModFlg.text = "R"
    End Select

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
 Screen.MousePointer = vbDefault
    If TransBegine = True Then
        TransBegine = False
        Cn.RollbackTrans
    End If
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    Msg = Msg & Chr(13) & Err.Description
    Msg = Msg & Chr(13) & Err.Number
    Msg = Msg & Chr(13) & Err.Source
    Msg = Msg & Chr(13) & Err.LastDllError
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    'Resume
    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
    End If

    If Not RsNotes Is Nothing Then
        If RsNotes.EOF Then RsNotes.CancelUpdate
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
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Msg = Msg & Chr(13) & Err.Description
        Msg = Msg & Chr(13) & Err.Number
        Msg = Msg & Chr(13) & Err.Source
        Msg = Msg & Chr(13) & Err.LastDllError
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If


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

Function save_cost_center()
'Exit Function
    'on error resume next
    If Not IsNumeric(TXTNoteID.text) Then Exit Function
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql_str As String

    'Rs.Open "", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    sql_str = "select * from marakes_taklefa_temp where kedno=" & val(TXTNoteID.text)
    rs.Open sql_str, Cn, adOpenStatic, adLockOptimistic, adCmdText

    For i = 1 To rs.RecordCount
        rs("ok").value = 1
        rs("NoteDate").value = XPDtbBill.value
        rs("NoteSerial").value = TxtNoteSerial.text
        rs("Remark").value = "”‰œ ’—› „Ê«œ Œ«„ —ﬁ„  " & TxtNoteSerial1 & "    " & TxtBillComment.text
 
        rs.update
        rs.MoveNext
    Next i

End Function

Public Function save_General_cost_center(cost_center_id As String, _
                                         cost_center, _
                                         opr_type As String, _
                                         record_date As Date) 'As String, value As Double, depit_or_credit As Boolean, opr_id As Double, opr_type As String, account_no As String, account_name As String, line_no As Double, recorddate As Date)
    Dim i As Integer
    Dim rs As New ADODB.Recordset
 
    Dim StrSQL As String

    StrSQL = "Delete  marakes_taklefa_temp  where general_des=1 AND kedno =" & val(TXTNoteID.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
 
    If Me.DcCostCenter.BoundText = "" Then
        Exit Function
    End If
        
    rs.Open "marakes_taklefa_temp", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    '??? ?I??
    rs.AddNew
    rs("general_des").value = 1
    rs("cost_center_id").value = cost_center_id
    rs("cost_center").value = cost_center
    rs("value").value = LblTotal.Caption
    rs("depit_or_credit").value = "„œÌ‰"
    rs("opr_id").value = general_noteid
    rs("kedno").value = general_noteid
        
    rs("opr_type").value = opr_type
    rs("account_name").value = Get_Account_Name(, DebitAccount)
    rs("account_no").value = DebitAccount
    rs("line_no").value = Line1
    rs("record_date").value = record_date
    rs.update
    Exit Function
        
    rs.AddNew
    rs("general_des").value = 1
    rs("cost_center_id").value = cost_center_id
    rs("cost_center").value = cost_center
    rs("value").value = LblTotal.Caption
    rs("depit_or_credit").value = "œ«∆‰"
    rs("opr_id").value = general_noteid
    rs("kedno").value = general_noteid
        
    rs("opr_type").value = opr_type
    rs("account_name").value = Get_Account_Name(, CreditAccount)
    rs("account_no").value = CreditAccount
    rs("line_no").value = Line2
    rs("record_date").value = record_date
    rs.update
    '??? IC??
    '    rs.AddNew
    '    rs("cost_center_id").value = cost_center_id
    '    rs("cost_center").value = cost_center
    '    rs("value").value = XPTxtVal.text
    '    rs("depit_or_credit").value = "IC??"
    '    rs("opr_id").value = Me.Text1.text
    '    rs("kedno").value = Me.Text1.text
    '
    '    rs("opr_type").value = opr_type
    '    rs("account_name").value = DcboCreditSide.text
    '    rs("account_no").value = DcboCreditSide.BoundText
    '    rs("line_no").value = Line2
    '    rs("record_date").value = record_date
    '    rs.update
 
    rs.Close
End Function

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
'
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
 

Private Sub XPDtbBill_Change()
   TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
    
    If Voucher_coding(val(dcBranch.BoundText), XPDtbBill.value, 10, 180, , 19) = "" Then Exit Sub
    If Trim(TxtNoteSerial1.text) <> "" Then
        oldtxtNoteSerial1.text = TxtNoteSerial1.text
    End If

 
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

    Dim BuyReport As ClsBuyReport
    On Error GoTo ErrTrap
    Dim ShowType As Boolean
    ShowType = GetSetting(StrAppRegPath, "View_Type", "ReportType", True)

    If ShowType = True Then
        If Not XPTxtBillID.text Then
            Set BuyReport = New ClsBuyReport
            BuyReport.ShowIssueVoucherData XPTxtBillID.text, , CBoBasedON.text, Me.DCboStoreName2.text, Me.DCCar.text, Me.DCDriver.text
        End If

    Else

        If Not XPTxtBillID.text Then
            Set BuyReport = New ClsBuyReport
            BuyReport.ShowIssueVoucherData XPTxtBillID.text, True, CBoBasedON.text
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
    If Trim(Me.TxtModFlg.text) = "" Then Exit Sub
    If Me.TxtModFlg.text <> "R" Then

        Select Case Me.TxtModFlg.text

            Case "N"
    
                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save " & Chr(13)
                    StrMSG = StrMSG & " the new data  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
        
                End If
        
            Case "E"

                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save  " & Chr(13)
                    StrMSG = StrMSG & " the Modifications  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
                
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

    On Error GoTo ErrTrap
    Dim fullcode As String
 
    GetCustomersDetail val(DBCboClientName.BoundText), , fullcode, 1
    TxtSearchCode.text = fullcode


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
    Msg = Err.Description & Chr(13) & ""
    Msg = Msg & Err.Source & Chr(13) & ""
    Msg = Msg & Me.Name & " DBCboClientName_Change:" & Chr(13) & ""
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub DBCboClientName_Click(Area As Integer)
    DBCboClientName_Change
End Sub

Private Sub XPTxtValue_Change(Index As Integer)
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
    If FG.TextMatrix(FG.Row, FG.ColIndex("Code")) <> "" And FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) <> "" Then
        StrSQL = "select * From ReplacedItems where ReturnID=" & XPTxtBillID.text
        StrSQL = StrSQL + " and ItemID=" & FG.TextMatrix(FG.Row, FG.ColIndex("Code"))
        StrSQL = StrSQL + " and ItemSerial='" & FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) & "'"
        Set RsReplace = New ADODB.Recordset
        RsReplace.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsReplace.EOF Or RsReplace.BOF) Then
            Msg = "·ﬁœ  „ «” »œ«· «·ﬁÿ⁄… : " & FG.cell(flexcpTextDisplay, FG.Row, FG.ColIndex("Name")) & Chr(13)
            Msg = Msg + "–«  «·”Ì—Ì«· : " & FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) & Chr(13)
            Msg = Msg + " »«·ﬁÿ⁄… –«  «·”Ì—Ì«· : " & RsReplace("newSerial").value & Chr(13)
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
txtManualDate.value = Date
txtRegDate.value = Date
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
Private Sub LoadCar()
Set Dcombos = New ClsDataCombos
Dim sql As String
  If CBoBasedON.ListIndex = 8 Then
  If val(Me.DcbType.ListIndex) = 1 Then
  If SystemOptions.UserInterface = ArabicInterface Then
  sql = " SELECT     dbo.tblordermaintenancetypes.PartID, dbo.FixedAssets.Name"
  Else
  sql = " SELECT     dbo.tblordermaintenancetypes.PartID, dbo.FixedAssets.NameE"
  End If
  sql = sql & " FROM         dbo.tblordermaintenancetypes LEFT OUTER JOIN"
  sql = sql & "                    dbo.FixedAssets ON dbo.tblordermaintenancetypes.PartID = dbo.FixedAssets.id"
  sql = sql & "  Where (dbo.tblordermaintenancetypes.OrderID = " & val(TXT_order_no.text) & ") And (dbo.tblordermaintenancetypes.TypeTrans = 2)"
  Dcombos.ClearMyDataCombo DCEquipments
  fill_combo DCEquipments, sql
  DoEvents
  DoEvents
  GetOrderMaintdet
    Else
     Dcombos.GetEquipments DCEquipments
     GetOrderMaint
    End If
   Else
   Dcombos.GetEquipments DCEquipments
   End If
End Sub
Private Sub LoadCombosData()
     Dcombos.GetEmpDepartments Me.DcboEmpDepartments
    Dcombos.GetEquipments DCEquipments
    Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetProjects Me.DcbProject
    Dcombos.GetEmployees Me.DcboEmp
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName
    Dcombos.GetStores Me.DCboStoreName
    Dcombos.GetStores Me.DCboStoreName2
      Dcombos.GetEmployees Me.DCDriver, , True
  Dcombos.GetCars Me.DCCar
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetDocTypebyid Me.DCDocTypes, 19, val(Me.dcBranch.BoundText)

    Set Dcombos = New ClsDataCombos

    If SystemOptions.UserInterface = ArabicInterface Then
        Dcombos.GetAccountingCodes DCExtraAccount, True
    Else
 
        Dcombos.GetAccountingCodesENg DCExtraAccount, True

    End If

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
lbl(45).Caption = "To Store"

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    lbl(67).Caption = "Type"
    Me.LblShortcutKeys.Caption = "(New F12 OR Enter) ,(Edit F11),(Save F10),(Undo F9),(Delete F8),(Search F7)"
    Me.Caption = "Issue Voucher"
     lbl(59).Caption = "Bseed On"
     sameCmd.Caption = "Same Copy"
lbl(65).Caption = "To Emp."
lbl(48).Caption = "Addition Account"
lbl(51).Caption = "GE Data"
lbl(66).Caption = "Project"

Command10.Caption = "Show"
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
lbl(4).Caption = "Car"
lbl(41).Caption = "Driver"

Label6.Caption = "Ge No."
lbl(58).Caption = "Acc."

    'Label4.Caption = "Doc Type"
'    Frame3.Caption = "GE Data"
    Cmd(10).Caption = "Print GE"
    'Frame1.Caption = "Account additions"
    lbl(57).Caption = "Value"
lbl(53).Caption = "Manual"
    Label5.Caption = "Notes"
    Label8.Caption = "CC"

    lbl(8).Caption = "Discount Value"
    lbl(22).Caption = "Profit Value"
    lbl(23).Caption = "Profit Perce"
    
    lbl(39).Caption = "Based On "
    lbl(60).Caption = "Type"
    lbl(62).Caption = "Equip"
    lbl(61).Caption = "Dept. "
    lbl(64).Caption = "Emp."
    
     
    


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
    lbl(32).Caption = "Recipt Number Maint."

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
    
    Me.XPTab301.TabCaption(3) = "Attachments"
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
       With FG
        .TextMatrix(0, .ColIndex("MintName")) = "Maintenance"
        .TextMatrix(0, .ColIndex("GroupMint")) = "Group"
    End With
    
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
    
    GroupBox1.Visible = False
    'Me.XPTab301.TabVisible(1) = False
    'Me.XPTab301.TabVisible(3) = False
    Me.XPTab301.TabVisible(0) = True
    'Me.XPTab301.TabVisible(2) = False
    'Me.XPTab301.TabVisible(4) = False
    Me.XPTab301.CurrTab = 0
End Sub

Private Sub XPTxtValue_KeyPress(Index As Integer, _
                                KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtValue(Index).text, 0)
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

Private Sub XPTxtValue_MouseMove(Index As Integer, _
                                 Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)

    If val(Me.XPTxtValue(Index).text) <> 0 Then
        Me.XPTxtValue(Index).ToolTipText = WriteNo(Me.XPTxtValue(Index).text, 1, True)
    Else
        Me.XPTxtValue(Index).ToolTipText = ""
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


Function fillapprovData()
Dim Num As Integer
 Dim RsDetails As New ADODB.Recordset
 Dim StrSQL As String
 
 
 StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
StrSQL = StrSQL + " FROM         dbo.ApprovalData left JOIN"
StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtBillID.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

 If Not (RsDetails.EOF Or RsDetails.BOF) Then
        GRID2.rows = RsDetails.RecordCount + 1
 

        For Num = 1 To RsDetails.RecordCount
        
       GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
    If GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = "1" Then
   GRID2.cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
   Else
    GRID2.cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
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

