VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{CFC0A331-9521-11D5-B9E6-5A06F6000000}#1.0#0"; "VDSCombo.DLL"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Begin VB.Form frmsalebillCompose 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "›« Ê—… „»Ì⁄«  „Ã„⁄Â"
   ClientHeight    =   9165
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   17325
   HelpContextID   =   160
   Icon            =   "FrmSaleBillCompose.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9165
   ScaleWidth      =   17325
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   9165
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   17325
      _cx             =   30559
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
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -1470
         TabIndex        =   296
         Top             =   615
         Visible         =   0   'False
         Width           =   1170
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1335
         Index           =   0
         Left            =   0
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   660
         Width           =   17295
         _cx             =   30506
         _cy             =   2355
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
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   334
            Top             =   0
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.TextBox TxtVATNO 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   9840
            MaxLength       =   55
            RightToLeft     =   -1  'True
            TabIndex        =   330
            Top             =   855
            Width           =   1215
         End
         Begin VB.TextBox TxtValueComm 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   314
            Top             =   480
            Width           =   1050
         End
         Begin VB.ComboBox DcbTypComm 
            Height          =   315
            Left            =   1890
            TabIndex        =   312
            Top             =   480
            Width           =   1230
         End
         Begin VB.TextBox TxtGratuity 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   180
            TabIndex        =   310
            Top             =   855
            Width           =   1050
         End
         Begin VB.TextBox TxtEmbarNo 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   4260
            TabIndex        =   309
            Top             =   480
            Width           =   1515
         End
         Begin VB.TextBox TxtCommission 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   180
            TabIndex        =   307
            Top             =   480
            Width           =   1080
         End
         Begin VB.TextBox TxtDriverName 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1890
            TabIndex        =   305
            Top             =   855
            Width           =   3885
         End
         Begin VB.TextBox TxtBoardNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6990
            TabIndex        =   303
            Top             =   855
            Width           =   1665
         End
         Begin VB.TextBox TxtSuplCode3 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   15120
            TabIndex        =   298
            Top             =   855
            Width           =   1035
         End
         Begin VB.Frame Frame5 
            Height          =   1110
            Left            =   -2985
            RightToLeft     =   -1  'True
            TabIndex        =   249
            Top             =   -120
            Visible         =   0   'False
            Width           =   3000
            Begin MSDataListLib.DataCombo DCCar 
               Height          =   315
               Left            =   120
               TabIndex        =   250
               Top             =   240
               Width           =   2475
               _ExtentX        =   4366
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCDriver 
               Height          =   315
               Left            =   120
               TabIndex        =   251
               Top             =   600
               Width           =   2475
               _ExtentX        =   4366
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Õœœ «·”«∆ﬁ"
               Height          =   285
               Index           =   82
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   253
               Top             =   600
               Width           =   855
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Õœœ «·„⁄œÂ/«·”Ì«—…"
               Height          =   285
               Index           =   81
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   252
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.TextBox TxtPurchaseBill 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   18060
            TabIndex        =   247
            Top             =   1455
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Frame Frame9 
            Caption         =   "Frame9"
            Height          =   990
            Left            =   -4260
            TabIndex        =   240
            Top             =   -735
            Visible         =   0   'False
            Width           =   4155
            Begin VB.ComboBox CboPOSBillType 
               Height          =   315
               Left            =   2265
               Style           =   2  'Dropdown List
               TabIndex        =   241
               Top             =   195
               Width           =   1635
            End
            Begin VB.Label lblSessionD 
               Caption         =   "0"
               Height          =   375
               Left            =   720
               TabIndex        =   245
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label LBLTable1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
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
               ForeColor       =   &H00FFFFFF&
               Height          =   435
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   244
               Top             =   1080
               Width           =   2685
            End
            Begin VB.Label LblSessionID 
               Height          =   375
               Left            =   480
               TabIndex        =   243
               Top             =   1200
               Width           =   2055
            End
            Begin VB.Label LblStableID 
               Caption         =   "-1"
               Height          =   375
               Left            =   3000
               TabIndex        =   242
               Top             =   720
               Width           =   855
            End
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   21300
            TabIndex        =   203
            Top             =   780
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  «·œ›⁄"
            Height          =   1470
            Left            =   -2220
            RightToLeft     =   -1  'True
            TabIndex        =   175
            Top             =   0
            Visible         =   0   'False
            Width           =   2250
            Begin VB.TextBox TxtRemainValue 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   178
               Top             =   960
               Width           =   1380
            End
            Begin VB.TextBox TxtPayedValue 
               Alignment       =   1  'Right Justify
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   120
               TabIndex        =   177
               Top             =   600
               Width           =   1380
            End
            Begin VB.TextBox TxtNetValue 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   176
               Top             =   240
               Width           =   1380
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„ »ﬁ”"
               Height          =   195
               Index           =   60
               Left            =   1440
               TabIndex        =   181
               Top             =   960
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„œ›Ê⁄"
               Height          =   315
               Index           =   59
               Left            =   1440
               TabIndex        =   180
               Top             =   600
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«Ã„«·Ì"
               Height          =   315
               Index           =   58
               Left            =   1440
               TabIndex        =   179
               Top             =   240
               Width           =   735
            End
         End
         Begin VB.CommandButton Command3 
            Caption         =   "«” ⁄·«„ ⁄‰ ’‰›"
            Height          =   255
            Left            =   4800
            TabIndex        =   101
            Top             =   2550
            Visible         =   0   'False
            Width           =   1965
         End
         Begin VB.TextBox TxtIssueSerial 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   0
            TabIndex        =   89
            Top             =   -240
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.TextBox TXTOrDer_no 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   -2055
            TabIndex        =   81
            Top             =   1095
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   14940
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            Caption         =   " ÕÊÌ· «·Ï «–‰ ’—›"
            Height          =   255
            Left            =   10980
            TabIndex        =   76
            Top             =   -240
            Visible         =   0   'False
            Width           =   2040
         End
         Begin VB.TextBox TxtCashCustomerName 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   17400
            TabIndex        =   4
            Top             =   1095
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.TextBox TxtEmployeeID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4875
            TabIndex        =   7
            Top             =   120
            Width           =   900
         End
         Begin VB.TextBox TxtStoreID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   21420
            TabIndex        =   5
            Top             =   1455
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.TextBox TxtCusID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   18255
            TabIndex        =   2
            Top             =   780
            Width           =   1005
         End
         Begin VB.ComboBox CboSaleType 
            Height          =   315
            Left            =   -2025
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   705
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   13860
            TabIndex        =   0
            Top             =   -300
            Visible         =   0   'False
            Width           =   1680
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   750
            Index           =   8
            Left            =   17370
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   1695
            Visible         =   0   'False
            Width           =   4515
            _cx             =   7964
            _cy             =   1323
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
               Left            =   5370
               TabIndex        =   31
               Top             =   165
               Width           =   2985
               _ExtentX        =   5265
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
               ButtonImage     =   "FrmSaleBillCompose.frx":038A
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
               Left            =   8235
               TabIndex        =   36
               Top             =   420
               Width           =   4290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·—»Õ"
               ForeColor       =   &H00C00000&
               Height          =   255
               Index           =   22
               Left            =   34695
               TabIndex        =   35
               Top             =   150
               Width           =   4320
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
               Left            =   3540
               TabIndex        =   34
               Top             =   390
               Width           =   5445
            End
            Begin VB.Label LblInvProfit1 
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
               Height          =   285
               Left            =   3540
               TabIndex        =   33
               Top             =   105
               Width           =   5445
            End
         End
         Begin VB.ComboBox XPCboDiscountType 
            Height          =   315
            Left            =   -2220
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   735
            Visible         =   0   'False
            Width           =   1770
         End
         Begin VB.TextBox XPTxtDiscountVal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   -2205
            TabIndex        =   10
            Top             =   735
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.ComboBox CboPayMentType 
            Height          =   315
            Left            =   11985
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   480
            Width           =   1680
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   13680
            TabIndex        =   3
            Top             =   1080
            Visible         =   0   'False
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            BoundColumn     =   ""
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   6990
            TabIndex        =   6
            Top             =   480
            Width           =   4065
            _ExtentX        =   7170
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   330
            Left            =   14940
            TabIndex        =   1
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   99024897
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton XPBtnNewClients 
            Height          =   360
            Left            =   22425
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   765
            Visible         =   0   'False
            Width           =   510
            _ExtentX        =   900
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
            ButtonImage     =   "FrmSaleBillCompose.frx":0724
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdCash 
            Height          =   390
            Index           =   0
            Left            =   10605
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   1395
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   688
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
            ButtonImage     =   "FrmSaleBillCompose.frx":0ABE
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdCash 
            Height          =   270
            Index           =   1
            Left            =   10470
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   1395
            Visible         =   0   'False
            Width           =   615
            _ExtentX        =   1085
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
            ButtonImage     =   "FrmSaleBillCompose.frx":0E58
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   10650
            TabIndex        =   82
            Top             =   120
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCPaymentNet 
            Height          =   315
            Left            =   17790
            TabIndex        =   84
            Top             =   1095
            Visible         =   0   'False
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   6990
            TabIndex        =   86
            Top             =   120
            Width           =   2610
            _ExtentX        =   4604
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCDocTypes 
            Height          =   315
            Left            =   -1725
            TabIndex        =   186
            Top             =   480
            Visible         =   0   'False
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   1830
            Left            =   -4635
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   120
            Width           =   4530
            _cx             =   7990
            _cy             =   3228
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
            Begin VB.TextBox TxtManualNo2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   120
               TabIndex        =   197
               Top             =   360
               Visible         =   0   'False
               Width           =   1140
            End
            Begin VB.TextBox TxtManualNo1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   2760
               TabIndex        =   195
               Top             =   0
               Visible         =   0   'False
               Width           =   1140
            End
            Begin VB.TextBox txt_Currency_rate 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   188
               Text            =   "1"
               Top             =   1440
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.Frame Frame2 
               Caption         =   " œ·«·«  «·«·Ê«‰"
               Height          =   735
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   720
               Visible         =   0   'False
               Width           =   2280
               Begin VB.Label Label5 
                  BackColor       =   &H000000FF&
                  Height          =   255
                  Left            =   1920
                  TabIndex        =   107
                  Top             =   240
                  Width           =   255
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  Caption         =   "»Ì⁄ «ﬁ· „‰ ”⁄— «· ﬂ·›Â"
                  Height          =   255
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   240
                  Width           =   1575
               End
               Begin VB.Label Label7 
                  BackColor       =   &H0000FFFF&
                  Height          =   255
                  Left            =   1920
                  TabIndex        =   105
                  Top             =   480
                  Width           =   255
               End
               Begin VB.Label Label8 
                  Alignment       =   1  'Right Justify
                  Caption         =   "»Ì⁄  »”⁄— «· ﬂ·›Â"
                  Height          =   255
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   480
                  Width           =   1215
               End
            End
            Begin MSDataListLib.DataCombo DcCurrency 
               Height          =   315
               Left            =   1140
               TabIndex        =   189
               Top             =   1440
               Visible         =   0   'False
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
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„  »Ê·Ì’… «·‘Õ‰"
               Height          =   195
               Index           =   67
               Left            =   1320
               TabIndex        =   198
               Top             =   360
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «–‰ «· ”·Ì„"
               Height          =   195
               Index           =   66
               Left            =   3840
               TabIndex        =   196
               Top             =   120
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "««·⁄„·…"
               Height          =   285
               Index           =   65
               Left            =   2265
               RightToLeft     =   -1  'True
               TabIndex        =   190
               Top             =   1440
               Visible         =   0   'False
               Width           =   540
            End
         End
         Begin VB.Frame Frame400 
            Height          =   495
            Left            =   18720
            RightToLeft     =   -1  'True
            TabIndex        =   199
            Top             =   1335
            Visible         =   0   'False
            Width           =   2835
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—»Õ «·›« Ê—…"
               ForeColor       =   &H00008000&
               Height          =   195
               Index           =   68
               Left            =   1680
               TabIndex        =   202
               Top             =   240
               Width           =   960
            End
            Begin VB.Label LblPrecenValuex 
               Caption         =   "0"
               Height          =   255
               Left            =   120
               TabIndex        =   201
               Top             =   840
               Width           =   1455
            End
            Begin VB.Label LblInvProfit 
               Alignment       =   2  'Center
               Caption         =   "0"
               ForeColor       =   &H00008000&
               Height          =   255
               Left            =   120
               TabIndex        =   200
               Top             =   240
               Width           =   1575
            End
         End
         Begin MSComCtl2.DTPicker DtpDelayDate 
            Height          =   285
            Left            =   -1320
            TabIndex        =   207
            Top             =   1215
            Visible         =   0   'False
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   503
            _Version        =   393216
            Format          =   99024897
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DCPreFix 
            Height          =   315
            Left            =   13755
            TabIndex        =   246
            Top             =   120
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbFarm 
            Height          =   315
            Left            =   11985
            TabIndex        =   299
            Top             =   885
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            BoundColumn     =   ""
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   315
            Left            =   180
            TabIndex        =   313
            Top             =   120
            Width           =   4605
            _ExtentX        =   8123
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ VAT"
            Height          =   270
            Index           =   105
            Left            =   11145
            RightToLeft     =   -1  'True
            TabIndex        =   331
            Top             =   900
            Width           =   645
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Œ“‰"
            Height          =   240
            Index           =   98
            Left            =   11010
            TabIndex        =   329
            Top             =   480
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„…"
            Height          =   210
            Index           =   86
            Left            =   930
            TabIndex        =   315
            Top             =   480
            Width           =   675
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«ﬂ—«„Ì…"
            Height          =   240
            Index           =   99
            Left            =   810
            TabIndex        =   311
            Top             =   855
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "›« Ê—… «·„Õ—Ã"
            Height          =   255
            Index           =   97
            Left            =   5895
            TabIndex        =   308
            Top             =   540
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·„—ﬂ»…"
            Height          =   240
            Index           =   96
            Left            =   8730
            TabIndex        =   306
            Top             =   900
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”«∆ﬁ"
            Height          =   240
            Index           =   95
            Left            =   5790
            TabIndex        =   304
            Top             =   900
            Width           =   960
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ê·…"
            Height          =   255
            Index           =   94
            Left            =   2940
            TabIndex        =   302
            Top             =   540
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Ê—œ "
            Height          =   270
            Index           =   93
            Left            =   15795
            TabIndex        =   301
            Top             =   900
            Width           =   1365
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì·"
            Height          =   270
            Index           =   7
            Left            =   17595
            TabIndex        =   300
            Top             =   825
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "› „ —ﬁ„ "
            Height          =   240
            Index           =   80
            Left            =   17745
            TabIndex        =   248
            Top             =   1455
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "  «·«” Õﬁ«ﬁ"
            Height          =   270
            Index           =   21
            Left            =   -1125
            TabIndex        =   208
            Top             =   1530
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·”‰œ"
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   0
            Left            =   -1320
            TabIndex        =   187
            Top             =   480
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.Label Label4 
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   270
            Left            =   -630
            TabIndex        =   88
            Top             =   480
            Visible         =   0   'False
            Width           =   705
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·Œ“‰…"
            Height          =   225
            Index           =   11
            Left            =   9615
            TabIndex        =   87
            Top             =   180
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ‰Ê⁄ «·‘»ﬂ…"
            Height          =   300
            Index           =   57
            Left            =   17670
            TabIndex        =   85
            Top             =   1140
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   12840
            TabIndex        =   83
            Top             =   75
            Width           =   705
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ÿ·»Ì…"
            Height          =   240
            Index           =   56
            Left            =   -1725
            TabIndex        =   80
            Top             =   1215
            Visible         =   0   'False
            Width           =   1005
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
            Left            =   -705
            TabIndex        =   77
            Top             =   855
            Visible         =   0   'False
            Width           =   345
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì· «·‰ﬁœÏ"
            Height          =   300
            Index           =   33
            Left            =   18420
            TabIndex        =   43
            Top             =   1155
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”Ì«”… «·»Ì⁄"
            Height          =   240
            Index           =   32
            Left            =   -1560
            TabIndex        =   39
            Top             =   1425
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·»«∆⁄"
            Height          =   285
            Index           =   25
            Left            =   6060
            TabIndex        =   29
            Top             =   75
            Width           =   690
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Œ’„"
            Height          =   195
            Index           =   10
            Left            =   -1320
            TabIndex        =   22
            Top             =   735
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·œ›⁄"
            Height          =   315
            Index           =   9
            Left            =   13905
            TabIndex        =   21
            Top             =   525
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„…"
            Height          =   330
            Index           =   8
            Left            =   -1200
            TabIndex        =   20
            Top             =   855
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   240
            Index           =   24
            Left            =   16350
            TabIndex        =   18
            Top             =   1500
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   300
            Index           =   6
            Left            =   14685
            TabIndex        =   17
            Top             =   540
            Width           =   2400
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·«‘⁄«—"
            Height          =   285
            Index           =   5
            Left            =   15330
            TabIndex        =   16
            Top             =   75
            Width           =   1755
         End
      End
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   5640
         Left            =   15
         TabIndex        =   13
         Top             =   2010
         Width           =   17295
         _cx             =   30506
         _cy             =   9948
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
         Caption         =   "«·√’‰«›|«·«ﬁ”«ÿ  Ê «·‘Ìﬂ« |„·«ÕŸ«  ⁄·Ï «·›« Ê—…|«·„—›ﬁ« "
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
         Picture(0)      =   "FrmSaleBillCompose.frx":11F2
         Picture(1)      =   "FrmSaleBillCompose.frx":158C
         Flags(1)        =   2
         Picture(2)      =   "FrmSaleBillCompose.frx":1926
         Flags(2)        =   2
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5175
            Index           =   19
            Left            =   18540
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   45
            Width           =   17205
            _cx             =   30348
            _cy             =   9128
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
            Begin VB.CheckBox ChecVAT 
               Alignment       =   1  'Right Justify
               Caption         =   " ÕœÌœ «·ﬂ·"
               Height          =   225
               Left            =   16020
               RightToLeft     =   -1  'True
               TabIndex        =   297
               Top             =   975
               Width           =   1125
            End
            Begin VB.TextBox TxtGTotal 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   210
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   294
               Top             =   4830
               Visible         =   0   'False
               Width           =   2145
            End
            Begin VB.TextBox TxtSAlGValue 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   210
               Left            =   13635
               RightToLeft     =   -1  'True
               TabIndex        =   292
               Top             =   4830
               Width           =   2145
            End
            Begin VB.TextBox TxtNoteSerial111 
               Alignment       =   1  'Right Justify
               Height          =   240
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   285
               Top             =   735
               Width           =   1950
            End
            Begin VB.CommandButton Command9 
               Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
               Height          =   405
               Left            =   2625
               RightToLeft     =   -1  'True
               TabIndex        =   283
               Top             =   525
               Width           =   1965
            End
            Begin VB.TextBox TxtNoteID111 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   1425
               RightToLeft     =   -1  'True
               TabIndex        =   282
               Top             =   585
               Visible         =   0   'False
               Width           =   900
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1335
               Index           =   6
               Left            =   9465
               TabIndex        =   265
               TabStop         =   0   'False
               Top             =   0
               Width           =   7680
               _cx             =   13547
               _cy             =   2355
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
               Caption         =   " ÕœÌœ «·› —… «·“„‰Ì…  ··›Ê« Ì—"
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
               Begin VB.TextBox TxtSuplCode 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   4890
                  RightToLeft     =   -1  'True
                  TabIndex        =   288
                  Top             =   240
                  Width           =   1500
               End
               Begin MSComCtl2.DTPicker DTPickerAccFrom 
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "dd/MM/yyyy"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   11265
                     SubFormatType   =   3
                  EndProperty
                  Height          =   345
                  Left            =   4890
                  TabIndex        =   266
                  ToolTipText     =   "„‰  «—ÌŒ ﬁœÌ„"
                  Top             =   600
                  Width           =   1500
                  _ExtentX        =   2646
                  _ExtentY        =   609
                  _Version        =   393216
                  CalendarBackColor=   -2147483624
                  CalendarTitleBackColor=   10383715
                  CheckBox        =   -1  'True
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   99024899
                  CurrentDate     =   37357
               End
               Begin MSComCtl2.DTPicker DTPickerAccTo 
                  BeginProperty DataFormat 
                     Type            =   1
                     Format          =   "dd/MM/yyyy"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   11265
                     SubFormatType   =   3
                  EndProperty
                  Height          =   345
                  Left            =   2760
                  TabIndex        =   267
                  ToolTipText     =   " ≈·Ï  «—ÌŒ √ÕœÀ"
                  Top             =   600
                  Width           =   1500
                  _ExtentX        =   2646
                  _ExtentY        =   609
                  _Version        =   393216
                  CalendarBackColor=   -2147483624
                  CalendarTitleBackColor=   10383715
                  CheckBox        =   -1  'True
                  CustomFormat    =   "yyyy/M/d"
                  Format          =   99024899
                  CurrentDate     =   37357
               End
               Begin ImpulseButton.ISButton Cmd1 
                  CausesValidation=   0   'False
                  Height          =   300
                  Left            =   1200
                  TabIndex        =   286
                  Top             =   960
                  Width           =   1185
                  _ExtentX        =   2090
                  _ExtentY        =   529
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "„—›ﬁ«  «·›« Ê—…"
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
               Begin ImpulseButton.ISButton ISButton1 
                  CausesValidation=   0   'False
                  Height          =   300
                  Left            =   3120
                  TabIndex        =   287
                  Top             =   960
                  Width           =   1065
                  _ExtentX        =   1879
                  _ExtentY        =   529
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "⁄—÷ «·›Ê« Ì—"
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
               Begin XtremeSuiteControls.CheckBox CheckBox1 
                  Height          =   315
                  Left            =   6600
                  TabIndex        =   289
                  Top             =   240
                  Width           =   975
                  _Version        =   786432
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  _StockProps     =   79
                  Caption         =   "Õœœ «·„Ê—œ"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbSuppler 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   290
                  Top             =   240
                  Width           =   4725
                  _ExtentX        =   8334
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox ChLoadVAT 
                  Height          =   255
                  Left            =   360
                  TabIndex        =   291
                  Top             =   600
                  Width           =   2295
                  _Version        =   786432
                  _ExtentX        =   4048
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   " Õ„Ì· «·›Ê« Ì— »«·ﬁÌ„… «·„÷«›…"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   285
                  Index           =   84
                  Left            =   4110
                  RightToLeft     =   -1  'True
                  TabIndex        =   269
                  Top             =   600
                  Width           =   555
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   285
                  Index           =   83
                  Left            =   6870
                  RightToLeft     =   -1  'True
                  TabIndex        =   268
                  Top             =   645
                  Width           =   555
               End
            End
            Begin VB.Frame Frame6 
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ«·Â «·«⁄ „«œ"
               Height          =   2895
               Left            =   -3810
               RightToLeft     =   -1  'True
               TabIndex        =   256
               Top             =   -2460
               Visible         =   0   'False
               Width           =   5355
               Begin VSFlex8UCtl.VSFlexGrid FGApproval 
                  Height          =   1725
                  Left            =   240
                  TabIndex        =   257
                  Tag             =   "1"
                  Top             =   360
                  Width           =   4935
                  _cx             =   8705
                  _cy             =   3043
                  Appearance      =   1
                  BorderStyle     =   1
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
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   300
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmSaleBillCompose.frx":1CC0
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
            Begin VB.Frame Frame3 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  ﬁÌœ «·›« Ê—Â"
               Height          =   1740
               Left            =   -3930
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   -1305
               Visible         =   0   'False
               Width           =   4290
               Begin VB.TextBox TxtNoteSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   98
                  Top             =   600
                  Width           =   2505
               End
               Begin ImpulseButton.ISButton Cmd 
                  CausesValidation=   0   'False
                  Height          =   375
                  Index           =   10
                  Left            =   240
                  TabIndex        =   99
                  Top             =   600
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   540
                  Index           =   8
                  Left            =   240
                  TabIndex        =   254
                  Top             =   840
                  Width           =   1785
                  _ExtentX        =   3149
                  _ExtentY        =   953
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "ÿ»«⁄… ⁄ﬁœ »Ì⁄ »«· ﬁ”Ìÿ"
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
                  Height          =   300
                  Index           =   9
                  Left            =   2520
                  TabIndex        =   255
                  Top             =   1320
                  Width           =   1785
                  _ExtentX        =   3149
                  _ExtentY        =   529
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "ÿ»«⁄Â ”‰œ ·√„—"
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
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·ﬁÌœ ··›« Ê—Â"
                  Height          =   435
                  Index           =   62
                  Left            =   2880
                  TabIndex        =   100
                  Top             =   240
                  Width           =   1215
               End
            End
            Begin VB.OptionButton BillBasedOn 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›« Ê—… „»Ì⁄« "
               Height          =   210
               Index           =   0
               Left            =   13755
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   315
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   4755
            End
            Begin VB.OptionButton BillBasedOn 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "√Ê«„— «·»Ì⁄"
               Height          =   165
               Index           =   2
               Left            =   17205
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   1455
               Visible         =   0   'False
               Width           =   705
            End
            Begin VB.OptionButton BillBasedOn 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”‰œ«  «·’—›"
               Height          =   255
               Index           =   1
               Left            =   11610
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   585
               Visible         =   0   'False
               Width           =   5595
            End
            Begin VB.TextBox TXTNoteID 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   0
               Visible         =   0   'False
               Width           =   1425
            End
            Begin VSFlex8UCtl.VSFlexGrid GRID1 
               Height          =   1110
               Left            =   21135
               TabIndex        =   90
               Tag             =   "1"
               Top             =   3045
               Visible         =   0   'False
               Width           =   8340
               _cx             =   14711
               _cy             =   1958
               Appearance      =   1
               BorderStyle     =   1
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
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmSaleBillCompose.frx":1D93
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
            Begin VSFlex8UCtl.VSFlexGrid GRID2 
               Height          =   2805
               Left            =   9465
               TabIndex        =   92
               Tag             =   "1"
               Top             =   1455
               Width           =   7680
               _cx             =   13547
               _cy             =   4948
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
               Cols            =   6
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmSaleBillCompose.frx":1EE0
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
            Begin ImpulseButton.ISButton Cmd 
               Height          =   510
               Index           =   11
               Left            =   -825
               TabIndex        =   258
               Top             =   2205
               Visible         =   0   'False
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   900
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "≈⁄ „«œ"
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
            Begin DBPIXLib.DBPix20 DBPix202 
               Height          =   495
               Index           =   0
               Left            =   -1305
               TabIndex        =   259
               Top             =   2790
               Visible         =   0   'False
               Width           =   1665
               _Version        =   131072
               _ExtentX        =   2937
               _ExtentY        =   873
               _StockProps     =   1
               BackColor       =   12632256
               _Image          =   "FrmSaleBillCompose.frx":1FD6
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
            Begin DBPIXLib.DBPix20 DBPix202 
               Height          =   510
               Index           =   1
               Left            =   -480
               TabIndex        =   260
               Top             =   3360
               Visible         =   0   'False
               Width           =   1665
               _Version        =   131072
               _ExtentX        =   2937
               _ExtentY        =   900
               _StockProps     =   1
               BackColor       =   12632256
               _Image          =   "FrmSaleBillCompose.frx":1FEE
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
            Begin DBPIXLib.DBPix20 DBPix202 
               Height          =   495
               Index           =   2
               Left            =   -240
               TabIndex        =   261
               Top             =   3945
               Visible         =   0   'False
               Width           =   1665
               _Version        =   131072
               _ExtentX        =   2937
               _ExtentY        =   873
               _StockProps     =   1
               BackColor       =   12632256
               _Image          =   "FrmSaleBillCompose.frx":2006
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
            Begin DBPIXLib.DBPix20 DBPix202 
               Height          =   495
               Index           =   3
               Left            =   -1425
               TabIndex        =   262
               Top             =   2790
               Visible         =   0   'False
               Width           =   1665
               _Version        =   131072
               _ExtentX        =   2937
               _ExtentY        =   873
               _StockProps     =   1
               BackColor       =   12632256
               _Image          =   "FrmSaleBillCompose.frx":201E
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
            Begin DBPIXLib.DBPix20 DBPix202 
               Height          =   510
               Index           =   4
               Left            =   -120
               TabIndex        =   263
               Top             =   3360
               Visible         =   0   'False
               Width           =   1665
               _Version        =   131072
               _ExtentX        =   2937
               _ExtentY        =   900
               _StockProps     =   1
               BackColor       =   12632256
               _Image          =   "FrmSaleBillCompose.frx":2036
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
            Begin DBPIXLib.DBPix20 DBPix202 
               Height          =   495
               Index           =   5
               Left            =   -120
               TabIndex        =   264
               Top             =   3945
               Visible         =   0   'False
               Width           =   1665
               _Version        =   131072
               _ExtentX        =   2937
               _ExtentY        =   873
               _StockProps     =   1
               BackColor       =   12632256
               _Image          =   "FrmSaleBillCompose.frx":204E
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
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«Ã„«·Ì"
               Height          =   195
               Index           =   92
               Left            =   2385
               TabIndex        =   295
               Top             =   4830
               Visible         =   0   'False
               Width           =   780
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ã„«·Ì «·„»Ì⁄« "
               Height          =   195
               Index           =   89
               Left            =   15900
               TabIndex        =   293
               Top             =   4830
               Width           =   1125
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·ﬁÌœ"
               Height          =   300
               Index           =   35
               Left            =   300
               RightToLeft     =   -1  'True
               TabIndex        =   284
               Top             =   420
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄„Ê·…"
               Height          =   195
               Index           =   85
               Left            =   8400
               TabIndex        =   281
               Top             =   210
               Width           =   645
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›« Ê—Â »‰«¡ ⁄·Ï"
               Height          =   195
               Index           =   61
               Left            =   12450
               TabIndex        =   94
               Top             =   150
               Visible         =   0   'False
               Width           =   2490
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5175
            Index           =   15
            Left            =   18240
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   45
            Width           =   17205
            _cx             =   30348
            _cy             =   9128
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
            _GridInfo       =   $"FrmSaleBillCompose.frx":2066
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1950
               Index           =   18
               Left            =   15
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   255
               Visible         =   0   'False
               Width           =   17175
               _cx             =   30295
               _cy             =   3440
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
               Begin VB.Frame Frame4 
                  Height          =   1230
                  Left            =   2385
                  TabIndex        =   235
                  Top             =   -75
                  Width           =   780
                  Begin VB.ComboBox CboPaymentType1 
                     Height          =   315
                     Left            =   0
                     RightToLeft     =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   237
                     Top             =   585
                     Width           =   2685
                  End
                  Begin VB.TextBox TxtAdvPaymnt 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   0
                     MaxLength       =   10
                     RightToLeft     =   -1  'True
                     TabIndex        =   236
                     Top             =   240
                     Width           =   2685
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "ÿ—Ìﬁ… «·ﬁ»÷"
                     Height          =   315
                     Index           =   79
                     Left            =   2850
                     RightToLeft     =   -1  'True
                     TabIndex        =   239
                     Top             =   585
                     Width           =   1275
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   "ﬁÌ„… «·œ›⁄Â"
                     Height          =   285
                     Index           =   78
                     Left            =   2850
                     RightToLeft     =   -1  'True
                     TabIndex        =   238
                     Top             =   255
                     Width           =   1275
                  End
               End
               Begin VB.Frame FraNote 
                  BackColor       =   &H00E2E9E9&
                  Height          =   1320
                  Left            =   1650
                  RightToLeft     =   -1  'True
                  TabIndex        =   223
                  Top             =   -120
                  Width           =   765
                  Begin VB.TextBox TxtChequeNumber 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   30
                     RightToLeft     =   -1  'True
                     TabIndex        =   225
                     Top             =   930
                     Width           =   2685
                  End
                  Begin VB.TextBox TXTBankName 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   224
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   2685
                  End
                  Begin MSComCtl2.DTPicker DtpChequeDueDate1 
                     Height          =   315
                     Left            =   30
                     TabIndex        =   226
                     Top             =   1260
                     Width           =   2685
                     _ExtentX        =   4736
                     _ExtentY        =   556
                     _Version        =   393216
                     Format          =   99024897
                     CurrentDate     =   39614
                  End
                  Begin MSDataListLib.DataCombo DcboBankName1 
                     Height          =   315
                     Left            =   30
                     TabIndex        =   227
                     Top             =   600
                     Width           =   2685
                     _ExtentX        =   4736
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DcboBox1 
                     Height          =   315
                     Left            =   30
                     TabIndex        =   228
                     Top             =   270
                     Width           =   2685
                     _ExtentX        =   4736
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DcChequeBox1 
                     Height          =   315
                     Left            =   0
                     TabIndex        =   229
                     Top             =   1680
                     Width           =   2685
                     _ExtentX        =   4736
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " «—ÌŒ «·≈” Õﬁ«ﬁ"
                     Height          =   285
                     Index           =   77
                     Left            =   2820
                     RightToLeft     =   -1  'True
                     TabIndex        =   234
                     Top             =   1260
                     Width           =   1215
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "—ﬁ„ «·‘Ìﬂ"
                     Height          =   285
                     Index           =   76
                     Left            =   2760
                     RightToLeft     =   -1  'True
                     TabIndex        =   233
                     Top             =   930
                     Width           =   1215
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«”„ «·»‰ﬂ"
                     Height          =   285
                     Index           =   75
                     Left            =   2790
                     RightToLeft     =   -1  'True
                     TabIndex        =   232
                     Top             =   630
                     Width           =   1215
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«”„ «·Œ“‰…"
                     Height          =   285
                     Index           =   74
                     Left            =   2790
                     RightToLeft     =   -1  'True
                     TabIndex        =   231
                     Top             =   300
                     Width           =   1215
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ«›Ÿ… «·‘Ìﬂ« "
                     Height          =   285
                     Index           =   73
                     Left            =   2760
                     RightToLeft     =   -1  'True
                     TabIndex        =   230
                     Top             =   1560
                     Width           =   1215
                  End
               End
               Begin VB.TextBox TxtTaxServiceValue 
                  Alignment       =   1  'Right Justify
                  Height          =   165
                  Left            =   1290
                  MaxLength       =   4
                  TabIndex        =   63
                  Top             =   45
                  Visible         =   0   'False
                  Width           =   240
               End
               Begin VB.CheckBox ChkTaxSerivce 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì… Œœ„…"
                  Height          =   105
                  Left            =   1785
                  TabIndex        =   58
                  Top             =   15
                  Width           =   225
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   135
                  Index           =   54
                  Left            =   945
                  TabIndex        =   75
                  Top             =   45
                  Visible         =   0   'False
                  Width           =   255
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
                  Height          =   135
                  Index           =   47
                  Left            =   1215
                  TabIndex        =   68
                  Top             =   45
                  Visible         =   0   'False
                  Width           =   75
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   165
                  Index           =   43
                  Left            =   1530
                  TabIndex        =   64
                  Top             =   45
                  Visible         =   0   'False
                  Width           =   90
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1710
               Index           =   17
               Left            =   15
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   495
               Visible         =   0   'False
               Width           =   17175
               _cx             =   30295
               _cy             =   3016
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
                  Height          =   135
                  Left            =   1290
                  MaxLength       =   4
                  TabIndex        =   62
                  Top             =   30
                  Width           =   240
               End
               Begin VB.CheckBox ChkTaxStamp 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "œ„€…"
                  Height          =   90
                  Left            =   1785
                  TabIndex        =   56
                  Top             =   75
                  Width           =   225
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   105
                  Index           =   53
                  Left            =   945
                  TabIndex        =   74
                  Top             =   30
                  Visible         =   0   'False
                  Width           =   255
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
                  Height          =   105
                  Index           =   48
                  Left            =   1215
                  TabIndex        =   69
                  Top             =   30
                  Width           =   60
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   120
                  Index           =   41
                  Left            =   1530
                  TabIndex        =   60
                  Top             =   30
                  Width           =   105
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1710
               Index           =   16
               Left            =   15
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   495
               Visible         =   0   'False
               Width           =   17175
               _cx             =   30295
               _cy             =   3016
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
                  Height          =   105
                  Left            =   1350
                  MaxLength       =   4
                  TabIndex        =   61
                  Top             =   15
                  Width           =   255
               End
               Begin VB.CheckBox ChkTaxAdd 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì»… Œ’„ Ê≈÷«›… (√—»«Õ  Ã«—Ì…)"
                  Height          =   135
                  Left            =   1740
                  TabIndex        =   54
                  Top             =   15
                  Width           =   390
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   90
                  Index           =   52
                  Left            =   975
                  TabIndex        =   73
                  Top             =   15
                  Visible         =   0   'False
                  Width           =   285
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
                  Index           =   46
                  Left            =   1275
                  TabIndex        =   67
                  Top             =   15
                  Width           =   75
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   90
                  Index           =   39
                  Left            =   1605
                  TabIndex        =   59
                  Top             =   15
                  Width           =   105
               End
            End
            Begin VB.TextBox TxtBillComment 
               Alignment       =   1  'Right Justify
               Height          =   2940
               Left            =   15
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   47
               Top             =   2220
               Width           =   17175
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   105
               Index           =   4
               Left            =   15
               TabIndex        =   49
               TabStop         =   0   'False
               Top             =   15
               Visible         =   0   'False
               Width           =   17175
               _cx             =   30295
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
                  Height          =   315
                  Left            =   150
                  TabIndex        =   51
                  Top             =   225
                  Width           =   30
               End
               Begin VB.TextBox XPTxtTaxValue 
                  Alignment       =   1  'Right Justify
                  Height          =   510
                  Left            =   105
                  MaxLength       =   4
                  TabIndex        =   50
                  Top             =   105
                  Width           =   30
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   360
                  Index           =   51
                  Left            =   15
                  TabIndex        =   72
                  Top             =   135
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
                  Height          =   360
                  Index           =   45
                  Left            =   105
                  TabIndex        =   66
                  Top             =   135
                  Width           =   15
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   240
                  Index           =   4
                  Left            =   120
                  TabIndex        =   52
                  Top             =   195
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
               Height          =   1710
               Index           =   44
               Left            =   15
               TabIndex        =   65
               Top             =   495
               Width           =   17175
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5175
            Index           =   7
            Left            =   45
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   45
            Width           =   17205
            _cx             =   30348
            _cy             =   9128
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
            Begin MSComctlLib.Toolbar TBar 
               Height          =   630
               Left            =   0
               TabIndex        =   37
               Top             =   2685
               Width           =   17205
               _ExtentX        =   30348
               _ExtentY        =   1111
               ButtonWidth     =   609
               ButtonHeight    =   1005
               Appearance      =   1
               _Version        =   393216
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   645
               Index           =   2
               Left            =   0
               TabIndex        =   142
               TabStop         =   0   'False
               Top             =   30
               Width           =   17205
               _cx             =   30348
               _cy             =   1138
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
               Begin VB.TextBox txtCustCode 
                  BackColor       =   &H00FFFFFF&
                  Height          =   315
                  Left            =   15825
                  TabIndex        =   343
                  Top             =   210
                  Width           =   1125
               End
               Begin VB.TextBox TxtPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   360
                  Left            =   585
                  MaxLength       =   10
                  TabIndex        =   146
                  Top             =   210
                  Width           =   1425
               End
               Begin VB.TextBox TxtSerial 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  Height          =   360
                  Left            =   4050
                  MaxLength       =   20
                  TabIndex        =   145
                  Top             =   210
                  Width           =   1275
               End
               Begin VB.TextBox TxtQuantity 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   360
                  Left            =   2220
                  MaxLength       =   10
                  TabIndex        =   144
                  Top             =   210
                  Width           =   1830
               End
               Begin VB.ComboBox CboItemCase 
                  Height          =   315
                  Left            =   5535
                  Style           =   2  'Dropdown List
                  TabIndex        =   143
                  Top             =   210
                  Width           =   1080
               End
               Begin MSDataListLib.DataCombo DCboItemsName 
                  Height          =   315
                  Left            =   6765
                  TabIndex        =   147
                  Top             =   210
                  Width           =   3420
                  _ExtentX        =   6033
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
                  Left            =   10215
                  TabIndex        =   148
                  Top             =   210
                  Width           =   2220
                  _ExtentX        =   3916
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdAdd 
                  Height          =   240
                  Left            =   150
                  TabIndex        =   149
                  Top             =   210
                  Width           =   345
                  _ExtentX        =   609
                  _ExtentY        =   423
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
                  ButtonImage     =   "FrmSaleBillCompose.frx":20D8
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
                  Height          =   180
                  Left            =   465
                  TabIndex        =   150
                  Top             =   45
                  Visible         =   0   'False
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   318
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
                  ButtonImage     =   "FrmSaleBillCompose.frx":2472
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo DcCustmer 
                  Height          =   315
                  Left            =   12480
                  TabIndex        =   346
                  Top             =   210
                  Width           =   3270
                  _ExtentX        =   5768
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
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·⁄„Ì·"
                  Height          =   180
                  Index           =   102
                  Left            =   15465
                  TabIndex        =   345
                  Top             =   15
                  Width           =   1905
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈”„ «·⁄„Ì·"
                  Height          =   180
                  Index           =   101
                  Left            =   13650
                  TabIndex        =   344
                  Top             =   0
                  Width           =   810
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”⁄—"
                  Height          =   180
                  Index           =   26
                  Left            =   870
                  TabIndex        =   156
                  Top             =   15
                  Width           =   855
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂ„Ì…"
                  Height          =   180
                  Index           =   27
                  Left            =   2550
                  TabIndex        =   155
                  Top             =   30
                  Width           =   990
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ì—Ì«·"
                  Height          =   180
                  Index           =   28
                  Left            =   4230
                  TabIndex        =   154
                  Top             =   15
                  Width           =   765
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·… «·’‰›"
                  Height          =   180
                  Index           =   29
                  Left            =   5595
                  TabIndex        =   153
                  Top             =   15
                  Width           =   825
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈”„ «·’‰›"
                  Height          =   180
                  Index           =   30
                  Left            =   8100
                  TabIndex        =   152
                  Top             =   15
                  Width           =   735
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·’‰›"
                  Height          =   180
                  Index           =   31
                  Left            =   10485
                  TabIndex        =   151
                  Top             =   30
                  Width           =   1920
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid FG 
               Height          =   1980
               Left            =   120
               TabIndex        =   157
               Top             =   690
               Width           =   17145
               _cx             =   30242
               _cy             =   3492
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
               Cols            =   25
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmSaleBillCompose.frx":280C
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
            Begin MSComctlLib.Toolbar Toolbar1 
               Height          =   630
               Left            =   0
               TabIndex        =   158
               Top             =   30
               Width           =   8610
               _ExtentX        =   15187
               _ExtentY        =   1111
               ButtonWidth     =   609
               ButtonHeight    =   1005
               Appearance      =   1
               _Version        =   393216
            End
            Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
               Height          =   1350
               Left            =   0
               TabIndex        =   316
               Top             =   3525
               Width           =   17205
               _cx             =   30348
               _cy             =   2381
               Appearance      =   1
               BorderStyle     =   1
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
               Cols            =   14
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmSaleBillCompose.frx":2C4C
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
               Begin VB.PictureBox PicDes 
                  BorderStyle     =   0  'None
                  Height          =   1635
                  Left            =   240
                  RightToLeft     =   -1  'True
                  ScaleHeight     =   1635
                  ScaleWidth      =   2925
                  TabIndex        =   317
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   2925
                  Begin VB.TextBox TxtDes 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000018&
                     BorderStyle     =   0  'None
                     Height          =   1125
                     Left            =   30
                     MultiLine       =   -1  'True
                     RightToLeft     =   -1  'True
                     ScrollBars      =   3  'Both
                     TabIndex        =   318
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   2115
                  End
                  Begin VB.Label LblDes 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H8000000C&
                     Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
                     ForeColor       =   &H0000C8FF&
                     Height          =   315
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   319
                     Top             =   0
                     Width           =   2445
                  End
               End
               Begin VDSCOMBOLibCtl.SmartCombo CboDes 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   320
                  ToolTipText     =   "ﬂ «»…  ⁄·Ìﬁ"
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   2955
                  _cx             =   1973752924
                  _cy             =   1973748268
                  Alignment       =   0
                  Appearance      =   3
                  AutoSearch      =   0   'False
                  BackColor       =   -2147483624
                  BackgroundColor =   -2147483633
                  BorderColor     =   0
                  BorderVisible   =   -1  'True
                  Caption         =   "SmartCombo1"
                  CaptionAlignment=   4
                  CaptionBackColor=   -2147483633
                  BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  CaptionForeColor=   -2147483630
                  CaptionHeight   =   15
                  CaptionOnTop    =   0   'False
                  CaptionMultiLine=   0
                  Checkbox3D      =   0   'False
                  CheckboxAlignment=   5
                  CheckboxBackColor=   16777215
                  CheckboxSize    =   13
                  CheckboxValue   =   0
                  BrowsePictureAlignment=   5
                  BrowsePictureStretchH=   0
                  BrowsePictureStretchV=   0
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
                  ForeColor       =   0
                  Gap             =   0
                  HideSelection   =   -1  'True
                  Locked          =   0   'False
                  MaxLength       =   0
                  MultiLine       =   0
                  OnFocus         =   3
                  PasswordChar    =   ""
                  Picture         =   "FrmSaleBillCompose.frx":2E7E
                  PictureAlignment=   5
                  PictureBackColor=   -2147483624
                  PictureStretchH =   0
                  PictureStretchV =   0
                  Redraw          =   -1  'True
                  ScrollBar       =   0
                  Style           =   0
                  Text            =   ""
                  UnderLine       =   0   'False
                  Enabled0        =   -1  'True
                  Position0       =   0
                  Tip0            =   "Caption"
                  Visible0        =   0   'False
                  Width0          =   90
                  Enabled1        =   -1  'True
                  Position1       =   1
                  Tip1            =   ""
                  Visible1        =   -1  'True
                  Width1          =   32
                  Enabled2        =   -1  'True
                  Position2       =   2
                  Tip2            =   "Check Box (Space, Ctrl + Space)"
                  Visible2        =   0   'False
                  Width2          =   16
                  Enabled3        =   -1  'True
                  Position3       =   3
                  Tip3            =   "ﬂ «»…  ⁄·Ìﬁ"
                  Visible3        =   -1  'True
                  Width3          =   145
                  Enabled4        =   -1  'True
                  Position4       =   4
                  Tip4            =   "Left Spinner (Alt + Left)"
                  Visible4        =   0   'False
                  Width4          =   16
                  Enabled5        =   -1  'True
                  Position5       =   5
                  Tip5            =   "Right Spinner (Alt + Right)"
                  Visible5        =   0   'False
                  Width5          =   16
                  Enabled6        =   -1  'True
                  Position6       =   6
                  Tip6            =   "Up Spinner (Ctrl + Up)"
                  Visible6        =   0   'False
                  Width6          =   16
                  Enabled7        =   -1  'True
                  Position7       =   7
                  Tip7            =   "Down Spinner (Ctrl + Down)"
                  Visible7        =   0   'False
                  Width7          =   16
                  Enabled8        =   -1  'True
                  Position8       =   8
                  Tip8            =   "Browse (Alt + Enter)"
                  Visible8        =   0   'False
                  Width8          =   16
                  Enabled9        =   -1  'True
                  Position9       =   9
                  Tip9            =   " (Alt + Down, F4)"
                  Visible9        =   -1  'True
                  Width9          =   16
                  Enabled10       =   -1  'True
                  Position10      =   10
                  Tip10           =   "Right Arrow (Alt + >)"
                  Visible10       =   0   'False
                  Width10         =   16
               End
            End
            Begin ImpulseButton.ISButton CmdDel 
               Height          =   495
               Left            =   16215
               TabIndex        =   321
               Top             =   4755
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   873
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› ”ÿ—"
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
               ButtonImage     =   "FrmSaleBillCompose.frx":3418
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„’«—Ì› «Œ—Ï"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   315
               Index           =   88
               Left            =   8730
               TabIndex        =   322
               Top             =   3285
               Width           =   1170
            End
            Begin VB.Label LblItemsCount 
               Alignment       =   2  'Center
               BackColor       =   &H00404040&
               ForeColor       =   &H0000FFFF&
               Height          =   645
               Left            =   0
               TabIndex        =   38
               Top             =   4620
               Width           =   180
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5175
            Index           =   5
            Left            =   17940
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   45
            Width           =   17205
            _cx             =   30348
            _cy             =   9128
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
            AutoSizeChildren=   0
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic2 
               Height          =   7575
               Left            =   0
               TabIndex        =   108
               TabStop         =   0   'False
               Top             =   0
               Width           =   20295
               _cx             =   35798
               _cy             =   13361
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
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   690
                  Index           =   11
                  Left            =   90
                  TabIndex        =   109
                  TabStop         =   0   'False
                  Top             =   90
                  Width           =   20115
                  _cx             =   35481
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
                  Begin VB.TextBox XPTxtSerial 
                     Alignment       =   1  'Right Justify
                     Height          =   360
                     Index           =   1
                     Left            =   15840
                     Locked          =   -1  'True
                     TabIndex        =   216
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   1635
                  End
                  Begin VB.CheckBox ChkInstall 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ”Ìÿ"
                     Height          =   195
                     Left            =   3300
                     TabIndex        =   214
                     Top             =   280
                     Width           =   930
                  End
                  Begin VB.CheckBox XPChkPayType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "¬Ã· "
                     Height          =   195
                     Index           =   1
                     Left            =   7155
                     TabIndex        =   212
                     Top             =   280
                     Width           =   1215
                  End
                  Begin VB.TextBox XPTxtValue 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Index           =   1
                     Left            =   4560
                     Locked          =   -1  'True
                     MaxLength       =   10
                     TabIndex        =   211
                     Top             =   225
                     Width           =   1500
                  End
                  Begin VB.TextBox XPTxtValue 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Index           =   0
                     Left            =   8820
                     Locked          =   -1  'True
                     MaxLength       =   10
                     TabIndex        =   112
                     Top             =   225
                     Width           =   1515
                  End
                  Begin VB.TextBox XPTxtSerial 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Index           =   0
                     Left            =   14430
                     Locked          =   -1  'True
                     TabIndex        =   111
                     Top             =   75
                     Visible         =   0   'False
                     Width           =   1530
                  End
                  Begin VB.CheckBox XPChkPayType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰ﬁœ«"
                     Height          =   195
                     Index           =   0
                     Left            =   11670
                     TabIndex        =   110
                     Top             =   280
                     Width           =   930
                  End
                  Begin ImpulseButton.ISButton CmdINSTALLMENT 
                     Height          =   390
                     Left            =   240
                     TabIndex        =   215
                     Top             =   240
                     Width           =   1530
                     _ExtentX        =   2699
                     _ExtentY        =   688
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
                     ButtonImage     =   "FrmSaleBillCompose.frx":39B2
                     ColorButton     =   14871017
                     ColorHighlight  =   16777215
                     ColorHoverText  =   16711680
                     ColorShadow     =   4210752
                     ColorOutline    =   0
                     DrawFocusRectangle=   0   'False
                     ColorToggledHoverText=   16711680
                     ColorTextShadow =   4210752
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„”·”·"
                     Height          =   375
                     Index           =   14
                     Left            =   15495
                     TabIndex        =   217
                     Top             =   315
                     Visible         =   0   'False
                     Width           =   630
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁÌ„…"
                     Height          =   195
                     Index           =   15
                     Left            =   6330
                     TabIndex        =   213
                     Top             =   280
                     Width           =   420
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
                     Height          =   225
                     Index           =   20
                     Left            =   12780
                     TabIndex        =   115
                     Top             =   250
                     Width           =   1410
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁÌ„…"
                     Height          =   225
                     Index           =   13
                     Left            =   10815
                     TabIndex        =   114
                     Top             =   285
                     Width           =   600
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„”·”·"
                     Height          =   225
                     Index           =   12
                     Left            =   15270
                     TabIndex        =   113
                     Top             =   45
                     Visible         =   0   'False
                     Width           =   810
                  End
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   90
                  Index           =   12
                  Left            =   90
                  TabIndex        =   116
                  TabStop         =   0   'False
                  Top             =   720
                  Width           =   20115
                  _cx             =   35481
                  _cy             =   159
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
               End
               Begin VSFlex8UCtl.VSFlexGrid FgInstallments 
                  Height          =   2010
                  Left            =   90
                  TabIndex        =   117
                  Top             =   870
                  Width           =   17385
                  _cx             =   30665
                  _cy             =   3545
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
                  Cols            =   6
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmSaleBillCompose.frx":3D4C
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
                  Height          =   690
                  Index           =   13
                  Left            =   90
                  TabIndex        =   118
                  TabStop         =   0   'False
                  Top             =   2700
                  Width           =   20115
                  _cx             =   35481
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
                  Begin VB.Label LblAdvPayment 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Left            =   10680
                     TabIndex        =   222
                     Top             =   240
                     Width           =   690
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " «·œ›⁄Â «·„ﬁœ„Â"
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
                     Height          =   285
                     Index           =   72
                     Left            =   11400
                     TabIndex        =   221
                     Top             =   240
                     Width           =   1125
                  End
                  Begin VB.Label LBLaDVpAY 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   165
                     Left            =   240
                     TabIndex        =   220
                     Top             =   480
                     Width           =   720
                  End
                  Begin VB.Label LblDiscount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   165
                     Left            =   120
                     TabIndex        =   210
                     Top             =   240
                     Width           =   720
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Œ’„  ﬁ”Ìÿ"
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
                     Height          =   285
                     Index           =   70
                     Left            =   1320
                     TabIndex        =   209
                     Top             =   240
                     Width           =   990
                  End
                  Begin VB.Label LblPrecenValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Left            =   14400
                     TabIndex        =   204
                     Top             =   240
                     Width           =   405
                  End
                  Begin VB.Label LblInstallmentType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     ForeColor       =   &H00FF0000&
                     Height          =   285
                     Left            =   5025
                     TabIndex        =   133
                     Top             =   285
                     Width           =   555
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
                     Height          =   285
                     Index           =   42
                     Left            =   5760
                     TabIndex        =   132
                     Top             =   285
                     Width           =   1170
                  End
                  Begin VB.Label LblFirstInstallDate 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Left            =   6990
                     TabIndex        =   131
                     Top             =   285
                     Width           =   870
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
                     Height          =   285
                     Index           =   40
                     Left            =   7920
                     TabIndex        =   130
                     Top             =   285
                     Width           =   885
                  End
                  Begin VB.Label LblInstallCount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Left            =   8940
                     TabIndex        =   129
                     Top             =   285
                     Width           =   375
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
                     Height          =   285
                     Index           =   38
                     Left            =   9570
                     TabIndex        =   128
                     Top             =   285
                     Width           =   960
                  End
                  Begin VB.Label LblInstallTotal 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Left            =   12585
                     TabIndex        =   127
                     Top             =   285
                     Width           =   690
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
                     Height          =   285
                     Index           =   36
                     Left            =   13425
                     TabIndex        =   126
                     Top             =   285
                     Width           =   885
                  End
                  Begin VB.Label LblPrecenType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Left            =   15525
                     TabIndex        =   125
                     Top             =   285
                     Width           =   780
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
                     Height          =   285
                     Index           =   34
                     Left            =   16455
                     TabIndex        =   124
                     Top             =   285
                     Width           =   780
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
                     Height          =   405
                     Index           =   35
                     Left            =   14640
                     TabIndex        =   123
                     Top             =   165
                     Width           =   750
                  End
                  Begin VB.Label LblPrecenValue1 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Left            =   13620
                     TabIndex        =   122
                     Top             =   285
                     Width           =   765
                  End
                  Begin VB.Label LblInstallSeprator 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     ForeColor       =   &H00FF0000&
                     Height          =   285
                     Left            =   5475
                     TabIndex        =   121
                     Top             =   285
                     Width           =   240
                  End
                  Begin VB.Label LblStartValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   165
                     Left            =   2550
                     TabIndex        =   120
                     Top             =   285
                     Width           =   720
                  End
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
                     Height          =   285
                     Index           =   37
                     Left            =   3285
                     TabIndex        =   119
                     Top             =   285
                     Width           =   1110
                  End
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   675
                  Index           =   14
                  Left            =   90
                  TabIndex        =   134
                  TabStop         =   0   'False
                  Top             =   3450
                  Width           =   20115
                  _cx             =   35481
                  _cy             =   1191
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
                  Begin VB.CheckBox XPChkPayType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‘Ìﬂ« "
                     Height          =   495
                     Index           =   2
                     Left            =   11820
                     TabIndex        =   135
                     Top             =   0
                     Width           =   915
                  End
                  Begin ImpulseButton.ISButton CmdCheque 
                     Height          =   510
                     Left            =   3690
                     TabIndex        =   136
                     Top             =   0
                     Width           =   1485
                     _ExtentX        =   2619
                     _ExtentY        =   900
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
                  Begin MSDataListLib.DataCombo Dcbanks 
                     Height          =   315
                     Left            =   13320
                     TabIndex        =   182
                     Top             =   0
                     Width           =   1500
                     _ExtentX        =   2646
                     _ExtentY        =   556
                     _Version        =   393216
                     ListField       =   "6"
                     BoundColumn     =   ""
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label Label2 
                     BackStyle       =   0  'Transparent
                     Caption         =   "«·»‰ﬂ"
                     Height          =   315
                     Left            =   15060
                     TabIndex        =   183
                     Top             =   0
                     Width           =   420
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   375
                     Index           =   19
                     Left            =   8370
                     TabIndex        =   140
                     Top             =   105
                     Width           =   1275
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
                     Height          =   495
                     Index           =   17
                     Left            =   9765
                     TabIndex        =   139
                     Top             =   105
                     Width           =   1260
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
                     Height          =   495
                     Index           =   16
                     Left            =   6270
                     TabIndex        =   138
                     Top             =   105
                     Width           =   1860
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   495
                     Index           =   18
                     Left            =   5175
                     TabIndex        =   137
                     Top             =   105
                     Width           =   1065
                  End
               End
               Begin VSFlex8UCtl.VSFlexGrid FgCheques 
                  Height          =   3300
                  Left            =   90
                  TabIndex        =   141
                  Top             =   4185
                  Width           =   17355
                  _cx             =   30612
                  _cy             =   5821
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
                  FormatString    =   $"FrmSaleBillCompose.frx":3E42
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
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   630
         Index           =   9
         Left            =   -105
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   15
         Width           =   17415
         _cx             =   30718
         _cy             =   1111
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
         Caption         =   "›« Ê—… „»Ì⁄«  „Ã„⁄Â"
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
         Begin VB.CommandButton Command2 
            Caption         =   " ÕÊÌ· «·Ï «–‰ ’—›"
            Height          =   330
            Left            =   7575
            Style           =   1  'Graphical
            TabIndex        =   194
            Top             =   240
            Visible         =   0   'False
            Width           =   4485
         End
         Begin VB.TextBox oldtxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   8340
            RightToLeft     =   -1  'True
            TabIndex        =   184
            Top             =   0
            Visible         =   0   'False
            Width           =   1230
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   285
            Left            =   7320
            TabIndex        =   71
            Top             =   0
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   285
            Left            =   6705
            TabIndex        =   70
            Top             =   0
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   11205
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   0
            Visible         =   0   'False
            Width           =   1950
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   360
            Index           =   0
            Left            =   2715
            TabIndex        =   25
            Top             =   30
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   635
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
            ButtonImage     =   "FrmSaleBillCompose.frx":3F77
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
            Height          =   360
            Index           =   3
            Left            =   1530
            TabIndex        =   26
            Top             =   30
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   635
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
            ButtonImage     =   "FrmSaleBillCompose.frx":4311
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
            Height          =   360
            Index           =   1
            Left            =   4080
            TabIndex        =   27
            Top             =   30
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   635
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
            ButtonImage     =   "FrmSaleBillCompose.frx":46AB
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
            Height          =   360
            Index           =   2
            Left            =   60
            TabIndex        =   28
            Top             =   30
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   635
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
            ButtonImage     =   "FrmSaleBillCompose.frx":4A45
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
            Height          =   360
            Left            =   10350
            TabIndex        =   40
            Top             =   120
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   635
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
            ButtonImage     =   "FrmSaleBillCompose.frx":4DDF
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdRetruns 
            Height          =   360
            Left            =   4755
            TabIndex        =   41
            Top             =   0
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   635
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
            ButtonImage     =   "FrmSaleBillCompose.frx":5179
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdInfo 
            Height          =   630
            Left            =   5910
            TabIndex        =   79
            Top             =   -120
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   1111
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
            ButtonImage     =   "FrmSaleBillCompose.frx":5713
            ButtonImageHover=   "FrmSaleBillCompose.frx":63ED
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
            Height          =   195
            Index           =   64
            Left            =   6765
            RightToLeft     =   -1  'True
            TabIndex        =   185
            Top             =   0
            Width           =   7050
         End
         Begin VB.Label LblShortcutKeys 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "ÃœÌœ F12 Or Enter ,  ⁄œÌ· F11 , Õ›Ÿ F10 ,  —«Ã⁄ F9 ,Õ–› F8 ,»ÕÀ F3 "
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
            Left            =   180
            TabIndex        =   42
            Top             =   405
            Width           =   10155
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   915
         Index           =   3
         Left            =   0
         TabIndex        =   159
         TabStop         =   0   'False
         Top             =   7665
         Width           =   17295
         _cx             =   30506
         _cy             =   1614
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
         Begin VB.Frame Frame7 
            Height          =   495
            Left            =   13050
            RightToLeft     =   -1  'True
            TabIndex        =   340
            Top             =   360
            Width           =   3465
            Begin VB.OptionButton optCommissionType 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄„Ê·… œ«Œ·Ì…"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   0
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   342
               Top             =   120
               Value           =   -1  'True
               Width           =   1695
            End
            Begin VB.OptionButton optCommissionType 
               Alignment       =   1  'Right Justify
               Caption         =   "⁄„Ê·… Œ«—ÃÌ…"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Index           =   1
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   341
               Top             =   120
               Width           =   1455
            End
         End
         Begin VB.TextBox TxtVATYou 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF0000&
            Enabled         =   0   'False
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
            Left            =   4725
            RightToLeft     =   -1  'True
            TabIndex        =   332
            Top             =   0
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.TextBox TXTFactoryExpenses 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF0000&
            Enabled         =   0   'False
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
            Left            =   10440
            RightToLeft     =   -1  'True
            TabIndex        =   325
            Top             =   0
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox TxtGVAT 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF0000&
            Enabled         =   0   'False
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
            Left            =   2385
            RightToLeft     =   -1  'True
            TabIndex        =   323
            Top             =   0
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.TextBox XPTxtSum 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Height          =   375
            Left            =   6705
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   160
            TabStop         =   0   'False
            Top             =   30
            Visible         =   0   'False
            Width           =   330
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   4335
            TabIndex        =   161
            Top             =   510
            Width           =   4320
            _ExtentX        =   7620
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.TextBox TxtNetValueComm 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            Enabled         =   0   'False
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
            Left            =   8670
            RightToLeft     =   -1  'True
            TabIndex        =   327
            Text            =   "5"
            Top             =   0
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label LblFinalView 
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
            TabIndex        =   339
            Top             =   0
            Width           =   1485
         End
         Begin VB.Label TXTFactoryExpensesView 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """#,###.##"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
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
            Left            =   10440
            TabIndex        =   338
            Top             =   0
            Width           =   1260
         End
         Begin VB.Label TxtGVATView 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """#,###.##"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
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
            Left            =   2370
            TabIndex        =   337
            Top             =   0
            Width           =   1275
         End
         Begin VB.Label TxtVATYouView 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """#,###.##"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
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
            Left            =   4740
            TabIndex        =   336
            Top             =   0
            Width           =   915
         End
         Begin VB.Label TxtNetValueCommView 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """#,###.##"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
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
            Left            =   8655
            TabIndex        =   335
            Top             =   0
            Width           =   1170
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰”»…«·VAT"
            Height          =   180
            Index           =   100
            Left            =   5625
            RightToLeft     =   -1  'True
            TabIndex        =   333
            Top             =   120
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ê·…"
            Height          =   195
            Index           =   87
            Left            =   9810
            TabIndex        =   328
            Top             =   0
            Width           =   615
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·„’—Ê›« "
            Height          =   195
            Index           =   90
            Left            =   11670
            TabIndex        =   326
            Top             =   0
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·VAT"
            Height          =   180
            Index           =   91
            Left            =   3765
            RightToLeft     =   -1  'True
            TabIndex        =   324
            Top             =   120
            Width           =   825
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«Ã„«·Ì"
            Height          =   330
            Index           =   71
            Left            =   1500
            RightToLeft     =   -1  'True
            TabIndex        =   219
            Top             =   105
            Width           =   750
         End
         Begin VB.Label LblFinal 
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
            TabIndex        =   218
            Top             =   0
            Width           =   1485
         End
         Begin VB.Label lblInstComm 
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
            Left            =   2265
            TabIndex        =   206
            Top             =   480
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·«÷«›« "
            Height          =   315
            Index           =   69
            Left            =   3480
            TabIndex        =   205
            Top             =   600
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.Label LblDiscountsTotalView 
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
            Left            =   13050
            TabIndex        =   193
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label LblTotalView 
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
            Left            =   6690
            TabIndex        =   192
            Top             =   0
            Width           =   1305
         End
         Begin VB.Label LblTotalAllView 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """#,###.##"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
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
            Left            =   15255
            TabIndex        =   191
            Top             =   0
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·ﬂ„ÌÂ"
            Height          =   195
            Index           =   63
            Left            =   11790
            TabIndex        =   174
            Top             =   585
            Width           =   1080
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
            Left            =   10440
            TabIndex        =   173
            Top             =   480
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·≈Ã„«·Ï"
            Height          =   285
            Index           =   3
            Left            =   16545
            RightToLeft     =   -1  'True
            TabIndex        =   172
            Top             =   75
            Width           =   690
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "/"
            Height          =   285
            Index           =   0
            Left            =   825
            RightToLeft     =   -1  'True
            TabIndex        =   171
            Top             =   555
            Width           =   195
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”Ã·"
            Height          =   285
            Index           =   2
            Left            =   1515
            RightToLeft     =   -1  'True
            TabIndex        =   170
            Top             =   555
            Width           =   795
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   285
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   169
            Top             =   555
            Width           =   435
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   285
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   168
            Top             =   555
            Width           =   555
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„” Œœ„"
            Height          =   330
            Index           =   1
            Left            =   8835
            RightToLeft     =   -1  'True
            TabIndex        =   167
            Top             =   555
            Width           =   600
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
            Left            =   6690
            RightToLeft     =   -1  'True
            TabIndex        =   166
            Top             =   -90
            Width           =   1200
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«›Ì"
            Height          =   285
            Index           =   49
            Left            =   7845
            RightToLeft     =   -1  'True
            TabIndex        =   165
            Top             =   75
            Width           =   780
         End
         Begin VB.Label LblTotalAll 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """#,###.##"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
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
            Left            =   15255
            RightToLeft     =   -1  'True
            TabIndex        =   164
            Top             =   -90
            Width           =   1140
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œ’Ê„« "
            Height          =   285
            Index           =   50
            Left            =   14400
            RightToLeft     =   -1  'True
            TabIndex        =   163
            Top             =   120
            Width           =   750
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
            Left            =   13050
            RightToLeft     =   -1  'True
            TabIndex        =   162
            Top             =   -90
            Width           =   1365
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   540
         Index           =   1
         Left            =   15
         TabIndex        =   270
         TabStop         =   0   'False
         Top             =   8610
         Width           =   17295
         _cx             =   30506
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
            Height          =   540
            Index           =   0
            Left            =   15405
            TabIndex        =   271
            Top             =   0
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   953
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
            Height          =   540
            Index           =   1
            Left            =   13305
            TabIndex        =   272
            Top             =   0
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   953
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
            Height          =   540
            Index           =   2
            Left            =   11460
            TabIndex        =   273
            Top             =   0
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   953
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
            Height          =   540
            Index           =   3
            Left            =   9690
            TabIndex        =   274
            Top             =   0
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   953
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
            Height          =   540
            Index           =   4
            Left            =   7875
            TabIndex        =   275
            Top             =   0
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   953
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
            Height          =   540
            Index           =   5
            Left            =   5865
            TabIndex        =   276
            Top             =   0
            Width           =   1605
            _ExtentX        =   2831
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
            Height          =   540
            Index           =   6
            Left            =   30
            TabIndex        =   277
            Top             =   0
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   953
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
            Height          =   540
            Index           =   7
            Left            =   4005
            TabIndex        =   278
            Top             =   0
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   953
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
            Height          =   540
            Left            =   1965
            TabIndex        =   279
            Top             =   0
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   953
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
         Begin VB.Label LBLGross 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """#,###.##"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   270
            Left            =   0
            TabIndex        =   280
            Top             =   0
            Visible         =   0   'False
            Width           =   1470
         End
      End
   End
End
Attribute VB_Name = "frmsalebillCompose"
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
Public invoiceSerach As Boolean
Public BolPrint As Boolean
Dim mNetValueComm As Long
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
Dim first_run As Boolean
Dim bank_account As String
Dim general_noteid As Long

Dim RsNotesGeneral As ADODB.Recordset
Dim CurrentVoucherNo As String
Dim CurrentVoucherSerialNo As String

Dim DateChanged As Boolean
Dim TxtNoteSerial1V As String

Function CuurentLogdata(Optional Currentmode As String)
   
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·›« Ê—…   " & TxtNoteSerial1.Text & CHR(13) & " «· «—ÌŒ " & XPDtbBill.value & CHR(13) & " «·Œ“Ì‰… " & DcboBox.Text & CHR(13) & " «·„Œ“‰  " & DCboStoreName.Text & CHR(13) & "  «·⁄„Ì· / «·„Ê—œ   " & DBCboClientName.Text & CHR(13) & "‰Ê⁄ «·”‰œ " & DCDocTypes & CHR(13) & "ÿ—Ìﬁ… «·œ›⁄ " & CboPayMentType & CHR(13) & "‰Ê⁄ «·Œ’„ " & XPCboDiscountType & CHR(13) & "ﬁÌ„… «·Œ’„ " & XPTxtDiscountVal & CHR(13) & "  «·«” Õﬁ«ﬁ " & DtpDelayDate & CHR(13) & " «·⁄„·Â " & DcCurrency & CHR(13) & "—ﬁ„ «·ﬁÌœ " & TxtNoteSerial & CHR(13) & "—ﬁ„ «·ÿ·»Ì… " & TXTOrDer_no
                     
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Bill No " & TxtNoteSerial1.Text & CHR(13) & " Date " & XPDtbBill.value & CHR(13) & " Box " & DcboBox.Text & CHR(13) & " Store  " & DCboStoreName.Text & CHR(13) & " Supplier/Cuxtomer" & DBCboClientName.Text & CHR(13) & "Doc Type" & DCDocTypes & CHR(13) & "Payment Type" & CboPayMentType & CHR(13) & "Discount Type  " & XPCboDiscountType & CHR(13) & " Discount Vaalue   " & XPTxtDiscountVal & CHR(13) & "Due Date " & DtpDelayDate & CHR(13) & " Currency " & DcCurrency & CHR(13) & " GE NO" & TxtNoteSerial & CHR(13) & "Order No " & TXTOrDer_no
                           
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", , val(TxtNoteSerial), val(TxtNoteSerial1)
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", , val(TxtNoteSerial), val(TxtNoteSerial1)
    End If
    
End Function

Function CheckBillType() As Integer
    ' ›Ê„ » ŒœÌœ Â· «·ﬁ« Ê—… «’‰«› «„ Œœ„«  «„ „Ã„⁄ «’‰«ﬁ ÊŒœ„« 
    Dim DblTempItemsGoodType As Double
    Dim DblTempItemsServiceType As Double

    DblTempItemsGoodType = NewGrid.GetItemsTotal(ItemsGoodType)
    DblTempItemsServiceType = NewGrid.GetItemsTotal(ItemsServiceType)

    If DblTempItemsGoodType = 0 And DblTempItemsServiceType > 0 Then  'Œœ„« 
        CheckBillType = 0
    ElseIf DblTempItemsServiceType > 0 And DblTempItemsGoodType > 0 Then ' Ê ·’‰«›   'Œœ„« 
        CheckBillType = 1
    ElseIf DblTempItemsServiceType = 0 And DblTempItemsGoodType > 0 Then 'Ê ·’‰«›   '
        CheckBillType = 2
      
    End If

End Function

Function CheckAccounts() As Boolean
 
End Function

Private Sub BillBasedOn_Click(Index As Integer)

    Select Case Index

        Case 1

            If BillBasedOn(1).value = True Then
                
                FillVoucherGrid
                GRID1.Enabled = True
            End If

        Case 2

            If BillBasedOn(2).value = True Then
                
                ' FillOrderGrid
                GRID2.Enabled = True
            End If

    End Select

End Sub

Private Sub CboPayMentType_Change()
    On Error GoTo ErrTrap

    'If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
    If CboPayMentType.ListIndex = 0 Then '‰ﬁœÌ
        XPChkPayType(0).Enabled = False
        XPChkPayType(1).Enabled = False
        XPChkPayType(2).Enabled = False
        XPChkPayType(0).value = Checked
        XPChkPayType(1).value = Unchecked
        XPChkPayType(2).value = Unchecked
        XPTxtValue(0).Text = XPTxtSum.Text
        XPTxtValue(1).Text = ""
        DcboBox.Enabled = True
      '  Frame1.Visible = True
        DCPaymentNet.Enabled = True
    Else
        XPChkPayType(0).Enabled = True
        XPChkPayType(1).Enabled = True
        XPChkPayType(2).Enabled = True
        XPChkPayType(0).value = Unchecked
        XPChkPayType(1).value = Checked
        XPChkPayType(2).value = Unchecked
        XPTxtValue(1).Text = XPTxtSum.Text
        XPTxtValue(0).Text = ""
        DcboBox.BoundText = ""
        DcboBox.Enabled = False
     '   Frame1.Visible = False
        DCPaymentNet.Enabled = False
    End If

    'End If
    Exit Sub
ErrTrap:
End Sub

Private Sub CboPayMentType_Click()

    If CboPayMentType.ListIndex = 0 Then
        DCPaymentNet.BoundText = 1
    Else
        DCPaymentNet.Text = ""
    End If

    CboPayMentType_Change
 
End Sub

Private Sub CheckBox1_Click()
TxtSuplCode.Enabled = False
DcbSuppler.Enabled = False
If CheckBox1.value = vbChecked Then
TxtSuplCode.Enabled = True
DcbSuppler.Enabled = True
Else
DcbSuppler.BoundText = 0
TxtSuplCode.Text = ""
End If
End Sub

Private Sub ChecVAT_Click()
  Dim i As Integer
If Me.TxtModFlg.Text <> "R" Then
    If ChecVAT.value = vbChecked Then

        With Me.GRID2
 
            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Select")) = True
            Next i

        End With

    Else

        With Me.GRID2

            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Select")) = False
            Next i

        End With

    End If
    GRID2_Click
    ReLineGrid
    End If
End Sub

Private Sub ChkInstall_Click()

    If ChkInstall.value = vbChecked Then
        Me.CmdINSTALLMENT.Enabled = True
        XPTxtValue(1).Text = LblTotal.Caption
    Else
        Me.CmdINSTALLMENT.Enabled = False

        With Me.FgInstallments
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
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
        TxtTaxAddValue.Text = ""
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
        TxtTaxServiceValue.Text = ""
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
        TxtTaxStampValue.Text = ""
        TxtTaxStampValue.Enabled = False
        lbl(41).Enabled = False
        lbl(48).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Function CloseIssueVoucher()
    On Error Resume Next
    Dim i As Integer
    Dim sql As String
 
    If BillBasedOn(1).value = False Then Exit Function

    With GRID1

        For i = 1 To .Rows - 1
     
            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
           
                sql = "update transactions set closed=1" & ",nots=" & val(Me.XPTxtBillID.Text) & ",nots2=" & Me.TxtNoteSerial1.Text & " where  Transaction_ID= " & val(.TextMatrix(i, .ColIndex("Transaction_ID")))
            Else
                sql = "update transactions set closed=0 ,nots='' ,nots2='' where  Transaction_ID=" & val(.TextMatrix(i, .ColIndex("Transaction_ID")))
               
            End If
       
            Cn.Execute sql
 
        Next
       
    End With
       
End Function

Function DeleteLinkTOIssueVoucher()
    On Error Resume Next
    Dim i As Integer
    Dim sql As String
 
    If BillBasedOn(1).value = False Then Exit Function

    With GRID1

        For i = 1 To .Rows - 1
     
            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then

                sql = "update transactions set closed=0 ,nots='' ,nots2='' where  Transaction_ID=" & val(.TextMatrix(i, .ColIndex("Transaction_ID"))) ' & "nots=" & "" & "nots2=" & ""
               
            End If
       
            Cn.Execute sql
 
        Next
       
    End With
       
End Function

Private Sub Cmd_Click(Index As Integer)
    Dim AskOption As Boolean
    Dim intDef As Integer
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTest As ADODB.Recordset
    Dim RsOptions As ADODB.Recordset
    BolPrint = True

    ' On Error GoTo ErrTrap
    If Voucher_coding(val(dcBranch.BoundText), XPDtbBill.value, 7, 170, 21) = "" Then
        TxtNoteSerial1.locked = False
    Else
        TxtNoteSerial1.locked = True
 
    End If
 
    Select Case Index

        Case 0
            
            
          '  mNetValueComm = IIf(val(TxtValueComm) = 0, 5, val(TxtValueComm))
            
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
                        Msg = Msg & CHR(13) & ""
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        Exit Sub
                    End If
                End If
            End If
        
            clear_all Me
            ClearNotes
            CboPayMentType.ListIndex = 1
            fillExpensesFactoryGrid
            
            TxtModFlg.Text = "N"
            XPTxtBillID.Text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            SetDefaults
            NewGrid.GridDefaultValue 1
            Me.DCboUserName.BoundText = user_id
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultClient", 2)
            DBCboClientName.BoundText = intDef
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSaleStore", 1)
            DCboStoreName.BoundText = intDef
             Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Fg_Journal.Rows = 2
            Fg_Journal.Enabled = True
            GRID2.Clear flexClearScrollable, flexClearEverything
            GRID2.Rows = 2
            GRID2.Enabled = True
            
            Set RsOptions = New ADODB.Recordset
            RsOptions.Open "tbloptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable

            If Not (RsOptions.BOF Or RsOptions.EOF) Then
                Me.DcboBox.BoundText = IIf(IsNull(RsOptions("SalesBoxID").value), "", RsOptions("SalesBoxID").value)
            End If

            XPTab301.CurrTab = 0
            '------------------
            '        Me.XPDtbBill.SetFocus
            '   customer_screen.Show
            '--------------------
        
            DcCurrency.BoundText = 1
        
            Me.dcBranch.BoundText = Current_branch
            Dim dstore As Integer
            Dim dBox As Integer
            Dim usertype As Integer
            Dim EmpID As Integer
            Dim userbranchid As Integer
            'GetBranchData branch_id, dstore, dBox
                 
            GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID
     
            If usertype <> 0 Then 'admin
                dcBranch.Enabled = False
                DcboBox.Enabled = False
                DCboStoreName.Enabled = True
                DcboEmp.Enabled = True
          
                Me.dcBranch.BoundText = userbranchid
                Me.DCboStoreName.BoundText = dstore
                Me.DcboBox.BoundText = dBox
                Me.DcboEmp.BoundText = EmpID
            Else
                dcBranch.Enabled = True
                DcboBox.Enabled = True
                DCboStoreName.Enabled = True
                DcboEmp.Enabled = True
                Me.dcBranch.BoundText = ""
                Me.DCboStoreName.BoundText = ""
                Me.DcboBox.BoundText = ""
                Me.DcboEmp.BoundText = ""

            End If
          
            BillBasedOn(0).value = True
 
            If Current_branch = 0 Then
                'branch_id = my_branch
                Me.dcBranch.BoundText = Current_branch
            End If
 
            BillBasedOn(1).Enabled = True
            DCboItemsCode.SetFocus
            CboPOSBillType.ListIndex = 0
            LblStableID.Caption = -1
            lblSessionD.Caption = -1
            Command2.Caption = ""
            DcbTypComm.ListIndex = 1
            If mNetValueComm = 0 Then TxtValueComm = 5 Else TxtValueComm = mNetValueComm
            optCommissionType(0) = True
            FillGridAuto
            TxtCustCode.SetFocus
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
                   
            '           If SystemOptions.usertype = UserNormal Then
            '    Msg = "·Ì” ·ﬂ Õﬁ  ⁄œÌ· ›Ï «·›Ê« Ì—"
            '    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.Title
            '    Exit Sub
            'End If
        
            'If AvailableDeal = True Then
            '«·√ﬁ”«ÿ «·„”Ã·… ⁄·Ï «·›« Ê—…
            If XPTxtValue(1).Tag <> "" Then
                StrSQL = "select * From InstallMent where NoteID=" & XPTxtValue(1).Tag
                Set RsTest = New ADODB.Recordset
                RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (RsTest.EOF Or RsTest.BOF) Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "·ﬁœ  „  ﬁ”Ìÿ «·ﬁÌ„ «·¬Ã·… ⁄·Ï Â–Â «·›« Ê—…" & CHR(13)
                        Msg = Msg + " ⁄œÌ· «·›« Ê—… ”ÌƒœÌ ≈·Ï Õ–› Â–Â «·√ﬁ”«ÿ" & CHR(13)
                        Msg = Msg + "Â·  —€» ›Ì  ⁄œÌ· Â–Â «·›« Ê—…ø"
                    Else
                
                        Msg = "this bill was linked With Installment and edit will Delete this Installment Confirm Edit?" & CHR(13)
                    End If

                    If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbNo Then
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
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "·ﬁœ  „  Õ’Ì· »⁄÷ «·√ﬁ”«ÿ «·„”Ã·… ⁄·Ï Â–Â «·›« Ê—…" & CHR(13)
                        Msg = Msg + "Ê·« Ì„ﬂ‰  ⁄œÌ· »Ì«‰« Â«" & CHR(13)
                        Msg = Msg + "≈–« ﬂ‰   —€» ›Ì  ⁄œÌ· »Ì«‰«  Â–Â «·›« Ê—…" & CHR(13)
                        Msg = Msg + "ÌÃ» Õ–› ⁄„·Ì«  «· Õ’Ì· «·Œ«’… »Â«"
                    Else
                        Msg = "Some premiums were collected  on this bill You Must delete Collected  premiums according to this bill First" & CHR(13)
                    End If

                    MsgBox Msg, vbOKOnly + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If
            End If

            '⁄„·Ì«  «·’Ì«‰… «·„— »ÿ… »«·›« Ê—…
            StrSQL = "select * From MaintenanceJuncTransaction where Transaction_ID=" & Trim(XPTxtBillID.Text)
            Set RsTest = New ADODB.Recordset
            RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsTest.EOF Or RsTest.BOF) Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·ﬁœ  „ ≈Ã—«¡ »⁄÷ ⁄„·Ì«  «·’Ì«‰… ⁄·Ï Â–Â «·›« Ê—… Ê·« Ì„ﬂ‰  ⁄œÌ·Â«"
                    Msg = Msg + "≈–« ﬂ‰   —€» ›Ì  ⁄œÌ· »Ì«‰«  Â–Â «·›« Ê—…" & CHR(13)
                    Msg = Msg + "ÌÃ» Õ–› ⁄„·Ì«  «·’Ì«‰… «·Œ«’… »Â«"
                Else
                    Msg = "this Bill Linked with Maintenance Operation You must Delete This Operation First"
            
                End If

                MsgBox Msg, vbOKOnly + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If

            Me.Retrive val(Me.XPTxtBillID.Text)
             
            TxtModFlg.Text = "E"
            DateChanged = False
            Me.DCboUserName.BoundText = user_id
              Fg_Journal.Rows = Fg_Journal.Rows + 1
            Fg_Journal.Enabled = True
            CuurentLogdata

            '    txtorder_no_Change
            'End If
        Case 2

            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Departement"
                Else
                    Msg = "Õœœ «·›—⁄ «Ê·« "
                End If
              
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                dcBranch.SetFocus
                 SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            If Trim(DCPaymentNet.BoundText) = "" And CboPayMentType.ListIndex = 0 Then
             '   If SystemOptions.UserInterface = EnglishInterface Then
             '       Msg = "SpecifY Payment Value"
             '   Else
                    Msg = "Õœœ ÿ—Ìﬁ… «·œ›⁄  «Ê·« "
             '   End If
             '
             '   MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'            '    DCPaymentNet.SetFocus
             '   SendKeys "{F4}"
             '   Screen.MousePointer = vbDefault
             '   Exit Sub
            End If

            If CboPayMentType.ListIndex = 0 Then

                If val(TxtRemainValue.Text) < 0 Then
                    If SystemOptions.UserInterface = EnglishInterface Then
                        Msg = "Enter Correct Payed Value"
                    Else
                        Msg = "  ﬁÌ„Â «·„œ›Ê⁄ €Ì— ’ÕÌÕÂ "
                    End If
             
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
  
                    Exit Sub
                End If
            End If

            If CboPayMentType.ListIndex = 1 And XPChkPayType(0).value = Unchecked And XPChkPayType(2).value = Unchecked Then
                XPTxtValue(1).Text = LblTotal.Caption
            End If
 
            Set RsNotesGeneral = New ADODB.Recordset
      '      RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
       StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      
      
            
            '    my_branch = Me.Dcbranch.BoundText
      
            '    If Me.TxtModFlg.text = "E" Then
             
            '     TxtInvID
            '     End If

    
             my_branch = val(Me.dcBranch.BoundText)
               Dim Account_Code_dynamic82 As String
         If val(TxtNetValueComm.Text) <> 0 Then
         If optCommissionType(0).value = True Then
                            Account_Code_dynamic82 = get_account_code_branch(150, my_branch)
                            If Account_Code_dynamic82 = "NO account" Then
                                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«» ⁄„Ê·…  „»Ì⁄«   Õ  «· ’—Ì›", vbCritical
                                                            Else
                                                                MsgBox "Please Select  Account", vbCritical
                                                            End If
                            
                                                GoTo ErrTrap
                              End If
          ElseIf optCommissionType(1).value = True Then
                                    Account_Code_dynamic82 = get_account_code_branch(157, my_branch)
                            If Account_Code_dynamic82 = "NO account" Then
                                                            If SystemOptions.UserInterface = ArabicInterface Then
                                                                MsgBox "·„ Ì „  ÕœÌœ Õ”«» ⁄„Ê·… Œ«—ÃÌ…    ", vbCritical
                                                            Else
                                                                MsgBox "Please Select  Account", vbCritical
                                                            End If
                            
                                                GoTo ErrTrap
                              End If
          End If
                              
          End If

If SumVAT() > 0 Then
If GetValueAddedAccount(XPDtbBill.value, , , 1, 21) = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›… ··„»Ì⁄« "
Else
MsgBox "Value added account not specified"
End If
Exit Sub
End If
End If
If val(Me.TxtGVAT.Text) > 0 Then
If GetValueAddedAccount(XPDtbBill.value, , , 1, 22) = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›… ··„‘ —Ì« "
Else
MsgBox "Value added account not specified"
End If
Exit Sub
End If
End If

     If CheckGrid() = True Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "Ì—ÃÏ «Œ Ì«— «·Õ”«» ›Ì «·„’—Ê›« "
     Else
     MsgBox "Please Select Account"
     End If
     Exit Sub
     End If
    If val(DcbFarm.BoundText) = 0 Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "Ì—ÃÏ «Œ Ì«— «·„Ê—œ"
     Else
     MsgBox "Please Select Vendor"
     End If
     DcbFarm.SetFocus
     Exit Sub
     End If
    If val(DCboStoreName.BoundText) = 0 Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "Ì—ÃÏ «Œ Ì«— «·„Œ“‰"
     Else
     MsgBox "Please Select Store"
     End If
     DCboStoreName.SetFocus
     Exit Sub
     End If
        If val(DcboEmp.BoundText) = 0 Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "Ì—ÃÏ «Œ Ì«— «·„‰œÊ»"
     Else
     MsgBox "Please Select Supervisor"
     End If
     DcboEmp.SetFocus
     Exit Sub
     End If
     Dim i As Integer
     If CheckCustomerInGrid(i) = True Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox i & "Ì—ÃÏ «Œ Ì«— «·⁄„Ì· ›Ì «·”ÿ— —ﬁ„ "
     Else
     MsgBox "Please Select Customer in line " & i
     End If
     Exit Sub
     End If

     
            SaveData
            ' Unload customer_screen
            '  Load customer_screen
            '  customer_screen.Show
        
        Case 3
            Dim mListIndex As Long
            mListIndex = DcbTypComm.ListIndex
          '  mNetValueComm = IIf(val(TxtValueComm) = 0, 5, val(TxtValueComm))
            
            Undo
            If mListIndex = 1 Then
                DcbTypComm.ListIndex = mListIndex
              '  TxtValueComm = mNetValueComm
            End If
        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            If SystemOptions.usertype = UserNormal Then
                Msg = "·Ì” ·ﬂ Õﬁ Õ–› ›Ï «·›Ê« Ì—"
                MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If
   
            Del_TransAction

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            If m_FrmSearch Is Nothing Then

                Set m_FrmSearch = New FrmBuySearch
                m_FrmSearch.DealingForm = InvoiceTransactionCompose
                m_FrmSearch.Caption = "«·»ÕÀ ⁄‰ ⁄„·Ì… »Ì⁄"
                Set m_FrmSearch.RetrunFrm = Me
                m_FrmSearch.show vbModal

            Else
                Msg = "Â‰«ﬂ ‘«‘… »ÕÀ Œ«’… »‘«‘… ›« Ê—… «·»Ì⁄ «·Õ«·Ì…"
                Msg = Msg & CHR(13) & "Ÿ«Â—… «„«„ﬂ ›⁄·«...·«Ì„ﬂ‰ ⁄—÷ «ﬂÀ— „‰ ‘«‘… »ÕÀ ·ﬂ· ‘«‘… ›« Ê—…"
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                m_FrmSearch.ZOrder 0
                'm_FrmSearch.SetFocus
            End If

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If Me.XPTxtBillID.Text = "" Then
                Msg = "·« ÊÃœ ›Ê« Ì— ·Ì „ ÿ»«⁄ Â«"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If

            AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowMe", False)

            If AskOption = False Then
'
                FrmSallReportOptions.show vbModal

                If FrmSallReportOptions.UserCanceled = True Then
                    Unload FrmSallReportOptions
                    Exit Sub
                End If

                Unload FrmSallReportOptions

            End If

            PrintReport
        
        Case 8
        
            PrintReport , 1, LblAdvPayment, LblInstallCount, LblInstallTotal, LblFirstInstallDate
        
        Case 9
        
            PrintReport , 2, LblAdvPayment, LblInstallCount, LblPrecenValue, LblFirstInstallDate
        
        Case 6
            Unload Me

        Case 10
            ShowGL_cc TxtNoteSerial.Text, , 200, val(Me.TXTNoteID.Text)

            'ShowGL_cc TxtNoteSerial.text, , 200
        Case 11
            Dim sql As String
            sql = "update TblTransactionsApproval set Approval_Date=" & SQLDate(Date, True) & " where Transaction_Type=170 AND Transactionid=" & val(Me.XPTxtBillID.Text) & " AND UserID=" & user_id
            Cn.Execute sql

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ «·«⁄ „«œ", vbInformation
            Else
                MsgBox "Approved", vbInformation
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub
Function CheckGrid() As Boolean
Dim i As Integer
CheckGrid = False
     With Fg_Journal
          If .Rows > 1 Then
          For i = 1 To .Rows - 1
          If .TextMatrix(i, .ColIndex("AccountName")) <> "" And val(.TextMatrix(i, .ColIndex("Price"))) Then
          If .TextMatrix(i, .ColIndex("Accountcode2")) = "" Then
          CheckGrid = True
          Exit Function
          End If
          End If
          Next i
          End If
          End With
End Function
Function LoadSigns(Transactionid As Double, Transaction_Type As Double)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim i As Integer
 
    sql = " SELECT     UserID, Transactionid, Approval_Date, Transaction_Type"
    sql = sql & " from dbo.TblTransactionsApproval"
    sql = sql & "  WHERE     (Transactionid = " & Transactionid & ") AND (NOT (Approval_Date IS NULL)) AND (Transaction_Type = " & Transaction_Type & ") "
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
    For i = 0 To 5
        DBPix202(i).ImageClear
    Next i
 
    If rs.RecordCount > 0 Then

        If Not IsNull(rs("UserID").value) Then
             
            For i = 0 To rs.RecordCount - 1

                If Dir(App.path & "\images\sign\sign" & val(rs("UserID").value) & ".JPG") <> "" Then
                    DBPix202(i).ImageLoadFile (App.path & "\images\sign\sign" & val(rs("UserID").value) & ".JPG")
                End If
           
                rs.MoveNext
            Next i

        End If
    End If

End Function

Function Retrive_orders_data(Transaction_ID As Integer)
    Dim StrSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim row_count As Integer
    Dim Num As Integer
    'StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & _
    '"ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    'StrSQL = StrSQL + " where Transaction_ID=" & Transaction_ID

    StrSQL = "SELECT     *, dbo.Transactions.Transaction_Date AS Transaction_Datesub, dbo.Transactions.NoteSerial1 AS Expr3 ,dbo.Transaction_Details.Vat,dbo.Transaction_Details.Vatyo"
    StrSQL = StrSQL + " FROM         dbo.TblItems INNER JOIN"
    StrSQL = StrSQL + "   dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID INNER JOIN"
    StrSQL = StrSQL + "  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID INNER JOIN"
    StrSQL = StrSQL + "  dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
    StrSQL = StrSQL + " WHERE     (dbo.Transaction_Details.Transaction_ID = " & Transaction_ID & ")"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.Text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        row_count = FG.Rows
    
        If FG.TextMatrix(row_count - 1, FG.ColIndex("Code")) = "" Then
            row_count = row_count - 1
        End If
     
        FG.Rows = RsDetails.RecordCount + row_count

        For Num = row_count To FG.Rows - 1 'RsDetails.RecordCount
     
            FG.TextMatrix(Num, FG.ColIndex("Transaction_Date")) = IIf(IsNull(RsDetails("Transaction_Datesub")), "", (RsDetails("Transaction_Datesub").value))
            FG.TextMatrix(Num, FG.ColIndex("order_no")) = GRID2.TextMatrix(GRID2.Row, GRID2.ColIndex("order_no"))
 
            FG.TextMatrix(Num, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("NoteSerial1")), "", (RsDetails("NoteSerial1").value))
            FG.TextMatrix(Num, FG.ColIndex("OrderArrivalDate")) = IIf(IsNull(RsDetails("OrderArrivalDate")), "", (RsDetails("OrderArrivalDate").value))
            FG.TextMatrix(Num, FG.ColIndex("FoxyNo")) = IIf(IsNull(RsDetails("FoxyNo")), "", (RsDetails("FoxyNo").value))
        
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
        
            '   FG.TextMatrix(Num, FG.ColIndex("Count")) = items_qty_not_recieved_in_order(FG.TextMatrix(Num, FG.ColIndex("Code")), FG.TextMatrix(Num, FG.ColIndex("order_no")))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))
        
            '   FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("SallingPrice")), "", (RsDetails("SallingPrice").value))
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showprice")), "", (RsDetails("showprice").value))
        
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        If ChLoadVAT.value = vbChecked Then
            FG.TextMatrix(Num, FG.ColIndex("Vat")) = IIf(IsNull(RsDetails("Vat")), 0, (RsDetails("Vat").value))
            FG.TextMatrix(Num, FG.ColIndex("Vatyo")) = IIf(IsNull(RsDetails("Vatyo")), 0, (RsDetails("Vatyo").value))
        End If
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        
            RsDetails.MoveNext
            ' Debug.Print Num
            ' If FG.Rows > 10 Then
            '     If Num = 8 Then FG.Refresh
            ' End If
        Next Num

    End If

End Function
 
Private Sub Cmd1_Click()
    On Error Resume Next

    If TxtNoteSerial1.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
         
            MsgBox "·«»œ „‰ «Õ Ì«—  ”‰œ  «Ê·«": Exit Sub
        Else
            MsgBox "Select Voucher Firstly": Exit Sub
        End If
 
    End If

    Unload imaged
    imaged.show

    If SystemOptions.UserInterface = EnglishInterface Then

        imaged.Label9.Caption = "Sales Invoice  #"
        imaged.Caption = "Sales Invoice  Attachment"
        imaged.txtopeation_type = "1001"
        imaged.SUBJECT_NO = TxtNoteSerial1.Text
        imaged.Label6.Caption = "Sales Invoice  #"
    Else

        imaged.Label9.Caption = "„—›ﬁ«  ›« Ê—… «·»Ì⁄ —ﬁ„"
        imaged.Caption = "„—›ﬁ«  ›« Ê—… «·»Ì⁄ —ﬁ„    "
        imaged.txtopeation_type = "1001"
        imaged.SUBJECT_NO = TxtNoteSerial1.Text
        imaged.Label6.Caption = "„—›ﬁ«  ›« Ê—… «·»Ì⁄ —ﬁ„"

    End If

    imaged.Adodc1.CommandType = adCmdText
    imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type ='1001'  and subject_no='" & TxtNoteSerial1.Text & "'"
    imaged.Adodc1.Refresh

    If imaged.Adodc1.Recordset.RecordCount > 0 Then

        imaged.DBPix201.Visible = True
    Else
        imaged.DBPix201.Visible = False
    End If

End Sub

Private Sub cmdAdd_Click()
'TxtCustCode.SetFocus
'TxtCustCode = ""
'DcCustmer.Text = ""
End Sub

Private Sub CmdCash_Click(Index As Integer)

    Select Case Index

        Case 0

        Case 1
    End Select

End Sub

Private Sub cmdCommand1_Click()
End Sub

Private Sub CmdDel_Click()
RemoveFactoryExpenses
End Sub
Function RemoveFactoryExpenses()

    With Me.Fg_Journal
  
        If .Row <= 0 Then Exit Function
        .RemoveItem .Row
    End With

      ReLineGrid

End Function
Private Sub CmdHelp_Click()
'    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
'    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd

Command9_Click

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
    XPTxtValue(1).Text = LblTotal.Caption
    'If Me.TxtModFlg = "R" Then Exit Sub

    If XPTxtValue(1).Text = "" Then
        Msg = "ÌÃ»  ÕœÌœ «·ﬁÌ„… «·¬Ã·… ﬁ»·  ”ÃÌ· «·√ﬁ”«ÿ"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

        If XPTxtValue(1).Enabled = True Then
            XPTxtValue(1).SetFocus
        End If

        Exit Sub
    End If

    Load FrmInstallMent
    Set FrmInstallMent.Frm = Me

    With FrmInstallMent

        If Me.TxtModFlg.Text = "E" Then
            .Tag = "E"
        
            .Retrive val(XPTxtValue(1).Tag)
            .Txt(1).Text = XPTxtValue(1).Text
        ElseIf Me.TxtModFlg.Text = "R" Then
  
            .Tag = "R"
            .Retrive val(XPTxtValue(1).Tag)
            '              .OptInt(1).value = True
            '.Txt(7).text = 1
            '.Txt(5).text = 12
        Else
            .Tag = "N"
            .Txt(1).Text = XPTxtValue(1).Text
            Me.CmdINSTALLMENT.Enabled = True
    
            .LblNoteID.Caption = XPTxtSerial(1).Text
            .CboPrecenType.ListIndex = val(Me.LblPrecenType.Tag)
            .Txt(3).Text = val(LblPrecenValue.Caption)
            .Txt(5).Text = val(LblInstallCount.Caption)
            .OptInt(1).value = True
            .Txt(7).Text = 1
            .Txt(5).Text = 12

            If IsDate(Me.LblFirstInstallDate.Caption) Then
                .Dtp_First.value = Me.LblFirstInstallDate.Caption
            End If

            '        .Txt(7).text = Val(LblInstallSeprator.Caption)
            If val(LblInstallmentType.Tag) = 0 Then
                '        .OptInt(0).value = True
            ElseIf val(LblInstallmentType.Tag) = 1 Then
                '        .OptInt(1).value = True
            ElseIf val(LblInstallmentType.Tag) = 2 Then
                '        .OptInt(2).value = True
            End If

            With .FG
                .Rows = Me.FgInstallments.Rows

                For i = 1 To Me.FgInstallments.Rows - 1
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
    ShowRelatedNotes val(Me.XPTxtBillID.Text), 1
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
    ShowRelatedTransactions val(Me.XPTxtBillID.Text), 1
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
    Dim RowNum As Integer
    Dim Frm As Form
    Dim Msg As String
    On Error GoTo ErrTrap

    If Text1.Text <> "" Then
        Msg = " „  ÕÊÌ· Â–… «·›« Ê—… „‰ ﬁ»· Ê·« Ì„ﬂ‰  ÕÊÌ·Â« „—… «Œ—Ï  …  "
        MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
    Set Frm = New FrmOut

    With Frm

        .Convert
        '    .XPTxtBillID.Text = XPTxtBillID.Text
        .XPDtbBill.value = XPDtbBill.value
        .DBCboClientName.BoundText = DBCboClientName.BoundText
        .DCboStoreName.BoundText = DCboStoreName.BoundText
        .Text2.Text = TxtTransSerial.Text
        .CboPayMentType.ListIndex = CboPayMentType.ListIndex

        For RowNum = 1 To FG.Rows - 1

            If .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Code")) <> "" Then
                .FG.Rows = .FG.Rows + 1
            End If

            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Name")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Name")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Code")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Code")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Code")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("ItemCase")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("HaveSerial")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Count")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Count")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Count")))
            ' .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Price")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Price")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Price")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(.FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod)
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("DiscountType")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")))
            Dim StrSQL As String
            Dim RsUnit As New ADODB.Recordset
            StrSQL = "SELECT TOP 100 PERCENT dbo.TblItemsUnits.UnitID, dbo.TblUnites.UnitName, dbo.Transactions.Transaction_Serial,dbo.Transactions.Transaction_Type FROM dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN dbo.TblUnites INNER JOIN dbo.TblItemsUnits ON dbo.TblUnites.UnitID = dbo.TblItemsUnits.UnitID ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID AND dbo.Transaction_Details.Item_ID = dbo.TblItemsUnits.ItemID WHERE (dbo.Transactions.Transaction_Serial = '" & TxtTransSerial & "') AND (dbo.Transactions.Transaction_Type = 21) AND (dbo.TblItemsUnits.ItemID = " & FG.TextMatrix(RowNum, FG.ColIndex("Code")) & ") ORDER BY dbo.TblItemsUnits.SecOrder"
            Set RsUnit = New ADODB.Recordset
            RsUnit.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
            .FG.Cell(flexcpData, RowNum, .FG.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitID")), "", (RsUnit("UnitID").value))
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

Function CREATE_VOUCHER_GE(Transaction_ID As Long, TxtNoteSerialV As String, TxtNoteSerial1V As String, general_noteid As Long, BranchID As Integer)
    Dim usedaccount As Integer
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
    Dim TOTAL_COST As Double
    Dim LngCurItemID As Integer
    Dim LngUnitID As Long
    Dim UnitFactor As Double

    With FG

        For i = 1 To FG.Rows - 1

            If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" And val(FG.TextMatrix(i, FG.ColIndex("ItemType"))) <> 1 Then
                LngCurItemID = val(FG.TextMatrix(i, FG.ColIndex("Code")))
                LngUnitID = val(FG.Cell(flexcpData, i, FG.ColIndex("UnitID")))
            
                GetUnitNoOfItems LngCurItemID, LngUnitID, UnitFactor
                    
                '           TOTAL_COST = TOTAL_COST + (FG.TextMatrix(i, FG.ColIndex("Count")) * ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, , LngUnitID))
                                                                                                                                                      
                TOTAL_COST = TOTAL_COST + val(FG.TextMatrix(i, FG.ColIndex("ItemCostPrice"))) * FG.TextMatrix(i, FG.ColIndex("Count"))
            End If

        Next i

    End With

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·œ«∆‰
    SngTemp = NewGrid.GetItemsTotal(ItemsGoodType)
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
            StrTempDes = "”‰œ ’—› —ﬁ„ " & Me.TxtTransSerial.Text
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
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
                StrTempDes = "”‰œ    ’—› —ﬁ„ " & TxtNoteSerial1V
            Else
                StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V
            End If

            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value As Single

            With FG

                For i = 1 To FG.Rows - 1

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

                        line_value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "”‰œ    ’—› —ﬁ„ " & TxtNoteSerial1V
                        Else
                            StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V
                        End If
            
                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

                Next i

            End With

        End If

        '«·ÿ—› «·„œÌ‰
        SngTemp = NewGrid.GetItemsTotal(ItemsGoodType)

        If TOTAL_COST > 0 Then
            If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Then

                Account_Code_dynamic = get_account_code_branch(1, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ  ﬂ·›… «·„»Ì⁄«  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
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
                        StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄« 
        
                    End If

                Else
                    StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄« 
                End If
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ    ’—› —ﬁ„ " & TxtNoteSerial1V
                Else
                    StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V
                End If
            
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
         
            ElseIf detect_inventory_work_type = 3 Then

                With FG

                    For i = 1 To FG.Rows - 1

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

                            line_value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
                            If SystemOptions.UserInterface = ArabicInterface Then
                                StrTempDes = "”‰œ    ’—› —ﬁ„ " & TxtNoteSerial1V
                            Else
                                StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V
                            End If
            
                            LngDevNO = LngDevNO + 1

                            If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                                GoTo ErrTrap
                            End If
    
                        End If

                    Next i

                End With

            End If
        End If
    End If

    Dim StrSQL  As String
    StrSQL = "UPDATE Transactions SET NOTS=" & Transaction_ID & " WHERE Transaction_ID=" & val(Me.XPTxtBillID.Text)
    Cn.Execute StrSQL
ErrTrap:
End Function

Private Function CheckCostForAllitems() As Boolean
    CheckCostForAllitems = True
End Function

Private Sub CreateIssueVoucher()
    'On Error GoTo errortrap
    'DeleteTransactiomsVoucher Val(Text1.text)

    If BillBasedOn(1).value = True Then Exit Sub

    If CheckBillType = 0 Then ' Œœ„« 
        Exit Sub
    ElseIf CheckBillType = 1 Then ' Ê«’‰«›  ' Œœ„« 

    ElseIf CheckBillType = 2 Then ' «’‰«›

    End If

    Dim i As Long
    Dim LngCurItemID As Integer
    Dim LngUnitID As Long
    Dim UnitFactor As Double

    '›Ì Õ«·… «·«‰ «Ã «·‰„ÿÌ
    If SystemOptions.TypicalProduction = True Then
        GoTo ll
    End If

    'GoTo ll
    '/////////////////////////////////////////////////////////////CACELED
    '      With Fg

    '                For i = 1 To Fg.Rows - 1

    '                            If Fg.TextMatrix(i, Fg.ColIndex("Code")) <> "" And val(Fg.TextMatrix(i, Fg.ColIndex("ItemType"))) <> 1 Then
    '                            LngCurItemID = val(Fg.TextMatrix(i, Fg.ColIndex("Code")))
    '                             LngUnitID = val(Fg.Cell(flexcpData, i, Fg.ColIndex("UnitID")))
    '                             GetUnitNoOfItems LngCurItemID, LngUnitID, UnitFactor
                     
    '                                          If ModItemCostPrice.GetCostItemPrice(Fg.TextMatrix(i, Fg.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Me.Text1.text), LngUnitID) = 0 Then
    '          If SystemOptions.UserInterface = ArabicInterface Then
    '               MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ  ﬂ·›Â «·»Ì⁄  Õ ·Ï  «—ÌŒ…  ·⁄œ„ ÊÃÊœ ﬂ„Ì… „‰… ›Ï Â–« «·„Œ“‰ ·« Ì„ﬂ‰ «‰‘«¡ ”‰œ «·’—› "
    '           Else
    '               MsgBox "Item in line no " & i & "Have No Qty "
    '           End If
                            
    '                                          Exit Sub
    '                                          End If
    '                            End If
    '                Next i
    '     End With
    '/////////////////////////////////////////////////////////////CACELED
     
ll:

    With Me.GRID1
        .Rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove
 
    End With

    Text1.Text = ""
    Text1_Change

    Dim groupAccount  As String

    If detect_inventory_work_type = 3 Then
   
        With FG

            For i = 1 To FG.Rows - 1

                If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
                
                    ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                    groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 0)

                    If groupAccount = "Error" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”⁄·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                        Else
                            MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                        End If

                        Exit Sub
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
        Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
        Text1.Text = Transaction_ID
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        Dim general_noteid As Long
        Dim RsNotesGeneral As ADODB.Recordset
        Dim TxtNoteSerialV As String
            
        my_branch = Me.dcBranch.BoundText

        If TxtNoteSerialV = "" Then
            If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
            Else
                       
                If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                Else
                    TxtNoteSerialV = Notes_coding(val(my_branch), XPDtbBill.value)
                End If
            End If
        End If
        
        If TxtNoteSerial1V = "" Then
            If Voucher_coding(val(my_branch), XPDtbBill.value, 10, 180, , 19) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ  ’—› ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
            Else
                       
                If Voucher_coding(val(my_branch), XPDtbBill.value, 10, 180, , 19) = "" Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ”‰œ  «·’—› ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                Else
                    TxtNoteSerial1V = Voucher_coding(val(my_branch), XPDtbBill.value, 10, 180, , 19)
                End If
            End If
        End If
             
        If SystemOptions.TypicalProduction = True Then
            TxtNoteSerialV = ""
 
        End If
 
        If Trim(CurrentVoucherNo) <> "" And DateChanged <> True Then
            TxtNoteSerialV = CurrentVoucherNo '—ﬁ„ «·ﬁÌœ
            TxtNoteSerial1V = Trim(CurrentVoucherSerialNo)
        End If

        Dim sql As String

        sql = "INSERT INTO  Transactions (Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,CusID,StoreID,UserID,Emp_ID,nots,nots2,NoteSerial,NoteSerial1,NoteId,BranchId,Closed)SELECT " & Transaction_ID & "," & MYTEXT & ",Transaction_Date,Transaction_Type = 19,CusID,StoreID,UserID,Emp_ID,nots=" & val(XPTxtBillID.Text) & ",nots2=" & TxtNoteSerial1.Text & " ,NoteSerial=' " & TxtNoteSerialV & "',NoteSerial1='" & TxtNoteSerial1V & "',NoteId=" & general_noteid & ",BranchId,1From Transactions Where  Transaction_ID =" & val(XPTxtBillID.Text) & " And Transaction_Type = 21"
        Cn.Execute sql
        '
        Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,ItemSize,UnitId,ShowQty,QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID,ProductionDate,ExpiryDate,LotNO)SELECT  costprice,guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , Quantity, costprice/ QtyBySmalltUnit ,ColorID,ItemSize, UnitId, ShowQty, QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID ,ProductionDate,ExpiryDate,LotNO From dbo.Transaction_Details Where SavedItemType=0 and   Transaction_ID = " & XPTxtBillID.Text
        Text1.Text = Transaction_ID
        'TxtIssueSerial.text = TxtNoteSerial1V
 
        StrSQL = "UPDATE Transactions SET NOTS=" & Transaction_ID & " WHERE Transaction_ID=" & val(Me.XPTxtBillID.Text)
        Cn.Execute StrSQL

        If SystemOptions.TypicalProduction = True Then
            Exit Sub
        End If

        'Create big notes
        Set RsNotesGeneral = New ADODB.Recordset
        RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
     StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText


        If Me.TxtModFlg.Text = "N" Then
    
        Else
        
            general_noteid = val(TXTNoteID.Text)
        End If
        
        RsNotesGeneral.AddNew
        RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
        general_noteid = RsNotesGeneral("NoteID").value
        TXTNoteID.Text = general_noteid
        RsNotesGeneral("Transaction_ID").value = Transaction_ID
        RsNotesGeneral("NoteDate").value = XPDtbBill.value
        RsNotesGeneral("NoteType").value = 180
        RsNotesGeneral("Note_Value").value = Null
        RsNotesGeneral("NoteSerial").value = IIf(Trim(TxtNoteSerialV) = "", Null, Trim(TxtNoteSerialV))
        RsNotesGeneral("NoteSerial1").value = IIf(Trim(TxtNoteSerial1V) = "", Null, Trim(TxtNoteSerial1V))
        RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
        RsNotesGeneral("numbering_type1").value = sand_numbering_type(10) '«–‰ wvt
        RsNotesGeneral("sanad_year").value = year(XPDtbBill.value)
        RsNotesGeneral("sanad_month").value = Month(XPDtbBill.value)
        RsNotesGeneral("branch_no").value = val(Me.dcBranch.BoundText)
        'RsNotes("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
        RsNotesGeneral.update
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

        CREATE_VOUCHER_GE Transaction_ID, TxtNoteSerialV, TxtNoteSerial1V, general_noteid, val(Me.dcBranch.BoundText)

    End If
 
    '
 
ErrTrap:

End Sub

Private Sub Command2_Click()

    If Me.TxtModFlg = "R" Then
        Cmd_Click (1)
        Cmd_Click (2)
        CreateIssueVoucher
    End If

End Sub

Private Sub Command3_Click()
    FrmSearchSerial.show vbModal
End Sub

Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial111.Text, , 200
End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.show vbModal
        FrmCustemerSearch.SearchType = 2
    End If

End Sub
 Function createVoucher()

Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
If SystemOptions.UserInterface = ArabicInterface Then
            des = "  ›« Ê—…  „“—⁄…  —ﬁ„ " & TxtNoteSerial1.Text
Else
            des = " Farm invoice  No " & TxtNoteSerial1.Text
End If
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
Dim sql As String
tablename = "Transactions"
Filedname = "Transaction_ID"
NoteSerial1 = val(XPTxtBillID.Text)
Notevalue = 0
 notytype = 9084
Notevalue = val(LblFinal.Caption)
BranchID = val(dcBranch.BoundText)
NoteDate = (XPDtbBill.value)
 
If Notevalue <> 0 Then
                              
                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des
                                              TxtNoteID111.Text = NoteID
                                              TxtNoteSerial111.Text = NoteSerial
                            
CREATE_VOUCHER_GE2 val(TxtNoteID111.Text), BranchID, user_id, NoteDate


 
updateNotesValueAndNobytext val(TxtNoteID111.Text)
     End If
    
  rs.Resync adAffectCurrent
End Function
Function SumVAT() As Double
Dim i As Integer
Dim SmValu As Double
SmValu = 0
With FG
For i = 1 To .Rows - 1
SmValu = SmValu + val(.TextMatrix(i, .ColIndex("Vat")))
Next i
End With
SumVAT = SmValu
End Function
Public Function CREATE_VOUCHER_GE2(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)

 Dim Notevalue As Single
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempCustomerCode As String
    Dim AccountVATCreit  As String
    Dim Note_Value As Double
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim Note_Value2 As Double
    Dim i As Integer
 Dim valuee As Double
 Dim StrSQL As String
 Dim Account_code As String
 
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
       LngDevNO = 0
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„Ì‰
    my_branch = BranchID
''''

LngDevNO = 0

         GetValueAddedAccount XPDtbBill.value, AccountVATCreit, , 1, 22
         If val(TxtGVAT.Text) > 0 And AccountVATCreit <> "" Then
             valuee = val(TxtGVAT.Text)
             LngDevNO = LngDevNO + 1
             
           '  AccountVATCreit = get_account_code_branch(145, my_branch)
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, valuee, 0, StrTempDes & "Õ”«» «·ﬁÌ„… «·„÷«›… ··„‘ —Ì«  ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
         End If
       With Fg_Journal

            For i = 1 To .Rows - 1

                If .TextMatrix(i, .ColIndex("AccountCode")) <> "" And val(.TextMatrix(i, .ColIndex("value"))) <> 0 And .TextMatrix(i, .ColIndex("AccountCode2")) <> "" Then
            
                    If SystemOptions.UserInterface = ArabicInterface Then
                        StrTempDes = " »‰«¡ ⁄·Ï ›« Ê—… „“—⁄…    —ﬁ„ " & Me.TxtNoteSerial1.Text
                    Else
                        StrTempDes = " Based On Farm  Invoice NO:" & Me.TxtNoteSerial1.Text
                    End If
            
                    LngDevNO = LngDevNO + 1
                    Account_code = .TextMatrix(i, .ColIndex("AccountCode"))
                    Note_Value = val(.TextMatrix(i, .ColIndex("value"))) * val(txt_Currency_rate.Text)
                    Note_Value2 = val(.TextMatrix(i, .ColIndex("value")))
                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_code, Note_Value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(XPTxtBillID.Text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    LngDevNO = LngDevNO + 1
                    Account_code = .TextMatrix(i, .ColIndex("AccountCode2"))
                     If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_code, Note_Value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(XPTxtBillID.Text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    
                End If
        
            Next

        End With
     '/////////
            With FG

            For i = 1 To .Rows - 1

                If .TextMatrix(i, .ColIndex("Code")) <> "" And val(.TextMatrix(i, .ColIndex("Valu"))) <> 0 And .TextMatrix(i, .ColIndex("CusID2")) <> "" Then
            
                    If SystemOptions.UserInterface = ArabicInterface Then
                        StrTempDes = " »‰«¡ ⁄·Ï ›« Ê—… „“—⁄…    —ﬁ„ " & Me.TxtNoteSerial1.Text
                    Else
                        StrTempDes = " Based On Farm  Invoice NO:" & Me.TxtNoteSerial1.Text
                    End If
            
                    LngDevNO = LngDevNO + 1
                    Account_code = GetMyAccountCode("TblCustemers", "CusID", val(.TextMatrix(i, .ColIndex("CusID2"))))
                    Note_Value = (val(.TextMatrix(i, .ColIndex("Valu"))) + val(.TextMatrix(i, .ColIndex("Vat")))) * val(txt_Currency_rate.Text)
                    Note_Value2 = val(.TextMatrix(i, .ColIndex("Valu"))) + val(.TextMatrix(i, .ColIndex("Vat")))
                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_code, Note_Value, 0, StrTempDes & "Õ”«» «·⁄„Ì·", general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(XPTxtBillID.Text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    
                End If
        
            Next

        End With
'''/////////////////
     valuee = SumVAT
         If valuee > 0 Then
             LngDevNO = LngDevNO + 1
             GetValueAddedAccount XPDtbBill.value, , AccountVATCreit, 1, 21
            ' AccountVATCreit = get_account_code_branch(145, my_branch)
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, valuee, 1, StrTempDes & "Õ”«»  «·ﬁÌ„… «·„÷«›… „»Ì⁄«  ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
         End If
         With Fg_Journal

            For i = 1 To .Rows - 1

                If .TextMatrix(i, .ColIndex("AccountCode")) <> "" And val(.TextMatrix(i, .ColIndex("value"))) <> 0 And .TextMatrix(i, .ColIndex("AccountCode2")) <> "" Then
            
                    If SystemOptions.UserInterface = ArabicInterface Then
                        StrTempDes = " »‰«¡ ⁄·Ï ›« Ê—… „»Ì⁄«  „Ã„⁄…  —ﬁ„ " & Me.TxtNoteSerial1.Text
                    Else
                        StrTempDes = " Based On Sales  Invoice NO:" & Me.TxtNoteSerial1.Text
                    End If
            
                    LngDevNO = LngDevNO + 1
                    
                    Note_Value = val(.TextMatrix(i, .ColIndex("value"))) * val(txt_Currency_rate.Text)
                    Note_Value2 = val(.TextMatrix(i, .ColIndex("value")))
                    Account_code = .TextMatrix(i, .ColIndex("AccountCode"))
                     If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_code, Note_Value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(XPTxtBillID.Text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                    
                End If
        
            Next

        End With


         
      If val(TxtNetValueComm.Text) > 0 Then
      valuee = val(TxtNetValueComm.Text)
       LngDevNO = LngDevNO + 1
       If optCommissionType(0).value = True Then
         AccountVATCreit = get_account_code_branch(150, my_branch)
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, valuee, 1, StrTempDes & "Õ”«» «·⁄„Ê·… œ«Œ·Ì…", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
       ElseIf optCommissionType(1).value = True Then
            AccountVATCreit = get_account_code_branch(157, my_branch)
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, valuee, 1, StrTempDes & "Õ”«» «·⁄„Ê·… Œ«—ÃÌ…", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
                GoTo ErrTrap
            End If
       End If
     End If
     
     If SystemOptions.UserInterface = ArabicInterface Then
                        StrTempDes = " »‰«¡ ⁄·Ï ›« Ê—…  „“—⁄…  —ﬁ„ " & Me.TxtNoteSerial1.Text
                    Else
                        StrTempDes = " Based On Farm  Invoice NO:" & Me.TxtNoteSerial1.Text
                    End If
                    
    valuee = val(LblFinal.Caption)
    If valuee <> 0 Then
      LngDevNO = LngDevNO + 1
             AccountVATCreit = GetMyAccountCode("TblCustemers", "CusID", val(Me.DcbFarm.BoundText))
            If valuee < 0 Then
                If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, Abs(valuee), 0, StrTempDes & "Õ”«» «·„Ê—œ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then GoTo ErrTrap
            Else
                 If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, valuee, 1, StrTempDes & "Õ”«» «·„Ê—œ", general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then GoTo ErrTrap
                
            End If
            
    End If


ErrTrap:
End Function
Private Sub DBCboClientName_MouseUp(Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    Y As Single)

    If Button = vbRightButton Then
        mdifrmmain.MnuCusTools.Tag = Me.DBCboClientName.BoundText
        Me.PopupMenu mdifrmmain.MnuCusTools
    End If

End Sub

Private Sub DcbFarm_Change()
DcbFarm_Click (0)
End Sub

Private Sub DcbFarm_Click(Area As Integer)
Dim Fullcode As String
Dim StrSQL As String
Dim RsTemp As ADODB.Recordset
Set RsTemp = New ADODB.Recordset
GetCustomersDetail val(DcbFarm.BoundText), , Fullcode, 2
TxtSuplCode3.Text = Fullcode
StrSQL = " Select * From TblCustemers Where CusID=" & val(DcbFarm.BoundText)
RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If RsTemp.RecordCount > 0 Then
Me.TxtVATNO.Text = IIf(IsNull(RsTemp("VATNO").value), "", RsTemp("VATNO").value)
Else
Me.TxtVATNO.Text = ""
End If
End Sub

Private Sub DcboEmp_Change()
    Dim StoreId As Integer
 'If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
         If val(Me.DcboEmp.BoundText) = 0 Then Exit Sub
           Me.TxtEmployeeID.Text = get_EMPLOYEE_Data(val(Me.DcboEmp.BoundText), "Fullcode")
        'DCEmP.text = DCEmP.text
'End If
 If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
  StoreId = get_StoreBYPurchasePerson(val(Me.DcboEmp.BoundText))
 If StoreId <> 0 Then
 DCboStoreName.BoundText = StoreId
 End If
 
 End If
End Sub

Private Sub DCboItemsCode_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF3 Then

        Load FrmItemSearch
        FrmItemSearch.RetrunType = 5
        FrmItemSearch.show vbModal

    End If

    If KeyCode = vbKeyF9 Then

        FrmSearchSerial.XPTxtCode.Text = DCboItemsCode.Text
        FrmSearchSerial.show
        FrmSearchSerial.Cmd_Click (0)

    End If

End Sub

Private Sub DCboItemsName_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF3 Then
        
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 5
        FrmItemSearch.show vbModal
    End If

    If KeyCode = vbKeyF9 Then
                    
        FrmSearchSerial.XPTxtCode.Text = DCboItemsCode.Text
        FrmSearchSerial.show
        FrmSearchSerial.Cmd_Click (0)
                    
    End If

End Sub

Private Sub Dcbranch_Change()

    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
        Dcombos.GetDocTypebyid Me.DCDocTypes, 21, val(Me.dcBranch.BoundText)
    End If

    If dcBranch.BoundText = "" Then TxtNoteSerial1.locked = True: Exit Sub

    If Voucher_coding(val(dcBranch.BoundText), XPDtbBill.value, 7, 170, 21) = "" Then
        TxtNoteSerial1.locked = False
    Else
        TxtNoteSerial1.locked = True
 
    End If
 
End Sub

Private Sub Dcbranch_Click(Area As Integer)
    Dcbranch_Change

    If Voucher_coding(val(val(Me.dcBranch.BoundText)), XPDtbBill.value, 7, 170, , 21) = "" Then Exit Sub
    TxtNoteSerial.Text = ""
    TxtNoteSerial1.Text = ""
End Sub

Private Sub DcbSuppler_Change()
DcbSuppler_Click (0)
End Sub

Private Sub DcbSuppler_Click(Area As Integer)
Dim Fullcode As String
Dim StrSQL As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
GetCustomersDetail val(DcbSuppler.BoundText), , Fullcode, 2
TxtSuplCode.Text = Fullcode
End Sub

Private Sub DcbTypComm_Change()
CalCulComm
End Sub
Sub CalCulComm()
If val(DcbTypComm.ListIndex) = 0 Then
TxtNetValueComm.Text = val(TxtValueComm.Text)
lbl(86).Caption = DcbTypComm.Text
ElseIf val(DcbTypComm.ListIndex) = 1 Then
If val(TxtSAlGValue.Text) <> 0 Then
TxtNetValueComm.Text = (val(TxtValueComm.Text) * val(LblTotalAll.Caption)) / 100
Else

TxtNetValueComm.Text = 0

TxtNetValueComm.Text = mNetValueComm
End If
Else
TxtNetValueComm.Text = 0
End If
lbl(86).Caption = DcbTypComm.Text
LblTotal.Caption = val(LblTotalAll.Caption) - val(LblDiscountsTotal.Caption) - val(TXTFactoryExpenses.Text) - val(TxtNetValueComm.Text)
TxtGVAT.Text = (val(txtVatYou.Text) * val(LblTotal.Caption)) / 100
LblFinal.Caption = val(TxtGVAT.Text) + val(LblTotal.Caption)
End Sub

Private Sub DcbTypComm_Click()
DcbTypComm_Change
End Sub

Private Sub DcCurrency_Change()

    If Me.TxtModFlg.Text = "" Or Me.TxtModFlg.Text = "R" Then Exit Sub
    If Me.DcCurrency.BoundText <> "" Then
        txt_Currency_rate.Text = get_currency_rate(Me.DcCurrency.BoundText)
    Else
        txt_Currency_rate.Text = 1
    End If

End Sub

Private Sub DcCurrency_Click(Area As Integer)
    DcCurrency_Change
End Sub

Private Sub DcCustmer_KeyPress(KeyAscii As Integer)
' If KeyAscii = vbKeyReturn Then
'        DCboItemsCode.SetFocus
' End If
End Sub

Private Sub DcCustmer_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 605
        FrmCustemerSearch.show vbModal
    End If
End Sub

Private Sub DCPaymentNet_Click(Area As Integer)

    If val(DCPaymentNet.BoundText) <> 1 Then
        DcboBox.Text = ""
    End If

End Sub

Function FillOrderGrid(Optional BegineDate As Date, Optional EndDate As Date)
    ' ⁄»∆… «Ê«„— «·‘—«¡ Ê «·»Ì⁄

    With Me.GRID2
        .Rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove

        '    .AutoSize 0, .Cols - 1, False
    End With

    Dim i As Integer
    Dim RsExp As ADODB.Recordset
    Dim My_SQL As String

    Set RsExp = New ADODB.Recordset
    My_SQL = "SELECT Transactions.NoteSerial1 , dbo.Transactions.closed,dbo.Transactions.Transaction_ID,dbo.Transactions.order_no , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where  Transaction_Type=21  AND CLOSED= 0 "
    'and   dbo.TblCustemers.CusID=" & val(DBCboClientName.BoundText)
 
    My_SQL = My_SQL + " and Transaction_Date >=" & SQLDate(BegineDate, True) & ""
 
    My_SQL = My_SQL + " and Transaction_Date <=" & SQLDate(EndDate, True) & ""
    If val(DcbSuppler.BoundText) <> 0 Then
    My_SQL = My_SQL + " and SupplerID =" & val(DcbSuppler.BoundText) & ""
    End If
If Me.TxtModFlg.Text = "R" Or Me.TxtModFlg.Text = "" Then
My_SQL = My_SQL + " and TransGorupID =" & val(XPTxtBillID.Text) & ""
End If
    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.GRID2
        .Rows = 2
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .Rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .Rows - 1
            If Me.TxtModFlg.Text = "" Or Me.TxtModFlg.Text = "R" Then
                    .TextMatrix(i, .ColIndex("Select")) = 1
                Else
                .TextMatrix(i, .ColIndex("Select")) = 0
             End If
                  '  IIf(IsNull(RsExp.Fields("closed").value), _
                  '    0, RsExp.Fields("closed").value)
         
                '          .TextMatrix(i, .ColIndex("order_no")) = IIf(IsNull(RsExp.Fields("order_no").value), _
                           "", RsExp.Fields("order_no").value)
               
                .TextMatrix(i, .ColIndex("order_no")) = IIf(IsNull(RsExp.Fields("NoteSerial1").value), "", RsExp.Fields("NoteSerial1").value)
               
                .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(RsExp.Fields("Transaction_Date").value), "", RsExp.Fields("Transaction_Date").value)
           
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(RsExp.Fields("CusName").value), "", RsExp.Fields("CusName").value)
                .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(RsExp.Fields("Transaction_ID").value), "", RsExp.Fields("Transaction_ID").value)

                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300
    End With

    GRID2.Visible = True

End Function

Function FillVoucherGrid()
    ' ⁄»∆…  ”‰œ«   «·’—›
    On Error Resume Next

    With Me.GRID1
        .Rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove

        '    .AutoSize 0, .Cols - 1, False
    End With

    Dim i As Integer
    Dim RsExp As ADODB.Recordset
    Dim My_SQL As String

    Set RsExp = New ADODB.Recordset

    'My_SQL = "SELECT dbo.Transactions.closed,dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_ID,dbo.Transactions.order_no , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where  Transaction_Type=19   and   dbo.TblCustemers.CusID=" & Val(DBCboClientName.BoundText)
    If BillBasedOn(0).value = True Then
        My_SQL = "SELECT dbo.Transactions.closed,dbo.Transactions.NoteSerial1,dbo.Transactions.NoteSerial, dbo.Transactions.Transaction_ID,dbo.Transactions.order_no , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where   ( (nots='" & Me.XPTxtBillID.Text & "' and  Transaction_Type=19)   and  (dbo.TblCustemers.CusID=" & val(DBCboClientName.BoundText) & ")) "
    Else
        My_SQL = "SELECT dbo.Transactions.closed,dbo.Transactions.NoteSerial1,dbo.Transactions.NoteSerial, dbo.Transactions.Transaction_ID,dbo.Transactions.order_no , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where   ( (nots='" & Me.XPTxtBillID.Text & "' and  Transaction_Type=19) or ( Transaction_Type=19   and  closed =0 ) and  (dbo.TblCustemers.CusID=" & val(DBCboClientName.BoundText) & ")) "
    End If
 
    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.GRID1
        .Rows = 2
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .Rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .Rows - 1
             
                .TextMatrix(i, .ColIndex("Select")) = IIf(IsNull(RsExp.Fields("closed").value), 0, RsExp.Fields("closed").value)
              
                .TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(RsExp.Fields("NoteSerial").value), "", RsExp.Fields("NoteSerial").value)
              
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(RsExp.Fields("NoteSerial1").value), "", RsExp.Fields("NoteSerial1").value)
               
                .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(RsExp.Fields("Transaction_Date").value), "", RsExp.Fields("Transaction_Date").value)
           
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(RsExp.Fields("CusName").value), "", RsExp.Fields("CusName").value)
                .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(RsExp.Fields("Transaction_ID").value), "", RsExp.Fields("Transaction_ID").value)

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("P1")) = "⁄—÷ «·”‰œ"
                    .TextMatrix(i, .ColIndex("P2")) = "ÿ»«⁄Â  «·ﬁÌœ"
                Else
                    .TextMatrix(i, .ColIndex("P1")) = "View VCHR"
                    .TextMatrix(i, .ColIndex("P2")) = "Print GE"
                End If

                RsExp.MoveNext
            Next
       
        End If
         
        .RowHeight(-1) = 300
    End With

    RsExp.Close
    GRID1.Visible = True

End Function

Private Sub DTPickerAccFrom_Change()
 
    ISButton1_Click
 
End Sub

Private Sub DTPickerAccTo_Change()
 
    ISButton1_Click
 
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

Public Sub Fg_Journal_AfterEdit(ByVal Row As Long, _
                                ByVal Col As Long)
 
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With Fg_Journal

        Select Case .ColKey(Col)
           Case "Account_Name2"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Accountcode2"), False, True)
                .TextMatrix(Row, .ColIndex("Accountcode2")) = StrAccountCode
 
            Case "AccountName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("ExpensesID")) = get_Expenses_id(StrAccountCode)
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "select * from Expenses_accounts where Account_Code='" & StrAccountCode & "'"
                Else
                    StrSQL = "select * from Expenses_accounts_eng where Account_Code='" & StrAccountCode & "'"
                End If
            
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     
                If rs.RecordCount > 0 Then
                    .TextMatrix(Row, .ColIndex("des")) = IIf(IsNull(rs("parent_account").value), "", rs("parent_account").value)
                Else
                    .TextMatrix(Row, .ColIndex("des")) = ""
                End If

            Case "value", "Price", "ChSameCurrncy"
                Dim sgl As String
           
                .TextMatrix(Row, .ColIndex("value")) = val(.TextMatrix(Row, .ColIndex("Price")))
              Me.TXTFactoryExpenses.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
                '    sgl = "update  marakes_taklefa_temp  set value=0 where  line_no=" & Val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                '     Cn.Execute sgl, , adExecuteNoRecords
        
                '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
                 Dim StrTempAccountCode As String
                 Dim mBoxID  As Integer, s As String
                 Dim rsDummy As New ADODB.Recordset
                 s = "Select BoxID From TblUsers Where UserId = " & user_id
                 rsDummy.Open s, Cn, adOpenStatic, adLockOptimistic, adCmdText
                 If Not rsDummy.EOF Then
                    
                    StrTempAccountCode = GetBoxAccount(val(rsDummy!BoxID & ""))
                    
                    StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", val(rsDummy!BoxID & ""))
                    
                    
                    'LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Accountcode2"), False, True)
                    .TextMatrix(Row, .ColIndex("Accountcode2")) = StrTempAccountCode
                    .TextMatrix(Row, .ColIndex("Account_Name2")) = Get_Account_name(, StrTempAccountCode)
                    
                 End If
                 
            '     Account_Name2
                 
           

          '  mBoxID = 2
          
            
        End Select

       Me.TXTFactoryExpenses.Text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))

        ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        'to Add new row if needed
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid
  
End Sub
Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter As Integer
Dim SumValue As Double
SumValue = 0
    With Fg_Journal

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
                                Dim sgl As String
 
                .TextMatrix(i, .ColIndex("value")) = val(.TextMatrix(i, .ColIndex("Price")))
                       
            End If
            If val(.TextMatrix(i, .ColIndex("Value"))) <> 0 Then
            SumValue = SumValue + val(.TextMatrix(i, .ColIndex("Value")))
           End If
        Next i
Me.TXTFactoryExpenses.Text = SumValue
    End With
'''//////////////
Dim SumVAT As Double
SumVAT = 0
SumValue = 0
    With FG
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("Code")) <> "" Then
             SumValue = SumValue + val(.TextMatrix(i, .ColIndex("Valu")))
             SumVAT = SumVAT + val(.TextMatrix(i, .ColIndex("Vat")))
            End If
        Next i
Me.TxtSAlGValue.Text = SumValue
If ChLoadVAT.value = vbChecked Then
Me.TxtGVAT.Text = SumVAT
Else
Me.TxtGVAT.Text = 0
End If
    End With
 CalCulComm
End Sub

Private Sub LblFinal_Change()
Me.LblFinalView.Caption = Format(val(LblFinal.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub TxtCustCode_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer

'    If KeyAscii = vbKeyReturn Then
'        GetTblCustemersCode TxtCustCode.Text, EmpID, , , , branch_id
'        DcCustmer.BoundText = EmpID
'        DcCustmer.SetFocus
'    End If

End Sub
Private Sub DcCustmer_Change()
DcCustmer_Click 0
End Sub

Private Sub DcCustmer_Click(Area As Integer)
  If DcCustmer.BoundText = "" Then TxtCustCode = ""
  If val(DcCustmer.BoundText) = 0 Then Exit Sub
  

    Dim EmpCode  As String
    GetTblCustemersCode , , DcCustmer.BoundText, EmpCode, , branch_id
    'Me.Text9.Text = EmpCode
    TxtCustCode = EmpCode
If Me.TxtModFlg.Text <> "R" Then

If val(DcCustmer.BoundText) <> 0 Then
'DBCboClientName.BoundText = DcCustmer.BoundText

'DBCboClientName_Click 0
'GetInformationCustomer (DcCustmer.BoundText)

End If
End If
End Sub

Private Sub TxtCustCode_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 604
        FrmCustemerSearch.show vbModal
    End If
End Sub

Private Sub TXTFactoryExpenses_Change()
TXTFactoryExpensesView.Caption = Format(val(TXTFactoryExpenses.Text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub TxtGVAT_Change()
TxtGVATView.Caption = Format(val(TxtGVAT.Text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub TxtNetValueComm_Change()
TxtNetValueCommView.Caption = Format(val(TxtNetValueComm.Text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub TxtSuplCode_KeyPress(KeyAscii As Integer)
    Dim CUSTID As Integer
    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSuplCode.Text, 2
        DcbSuppler.BoundText = CUSTID
    End If
End Sub

Private Sub Ele_KeyUp(Index As Integer, _
                      KeyCode As Integer, _
                      Shift As Integer)

    If Me.TxtModFlg.Text = "R" And Not (Me.ActiveControl Is TxtTransSerial) Then
        '        Cmd_Click (0)
    Else
        SendKeys "{TAB}"
    End If

End Sub

Sub FillGridAuto()
Dim i As Integer
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim sql As String
 Fg_Journal.Clear flexClearScrollable, flexClearEverything
   Fg_Journal.Rows = 2
 sql = "select * from Expenses_accounts where ComposeExpenses=1"
 rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If rs2.RecordCount > 0 Then
With Fg_Journal
.Rows = 1 + rs2.RecordCount
rs2.MoveFirst
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("LineNo")) = i
.TextMatrix(i, .ColIndex("Price")) = 0
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(rs2("Account_Name").value), "", rs2("Account_Name").value)
Else
.TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(rs2("Account_NameEng").value), "", rs2("Account_NameEng").value)
End If
.TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(rs2("Account_Code").value), "", rs2("Account_Code").value)
.TextMatrix(i, .ColIndex("ExpensesID")) = get_Expenses_id(.TextMatrix(i, .ColIndex("AccountCode")))
.TextMatrix(i, .ColIndex("LineNo1")) = setfoxy_Line
rs2.MoveNext
Next i
End With
End If
End Sub
Private Sub Fg_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)
                       '  NewGrid.Calculate 1, , , True
ReLineGrid
DcbTypComm_Change
    If Me.TxtModFlg <> "E" Then Exit Sub
    If val(Me.TxtNoteSerial.Text) = 0 Or val(Me.TxtNoteSerial1.Text) = 0 Then GoTo ll

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    If Col = FG.ColIndex("Code") Or Col = FG.ColIndex("Name") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 170
    ElseIf Col = FG.ColIndex("UnitID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("UnitID")), , , , , , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 170
    ElseIf Col = FG.ColIndex("Count") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , (FG.TextMatrix(Row, FG.ColIndex("Count"))), , , , , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 170
    ElseIf Col = FG.ColIndex("Price") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , (FG.TextMatrix(Row, FG.ColIndex("Price"))), , , , , , , val(Me.TxtNoteSerial), val(Me.TxtNoteSerial1), 170
    ElseIf Col = FG.ColIndex("ColorID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ColorID")), , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 170
    ElseIf Col = FG.ColIndex("ItemSize") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ItemSize")), , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 170
    ElseIf Col = FG.ColIndex("ClassId") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ClassId")), , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 170
    ElseIf Col = FG.ColIndex("DiscountType") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("DiscountType")), , Me.TxtNoteSerial, Me.TxtNoteSerial1, 170
    ElseIf Col = FG.ColIndex("DiscountVal") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , FG.TextMatrix(Row, FG.ColIndex("DiscountVal")), Me.TxtNoteSerial, Me.TxtNoteSerial1, 170

    End If

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
ll:
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

    'If first_run = True Then
    'Me.left = Me.left + 1420
    'Cmd_Click (0)
    'first_run = False
    'End If
    Ele(2).Enabled = True
End Sub

Private Sub Grid1_Click()

    With GRID1

        Select Case .Col

            Case 2
 
                With FG
                    .Clear flexClearScrollable, flexClearEverything
                    .Rows = 1
       
                End With
 
                fillVchr

            Case 7

                FrmOut.Retrive val(.TextMatrix(.Row, 1))

            Case 8
                ShowGL_cc val(.TextMatrix(.Row, .ColIndex("NoteSerial"))), , 200

        End Select

    End With

End Sub

Private Sub GRID2_Click()

    With FG
        .Clear flexClearScrollable, flexClearEverything
        .Rows = 1
       
    End With
 
    fillOrders
    NewGrid.Calculate 1, , , True
    showComm
   ' CalCulComm
    ReLineGrid
End Sub
Private Sub Fg_Journal_BeforeEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)

    With Fg_Journal

        If Row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel = True
            '  End If
        End If

        Select Case .ColKey(Col)

            Case "value"
               Cancel = True
              Case "Price"
                .ComboList = ""
              Case "NoteSerial"
                .ComboList = ""

            Case "des"
                .ComboList = ""
        
            Case "Order_No"
                .ComboList = ""
        
                '  Cancel = True
            
        End Select

    End With

End Sub
Function fillExpensesFactoryGrid()
    With Me.Fg_Journal
        .Rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove

        '    .AutoSize 0, .Cols - 1, False
    End With

    Dim i As Integer
    Dim RsExp As ADODB.Recordset
    Dim My_SQL As String

    Set RsExp = New ADODB.Recordset
My_SQL = " SELECT     dbo.TblSalesGExpepenses.id, dbo.TblSalesGExpepenses.Transaction_ID, dbo.TblSalesGExpepenses.[Value],"
My_SQL = My_SQL & "                       dbo.TblSalesGExpepenses.des, dbo.TblSalesGExpepenses.Price,"
My_SQL = My_SQL & "                      dbo.TblSalesGExpepenses.Accountcode, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng,"
My_SQL = My_SQL & "                      dbo.TblSalesGExpepenses.Accountcode2, ACCOUNTS_1.Account_Name AS Account_Name2, ACCOUNTS_1.Account_Serial AS Account_Serial2,"
My_SQL = My_SQL & "                      ACCOUNTS_1.Account_NameEng AS Account_NameE2 ,dbo.TblSalesGExpepenses.NoteSerial"
My_SQL = My_SQL & " FROM         dbo.TblSalesGExpepenses LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.ACCOUNTS ACCOUNTS_1 ON dbo.TblSalesGExpepenses.Accountcode2 = ACCOUNTS_1.Account_Code LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.ACCOUNTS ON dbo.TblSalesGExpepenses.Accountcode = dbo.ACCOUNTS.Account_Code"
My_SQL = My_SQL & " Where (dbo.TblSalesGExpepenses.Transaction_ID = " & val(XPTxtBillID.Text) & ")"
    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    Dim StrSQL  As String

    With Me.Fg_Journal
        .Rows = 1
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .Rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("LineNo")) = i
                .TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(RsExp.Fields("NoteSerial").value), "", RsExp.Fields("NoteSerial").value)
                .TextMatrix(i, .ColIndex("Accountcode2")) = IIf(IsNull(RsExp.Fields("Accountcode2").value), "", RsExp.Fields("Accountcode2").value)
                .TextMatrix(i, .ColIndex("Accountcode")) = IIf(IsNull(RsExp.Fields("Accountcode").value), "", RsExp.Fields("Accountcode").value)
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Account_Name2")) = IIf(IsNull(RsExp.Fields("Account_Name2").value), "", RsExp.Fields("Account_Name2").value)
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsExp.Fields("Account_Name").value), "", RsExp.Fields("Account_Name").value)
            Else
                .TextMatrix(i, .ColIndex("Account_Name2")) = IIf(IsNull(RsExp.Fields("Account_NameE2").value), "", RsExp.Fields("Account_NameE2").value)
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsExp.Fields("Account_NameEng").value), "", RsExp.Fields("Account_NameEng").value)
            End If
               
                .TextMatrix(i, .ColIndex("value")) = IIf(Not IsNumeric(RsExp.Fields("value").value), 0, RsExp.Fields("value").value)
                .TextMatrix(i, .ColIndex("Price")) = IIf(Not IsNumeric(RsExp.Fields("Price").value), 0, RsExp.Fields("Price").value)
               ' .TextMatrix(i, .ColIndex("ChSameCurrncy")) = IIf(Not IsNumeric(RsExp.Fields("ChSameCurrncy").value), 0, RsExp.Fields("ChSameCurrncy").value)
            
                .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsExp.Fields("des").value), "", RsExp.Fields("des").value)
                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300
    End With

    With Me.Fg_Journal
        Me.TXTFactoryExpenses.Text = .Aggregate(flexSTSum, .FixedRows - 1, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
    End With
 
End Function
Private Sub Fg_Journal_StartEdit(ByVal Row As Long, _
                                 ByVal Col As Long, _
                                 Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset

    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim StrComboList1 As String

    Dim Msg As String


    With Fg_Journal

        Select Case .ColKey(Col)
          Case "Account_Name2"
                If SystemOptions.UserInterface = EnglishInterface Then
                
                    StrSQL = "SELECT Account_Code, Account_NameEng  from ACCOUNTS    "
           
                        StrSQL = StrSQL + " And(last_account=1)"
                    
                    StrSQL = StrSQL & GetAccountByBarnchUser
                    StrSQL = StrSQL & GetAccountCodeHiding
                Else
                
                    StrSQL = "SELECT Account_Code, Account_Name from ACCOUNTS   "
                     StrSQL = StrSQL + " where (last_account=1)"

                   StrSQL = StrSQL & GetAccountByBarnchUser
                  StrSQL = StrSQL & GetAccountCodeHiding
                
                End If
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                StrComboList = .BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                StrComboList = .BuildComboList(rs, "Account_NameEng", "Account_Code")
                End If
                
                
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
            Case "AccountName"

                '      StrSQL = "select * from Expenses_accounts"
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "select * from Expenses_accounts where ComposeExpenses=1"
                Else
                    StrSQL = "select * from Expenses_accounts_eng ComposeExpenses=1 "
                End If
                 
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                'StrComboList = Fg_Journal.BuildComboList(rs, "Account_Name", "Account_Code")
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "Account_NameEng", "Account_Code")
                End If
            
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList

            Case "opr_fullcode"
                StrSQL = "  select fullcode,name from terms_operations "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList1 = Fg_Journal.BuildComboList(rs, "fullcode", "fullcode")

                If StrComboList1 <> "" Then
                    StrComboList1 = "|" & StrComboList1
                End If

                .ComboList = StrComboList1
         
        End Select

    End With

End Sub
Function fillVchr()
    Dim i As Integer
        
    With GRID1

        For i = 1 To .Rows - 1

            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
                Retrive_orders_data (val(.TextMatrix(i, .ColIndex("Transaction_ID"))))
            
            End If

        Next i

    End With

End Function

Function fillOrders()
    Dim i As Integer

    With GRID2

        For i = 1 To .Rows - 1

            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
                Retrive_orders_data (val(.TextMatrix(i, .ColIndex("Transaction_ID"))))
            
            End If

        Next i

    End With
NewGrid.Calculate 1, , , True
End Function

Private Sub Label9_Click()

End Sub

Private Sub ISButton1_Click()

If CheckBox1.value = vbChecked Then
If val(Me.DcbSuppler.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„Ê—œ"
Else
MsgBox "Please Select Vendor"
End If
Exit Sub
End If
End If
If IsNull(DTPickerAccFrom.value) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ  ÕœÌœ «· «—ÌŒ"
Else
MsgBox "Please Select Date"
End If
Exit Sub
End If
If IsNull(DTPickerAccTo.value) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ  ÕœÌœ «· «—ÌŒ"
Else
MsgBox "Please Select Date"
End If
Exit Sub
End If
    FillOrderGrid DTPickerAccFrom.value, DTPickerAccTo.value
    NewGrid.Calculate 1, , , True
    showComm
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

Private Sub LblDiscountsTotal_Change()
    LblDiscountsTotalView.Caption = Format(val(LblDiscountsTotal.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
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

Private Sub LblTotal_Change()
ReLineGrid
    LblTotalView.Caption = Format(val(LblTotal.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))

    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
        TxtNetValue.Text = val(LblTotal.Caption)
        TxtPayedValue.Text = TxtNetValue.Text
 
        With Me.FgInstallments
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            LblPrecenType.Caption = ""
            LblPrecenValue.Caption = ""
            LblInstallTotal.Caption = ""
            LblInstallCount.Caption = ""
            LblFirstInstallDate.Caption = ""
            LblInstallmentType.Caption = ""
        End With

    End If
  
End Sub

Function showComm()

    If val(LblInstallTotal.Caption) > 0 Then
        lblInstComm.Caption = val(LblInstallTotal.Caption) - val(LblTotal.Caption)
  
    Else
        lblInstComm.Caption = 0
        '  Me.LblFinal = 0
    End If

    Me.LblFinal = val(lblInstComm.Caption) + val(LblTotal.Caption)
    'Me.lblInstComm.Caption = Format(Val(lblInstComm.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
 DcbTypComm_Change
    

End Function

Private Sub LblTotal_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    LblTotal.ToolTipText = WriteNo(LblTotal.Caption, 0, True)

End Sub

Private Sub LblTotalAll_Change()
DcbTypComm_Change
    LblTotalAllView.Caption = Format(val(LblTotalAll.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
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

    If Me.TxtModFlg.Text <> "R" Then
        Msg = " ÕœÌÀ «·»Ì«‰«  €Ì— „ «Õ ≈·« «‰  ﬂÊ‰ «·‘«‘… ›Ï Õ«·… «·⁄—÷ ›ﬁÿ..!"
        'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        'Exit Sub
    End If

    LoadCombosData
    NewGrid.FillGrid , , False
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

Private Sub Text1_Change()

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        Command2.backcolor = vbYellow
        Command2.Enabled = False

        'Exit Sub
    End If

    If Text1.Text = "" Then
        Command2.backcolor = vbGreen
        Command2.Enabled = True

        If SystemOptions.UserInterface = ArabicInterface Then
            Command2.Caption = "  ·„ Ì „ «‰‘«¡ «–‰ «·’—›- «÷€ÿ  ·«‰‘«¡ «–‰ ’—› «·Ì"
        Else
            Command2.Caption = "Create Issue Voucher"
        End If
        
    Else
        Command2.backcolor = &HC0C0C0
        Command2.Enabled = False

        If SystemOptions.UserInterface = ArabicInterface Then
            Command2.Caption = "  „ «‰‘«¡ «–‰ «·’—› "
        Else
            Command2.Caption = "Voucher Was Created"
        
        End If
    End If

    If BillBasedOn(1).value = True Then
        Command2.backcolor = &HC0C0C0
        Command2.Enabled = False

        If SystemOptions.UserInterface = ArabicInterface Then
            Command2.Caption = "·« Ì„ﬂ‰ «‰‘«¡ «·”‰œ ·«‰ «·›« Ê—Â  „ —»ÿÂ« »⁄œÂ ”‰œ«  "
        Else
            Command2.Caption = "Can't Create Voucher  "
        End If
    End If

End Sub

Private Sub TxtFillData_Change()

    If TxtFillData.Text = "F" Then
        NewGrid.Calculate 1, , , True
    End If

End Sub

Public Sub RetriveSerials(ItemID As String, _
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
        FG.Rows = FG.Rows + 1
 
        Num = Num + 1
    Next
 
    TxtFillData.Text = "F"
    TxtFillData_Change
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Public Sub RetriveOrder(Optional order_no As String = "", _
                        Optional Transaction_Type As Integer = 0)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh

    If Transaction_Type = 0 Then
        StrSQL = "Select * from transactions where  Transaction_Type=6 and order_no='" & order_no & "'"
    Else
        StrSQL = "Select * from transactions where  Transaction_Type=22 and NoteSerial1='" & order_no & "'"
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount < 1 Then
 
        Exit Sub
    Else
        DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
        Me.DcCurrency.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
        Me.DCboStoreName.BoundText = IIf(IsNull(rs("storeid").value), "", rs("storeid").value)
        Me.dcBranch.BoundText = IIf(IsNull(rs("Branchid").value), "", rs("Branchid").value)

        'txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass

    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.Text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))

            'FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
            If Transaction_Type = 0 Then
                FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("ShowPrice")), 0, (RsDetails("ShowPrice").value)) ' GET_COST_PRICE_FOR_PRODUCT_ITEM(Val(FG.TextMatrix(Num, FG.ColIndex("Code"))))
            End If
      
            '  FG.TextMatrix(Num, FG.ColIndex("Expenses")) = IIf(IsNull(RsDetails("Lineexpenses")), "", (RsDetails("Lineexpenses").value))
         
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemType")) = IIf(IsNull(RsDetails("ItemType")), 0, (RsDetails("ItemType").value))
         
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        
            RsDetails.MoveNext
            Debug.Print Num

            If FG.Rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num

    End If

    TxtFillData.Text = "F"
    Screen.MousePointer = vbDefault
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Private Sub TxtNetValue_Change()
    TxtRemainValue.Text = val(Me.TxtPayedValue.Text) - val(Me.TxtNetValue.Text)
End Sub

Private Sub TxtNetValue_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    TxtNetValue.ToolTipText = WriteNo(LblTotal.Caption, 0, True)
End Sub

Private Sub TXTOrDer_no_Change()

    If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
        RetriveOrder Me.TXTOrDer_no
    End If

End Sub

Public Function NewBillFromOrder(orderNo As String)

    If Me.TxtModFlg = "R" Then
        Cmd_Click (0)
        Me.TXTOrDer_no.Text = orderNo
        'txtorder_no_Change
        'RetriveOrder orderNo
    End If

End Function

Private Sub TXTOrDer_no_KeyUp(KeyCode As Integer, _
                              Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Order_no_search.show
        Order_no_search.RetrunType = 8

        If val(Me.DBCboClientName.BoundText) <> 2 Then
        
            Order_no_search.DBCboClientName.BoundText = Me.DBCboClientName.BoundText
        End If
    End If

End Sub

Private Sub TxtPayedValue_Change()
    TxtRemainValue.Text = val(Me.TxtPayedValue.Text) - val(Me.TxtNetValue.Text)
End Sub

Private Sub TxtPayedValue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtPayedValue.Text, 0)
End Sub

Private Sub TxtPurchaseBill_Change()

    If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
        RetriveOrder Me.TxtPurchaseBill, 22
    End If

End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.Text, 1
        DBCboClientName.BoundText = CUSTID
    End If

End Sub

Private Sub TxtSuplCode3_KeyPress(KeyAscii As Integer)
    Dim CUSTID As Integer
    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSuplCode3.Text, 2
        DcbFarm.BoundText = CUSTID
    End If
End Sub

Private Sub TxtTransSerial_Change()
    FillVoucherGrid
End Sub

Private Sub TxtTransSerial_KeyDown(KeyCode As Integer, _
                                   Shift As Integer)
    Dim StrSearch As String
    Dim VarBookMark As Variant
    Dim Msg As String

    If Me.TxtModFlg.Text = "R" Then
        If KeyCode = vbKeyReturn Then
            If Trim$(TxtTransSerial.Text) <> "" Then
                StrSearch = Trim$(TxtTransSerial.Text)

                If Not (rs.BOF Or rs.EOF) Then
                    If rs.EditMode = adEditNone Then
                        VarBookMark = rs.Bookmark
                        rs.find "Transaction_Serial='" & StrSearch & "'", , adSearchForward, adBookmarkFirst

                        If Not (rs.BOF Or rs.EOF) Then
                            Me.Retrive rs("Transaction_ID").value
                        Else
                            rs.Bookmark = VarBookMark
                            Msg = "Â–Â «·›« Ê—… €Ì— „ÊÃÊœ…...!!!"
                            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        End If
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Sub TxtTransSerial_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtTransSerial.Text, 1)
End Sub

Private Sub TxtValueComm_Change()
CalCulComm
End Sub

Private Sub TxtValueComm_LostFocus()
mNetValueComm = IIf(val(TxtValueComm) = 0, 5, val(TxtValueComm))
End Sub

Private Sub TxtVATNO_Change()
If TxtVATNO.Text <> "" Then
txtVatYou.Text = 5
Else
txtVatYou.Text = 0
End If
End Sub

Private Sub TxtVATYou_Change()
TxtVATYouView.Caption = Format(val(txtVatYou.Text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
CalCulComm
End Sub

Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap
invoiceSerach = False
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

'
Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    'Exit Sub
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" And Not (Me.ActiveControl Is TxtTransSerial) Then
            '        Cmd_Click (0)
        Else
            '    SendKeys "{TAB}"
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
        '    If Cmd(3).Enabled = False Then Exit Sub
        '    Cmd_Click (3)
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
        If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
            'XPBtnAdd_Click
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
            'XPBtnRemove_Click
        End If
    End If

    If KeyCode = vbKeyDelete Then
        If Me.ActiveControl Is FG Then
            If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
                'XPBtnRemove_Click
            End If
        End If
    End If

    If KeyCode = vbKeyF5 Then
        If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
            XPBtnNewClients_Click
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeySpace Then
            If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
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
    On Error GoTo ErrTrap
    invoiceSerach = True
    Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2
    ScreenNameArabic = " ›« Ê—… «·„»Ì⁄«  "
    ScreenNameEnglish = " Sales Invoice"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
  
    Command2.Caption = ""
    first_run = True
    Dim StrSQL As String
    Dim Num As Integer
    Dim StrList As String
    Dim BGround As New ClsBackGroundPic
    Dim ShowTax As Boolean

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    If mdifrmmain.TransporterMain.Visible = False Then
        Frame5.Visible = False
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
    'fill_combo dcBranch, My_SQL
  
    If SystemOptions.usertype <> UserAdminAll Then
 
        Me.dcBranch.Enabled = True
        ' XPDtbBill.Enabled = False
    End If

    Set NewGrid.Grid = FG

    ShowTax = GetSetting(StrAppRegPath, "SallBill", "HaveTaxOnSalles", False)
    Ele(4).Visible = ShowTax
    NewGrid.GridTrans = InvoiceTransaction
    NewGrid.GridTrans2 = InvoiceTransactionCompose
    
    Set NewGrid.TxtInvID = Me.XPTxtBillID
    Set NewGrid.TxtNots = Me.Text1
    Set NewGrid.Branch = Me.dcBranch
    Set NewGrid.TxtModFlag = TxtModFlg
    Set NewGrid.txtTotal = XPTxtSum
    Set NewGrid.CboDiscount_Type = XPCboDiscountType
    Set NewGrid.TxtDiscount_Val = XPTxtDiscountVal
    Set NewGrid.TxtValueCash = XPTxtValue(0)
    Set NewGrid.TxtValueDelay = XPTxtValue(1)
    Set NewGrid.TxtValuechque = XPTxtValue(2)
    Set NewGrid.txt_Currency_rate = txt_Currency_rate
    Set NewGrid.LBLGross = LBLGross
    Set NewGrid.Customer = Me.DcCustmer
    '--------------------------------------
    Set NewGrid.TxtTaxValue = Me.XPTxtTaxValue
    Set NewGrid.TxtAddTax = Me.TxtTaxAddValue
    Set NewGrid.TxtStampTax = Me.TxtTaxStampValue
    Set NewGrid.TxtServiceTax = Me.TxtTaxServiceValue
    '------------------------------------------------
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.StoreName = Me.DCboStoreName
    Set NewGrid.DtpBillDate = Me.XPDtbBill
    Set NewGrid.CmdAddSerialLIst = Me.CmdSearch
    'Set NewGrid.CboDiscountType = CboDiscountType
    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    
   'Set NewGrid.Customer2 = DcCustmer
   
    Set NewGrid.TxtCustCode = TxtCustCode
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

    Set NewGrid.LblTotalQty = Me.LblTotalQty

    Set NewGrid.LblTaxSalesValue = Me.lbl(51)
    Set NewGrid.LblTaxAddValue = Me.lbl(52)
    Set NewGrid.LblTaxStampValue = Me.lbl(53)
    Set NewGrid.LblTaxServiceValue = Me.lbl(54)

    NewGrid.FillGrid , , False
    StrSQL = " select id,code from currency"
 
    fill_combo Me.DcCurrency, StrSQL

    FG.WallPaper = BGround.Picture
    AddTip
    XPTab301.CurrTab = 0
    XPDtbBill.value = Date

    If SystemOptions.UserInterface = ArabicInterface Then
        With DcbTypComm
        .Clear
        .AddItem "ﬁÌ„…"
        .AddItem "‰”»…"
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

        With CboPOSBillType
            .Clear
            .AddItem "«·ÿ«Ê·…"
            .AddItem "ÿ·»«  Œ«—ÃÌ…"
            .AddItem " Œœ„…  Ê’Ì· "
        End With
NewGrid.ReturnTyp = 10
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        With DcbTypComm
        .Clear
        .AddItem "Value"
        .AddItem "Percentage"
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
            .AddItem "Credit"
        End With

        With Me.CboSaleType
            .Clear
            .AddItem "Retail"
            .AddItem "WholeSale"
        End With
    
        With CboPOSBillType
            .Clear
            .AddItem "table"
            .AddItem "Take Away"
            .AddItem " Delivery  "
        End With

    End If

    '--------------------------------
    Set Dcombos = New ClsDataCombos
    LoadCombosData

    '--------------------------------
    If SystemOptions.UserInvoiceShowProfit = 0 Then
        '   Me.Ele(8).Visible = False
        Frame400.Visible = False
    Else
        Frame400.Visible = True
        'Me.Ele(8).Visible = True
    End If

    SetDtpickerDate Me.XPDtbBill
    '----------------------------
    SetDtpickerDate Me.DtpDelayDate
    '≈⁄œ«œ Ã—œ «·√ﬁ”«ÿ
    ChkInstall.value = Unchecked
    ChkInstall.Enabled = False

    With Me.FgInstallments
        .Rows = .FixedRows
        Set .WallPaper = BGround.Picture
        .RowHeightMin = 300
        .AutoSize 0, .Cols - 1, False
    End With

    With Me.FgCheques
        .Rows = .FixedRows
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
    'Resize_Form Me, TransactionSize
    Me.Height = 10000
    Me.Width = 17595
    Me.Top = (mdifrmmain.ScaleHeight - Me.Height) / 2
    Me.Left = (mdifrmmain.ScaleWidth - Me.Width) / 2

    '----------------------------
    'DB_CreateField "Transactions", "TransactionComment", adVarWChar, adColNullable, 255, , " ”ÃÌ· „·«ÕŸ«  ⁄·Ï «·›« Ê—…", False, True
    '----------------------------
    Dim rsOut As New ADODB.Recordset
    Dim Msg As String
    Set rsOut = New ADODB.Recordset
    rsOut.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    If Not (rsOut.EOF Or rsOut.BOF) Then
 
        If rsOut!checkout = True Then
            StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=40 "
     StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"
            If SystemOptions.usertype <> UserAdminAll Or val(Current_branch) <> 0 Then
             '   StrSQL = StrSQL & " AND   BranchId=" & Current_branch
            End If

            StrSQL = StrSQL & " Order by Transaction_ID"
                
            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                rs.MoveLast
            End If

            Retrive
            Me.TxtModFlg.Text = "R"
            InvType = 40
        Else
 
            StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=40  "

            If SystemOptions.usertype <> UserAdminAll Or val(Current_branch) <> 0 Then
                StrSQL = StrSQL & "  AND   BranchId=" & Current_branch
            End If

            StrSQL = StrSQL & " Order by Transaction_ID"

            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                rs.MoveLast
            End If

            Retrive
            Me.TxtModFlg.Text = "R"
            InvType = 40
        End If
    End If
mNetValueComm = 5
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    Dim i As Integer
    SelectedIssueVoucher = False
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish

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

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "›« Ê—…«·»Ì⁄"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.Caption = "Bill Invoice"
            End If

            BillBasedOn(1).Enabled = False
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.Cmd(7).Enabled = True
            Me.DcboEmp.Enabled = True
            GRID1.Enabled = True
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

                    If SystemOptions.UserInterface = ArabicInterface Then
                        CmdINSTALLMENT.Caption = "⁄—÷ «·√ﬁ”«ÿ «·„”Ã·…"
                    Else
                        CmdINSTALLMENT.Caption = "View"
                    End If

                Else
                    CmdINSTALLMENT.Enabled = False

                    If SystemOptions.UserInterface = ArabicInterface Then
                        CmdINSTALLMENT.Caption = " ﬁ”Ìÿ «·ﬁÌ„… «·¬Ã·…"
                    Else
                        CmdINSTALLMENT.Caption = "Calc"
                    End If
                End If
            End If

            Ele(2).Enabled = False
            DcboEmp.Enabled = True
            XPChkTAX.Enabled = False
            ChkTaxAdd.Enabled = False
            ChkTaxSerivce.Enabled = False
            ChkTaxStamp.Enabled = False

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "›« Ê—…«·»Ì⁄( ÃœÌœ )"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.Caption = "Bill Invoice(New)"
            End If

            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
            Me.DcboEmp.Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                CmdINSTALLMENT.Caption = " ﬁ”Ìÿ «·ﬁÌ„… «·¬Ã·…"
            Else
                CmdINSTALLMENT.Caption = "Calc Installments"
            End If
               
            '        Me.XPBtnMove(0).Enabled = False
            '        Me.XPBtnMove(1).Enabled = False
            '        Me.XPBtnMove(2).Enabled = False
            '        Me.XPBtnMove(3).Enabled = False
            XPBtnNewClients.Enabled = True
            FG.Enabled = True
            FG.Rows = FG.FixedRows
            FG.Rows = 2
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
            ChkTaxAdd.Enabled = True
            ChkTaxStamp.Enabled = True
            ChkTaxSerivce.Enabled = True
            XPTxtTaxValue.Text = ""
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
            ChkTaxStamp.Enabled = True
            ChkTaxSerivce.Enabled = True
            ChkTaxStamp.Enabled = True
            ChkTaxSerivce.Enabled = True

            '        ChkTaxSerivce.Enabled = True
            '        ChkTaxStamp.Enabled = True
        Case "E"
            BillBasedOn(1).Enabled = False
    
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "›« Ê—…«·»Ì⁄(   ⁄œÌ· )"
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.Caption = "Bill Invoice( Edit )"
            End If

            XPDtbBill.Enabled = False
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
                If XPTxtValue(1).Text <> "" Then
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
            ChkTaxSerivce.Enabled = True
            ChkTaxStamp.Enabled = True
            '        ChkTaxSerivce.Enabled = True
            '        ChkTaxStamp.Enabled = True

    End Select

    If SystemOptions.usertype <> UserAdminAll Then
 
        Me.dcBranch.Enabled = True
        'XPDtbBill.Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0, _
                   Optional NoteSerial1 As String)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTest  As ADODB.Recordset
    Dim RsReplace As ADODB.Recordset
    Dim LngPartID As Long
    Dim RsPartDetails As ADODB.Recordset
    Dim i As Long

    On Error GoTo ErrTrap
    '---------------------------------------------
    'Here We Reset all Setting

    With Me.FgInstallments
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows
        LblPrecenType.Caption = ""
        LblPrecenValue.Caption = ""
        LblInstallTotal.Caption = ""
        LblInstallCount.Caption = ""
        LblFirstInstallDate.Caption = ""
        LblInstallmentType.Caption = ""
    End With
    
    Me.CmdNotes.Visible = False
    Me.CmdNotes.Tag = ""
    Me.CmdRetruns.Visible = False
    Me.CmdRetruns.Tag = ""

    ChkTaxAdd.value = vbUnchecked
    Me.TxtTaxAddValue.Text = ""
    ChkTaxStamp.value = vbUnchecked
    Me.TxtTaxStampValue.Text = ""
    ChkTaxStamp.value = vbUnchecked
    Me.TxtTaxStampValue.Text = ""
    ChkTaxSerivce.value = vbUnchecked
    Me.TxtTaxServiceValue.Text = ""

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
        rs.find "Transaction_ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If

    If NoteSerial1 <> "" Then

        rs.find "noteserial1='" & NoteSerial1 & "'", , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If

    TxtFillData.Text = "T"
    Screen.MousePointer = vbArrowHourglass
    ' »Ì«‰«  ÃœÌœ…
    DTPickerAccFrom.value = IIf(IsNull(rs.Fields("DTPickerAccFrom").value), Date, rs.Fields("DTPickerAccFrom").value)
    DTPickerAccTo.value = IIf(IsNull(rs.Fields("DTPickerAccTo").value), Date, rs.Fields("DTPickerAccTo").value)
    
    TxtNoteSerial111.Text = IIf(IsNull(rs.Fields("NoteSerial").value), "", rs.Fields("NoteSerial").value)
    Me.TxtNoteID111.Text = IIf(IsNull(rs.Fields("NoteID").value), "", rs.Fields("NoteID").value)
    Me.DCPaymentNet.BoundText = IIf(IsNull(rs("PaymentNetid").value), "", rs("PaymentNetid").value)
    TxtNetValue.Text = IIf(IsNull(rs("NetValue").value), "", (rs("NetValue").value))
    TxtPayedValue.Text = IIf(IsNull(rs("PayedValue").value), "", (rs("PayedValue").value))
    TxtRemainValue.Text = IIf(IsNull(rs("RemainValue").value), "", (rs("RemainValue").value))
    Me.DcbTypComm.ListIndex = IIf(IsNull(rs("TypComm").value), -1, rs("TypComm").value)
    Me.TxtValueComm.Text = IIf(IsNull(rs("ValueComm").value), "", rs("ValueComm").value)
    mNetValueComm = val(Me.TxtValueComm.Text)
    Me.TxtNetValueComm.Text = IIf(IsNull(rs("NetValueComm").value), "", rs("NetValueComm").value)
    Me.TxtGVAT.Text = IIf(IsNull(rs("GVAT").value), "", rs("GVAT").value)
    Me.TxtSAlGValue.Text = IIf(IsNull(rs("SAlGValue").value), "", rs("SAlGValue").value)
    Me.TXTFactoryExpenses.Text = IIf(IsNull(rs("FactoryExpenses").value), "", rs("FactoryExpenses").value)
    Me.TxtGTotal.Text = IIf(IsNull(rs("GTotal").value), "", rs("GTotal").value)
    TxtVATNO.Text = IIf(IsNull(rs("VATNO").value), "", (rs("VATNO").value))
    txtVatYou.Text = IIf(IsNull(rs("VATYou").value), "", (rs("VATYou").value))
    TxtGratuity.Text = IIf(IsNull(rs("Gratuity").value), "", (rs("Gratuity").value))
    If Not IsNull(rs("ChLoadVAT").value) Then
    If rs("ChLoadVAT").value = 1 Then
      ChLoadVAT.value = vbChecked
    Else
      ChLoadVAT.value = vbUnchecked
    End If
    Else
    ChLoadVAT.value = vbUnchecked
   End If
   
   optCommissionType(0) = IIf(IsNull(rs("CommissionType").value), True, (rs("CommissionType").value))
   optCommissionType(1) = Not optCommissionType(0)
   
    TxtManualNo1.Text = IIf(IsNull(rs("ManualNo1").value), "", (rs("ManualNo1").value))
    TxtManualNo2.Text = IIf(IsNull(rs("ManualNo2").value), "", (rs("ManualNo2").value))
 
    '‰ﬁ«ÿ «·»Ì⁄
    If Not IsNull(rs("POSBillType").value) Then
        CboPOSBillType.ListIndex = val(rs("POSBillType").value)
        LblStableID.Caption = IIf(IsNull(rs("STableID").value), -1, (rs("STableID").value))

    Else
        CboPOSBillType.ListIndex = -1
        LblStableID.Caption = -1

    End If
 
    Me.DCCar.BoundText = IIf(IsNull(rs("CarId").value), "", rs("CarId").value)
    Me.DCDriver.BoundText = IIf(IsNull(rs("DriverId").value), "", rs("DriverId").value)

    lblSessionD.Caption = IIf(IsNull(rs("SessionD").value), -1, (rs("SessionD").value))

    If Not IsNull(rs("BillBasedOn").value) Then

        If rs("BillBasedOn").value = 0 Then
            BillBasedOn(0).value = True
            '   BillBasedOn_Click (0)
        ElseIf rs("BillBasedOn").value = 1 Then
            BillBasedOn(1).value = True
            '      BillBasedOn_Click (1)
        ElseIf rs("BillBasedOn").value = 2 Then
            BillBasedOn(2).value = True
            '      BillBasedOn_Click (2)
        End If
    
    Else

        BillBasedOn(0).value = True
        '  BillBasedOn_Click (0)
    End If
''///////////New Data Farm
DcbFarm.BoundText = IIf(IsNull(rs("FarmID").value), "", rs("FarmID").value)
TxtBoardNO.Text = IIf(IsNull(rs("BoardNo").value), "", rs("BoardNo").value)
Txtcommission.Text = IIf(IsNull(rs("Commission").value), "", rs("Commission").value)
TxtEmbarNo.Text = IIf(IsNull(rs("EmbarNo").value), "", rs("EmbarNo").value)
TxtDriverName.Text = IIf(IsNull(rs("DriverName").value), "", rs("DriverName").value)
''/////////////
    dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
    DCDocTypes.BoundText = IIf(IsNull(rs("Doctype").value), "", rs("Doctype").value)
    Me.DcCurrency.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
    txt_Currency_rate.Text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
 
    Me.TxtNoteSerial.Text = IIf(IsNull(rs("NoteSerial").value), "", (rs("NoteSerial").value))
    Me.TxtNoteSerial1.Text = IIf(IsNull(rs("NoteSerial1").value), "", (rs("NoteSerial1").value))
    DCPreFix.Text = IIf(IsNull(rs("Prefix").value), "", rs("Prefix").value)
    Me.oldtxtNoteSerial1.Text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)

    lbl(64).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

    Me.TXTNoteID.Text = IIf(IsNull(rs("NoteID").value), "", (rs("NoteID").value))
    Text1.Text = IIf(IsNull(rs("NotS").value), "", (rs("NotS").value))

    XPTxtBillID.Text = IIf(IsNull(rs("Transaction_ID").value), "", val(rs("Transaction_ID").value))
    DcbSuppler.BoundText = IIf(IsNull(rs("SupplerID").value), "", rs("SupplerID").value)
    TxtTransSerial.Text = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
    XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", (rs("Transaction_Date").value))
    XPCboDiscountType.ListIndex = IIf(IsNull(rs("Trans_DiscountType").value), -1, val(rs("Trans_DiscountType").value))
    CboPayMentType.ListIndex = IIf(IsNull(rs("PaymentType").value), 0, rs("PaymentType").value)
    XPTxtDiscountVal.Text = IIf(IsNull(rs("Trans_Discount").value), "", (rs("Trans_Discount").value))
    Me.DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    FG.Clear flexClearScrollable, flexClearEverything
    Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
    Me.DcboEmp.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
    XPTxtTaxValue.Text = IIf(IsNull(rs("TaxValue").value), "", (rs("TaxValue").value))
    XPChkTAX.value = IIf(rs("TaxFound") = True, Checked, Unchecked)
    'Text1.text = IIf(IsNull(rs("nots2").value), "", (rs("nots2").value))
    Me.TXTOrDer_no.Text = IIf(IsNull(rs("order_no").value), "", (rs("order_no").value))

    TxtPurchaseBill.Text = IIf(IsNull(rs("PurchaseBill").value), "", (rs("PurchaseBill").value))
 
    If IsNull(rs("BoxID").value) Then
        Me.DcboBox.BoundText = ""
    Else
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
    End If

    If IsNull(rs("SaleType").value) Then
        Me.CboSaleType.ListIndex = 0
    Else
        Me.CboSaleType.ListIndex = IIf(rs("SaleType").value = 0, 0, 1)
    End If

    If Not (IsNull(rs("CashCustomerName").value)) Then
        Me.TxtCashCustomerName.Text = rs("CashCustomerName").value
    Else
        Me.TxtCashCustomerName.Text = ""
    End If

    'ChkInstall 11 10 2012
    If IsNull(rs("ChkInstall").value) Then
        Me.ChkInstall.value = vbUnchecked
    Else
        Me.ChkInstall.value = IIf(rs("ChkInstall").value = 0, vbUnchecked, vbChecked)
    End If

    '÷—»Ì… «·Œ’„ Ê«·≈÷«›…
    If Not IsNull(rs("TaxAddValue").value) Then
        If rs("TaxAddValue").value > 0 Then
            ChkTaxAdd.value = vbChecked
            Me.TxtTaxAddValue.Text = rs("TaxAddValue").value
        End If
    End If

    '÷—»Ì… «·œ„€…
    If Not IsNull(rs("TaxStampValue").value) Then
        If rs("TaxStampValue").value > 0 Then
            ChkTaxStamp.value = vbChecked
            Me.TxtTaxStampValue.Text = rs("TaxStampValue").value
        End If
    End If

    '÷—»Ì… «·Œœ„…
    If Not IsNull(rs("TaxServiceValue").value) Then
        If rs("TaxServiceValue").value > 0 Then
            ChkTaxSerivce.value = vbChecked
            Me.TxtTaxServiceValue.Text = rs("TaxServiceValue").value
        End If
    End If

    TxtBillComment.Text = IIf(IsNull(rs("TransactionComment").value), "", (rs("TransactionComment").value))

    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)
    StrSQL = StrSQL + "order by id"

    Set RsDetails = New ADODB.Recordset
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.Text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1

        For i = 1 To RsDetails.RecordCount
            FG.TextMatrix(i, FG.ColIndex("CusID2")) = IIf(IsNull(RsDetails("CusID2")), "", (RsDetails("CusID2").value))
           ' FG.TextMatrix(i, FG.ColIndex("CusID3")) = IIf(IsNull(RsDetails("CusID3")), "", (RsDetails("CusID3").value))
            FG.TextMatrix(i, FG.ColIndex("FoxyNo")) = IIf(IsNull(RsDetails("FoxyNo")), "", (RsDetails("FoxyNo").value))
            FG.TextMatrix(i, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no").value), "", RsDetails("order_no").value)
            FG.TextMatrix(i, FG.ColIndex("OrderArrivalDate")) = IIf(IsNull(RsDetails("OrderArrivalDate").value), "", RsDetails("OrderArrivalDate").value)
            FG.Cell(flexcpPicture, i, FG.ColIndex("Ser")) = ""
            FG.Cell(flexcpData, i, FG.ColIndex("Ser")) = ""
            FG.TextMatrix(i, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(i, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim$(RsDetails("Item_ID").value))
            FG.TextMatrix(i, FG.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial")), "", Trim(RsDetails("ItemSerial").value))

            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(i, FG.ColIndex("HaveSerial")) = True

                '«·»ÕÀ ⁄‰ ⁄„·Ì«  «·«” »œ«· «·Œ«’… »«·›« Ê—…
                If (RsDetails("Item_ID")) <> "" And RsDetails("ItemSerial") <> "" Then
                    StrSQL = "select * From ReplacedItems where ReturnID=" & XPTxtBillID.Text
                    StrSQL = StrSQL + " and ItemID=" & RsDetails("Item_ID")
                    StrSQL = StrSQL + " and ItemSerial='" & RsDetails("ItemSerial") & "'"
                    Set RsReplace = New ADODB.Recordset
                    RsReplace.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (RsReplace.EOF Or RsReplace.BOF) Then
                        FG.Cell(flexcpPicture, i, FG.ColIndex("Ser")) = mdifrmmain.ImgLstTree.ListImages("Request").Picture
                        FG.Cell(flexcpData, i, FG.ColIndex("Ser")) = "X"
                    End If
                End If
            End If

            FG.TextMatrix(i, FG.ColIndex("ItemType")) = IIf(IsNull(RsDetails("ItemType").value), "", (RsDetails("ItemType").value))

            If RsDetails("ItemType").value = 1 Then
                FG.Cell(flexcpPicture, i, FG.ColIndex("Ser")) = mdifrmmain.ImgLstTree.ListImages("Maintenance").Picture
            
            End If

            FG.TextMatrix(i, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("ShowQty")), "", (RsDetails("ShowQty").value))
            FG.TextMatrix(i, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice")), "", (RsDetails("showPrice").value))
        
            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                FG.TextMatrix(i, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            Else
                FG.TextMatrix(i, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("Transaction_Details.ItemCase")), "", (RsDetails("Transaction_Details.ItemCase").value))
            End If
            FG.TextMatrix(i, FG.ColIndex("Vat")) = IIf(IsNull(RsDetails("Vat")), "", (RsDetails("Vat").value))
            FG.TextMatrix(i, FG.ColIndex("Vatyo")) = IIf(IsNull(RsDetails("Vatyo")), "", (RsDetails("Vatyo").value))
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
                Me.FG.Cell(flexcpBackColor, i, 1, i, FG.Cols - 1) = vbYellow
            ElseIf val(FG.TextMatrix(i, FG.ColIndex("ItemProfit"))) < 0 Then
                Me.FG.Cell(flexcpBackColor, i, 1, i, FG.Cols - 1) = vbRed
            Else
                Me.FG.Cell(flexcpBackColor, i, 1, i, FG.Cols - 1) = 0
            End If

            FG.Cell(flexcpData, i, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
        
            If SystemOptions.UserInterface = ArabicInterface Then
                FG.TextMatrix(i, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            Else
                FG.TextMatrix(i, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitNamee")), "", (RsDetails("UnitNamee").value))
            End If

            FG.TextMatrix(i, FG.ColIndex("ProductionDate")) = IIf(IsNull(RsDetails("ProductionDate")), "", (RsDetails("ProductionDate").value))
            FG.TextMatrix(i, FG.ColIndex("ExpiryDate")) = IIf(IsNull(RsDetails("ExpiryDate")), "", (RsDetails("ExpiryDate").value))
            FG.TextMatrix(i, FG.ColIndex("LotNO")) = IIf(IsNull(RsDetails("LotNO")), "", (RsDetails("LotNO").value))
       
            FG.TextMatrix(i, FG.ColIndex("GranteeType")) = IIf(IsNull(RsDetails("GranteeType")), "", (RsDetails("GranteeType").value))
            FG.TextMatrix(i, FG.ColIndex("GranteeStartDate")) = IIf(IsNull(RsDetails("GranteeStartDate")), "", (RsDetails("GranteeStartDate").value))
            FG.TextMatrix(i, FG.ColIndex("GranteeEndDate")) = IIf(IsNull(RsDetails("GranteeEndDate")), "", (RsDetails("GranteeEndDate").value))
            FG.TextMatrix(i, FG.ColIndex("RegularMaintenancedates")) = IIf(IsNull(RsDetails("RegularMaintenancedates")), "", (RsDetails("RegularMaintenancedates").value))
       
            RsDetails.MoveNext
        
            If FG.Rows > 10 Then
                If i = 8 Then FG.Refresh
            End If

        Next i

        '----------------------------
        Me.LblInvProfit.Caption = FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("ItemProfit"), FG.Rows - 1, FG.ColIndex("ItemProfit"))

        If val(Me.LblInvProfit.Caption) > 0 Then
            Me.LblInvProfit.ForeColor = &H4000&
        ElseIf val(Me.LblInvProfit.Caption) = 0 Then
            Me.LblInvProfit.ForeColor = vbBlack
        ElseIf val(Me.LblInvProfit.Caption) < 0 Then
            Me.LblInvProfit.ForeColor = vbRed
        End If

        '---------------------------
        '    Fg.AutoSize 0, Fg.Cols - 1, False
    End If

    XPChkPayType(0).value = Unchecked
    XPChkPayType(1).value = Unchecked
    XPChkPayType(2).value = Unchecked
    XPTxtValue(0).Text = ""
    XPTxtValue(1).Text = ""
    XPTxtSerial(0).Text = ""
    XPTxtSerial(1).Text = ""
    XPTxtValue(1).Tag = ""
    DtpDelayDate.value = Date
    '----------------------------------------------------------------------------------------
    StrSQL = "Select * From Notes Where Transaction_ID=" & val(rs("Transaction_ID").value)
    RsNotes.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsNotes.EOF Or RsNotes.BOF) Then

        For i = 1 To RsNotes.RecordCount

            If RsNotes("NoteType").value = 170 Then
                XPChkPayType(0).value = Checked
                XPChkPayType_Click (0)
                XPTxtValue(0).Text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtSerial(0).Text = IIf(IsNull(RsNotes("NoteSerial").value), "", Trim$(RsNotes("NoteSerial").value))
                Me.DcboBox.BoundText = IIf(IsNull(RsNotes("BoxID").value), "", RsNotes("BoxID").value)
            End If

            If RsNotes("NoteType").value = 1 Then
                XPChkPayType(1).value = Checked
                XPChkPayType_Click (1)
                XPTxtValue(1).Text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtValue(1).Tag = IIf(IsNull(RsNotes("NoteID").value), "", (RsNotes("NoteID").value))
                XPTxtSerial(1).Text = IIf(IsNull(RsNotes("NoteSerial").value), "", (RsNotes("NoteSerial").value))
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
        .Rows = .FixedRows

        If Not (RsNotes.BOF Or RsNotes.EOF) Then
            .Rows = .FixedRows + RsNotes.RecordCount

            For i = .FixedRows To .Rows - 1
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
   
    TxtFillData.Text = "F"
    '-----------------------------------------------------------------------------------------------
    Dim SngRelatedNotesValues As Single
    Me.CmdNotes.Visible = ShowRelatedNotes(val(Me.XPTxtBillID.Text), 0, SngRelatedNotesValues)
    Me.CmdNotes.Tag = SngRelatedNotesValues

    SngRelatedNotesValues = 0
    Me.CmdRetruns.Visible = ShowRelatedTransactions(val(Me.XPTxtBillID.Text), 0, SngRelatedNotesValues)
    Me.CmdRetruns.Tag = SngRelatedNotesValues

    '-----------------------------------------------------------------------------------------------
    Screen.MousePointer = vbDefault
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    showComm
    FillVoucherGrid
    'FillOrderGrid
    LoadSigns val(XPTxtBillID.Text), 170

    '    Else
    '        CmdINSTALLMENT.Enabled = False
    '        CmdINSTALLMENT.Caption = " ﬁ”Ìÿ «·ﬁÌ„… «·¬Ã·…"
    
    '  End If
    'Else
    'FgInstallments.Clear

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
            'LblDiscount.Caption = IIf(IsNull(RsTest("Discount").value), "", RsTest("Discount").value)
            'Me.LblAdvPayment.Caption = IIf(IsNull(RsTest("AdvPayment").value), "", RsTest("AdvPayment").value)
        
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
            LblDiscount.Caption = IIf(IsNull(RsTest("Discount").value), "", RsTest("Discount").value)
            Me.LblAdvPayment.Caption = IIf(IsNull(RsTest("AdvPayment").value), "", RsTest("AdvPayment").value)
        
            Set RsPartDetails = New ADODB.Recordset
            StrSQL = "Select * From InstallMentDetails Where PartID=" & LngPartID
            RsPartDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            'fill Installments Grid
            If Not (RsPartDetails.BOF Or RsPartDetails.EOF) Then
                RsPartDetails.MoveFirst

                With Me.FgInstallments
                    .Rows = .FixedRows + RsPartDetails.RecordCount

                    For i = .FixedRows To .Rows - 1
                        .TextMatrix(i, .ColIndex("QestID")) = IIf(IsNull(RsPartDetails("QestID").value), "", RsPartDetails("QestID").value)
                        .TextMatrix(i, .ColIndex("Serial")) = IIf(IsNull(RsPartDetails("QeqtNum").value), "", RsPartDetails("QeqtNum").value)
                        .TextMatrix(i, .ColIndex("QeqtNum")) = IIf(IsNull(RsPartDetails("QeqtNum").value), "", RsPartDetails("QeqtNum").value)
                    
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

            showComm
        Else
            CmdINSTALLMENT.Enabled = False
            CmdINSTALLMENT.Caption = " ﬁ”Ìÿ «·ﬁÌ„… «·¬Ã·…"
    
        End If

    End If
fillExpensesFactoryGrid
ISButton1_Click
    '›« Ê—… «·Œœ„« 
    If CheckBillType = 0 Then
        Command2.backcolor = &HC0C0C0
        Command2.Enabled = False

        If SystemOptions.UserInterface = ArabicInterface Then
            Command2.Caption = "  ›« Ê—… Œœ„«  Ê·Ì” ·Â« ”‰œ ’—› "
        Else
            Command2.Caption = " Services Invoices"
        
        End If

        Exit Sub

    End If

    DoEvents
        
    Exit Sub

ErrTrap:
    Resume
    Screen.MousePointer = vbDefault
End Sub

Private Sub Undo()
    Dim Msg As String

    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                
                clear_all Me
                Me.TxtModFlg.Text = "R"
                XPBtnMove_Click (1)
            End If

        Case "E"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ⁄œÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                rs.find "Transaction_ID='" & val(XPTxtBillID.Text) & "'", , adSearchForward, adBookmarkFirst

                If rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.Text = "R"
                    Exit Sub
                End If

                If Not rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.Text = "R"
                    Retrive
                End If
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Sub SaveHeader(Optional NoteSerial1 As String, Optional ByRef TransectionID As Double, Optional CusID As Double)
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim StrSQL As String
    If NoteSerial1 = "" Then
    NoteSerial1 = Voucher_coding(val(val(Me.dcBranch.BoundText)), XPDtbBill.value, 7, 170, , 21, DCPreFix.Text, val(DCboStoreName.BoundText))
    
                If NoteSerial1 = "error" Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›…   ›« Ê—… „»Ì⁄«  ÃœÌœ… ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                Else
                               
                    If NoteSerial1 = "" Then
                        MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ›« Ê—… «·„»Ì⁄«   ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                    Else
                    End If
                End If
    End If
     StrSQL = "select * From Transactions where 1=-1"
     rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        TransectionID = CStr(new_id("Transactions", "Transaction_ID", "", True))
     rs2.AddNew
    rs2("Transaction_ID").value = TransectionID
    rs2("FlagFarm").value = 1
    rs2("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
    rs2("Doctype").value = IIf(Me.DCDocTypes.BoundText = "", Null, val(DCDocTypes.BoundText))
    rs2("DTPickerAccFrom").value = DTPickerAccFrom.value
    rs2("DTPickerAccTo").value = DTPickerAccTo.value
    rs2("CarId").value = IIf(Me.DCCar.BoundText = "", Null, (Me.DCCar.BoundText))
    rs2("DriverId").value = IIf(Me.DCDriver.BoundText = "", Null, (Me.DCDriver.BoundText))
    rs2("NoteSerial1").value = NoteSerial1
    If CboPayMentType.ListIndex = 0 Then
        rs2("BoxID").value = IIf(DcboBox.BoundText = "", Null, val(DcboBox.BoundText))
    Else
        rs2("BoxID").value = Null
      
    End If
      
    rs("Transaction_Serial").value = IIf(Trim(Me.TxtTransSerial.Text) = "", "", Trim(Me.TxtTransSerial.Text))
    rs2("Transaction_Date").value = XPDtbBill.value
    rs2("Transaction_Type").value = 21
    rs2("UserID").value = user_id
    rs2("nots").value = ""
    rs2("Currency_id").value = IIf(DcCurrency.BoundText = "", Null, val(DcCurrency.BoundText))
    rs2("Currency_rate").value = IIf(Not IsNumeric(txt_Currency_rate.Text), 1, txt_Currency_rate.Text)
    If XPCboDiscountType.ListIndex = -1 Then
        rs2("Trans_DiscountType").value = 0
    Else
        rs2("Trans_DiscountType").value = val(XPCboDiscountType.ListIndex)
    End If
    rs2("SupplerID").value = val(DcbFarm.BoundText)
    rs2("TypComm").value = val(DcbTypComm.ListIndex)
    rs2("NetValueComm").value = val(TxtNetValueComm.Text)
    rs2("ValueComm").value = val(TxtValueComm.Text)
    rs2("Trans_Discount").value = IIf(XPTxtDiscountVal.Text = "", Null, val(XPTxtDiscountVal.Text))
    rs2("CusID").value = CusID
    rs2("FarmID").value = val(DcbFarm.BoundText)
    rs2("GVAT").value = val(TxtGVAT.Text)
    rs2("SAlGValue").value = val(TxtSAlGValue.Text)
    rs2("FactoryExpenses").value = val(TXTFactoryExpenses.Text)
    rs2("GTotal").value = val(TxtGTotal.Text)
    If ChLoadVAT.value = vbChecked Then
    rs2("ChLoadVAT").value = 1
    Else
    rs2("ChLoadVAT").value = 0
    End If
    
    
    rs2("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, val(DCboStoreName.BoundText))
    rs2("order_no") = IIf(TXTOrDer_no.Text = "", Null, val(TXTOrDer_no.Text))
    rs2("PurchaseBill") = IIf(TxtPurchaseBill.Text = "", Null, val(TxtPurchaseBill.Text))

    If CboPayMentType.ListIndex = -1 Then
        rs2("PaymentType").value = 0
    Else
        rs2("PaymentType").value = val(CboPayMentType.ListIndex)
    End If

    rs2("TaxFound").value = IIf(XPChkTAX.value = Checked, True, False)
    rs2("TaxValue").value = IIf(XPTxtTaxValue.Text = "", Null, val(XPTxtTaxValue.Text))
    rs2("Emp_ID").value = IIf(DcboEmp.BoundText = "", Null, DcboEmp.BoundText)

    'ChkInstall 11 10 2012
    If ChkInstall.value = vbChecked Then
        rs2("ChkInstall").value = 1
    Else
        rs2("ChkInstall").value = 0
    End If

    If Me.CboSaleType.ListIndex = 0 Or Me.CboSaleType.ListIndex = -1 Then
        rs2("SaleType").value = 0
    Else
        rs2("SaleType").value = 1
    End If

    If Trim$(Me.TxtCashCustomerName.Text) <> "" Then
        rs2("CashCustomerName").value = Trim$(Me.TxtCashCustomerName.Text)
    Else
        rs2("CashCustomerName").value = Null
    End If

    rs2("TransactionComment").value = IIf(Trim$(TxtBillComment.Text) = "", Null, Trim$(TxtBillComment.Text))

    '÷—»Ì… Œ’„ Ê≈÷«›…
    If ChkTaxAdd.value = vbChecked And val(Me.TxtTaxAddValue.Text) > 0 Then
        rs2("TaxAddValue").value = val(Me.TxtTaxAddValue.Text)
    Else
        rs2("TaxAddValue").value = 0
    End If

    '÷—»Ì… œ„€…
    If ChkTaxStamp.value = vbChecked And val(Me.TxtTaxStampValue.Text) > 0 Then
        rs2("TaxStampValue").value = val(Me.TxtTaxStampValue.Text)
    Else
        rs2("TaxStampValue").value = 0
    End If

    '÷—»Ì… Œœ„…
    If ChkTaxSerivce.value = vbChecked And val(Me.TxtTaxServiceValue.Text) > 0 Then
        rs2("TaxServiceValue").value = val(Me.TxtTaxServiceValue.Text)
    Else
        rs2("TaxServiceValue").value = 0
    End If

    '»Ì«‰«  ÃœÌœ…
    rs2("PaymentNetid").value = IIf(DCPaymentNet.BoundText = "", Null, DCPaymentNet.BoundText)
    rs2("NetValue").value = IIf(TxtNetValue.Text = "", Null, val(TxtNetValue.Text))
    rs2("PayedValue").value = IIf(TxtPayedValue.Text = "", Null, val(TxtPayedValue.Text))
    rs2("RemainValue").value = IIf(TxtRemainValue.Text = "", Null, val(TxtRemainValue.Text))
  
    rs2("ManualNo1").value = IIf(TxtManualNo1.Text = "", Null, val(TxtManualNo1.Text))
    rs2("ManualNo2").value = IIf(TxtManualNo2.Text = "", Null, val(TxtManualNo2.Text))
  
    If BillBasedOn(0).value = True Then
        rs2("BillBasedOn").value = 0
    ElseIf BillBasedOn(1).value = True Then
        rs2("BillBasedOn").value = 1
    ElseIf BillBasedOn(2).value = True Then
        rs2("BillBasedOn").value = 2
    End If

    rs2.update
End Sub

Sub SaveBillByCustomerLine()
Dim sql As String
Dim rs2 As ADODB.Recordset
Dim TotalDiscountPerLine As Variant
Dim TotalBillDiscount As Double
Dim RsTemp As ADODB.Recordset
Dim SngTemp As Variant
Dim StrSQL As String
Dim RowNum As Integer
Dim NoteSerial1 As String
Dim TransectionID As Double

Dim RSTransDetails As ADODB.Recordset
Set RSTransDetails = New ADODB.Recordset
 StrSQL = "select * From Transaction_Details where 1=-1"
            RSTransDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    For RowNum = 1 To FG.Rows - 1
    TransectionID = 0
    If val(FG.TextMatrix(RowNum, FG.ColIndex("CusID2"))) <> 0 Then
    'TransectionID = CheckGetTransID(val(FG.TextMatrix(RowNum, FG.ColIndex("CusID2"))), NoteSerial1)
    End If
    If TransectionID = 0 Then
    SaveHeader NoteSerial1, TransectionID, val(FG.TextMatrix(RowNum, FG.ColIndex("CusID2")))
    End If
    
            RSTransDetails.AddNew
            RSTransDetails("OrderArrivalDate").value = IIf(Not IsDate(FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate"))), Null, FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")))
            RSTransDetails("Transaction_Date").value = IIf(Not IsDate(FG.TextMatrix(RowNum, FG.ColIndex("Transaction_Date"))), Null, FG.TextMatrix(RowNum, FG.ColIndex("Transaction_Date")))
            RSTransDetails("FoxyNo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")))
           ' RSTransDetails("CusID3").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("CusID3")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("CusID3")))
            RSTransDetails("order_no").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("order_no")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("order_no")))
            RSTransDetails("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
            RSTransDetails("Transaction_ID").value = TransectionID
            RSTransDetails("Transaction_ID2").value = val(XPTxtBillID.Text)
            RSTransDetails("FarmID2").value = val(DcbFarm.BoundText)
            RSTransDetails("CusID2").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("CusID2")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("CusID2")))
            
            RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
           ' RSTransDetails("CusID2").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("CusID2")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("CusID2"))))

            If Not FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "" Then
                StrSQL = "select * From TblItems where ItemID=" & FG.TextMatrix(RowNum, FG.ColIndex("Name"))
                Set RsTemp = New ADODB.Recordset
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (RsTemp.EOF Or RsTemp.BOF) Then
                    If RsTemp("HaveSerial").value = True Then
                        RSTransDetails("ItemSerial").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Serial")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("Serial")))
                    End If
                End If

                RsTemp.Close
            End If

            RSTransDetails("ShowPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
            RSTransDetails("ItemDiscountType").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType"))))
            RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
            RSTransDetails("Vat").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Vat")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("Vat"))))
            RSTransDetails("Vatyo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Vatyo")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("Vatyo"))))
            RSTransDetails("ItemDiscount").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal"))))
            
            RSTransDetails("guaranteeTime").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime"))))
            RSTransDetails("CostTransID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("PofTransID")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("PofTransID"))))
            RSTransDetails("ItemProfit").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemProfit")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemProfit"))))
        
            RSTransDetails("UnitID").value = IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
          
            If SystemOptions.TypicalProduction = False Then
  
                RSTransDetails("CostPrice").value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(RowNum, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Me.Text1.Text), RSTransDetails("UnitID").value)

                If RSTransDetails("CostPrice").value = 0 Then
                    RSTransDetails("CostPrice").value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(RowNum, FG.ColIndex("Code")), 0, , , LastPurPriceType, , , XPDtbBill.value, val(Me.Text1.Text), RSTransDetails("UnitID").value)
                    
                End If
                  
            Else
                RSTransDetails("CostPrice").value = 0
            
            End If

            FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = RSTransDetails("CostPrice").value
              
            RSTransDetails("SavedItemType").value = val(FG.TextMatrix(RowNum, FG.ColIndex("ItemType")))
               
            RSTransDetails("ShowQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
            Dim cnt As Double
            cnt = FG.TextMatrix(RowNum, FG.ColIndex("Count"))

            RSTransDetails("Remarks").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Remarks")) = ""), Null, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("Remarks"))))
                
            RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
            RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
            RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
            
            '«·ÊÕœ« 
           
            Dim RsUnitData As ADODB.Recordset
            Dim LngCurItemID As Long
            Dim LngUnitID As Long
            Dim DblQty As Double
        
            LngCurItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
            LngUnitID = val(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
            DblQty = val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))

            StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            Set RsUnitData = New ADODB.Recordset
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                RSTransDetails("OpeningSalesQty").value = RSTransDetails("Quantity").value
                RSTransDetails("OpeningSalesValue").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Valu")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Valu"))))
                RSTransDetails("Price").value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
            
            End If

            SngTemp = SngTemp + (val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice"))) * RSTransDetails("quantity").value)
         
            If Me.XPCboDiscountType.ListIndex = 1 Then
                TotalBillDiscount = IIf(XPTxtDiscountVal.Text = "", Null, (XPTxtDiscountVal.Text))
                     
            ElseIf XPCboDiscountType.ListIndex = 2 Then

                If XPTxtDiscountVal.Text <> "" Then
                    TotalBillDiscount = IIf(XPTxtDiscountVal.Text = "", Null, (XPTxtDiscountVal.Text)) * val(LblTotalAll.Caption) / 100
                             
                Else
                    TotalBillDiscount = 0
                End If
            End If

            TotalDiscountPerLine = ((RSTransDetails("SHOWprice") * RSTransDetails("SHOWQTY")) / LblTotalAll.Caption) * (TotalBillDiscount)
            RSTransDetails("TotalDiscountPerLine") = Round(TotalDiscountPerLine, 20)
                 
            RSTransDetails("ProductionDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate"))), "DD/mm/YYYY"))
            RSTransDetails("ExpiryDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate"))), "DD/mm/YYYY"))
            RSTransDetails("LotNO").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("LotNO")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("LotNO")))
        
            RSTransDetails("GranteeType").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("GranteeType")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("GranteeType")))
            RSTransDetails("GranteeStartDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("GranteeStartDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("GranteeStartDate"))), "DD/mm/YYYY"))
            RSTransDetails("GranteeEndDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("GranteeEndDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("GranteeEndDate"))), "DD/mm/YYYY"))
            RSTransDetails("RegularMaintenancedates").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("RegularMaintenancedates")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("RegularMaintenancedates")))
            RSTransDetails.update
            '-------------
sql = "Select * from  TransactionValueAdded where 1=-1"
Set rs2 = New ADODB.Recordset
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
rs2.AddNew
rs2("Transaction_ID2").value = val(XPTxtBillID.Text)
rs2("Transaction_ID").value = TransectionID
rs2("Transaction_Type").value = 21
rs2("ItemID").value = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
rs2("Vatyo").value = val(FG.TextMatrix(RowNum, FG.ColIndex("Vatyo")))
rs2("Vat").value = val(FG.TextMatrix(RowNum, FG.ColIndex("Vat")))
rs2("Valu").value = FG.TextMatrix(RowNum, FG.ColIndex("Price"))
rs2("selectd").value = 1
rs2.update

Cn.Execute "update Transactions set Transaction_NetValue=" & GetTotalValue(TransectionID) + GetTotalVAT(TransectionID) & "  ,VAT=" & GetTotalVAT(TransectionID) & " where Transaction_ID =" & TransectionID & " and CusID=" & val(FG.TextMatrix(RowNum, FG.ColIndex("CusID2"))) & ""
    Next RowNum
 

End Sub
Function GetTotalVAT(Optional Transaction_ID As Double) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     SUM(Vat) AS Valu"
sql = sql & " From dbo.TransactionValueAdded"
sql = sql & " Where (Transaction_ID = " & Transaction_ID & ")"
sql = sql & " GROUP BY Transaction_ID"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetTotalVAT = IIf(IsNull(rs2("Valu").value), 0, rs2("Valu").value)
Else
GetTotalVAT = 0
End If
End Function

Function GetTotalValue(Optional Transaction_ID As Double) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     SUM(ShowQty*showPrice) AS Valu"
sql = sql & " From dbo.Transaction_Details"
sql = sql & " Where (Transaction_ID = " & Transaction_ID & ")"
sql = sql & " GROUP BY Transaction_ID"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetTotalValue = IIf(IsNull(rs2("Valu").value), 0, rs2("Valu").value)
Else
GetTotalValue = 0
End If
End Function
Function CheckGetTransID(Optional CusID As Double, Optional ByRef NoteSerial1 As String) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     Transaction_ID ,NoteSerial1"
sql = sql & " From dbo.transactions"
sql = sql & " WHERE     (Transaction_Date = " & SQLDate(XPDtbBill.value, True) & ") AND (CusID = " & CusID & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
CheckGetTransID = IIf(IsNull(rs2("Transaction_ID").value), 0, rs2("Transaction_ID").value)
NoteSerial1 = IIf(IsNull(rs2("NoteSerial1").value), "", rs2("NoteSerial1").value)
Else
CheckGetTransID = 0
NoteSerial1 = ""
End If
End Function
Private Sub Del_TransAction()
    Dim Msg As String
    Dim StrSqlDel As String
    Dim RsTest As ADODB.Recordset
    Dim StrSQL As String
    Dim IntRes As Integer
    Dim i As Integer
    Dim RowNum As Long
    Dim BegainTrans As Boolean
    On Error GoTo ErrTrap

    If XPTxtBillID.Text = "" Then
        clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
            MsgBox Msg, vbOKOnly + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
    End If

    '⁄„·Ì«  «·’Ì«‰… «·„— »ÿ… »«·›« Ê—…
    StrSQL = "select * From MaintenanceJuncTransaction Where Transaction_ID=" & Trim(XPTxtBillID.Text)
    Set RsTest = New ADODB.Recordset
    RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTest.EOF Or RsTest.BOF) Then
        Msg = "·ﬁœ  „ ≈Ã—«¡ »⁄÷ ⁄„·Ì«  «·’Ì«‰… ⁄·Ï Â–Â «·›« Ê—… Ê·« Ì„ﬂ‰ Õ–›Â«"
        Msg = Msg + "≈–« ﬂ‰   —€» ›Ì Õ–› »Ì«‰«  Â–Â «·›« Ê—…" & CHR(13)
        Msg = Msg + "ÌÃ» Õ–› ⁄„·Ì«  «·’Ì«‰… «·Œ«’… »Â«"
        MsgBox Msg, vbOKOnly + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If Me.CboPayMentType.ListIndex = 0 Then

        '›« Ê—… ‰ﬁœÌ…
        If CheckBoxAccount(val(Me.DcboBox.BoundText), val(Me.XPTxtValue(0).Text), XPDtbBill.value, False) = False Then
            Msg = "·‰ Ì„ﬂ‰ «·”„«Õ »Õ–› Â–« «·⁄„·Ì…..!!!"
            Msg = Msg & CHR(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï Õ”«»«  «·Œ“‰…"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "  √ﬂÌœ Õ–›    »Ì«‰«  Â–Â «·⁄„·Ì…" & CHR(13)
        ' Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
    Else
        Msg = " Confirm Delete  " & CHR(13)
        '     Msg = Msg + "do you new Operation?"
       
    End If
 
    IntRes = MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)

    If IntRes = vbYes Then
        If Not rs.RecordCount < 1 Then
            Cn.BeginTrans
            BegainTrans = True
                With GRID2
    For i = 1 To .Rows - 1
    If val(.TextMatrix(i, .ColIndex("Transaction_ID"))) <> 0 Then
    Cn.Execute " Update Transactions    set TransGorupID=null  where Transaction_ID =" & val(.TextMatrix(i, .ColIndex("Transaction_ID"))) & " "
    End If
    Next i
    End With
    DeletedCustomer
           StrSqlDel = "delete From TransactionValueAdded where Transaction_ID2=" & val(Me.XPTxtBillID.Text) 'Val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
            StrSqlDel = "delete From Notes where Transaction_ID=" & val(Me.XPTxtBillID.Text)  'Val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS  " & "Where DOUBLE_ENTREY_VOUCHERS.Transaction_ID=" & rs("Transaction_ID").value
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSqlDel = "delete From TblSalesGExpepenses where Transaction_ID=" & val(Me.XPTxtBillID.Text)  'Val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
            StrSqlDel = "delete From DOUBLE_ENTREY_VOUCHERS where Notes_ID=" & val(TxtNoteID111.Text)   'Val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
             StrSqlDel = "delete From Transaction_Details where Transaction_ID2=" & val(Me.XPTxtBillID.Text)
             Cn.Execute StrSqlDel, , adExecuteNoRecords
        For RowNum = 1 To FG.Rows - 1
        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" And val(FG.TextMatrix(RowNum, FG.ColIndex("CusID2"))) <> 0 Then
        If CheckCustomerTrans(val(FG.TextMatrix(RowNum, FG.ColIndex("CusID2"))), XPDtbBill.value) = False Then
        DeletInvoiceofCustomer val(FG.TextMatrix(RowNum, FG.ColIndex("CusID2"))), XPDtbBill.value
        End If
        End If
      Next RowNum
        
            '                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS  " & _
            '         "Where DOUBLE_ENTREY_VOUCHERS.Transaction_ID=" & get_transaction_id(rs("nots").value, 19)
            '         Cn.Execute StrSQL, , adExecuteNoRecords
                
            '         StrSQL = "Delete From Transactions  " & _
            '         "Where Transaction_ID=" & get_transaction_id(rs("nots").value, 19)
            '         Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "delete From Notes where noteid=" & val(TXTNoteID.Text)
    
            Cn.Execute StrSQL, , adExecuteNoRecords
            DeleteLinkTOIssueVoucher
            DeleteTransactiomsVoucher val(Text1.Text)
            CuurentLogdata ("D")

            If CboPOSBillType.ListIndex = 0 And val(LblStableID.Caption) <> -1 Then
                Cn.Execute "update Stables set Status =Null   where id=" & val(LblStableID.Caption)
       
            End If
       
            rs.delete
            Cn.CommitTrans
            BegainTrans = False
            Msg = " „  ⁄„·Ì… «·Õ–› "
            MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
    Msg = Msg & CHR(13) & Err.description
    Msg = Msg & CHR(13) & Err.Source
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title

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
            .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… »Ì⁄ ÃœÌœ…" & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F12 OR Enter", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…" & Wrap & "„›« ÌÕ «·«Œ ’«— F6", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  ⁄„·Ì… «·»Ì⁄" & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F11", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  ⁄„·Ì… «·»Ì⁄ «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F10", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·»Ì⁄" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F9", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  ⁄„·Ì… »Ì⁄" & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F8", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄„·Ì… »Ì⁄" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F7", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & "  ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— Ctrl + X", BolRtl
        End With
    
        With TTP
            .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnNewClients, "≈÷«›… ⁄„Ì· ÃœÌœ ..." & Wrap & "· ”ÃÌ· »Ì«‰«  ⁄„Ì· ÃœÌœ" & Wrap & " «÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F5", BolRtl
        End With
    
        With TTP
            .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·»Ì⁄", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, BolRtl
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        BolRtl = False

        With TTP
            .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "New..." & Wrap & "Click here to add new Bill Invoice" & Wrap & "" & Wrap & "Shortcut (Enter Or F12)", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "Print..." & Wrap & "Print this Bill Invoice" & Wrap & "" & Wrap & "Shortcut (F6)", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), "Edit..." & Wrap & "Edit this Bill Invoice Record" & Wrap & "  " & Wrap & "Shortcut (F11)", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Save..." & Wrap & "Save the New Bill Invoice Or Save the edit" & Wrap & "in the current Bill Invoice" & Wrap & "" & Wrap & "Shortcut (F10)", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), "Undo..." & Wrap & "Undo in the New Bill Invoice" & Wrap & "Or Undo in the Editing" & Wrap & "" & Wrap & "Shortcut (F9)", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Delete..." & Wrap & "Delete this current Bill Invoice" & Wrap & "" & Wrap & "Shortcut (F8)", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "Search..." & Wrap & "Click here to display the search" & Wrap & "Screen" & Wrap & "Shortcut (F7)", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Exit..." & Wrap & "Close this Window", BolRtl
        End With
    
        With TTP
            .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnNewClients, "Add New Customer...." & Wrap & "To add New Customer Click here..." & Wrap & "Shortcut (F5)", BolRtl
        End With
    
        With TTP
            .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "First..." & Wrap & "Move to first Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "Previous..." & Wrap & "Move to Previous Record" & Wrap & " , BolRTL"
        End With

        With TTP
            .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "Next..." & Wrap & "Move to Next Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "Last..." & Wrap & "Move to Last Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hwnd, "Bill Invoice", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "Help..." & Wrap & "to View Help Files" & Wrap & "click Here" & Wrap & "Shortcut(F1)" & Wrap, BolRtl
        End With

    End If

    Exit Sub
ErrTrap:
End Sub

Function SaveGranteeData()
    Dim RsgGrantee    As New ADODB.Recordset
    Dim strInputString As String
    Dim strFilterText As String
    Dim astrSplitItems() As String
    Dim astrFilteredItems() As String
    Dim strFilteredString As String
    Dim intX As Integer
    Dim AllDate As String
    Dim RowNum As Integer
    strFilterText = ","
    Set RsgGrantee = New ADODB.Recordset
    Cn.Execute "delete TBLRegularMaint   where Transaction_ID= " & val(Me.XPTxtBillID.Text)
    RsgGrantee.Open "TBLRegularMaint", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    For RowNum = 1 To FG.Rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            If FG.TextMatrix(RowNum, FG.ColIndex("RegularMaintenancedates")) <> "" Then
                AllDate = FG.TextMatrix(RowNum, FG.ColIndex("RegularMaintenancedates"))
                astrSplitItems = Split(AllDate, strFilterText)
         
                For intX = 0 To UBound(astrSplitItems)
                        
                    If IsDate(astrSplitItems(intX)) Then
                        RsgGrantee.AddNew
                        RsgGrantee("DateOfRegularMaint").value = Format$(astrSplitItems(intX), "dd/mm/yyyy")
                        RsgGrantee("itemid").value = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                        RsgGrantee("Transaction_ID").value = val(Me.XPTxtBillID.Text)
                        RsgGrantee("GranteeType").value = val(FG.TextMatrix(RowNum, FG.ColIndex("GranteeType")))
                        RsgGrantee("GranteeStartDate").value = FG.TextMatrix(RowNum, FG.ColIndex("GranteeStartDate"))
                        RsgGrantee("GranteeEndDate").value = FG.TextMatrix(RowNum, FG.ColIndex("GranteeEndDate"))
                        RsgGrantee("ItemSerial").value = FG.TextMatrix(RowNum, FG.ColIndex("Serial"))
                  RsgGrantee("Count").value = FG.TextMatrix(RowNum, FG.ColIndex("Count"))

                        RsgGrantee("Done").value = 0
                        RsgGrantee.update
                    End If
                       
                Next intX
                    
            End If

        End If

    Next RowNum

End Function
Function CheckCustomerInGrid(Optional ByRef Row As Integer) As Boolean
Dim i As Integer
With FG
CheckCustomerInGrid = False
For i = 1 To .Rows - 1
If .TextMatrix(i, .ColIndex("Code")) <> "" Then
If val(.TextMatrix(i, .ColIndex("CusID2"))) = 0 Then
Row = i
CheckCustomerInGrid = True
Exit Function
End If
End If
Next i
End With
End Function
Sub DeletedCustomer()
Dim RowNum As Integer
Dim StrSqlDel As String
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     CusID2, Transaction_ID2"
sql = sql & " From dbo.Transaction_Details"
sql = sql & " GROUP BY CusID2, Transaction_ID2"
sql = sql & " Having (Not (CusID2 Is Null)) And (Transaction_ID2 = " & val(XPTxtBillID.Text) & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    StrSqlDel = "delete From TransactionValueAdded where Transaction_ID2=" & val(Me.XPTxtBillID.Text) 'Val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        StrSqlDel = "delete From Transaction_Details where Transaction_ID2=" & val(Me.XPTxtBillID.Text)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
If rs2.RecordCount > 0 Then
rs2.MoveFirst
       For RowNum = 1 To rs2.RecordCount
        If CheckCustomerTrans(IIf(IsNull(rs2("CusID2").value), 0, rs2("CusID2").value), XPDtbBill.value) = False Then
        DeletInvoiceofCustomer IIf(IsNull(rs2("CusID2").value), 0, rs2("CusID2").value), XPDtbBill.value
        End If
        
     rs2.MoveNext
      Next RowNum
End If
End Sub
Private Sub SaveData()
    Dim Msg As String
    Dim RowNum As Integer
    Dim RSTransDetails As ADODB.Recordset
    Dim RsNotes As ADODB.Recordset
    Dim RsTemp      As New ADODB.Recordset
    
    Dim RsTest      As New ADODB.Recordset
    Dim RsRepeat    As ADODB.Recordset
    Dim RsDetalis   As ADODB.Recordset
    Dim StrSQL      As String
    Dim StrSqlDel   As String
    Dim note_id As Long
    Dim TransBegine As Boolean
    Dim BolTemp As Boolean
    Dim LnItemID As Long
    Dim i As Integer
    Dim DblNotesTotal As Double
    Dim SngTemp As Variant
    Dim usedaccount As Integer
    Dim TotalDiscountPerLine As Variant
    Dim TotalBillDiscount As Double
    'On Error GoTo ErrTrap

    Me.FG.FinishEditing True

    DoEvents

    Screen.MousePointer = vbArrowHourglass

    '«· √ﬂœ „‰ ⁄œ„  ﬂ—«— —ﬁ„ «·›« Ê—…
    If Voucher_coding(val(my_branch), XPDtbBill.value, 40, 170, , 40, DCPreFix.Text) = "" Then
        If Me.TxtModFlg.Text = "N" Then
    
            BolTemp = UniqueNoteSerial1(Trim(Me.TxtNoteSerial1.Text), 40, , val(dcBranch.BoundText))
        ElseIf Me.TxtModFlg.Text = "E" Then
            BolTemp = UniqueNoteSerial1(Trim(Me.TxtNoteSerial1.Text), 40, val(Me.XPTxtBillID.Text), val(dcBranch.BoundText))
        End If
 
        If BolTemp = False Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "—ﬁ„ «·›« Ê—… „”Ã· „”»ﬁ« ›Ï «·»—‰«„Ã.." & CHR(13)
                Msg = Msg & "Ê·«Ì„ﬂ‰  ﬂ—«— —ﬁ„ «·›« Ê—…"
            Else
                Msg = "This Bill No Already Exist" & CHR(13)
        
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtNoteSerial1.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
     
    End If
    
    '‰Â«Ì… «· √ﬂœ
    
    If Trim(Me.TxtTransSerial.Text) = "" Then
        Msg = "ÌÃ» ≈œŒ«· —ﬁ„ «·›« Ê—…...!!"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtTransSerial.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    Else

        If Me.TxtModFlg.Text = "N" Then
    
            BolTemp = UniqueTransSerial(Trim(Me.TxtTransSerial.Text), 2)
        ElseIf Me.TxtModFlg.Text = "E" Then
            BolTemp = UniqueTransSerial(Trim(Me.TxtTransSerial.Text), 2, val(Me.XPTxtBillID.Text))
        End If

        BolTemp = True

        If BolTemp = False Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "—ﬁ„ «·›« Ê—… „”Ã· „”»ﬁ« ›Ï «·»—‰«„Ã.." & CHR(13)
                Msg = Msg & "Ê·«Ì„ﬂ‰  ﬂ—«— —ﬁ„ «·›« Ê—…"
            Else
                Msg = "This Bill No Already Exist" & CHR(13)
        
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtTransSerial.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If

    If DcCurrency.BoundText = "" Then
    
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "«Œ — «·⁄„·… «Ê·« "
        Else
            Msg = "Select Currency First"
    
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DcCurrency.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    
    End If
   
'    If val(DBCboClientName.BoundText) = 0 Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'            Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·⁄„Ì·"
'        Else
'            Msg = "Select Customer First"
'        End If
'
'        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'        'DBCboClientName.SetFocus
'        SendKeys "{F4}"
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If

   ' If Trim(DcboEmp.BoundText) = "" Then
   '     If SystemOptions.UserInterface = ArabicInterface Then
   '         Msg = "ÌÃ»  ÕœÌœ «”„ «·»«∆⁄/«·„‰œÊ»..!!!"
   '     Else
   '         Msg = "Must Specify SalesPerson/Saller..!!!"
   '     End If
'
'        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'        DcboEmp.SetFocus
'        SendKeys "{F4}"
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    End If

    If XPDtbBill.value = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ  «—ÌŒ «·»Ì⁄"
        Else
            Msg = "Specify Date"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPDtbBill.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If CboPayMentType.ListIndex = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ ÿ—Ìﬁ… «·œ›⁄"
        Else
            Msg = "Specify Payment Method"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CboPayMentType.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
 
    '----------------------------------------------
    If val(Me.XPTxtValue(1).Text) > 0 Then
        If ChkInstall.value = vbChecked Then
            If val(Me.LblInstallTotal.Caption) = 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» Õ”«» «·√ﬁ”«ÿ ﬁ»· ⁄„·Ì… «·Õ›Ÿ..!!!"
                Else
            
                    Msg = "Must Calculate Installment Before Save..!!!"
                End If

                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.XPTab301.CurrTab = 1
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            If val(Me.LblInstallTotal.Caption) <> val(Me.XPTxtValue(1).Text) Then
                Me.XPTxtValue(1).Text = val(Me.LblInstallTotal.Caption)
            End If
            
        End If
    End If

    '-----------------------------------------
    If XPChkPayType(2).value = vbChecked Then
        If val(Me.lbl(18).Caption) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈œŒ«· «·‘Ìﬂ«  ﬁ»· ⁄„·Ì… «·Õ›Ÿ..!!!"
            Else
                Msg = "Must Enter Cheque Before Save..!!!"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.XPTab301.CurrTab = 1
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        If Dcbanks.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = Msg + "ÌÃ»  ÕœÌœ «”„ «·»‰ﬂ Â–« «·Œ’„ " & CHR(13)
            Else
                Msg = Msg + "Specify Bank Name " & CHR(13)
            End If
        
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Dcbanks.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
    
            Dim rsbank As New ADODB.Recordset
            Set rsbank = New ADODB.Recordset
            rsbank.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
       
            If Not (rsbank.EOF Or rsbank.BOF) Then
                If rsbank!banks_Accounts = True Then
                    bank_account = get_bank_Account(val(Me.Dcbanks.BoundText), "Account_Code1")
                Else
                    bank_account = get_bank_Account(val(Me.Dcbanks.BoundText), "Account_Code")
                End If
            End If
        
        End If
    End If

    If XPChkTAX.value = Checked Then
        If XPTxtTaxValue.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» «œŒ«· ﬁÌ„… ÷—Ì»… «·„»Ì⁄« "
            Else
                Msg = "Enter Sales Tax"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPTxtTaxValue.SetFocus
            FG.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If

    If XPCboDiscountType.ListIndex = 1 Or XPCboDiscountType.ListIndex = 2 Then
        If XPTxtDiscountVal.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "≈–« ﬂ«‰ Â‰«ﬂ Œ’„ ⁄·Ï «·›« Ê—… " & CHR(13)
                Msg = Msg + "ÌÃ»  ÕœÌœ ﬁÌ„… Â–« «·Œ’„ " & CHR(13)
                Msg = Msg + "√Ê √Œ Ì«— ·« ÌÊÃœ Œ’„ "
            Else
                Msg = Msg + " Must Enter Discount Value " & CHR(13)
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPCboDiscountType.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If

    '--------------------------------
    '«·ﬂ‘› ⁄·Ï „œÌÊ‰Ì… «·⁄„Ì· '
'    If val(Me.DBCboClientName.BoundText) <> 1 Or val(Me.DBCboClientName.BoundText <> 2) Then
'        If Me.CboPayMentType.ListIndex = 1 Then
'            If val(Me.XPTxtValue(1).Text) > 0 Then
'                If CheckCusCredit(val(Me.DBCboClientName.BoundText), val(Me.XPTxtValue(1).Text), 0) = False Then
'                    Screen.MousePointer = vbDefault
'                    Exit Sub
'                End If
'            End If
'        End If
'    End If

    '--------------------------------
    
    Me.XPTab301.CurrTab = 0

    If CboPayMentType.ListIndex = 1 And Me.XPChkPayType(0).value = Unchecked And Me.XPChkPayType(2).value = Unchecked Then
 
    End If

    If BillBasedOn(1).value = True Then
        SelectedIssueVoucher = True
    Else
        SelectedIssueVoucher = False
    End If
 
    If NewGrid.CheckDataEntered = False Then
        Exit Sub
    End If

    If CheckCostForAllitems = False Then '«· √ﬂœ „‰ ÊÃÊœ  ﬂ·›… ·ﬂ· ««’‰«›
        Exit Sub
    End If

    '-------------------------------
    If NewGrid.Calculate(1, , False, True) = False Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    '-------------------------------
    If Me.XPChkPayType(0).value = vbChecked Then
        DblNotesTotal = val(Me.XPTxtValue(0).Text)
    End If

    If Me.XPChkPayType(1).value = vbChecked Then
        DblNotesTotal = DblNotesTotal + val(Me.XPTxtValue(1).Text)
    End If

    If Me.XPChkPayType(2).value = vbChecked Then
        DblNotesTotal = DblNotesTotal + val(Me.lbl(18).Caption)
    End If

    If CboPayMentType.ListIndex = 1 And Me.XPChkPayType(2).value = Unchecked Then
        XPChkPayType(1).value = 1
        '  XPTxtValue(1).text = Val(LblTotalAll.Caption)
        XPTxtValue(1).Text = val(LblTotal.Caption)

    Else

        If CboPayMentType.ListIndex = 1 And Me.XPChkPayType(2).value = vbChecked Then
            XPChkPayType(1).value = 0

        Else
            XPChkPayType(0).value = 1
            '  XPTxtValue(0).text = Val(LblTotalAll.Caption)
            XPTxtValue(0).Text = val(LblTotal.Caption)

        End If
    End If

    If Due_Date > DtpDelayDate.value Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "ÌÃ» «‰ ÌﬂÊ‰  «—ÌŒ «·«” Õﬁ«ﬁ «ﬂ»— „‰ «Ê Ì”«ÊÏ  «—ÌŒ «Œ— ﬁ”ÿ"
        Else
            MsgBox "Installment Date must be >= today date"
        End If

        Exit Sub
    End If

    CurrentVoucherNo = ""
    CurrentVoucherSerialNo = ""

    'Create big notes
    my_branch = val(Me.dcBranch.BoundText)
 
    my_branch = val(Me.dcBranch.BoundText)
    Dim currentSeial As String
    currentSeial = Voucher_coding(val(my_branch), XPDtbBill.value, 40, 170, , 40, DCPreFix.Text)

    If TxtNoteSerial1.Text = "" Then
        If currentSeial = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›…   ›« Ê—… „»Ì⁄«  „Ã„⁄Â ÃœÌœ… ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
        ElseIf currentSeial = "" Then
            MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ›« Ê—… «·„»Ì⁄«    «·„Ã„⁄Â ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
        Else
            TxtNoteSerial1.Text = currentSeial
        End If
    End If
     
    'Set RsNotesGeneral = New ADODB.Recordset
    'RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
    If Me.TxtModFlg.Text = "N" Then
        Me.oldtxtNoteSerial1.Text = Trim$(Me.TxtNoteSerial1.Text)
    Else
        'StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & Val(rs("Transaction_ID").value)
        'Cn.Execute StrSqlDel, , adExecuteNoRecords
        '        MsgBox Val(rs("Transaction_ID").value)
        '      StrSqlDel = "delete From Notes where Transaction_ID=" & val(Me.XPTxtBillID.text) ' Val(rs("Transaction_ID").value)
        '      Cn.Execute StrSqlDel, , adExecuteNoRecords
        
        '      StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & val(Me.XPTxtBillID.text)
        '      Cn.Execute StrSQL, , adExecuteNoRecords
        
        '      StrSQL = "delete From Notes where noteid=" & val(TXTNoteID.text)
        '      Cn.Execute StrSQL, , adExecuteNoRecords

        '   CurrentVoucherNo = GetVoucherGLNO(val(Text1.text), CurrentVoucherSerialNo)
        '     DeleteTransactiomsVoucher val(Text1.text)
        
        '    general_noteid = val(TXTNoteID.text)
    End If

    '      RsNotesGeneral.AddNew
    '      RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
    '      general_noteid = RsNotesGeneral("NoteID").value
    '      TXTNoteID.text = general_noteid
    '     ' RsNotesGeneral("Transaction_ID").value = Val(XPTxtBillID.text)
    '      RsNotesGeneral("NoteDate").value = XPDtbBill.value
    '      RsNotesGeneral("NoteType").value = 170
    '      RsNotesGeneral("Note_Value").value = val(LblTotal.Caption)
    my_branch = val(Me.dcBranch.BoundText)
                       
    '                     If TxtNoteSerial.text = "" Then
    '                         TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbBill.value)
    '                    End If
                      
    '                    If TxtNoteSerial1.text = "" Then
    '                                 TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbBill.value, 7, 170, , 21, DCPreFix.text)
    '                    End If
        
    '      RsNotesGeneral("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
    '      RsNotesGeneral("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
    '
    '      RsNotesGeneral("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '

    '      RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
    '      RsNotesGeneral("numbering_type1").value = sand_numbering_type(7) '  ›« Ê—… »Ì⁄
    '      RsNotesGeneral("sanad_year").value = year(XPDtbBill.value)
    '   RsNotesGeneral("sanad_month").value = Month(XPDtbBill.value)
    '   RsNotesGeneral("branch_no").value = val(Me.dcBranch.BoundText)
    'RsNotes("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
    '   RsNotesGeneral.update

    '---------------------------------
    Set RSTransDetails = New ADODB.Recordset
  '  RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
 StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    
    'Set RsNotes = New ADODB.Recordset
    'RsNotes.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    If SystemOptions.SysRegisterState <> Registered And SystemOptions.SysRegisterState <> DevelopVersion Then
        If rs.RecordCount > 50 Then
            'Exit Sub
        End If
    End If

    Screen.MousePointer = vbArrowHourglass
    Cn.BeginTrans
    TransBegine = True

    If Me.TxtModFlg.Text = "N" Then
        XPTxtBillID.Text = CStr(new_id("Transactions", "Transaction_ID", "", True))
        rs.AddNew
       
    ElseIf Me.TxtModFlg.Text = "E" Then
    DeletedCustomer
        
        StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(Me.XPTxtBillID.Text) 'Val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
         StrSqlDel = "delete From DOUBLE_ENTREY_VOUCHERS where Notes_ID=" & val(TxtNoteID111.Text)   'Val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        StrSqlDel = "delete From Notes where Transaction_ID=" & val(Me.XPTxtBillID.Text)  'Val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
           StrSqlDel = "delete From TblSalesGExpepenses where Transaction_ID=" & val(Me.XPTxtBillID.Text)  'Val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
    
 
       
        '    StrSqlDel = "delete From Notes where Transaction_ID=" & val(Me.XPTxtBillID.text)  'Val(rs("Transaction_ID").value)
        '    Cn.Execute StrSqlDel, , adExecuteNoRecords
        '    StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & val(Me.XPTxtBillID.text)
        '    Cn.Execute StrSQL, , adExecuteNoRecords
    End If

    rs("Transaction_ID").value = val(XPTxtBillID.Text)
    rs("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
    rs("Doctype").value = IIf(Me.DCDocTypes.BoundText = "", Null, val(DCDocTypes.BoundText))
    rs("DTPickerAccFrom").value = DTPickerAccFrom.value
    rs("DTPickerAccTo").value = DTPickerAccTo.value
    rs("CarId").value = IIf(Me.DCCar.BoundText = "", Null, (Me.DCCar.BoundText))
    rs("DriverId").value = IIf(Me.DCDriver.BoundText = "", Null, (Me.DCDriver.BoundText))
    ''///////////////////NewData Farm
    rs("FarmID").value = val(DcbFarm.BoundText)
    rs("BoardNo").value = TxtBoardNO.Text
    rs("Commission").value = val(Txtcommission.Text)
    rs("EmbarNo").value = TxtEmbarNo.Text
    rs("DriverName").value = TxtDriverName.Text
    ''////////////////
    ' rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
    rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.Text) = "", Null, Trim(Me.TxtNoteSerial1.Text))
    rs("Fullcode").value = IIf(DCPreFix.BoundText = "", Null, DCPreFix.Text) & IIf(Trim(TxtNoteSerial1.Text) = "", Null, TxtNoteSerial1.Text)
    rs("Prefix").value = IIf(DCPreFix.Text = "", Null, DCPreFix.Text)

    rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.Text) '

    If CboPayMentType.ListIndex = 0 Then
        rs("BoxID").value = IIf(DcboBox.BoundText = "", Null, val(DcboBox.BoundText))
    Else
        rs("BoxID").value = Null
      
    End If
      
    ' rs("NoteId").value = val(TXTNoteID.text)
    rs("Gratuity").value = val(TxtGratuity.Text)
    rs("VATYou").value = val(txtVatYou.Text)
    rs("VAT").value = val(TxtGVAT.Text)
    rs("VATNO").value = IIf(Trim(Me.TxtVATNO.Text) = "", Null, Trim(Me.TxtVATNO.Text))
    rs("Transaction_Serial").value = IIf(Trim(Me.TxtTransSerial.Text) = "", "", Trim(Me.TxtTransSerial.Text))
    rs("Transaction_Date").value = XPDtbBill.value
    rs("Transaction_Type").value = 40
    rs("UserID").value = user_id
    rs("nots").value = ""
     rs("CommissionType").value = optCommissionType(0)
    
    rs("Currency_id").value = IIf(DcCurrency.BoundText = "", Null, val(DcCurrency.BoundText))
    rs("Currency_rate").value = IIf(Not IsNumeric(txt_Currency_rate.Text), 1, txt_Currency_rate.Text)

    If XPCboDiscountType.ListIndex = -1 Then
        rs("Trans_DiscountType").value = 0
    Else
        rs("Trans_DiscountType").value = val(XPCboDiscountType.ListIndex)
    End If
    rs("SupplerID").value = val(DcbSuppler.BoundText)
    rs("TypComm").value = val(DcbTypComm.ListIndex)
    rs("NetValueComm").value = val(TxtNetValueComm.Text)
    rs("ValueComm").value = val(TxtValueComm.Text)
    rs("Trans_Discount").value = IIf(XPTxtDiscountVal.Text = "", Null, val(XPTxtDiscountVal.Text))
    rs("CusID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
    rs("GVAT").value = val(TxtGVAT.Text)
    rs("SAlGValue").value = val(TxtSAlGValue.Text)
    rs("FactoryExpenses").value = val(TXTFactoryExpenses.Text)
    rs("GTotal").value = val(TxtGTotal.Text)
    If ChLoadVAT.value = vbChecked Then
    rs("ChLoadVAT").value = 1
    Else
    rs("ChLoadVAT").value = 0
    End If
    
    
    rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, val(DCboStoreName.BoundText))
    rs("order_no") = IIf(TXTOrDer_no.Text = "", Null, val(TXTOrDer_no.Text))
    rs("PurchaseBill") = IIf(TxtPurchaseBill.Text = "", Null, val(TxtPurchaseBill.Text))

    If CboPayMentType.ListIndex = -1 Then
        rs("PaymentType").value = 0
    Else
        rs("PaymentType").value = val(CboPayMentType.ListIndex)
    End If

    rs("TaxFound").value = IIf(XPChkTAX.value = Checked, True, False)
    rs("TaxValue").value = IIf(XPTxtTaxValue.Text = "", Null, val(XPTxtTaxValue.Text))
    rs("Emp_ID").value = IIf(DcboEmp.BoundText = "", Null, DcboEmp.BoundText)

    'ChkInstall 11 10 2012
    If ChkInstall.value = vbChecked Then
        rs("ChkInstall").value = 1
    Else
        rs("ChkInstall").value = 0
    End If

    If Me.CboSaleType.ListIndex = 0 Or Me.CboSaleType.ListIndex = -1 Then
        rs("SaleType").value = 0
    Else
        rs("SaleType").value = 1
    End If

    If Trim$(Me.TxtCashCustomerName.Text) <> "" Then
        rs("CashCustomerName").value = Trim$(Me.TxtCashCustomerName.Text)
    Else
        rs("CashCustomerName").value = Null
    End If

    rs("TransactionComment").value = IIf(Trim$(TxtBillComment.Text) = "", Null, Trim$(TxtBillComment.Text))

    '÷—»Ì… Œ’„ Ê≈÷«›…
    If ChkTaxAdd.value = vbChecked And val(Me.TxtTaxAddValue.Text) > 0 Then
        rs("TaxAddValue").value = val(Me.TxtTaxAddValue.Text)
    Else
        rs("TaxAddValue").value = 0
    End If

    '÷—»Ì… œ„€…
    If ChkTaxStamp.value = vbChecked And val(Me.TxtTaxStampValue.Text) > 0 Then
        rs("TaxStampValue").value = val(Me.TxtTaxStampValue.Text)
    Else
        rs("TaxStampValue").value = 0
    End If

    '÷—»Ì… Œœ„…
    If ChkTaxSerivce.value = vbChecked And val(Me.TxtTaxServiceValue.Text) > 0 Then
        rs("TaxServiceValue").value = val(Me.TxtTaxServiceValue.Text)
    Else
        rs("TaxServiceValue").value = 0
    End If

    '»Ì«‰«  ÃœÌœ…
    rs("PaymentNetid").value = IIf(DCPaymentNet.BoundText = "", Null, DCPaymentNet.BoundText)
    rs("NetValue").value = IIf(TxtNetValue.Text = "", Null, val(TxtNetValue.Text))
    rs("PayedValue").value = IIf(TxtPayedValue.Text = "", Null, val(TxtPayedValue.Text))
    rs("RemainValue").value = IIf(TxtRemainValue.Text = "", Null, val(TxtRemainValue.Text))
  
    rs("ManualNo1").value = IIf(TxtManualNo1.Text = "", Null, val(TxtManualNo1.Text))
    rs("ManualNo2").value = IIf(TxtManualNo2.Text = "", Null, val(TxtManualNo2.Text))
  
    If BillBasedOn(0).value = True Then
        rs("BillBasedOn").value = 0
    ElseIf BillBasedOn(1).value = True Then
        rs("BillBasedOn").value = 1
    ElseIf BillBasedOn(2).value = True Then
        rs("BillBasedOn").value = 2
    End If
    
    '‰ﬁ«ÿ «·»Ì⁄
    rs("Printed").value = 1
   
    '   If CboPOSBillType.ListIndex = 0 And val(LblStableID.Caption) <> -1 Then
    '     Cn.Execute "update Stables set Status =Null   where id=" & val(LblStableID.Caption)
    '
    '      End If
       
    '‰ﬁ«ÿ «·»Ì⁄
    '   If CboPOSBillType.ListIndex = 0 Then
    '       rs("POSBillType").value = 0
    '        rs("STableID").value = val(LblStableID.Caption)
    '   Else
    '       rs("POSBillType").value = val(CboPOSBillType.ListIndex)
    '       rs("STableID").value = Null
    '   End If

    'rs("SessionD").value = lblSessionD

    rs.update

    For RowNum = 1 To FG.Rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then

            'Check Repeat Serial
            If FG.TextMatrix(RowNum, FG.ColIndex("Serial")) <> "" Then
                StrSQL = "select * From Transaction_Details where ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                StrSQL = StrSQL + " and Transaction_ID =" & XPTxtBillID.Text
                Set RsTemp = New ADODB.Recordset
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (RsTemp.EOF Or RsTemp.BOF) Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "«·”Ì—Ì«· «·Œ«’ »«·’‰›" & CHR(13)
                        Msg = Msg + FG.Cell(flexcpTextDisplay, RowNum, FG.ColIndex("name")) & CHR(13)
                        Msg = Msg + " „ √œŒ«·Â ·ﬁÿ⁄… √Œ—Ï ›Ì Â–Â «·›« Ê—…"
                    Else
                        Msg = "Item Serial " & CHR(13)
                        Msg = Msg + FG.Cell(flexcpTextDisplay, RowNum, FG.ColIndex("name")) & CHR(13)
                        Msg = Msg + "Duplicated in this Bill"
                    End If

                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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

            If IsEmpty(Me.FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))) Then
                If val(Me.FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))) = 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = " ÌÃ»  ÕœÌœ ÊÕœ… «·ﬂ„Ì… «·Œ«’… »«·’‰›" & CHR(13)
                    Else
                        Msg = " Must Select Item Unit For Item :" & CHR(13)
                    End If

                    Msg = Msg + FG.Cell(flexcpTextDisplay, RowNum, FG.ColIndex("Name")) & CHR(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    XPTab301.CurrTab = 0
                    FG.Row = RowNum
                    FG.Col = FG.ColIndex("UnitID")
                    FG.ShowCell RowNum, FG.ColIndex("UnitID")
                    FG.SetFocus
                    Screen.MousePointer = vbDefault
                    GoTo ErrTrap
                End If
            End If
            RSTransDetails.AddNew
            RSTransDetails("OrderArrivalDate").value = IIf(Not IsDate(FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate"))), Null, FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")))
            RSTransDetails("Transaction_Date").value = IIf(Not IsDate(FG.TextMatrix(RowNum, FG.ColIndex("Transaction_Date"))), Null, FG.TextMatrix(RowNum, FG.ColIndex("Transaction_Date")))
            RSTransDetails("FoxyNo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")))
            RSTransDetails("order_no").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("order_no")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("order_no")))
            RSTransDetails("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
            RSTransDetails("Transaction_ID").value = val(XPTxtBillID.Text)
            RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
            RSTransDetails("CusID2").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("CusID2")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("CusID2"))))
       
            'RSTransDetails("Quantity").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
            'RSTransDetails("ItemName").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Name")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Name"))))
         
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

            RSTransDetails("ShowPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
            RSTransDetails("ItemDiscountType").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType"))))
            RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
            RSTransDetails("Vat").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Vat")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("Vat"))))
            RSTransDetails("Vatyo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Vatyo")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("Vatyo"))))
            RSTransDetails("ItemDiscount").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal"))))
            
            RSTransDetails("guaranteeTime").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime"))))
            RSTransDetails("CostTransID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("PofTransID")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("PofTransID"))))
            RSTransDetails("ItemProfit").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemProfit")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemProfit"))))
        
            RSTransDetails("UnitID").value = IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
          
            If SystemOptions.TypicalProduction = False Then
  
                RSTransDetails("CostPrice").value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(RowNum, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Me.Text1.Text), RSTransDetails("UnitID").value)

                If RSTransDetails("CostPrice").value = 0 Then
                    RSTransDetails("CostPrice").value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(RowNum, FG.ColIndex("Code")), 0, , , LastPurPriceType, , , XPDtbBill.value, val(Me.Text1.Text), RSTransDetails("UnitID").value)
                    
                End If
                  
            Else
                RSTransDetails("CostPrice").value = 0
            
            End If

            FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = RSTransDetails("CostPrice").value
              
            RSTransDetails("SavedItemType").value = val(FG.TextMatrix(RowNum, FG.ColIndex("ItemType")))
               
            RSTransDetails("ShowQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
            Dim cnt As Double
            cnt = FG.TextMatrix(RowNum, FG.ColIndex("Count"))

            RSTransDetails("Remarks").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Remarks")) = ""), Null, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("Remarks"))))
                
            RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
            RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
            RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
            
            '«·ÊÕœ« 
           
            Dim RsUnitData As ADODB.Recordset
            Dim LngCurItemID As Long
            Dim LngUnitID As Long
            Dim DblQty As Double
        
            LngCurItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
            LngUnitID = val(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
            DblQty = val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))

            StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            Set RsUnitData = New ADODB.Recordset
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                RSTransDetails("OpeningSalesQty").value = RSTransDetails("Quantity").value
                RSTransDetails("OpeningSalesValue").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Valu")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Valu"))))
                RSTransDetails("Price").value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
            
            End If

            SngTemp = SngTemp + (val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice"))) * RSTransDetails("quantity").value)
         
            If Me.XPCboDiscountType.ListIndex = 1 Then
                TotalBillDiscount = IIf(XPTxtDiscountVal.Text = "", Null, (XPTxtDiscountVal.Text))
                     
            ElseIf XPCboDiscountType.ListIndex = 2 Then

                If XPTxtDiscountVal.Text <> "" Then
                    TotalBillDiscount = IIf(XPTxtDiscountVal.Text = "", Null, (XPTxtDiscountVal.Text)) * val(LblTotalAll.Caption) / 100
                             
                Else
                    TotalBillDiscount = 0
                End If
            End If

            TotalDiscountPerLine = ((RSTransDetails("SHOWprice") * RSTransDetails("SHOWQTY")) / LblTotalAll.Caption) * (TotalBillDiscount)
            RSTransDetails("TotalDiscountPerLine") = Round(TotalDiscountPerLine, 20)
                 
            RSTransDetails("ProductionDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate"))), "DD/mm/YYYY"))
            RSTransDetails("ExpiryDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate"))), "DD/mm/YYYY"))
            RSTransDetails("LotNO").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("LotNO")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("LotNO")))
        
            RSTransDetails("GranteeType").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("GranteeType")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("GranteeType")))
            RSTransDetails("GranteeStartDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("GranteeStartDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("GranteeStartDate"))), "DD/mm/YYYY"))
            RSTransDetails("GranteeEndDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("GranteeEndDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("GranteeEndDate"))), "DD/mm/YYYY"))
            RSTransDetails("RegularMaintenancedates").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("RegularMaintenancedates")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("RegularMaintenancedates")))
            
            RSTransDetails.update
            '-------------
        
        End If

    Next RowNum
 
    Cn.CommitTrans

    TransBegine = False
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
  Dim FactoryExpenses As ADODB.Recordset
    Set FactoryExpenses = New ADODB.Recordset
    StrSQL = "Select * from TblSalesGExpepenses where Transaction_ID=" & val(XPTxtBillID.Text)
    FactoryExpenses.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
   
    For RowNum = 1 To Fg_Journal.Rows - 2

        If Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("AccountName")) <> "" And val(Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("Price"))) <> 0 Then
            FactoryExpenses.AddNew
            FactoryExpenses("Transaction_ID").value = val(XPTxtBillID.Text)
            FactoryExpenses("Accountcode2").value = Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("Accountcode2"))
            FactoryExpenses("NoteSerial").value = Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("NoteSerial"))
            FactoryExpenses("Accountcode").value = Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("Accountcode"))
            FactoryExpenses("AccountName").value = Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("AccountName"))
            FactoryExpenses("value").value = val(Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("value")))
            FactoryExpenses("Price").value = val(Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("Price")))
            FactoryExpenses("des").value = Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("des"))
            FactoryExpenses.update
        End If
         
    Next RowNum
    With GRID2
    For i = 1 To .Rows - 1
    If val(.TextMatrix(i, .ColIndex("Transaction_ID"))) <> 0 Then
    If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
    Cn.Execute " Update Transactions    set TransGorupID=" & val(XPTxtBillID.Text) & "  where Transaction_ID =" & val(.TextMatrix(i, .ColIndex("Transaction_ID"))) & " "
    Else
    Cn.Execute " Update Transactions    set TransGorupID=null  where Transaction_ID =" & val(.TextMatrix(i, .ColIndex("Transaction_ID"))) & " "
    End If
    End If
    Next i
    End With
    '----------------------------------------------------------------
    '·√‰‰« ﬁ„‰« »≈÷«›… Õ—ﬂ… „‰ ‰Ê⁄ „Œ ·›…
    
    If invoiceSerach = True Then
 StrSQL = "SELECT * FROM Transactions WHERE Transaction_ID=" & val(Me.XPTxtBillID.Text) & "" ' & InvType
 Else
 StrSQL = "SELECT * FROM Transactions WHERE  Transaction_Type=40 "
 End If
 
    
    
         
    'If SystemOptions.usertype <> UserAdminAll Then
    StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"
    'StrSQL = StrSQL & "  AND   BranchId=" & Current_branch

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.Retrive val(Me.XPTxtBillID.Text)
    '----------------------------------------------------------------
SaveBillByCustomerLine
    CuurentLogdata
 createVoucher
    Select Case Me.TxtModFlg.Text
    
        Case "N"
            saveApprovalData val(Me.XPTxtBillID.Text), 170, val(TxtNoteSerial1.Text), Me.Name

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
                Msg = " Data Was Saved do you want Another Entry" & CHR(13)
                        
            End If
            
            XPBtnMove_Click (2)

            If SystemOptions.Save_options = 1 Or SystemOptions.Save_options = 2 Then
                PrintReport

                DoEvents
                DoEvents
                DoEvents
        
            ElseIf SystemOptions.Save_options = 3 Then
                PrintReport

                DoEvents
                DoEvents
                DoEvents
        
                Cmd_Click (0)
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton1, App.title) = vbYes Then
                Cmd_Click (0)
                Screen.MousePointer = vbDefault
            Else
                TxtModFlg.Text = "R"
            End If

  
 
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
                Msg = " changes Was Saved   & Chr(13)"
    
            End If

            lbl(64).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
       
            '    Me.Retrive Val(Me.XPTxtBillID.text)
            TxtModFlg.Text = "R"
    End Select
ISButton1_Click
    Screen.MousePointer = vbDefault

    'her
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
            Msg = Msg & CHR(13) & Err.description
            Msg = Msg & CHR(13) & Err.Number
            Msg = Msg & CHR(13) & Err.Source
            Msg = Msg & CHR(13) & Err.LastDllError
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Else
            Msg = "Can't Save error in Data" & CHR(13)
        End If

        Exit Sub
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)

        Msg = Msg & CHR(13) & Err.description
        Msg = Msg & CHR(13) & Err.Number
        Msg = Msg & CHR(13) & Err.Source
        Msg = Msg & CHR(13) & Err.LastDllError
    Else
        Msg = "Sorry........Error During Save " & CHR(13)

    End If

    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
        XPTxtDiscountVal.Text = ""
    Else
    
        XPTxtDiscountVal.Enabled = True
        XPTxtDiscountVal.Text = ""
    End If

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        If FG.TextMatrix(1, FG.ColIndex("Code")) <> "" Then
            NewGrid.Calculate 1, , , True
        End If
    End If

    Me.lbl(55).Visible = (Me.XPCboDiscountType.ListIndex = 2)

    'Me.lbl(21).Visible = (Me.XPCboDiscountType.ListIndex = 2)
    If XPCboDiscountType.ListIndex = 0 Then
        lbl(8).Visible = False
        XPTxtDiscountVal.Visible = False
        lbl(8).Visible = False
    Else
        lbl(8).Visible = True
        XPTxtDiscountVal.Visible = True
        lbl(8).Visible = True
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPChkPayType_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If XPChkPayType(0).value = Checked Then
                If Me.TxtModFlg.Text = "N" Then
                    XPTxtValue(0).Text = ""
                    XPTxtSerial(0).Text = ""
                End If

                If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
                    XPTxtValue(0).Enabled = True
                    '                XPTxtSerial(0).Enabled = True
                    XPTxtValue(0).locked = False
                    '                XPTxtSerial(0).Locked = False
                End If

            Else
                XPTxtValue(0).Enabled = False
                XPTxtValue(0).Text = ""
                '            XPTxtSerial(0).Enabled = False
            End If

        Case 1

            If XPChkPayType(1).value = Checked Then
                If Me.TxtModFlg.Text = "N" Then
                    XPTxtValue(1).Text = ""
                    XPTxtSerial(1).Text = ""
                    DtpDelayDate.value = Date
                    XPTxtSerial(1).Text = CStr(new_id("Notes", "NoteSerial", "", True))
                End If

                If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
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
                XPTxtValue(1).Text = ""
                Me.ChkInstall.Enabled = False
            End If

        Case 2

            If XPChkPayType(2).value = Checked And Me.TxtModFlg.Text <> "R" Then
                Me.CmdCheque.Enabled = True
            Else
                Me.CmdCheque.Enabled = False
                Me.lbl(18).Caption = 0
                Me.lbl(19).Caption = 0
                Me.FgCheques.Rows = Me.FgCheques.FixedRows
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
        XPTxtTaxValue.Text = ""
        XPTxtTaxValue.Enabled = False
        lbl(4).Enabled = False
        lbl(45).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPDtbBill_Change()

    If Trim(TxtNoteSerial1.Text) <> "" Then
        oldtxtNoteSerial1.Text = TxtNoteSerial1.Text
    End If

    TxtNoteSerial.Text = ""
    TxtNoteSerial1.Text = ""
    CurrentVoucherNo = ""
    DateChanged = True
    'updateProfit
End Sub

Private Sub XPTxtDiscountVal_Change()
    Dim Msg As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        NewGrid.Calculate 1, , , True
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub PrintReport(Optional PrinterTarget As Boolean = False, _
                        Optional reportid As Integer, _
                        Optional AdvPayment As String, _
                        Optional LblInstallCount As String, _
                        Optional LblPrecenValue As String, _
                        Optional LblFirstInstallDate As String)

    Dim ShowType As Integer
    'Dim clrep As ClsReportProp
    Dim StrPath As String
    Dim Msg As String
    Dim P_Target As PrintTarget

    On Error GoTo ErrTrap

    'If MDIFrmMain.MnuInvPrintDirect.Checked = True Then
    '    P_Target = PrinterTarget

    'End If
    Dim RowNum As Integer
    Dim PayDes As String
    PayDes = ""
    For RowNum = 1 To Fg_Journal.Rows - 1
   If val(Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("Price"))) <> 0 And Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("AccountName")) <> "" Then
   If PayDes <> "" Then
          PayDes = PayDes & CHR(13) & Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("AccountName")) & "  : " & Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("Price"))
   Else
           PayDes = Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("AccountName")) & "  : " & Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("Price"))
  End If
  If RowNum = Fg_Journal.Rows - 1 Then
  PayDes = PayDes & CHR(13)
  End If
  End If
  Next RowNum
  
    If SystemOptions.Save_options = 2 Or SystemOptions.Save_options = 3 Then
        P_Target = PrinterTarget
    Else
        P_Target = WindowTarget
    End If

    ShowType = GetSetting(StrAppRegPath, "View_Type", "SallReportType", 1)

    If reportid = 1 Then
        If XPTxtBillID.Text <> "" Then
            Set SaleReport = New ClsSaleReport
            Dim Nationality As String
            Dim GID As Double
        
            GetCustomerAllData val(DBCboClientName.BoundText), , , , , , , , , , , , , , , Nationality, , GID
            SaleReport.ShowSallingDataDetailed XPTxtBillID.Text, , , , val(lblInstComm.Caption) + val(LblTotal.Caption), TxtSearchCode.Text, TxtBillComment.Text, val(lblInstComm.Caption), 1, Nationality, GID, XPDtbBill.value, AdvPayment, LblInstallCount, LblPrecenValue, LblFirstInstallDate
 
        End If
    
        Exit Sub
    End If

    If reportid = 2 Then
        If XPTxtBillID.Text <> "" Then
            Set SaleReport = New ClsSaleReport
 
            Dim Madyna As String
            Dim hay As String
            Dim Address As String
        
            GetCustomerAllData val(DBCboClientName.BoundText), , , , , , , , , , , , , Madyna, hay, Nationality, , GID, Address
            SaleReport.ShowSallingDataDetailed XPTxtBillID.Text, , , , LblTotal, TxtSearchCode.Text, TxtBillComment.Text, val(lblInstComm.Caption), 2, Nationality, GID, XPDtbBill.value, AdvPayment, LblInstallCount, LblPrecenValue, LblFirstInstallDate, Madyna, hay, Address, val(val(lblInstComm.Caption) + val(LblTotal.Caption)) - val(AdvPayment)
 
        End If
    
        Exit Sub
    End If

    If ShowType = 1 Then
        If XPTxtBillID.Text <> "" Then
            Set SaleReport = New ClsSaleReport
            SaleReport.ShowComposedSallingDataDetailed XPTxtBillID.Text, , , , val(LblTotal.Caption) + val(TxtGVAT.Text), TxtSearchCode.Text, TxtBillComment.Text, val(lblInstComm.Caption), , , , , , , , , , , , , val(Me.dcBranch.BoundText), PayDes
        
            '  If MDIFrmMain.MnuInvPrintReceipt.Checked = True Then
            '      SaleReport.PrintInvoiceReceipt Val(XPTxtBillID.text), P_Target
            '  End If
        End If
    
    ElseIf ShowType = 40 Then

        If XPTxtBillID.Text <> "" Then
            Set SaleReport = New ClsSaleReport
            SaleReport.ShowSallingDataDetailed2 XPTxtBillID.Text, , , , LblTotal, TxtSearchCode.Text, TxtBillComment.Text, val(lblInstComm.Caption)
        
            '  If MDIFrmMain.MnuInvPrintReceipt.Checked = True Then
            '      SaleReport.PrintInvoiceReceipt Val(XPTxtBillID.text), P_Target
            '  End If
        End If
    
    ElseIf (ShowType = 2) Or (ShowType = 4) Then
        '    P_Target = IIf(MDIFrmMain.MnuInvPrintSave.Checked = True, PrintTarget.PrinterTarget, PrintTarget.WindowTarget)

        If SystemOptions.Save_options = 2 Or SystemOptions.Save_options = 3 Then
            P_Target = PrinterTarget
        Else
            P_Target = WindowTarget
        End If

        If XPTxtBillID.Text <> "" Then
            Set SaleReport = New ClsSaleReport
            SaleReport.ShowComposedSallingDataDetailed XPTxtBillID.Text, , , , LblTotal, TxtSearchCode.Text, TxtBillComment.Text, val(lblInstComm.Caption), 2
        
        End If

    ElseIf ShowType = 3 Then

        If XPTxtBillID.Text <> "" Then
            StrPath = GetSetting(StrAppRegPath, "PrintReport", "ReportPath", App.path & "\Bill_Template\SaleMain.drp")

            If StrPath = "" Then
                Msg = "⁄›Ê« : Â‰«ﬂ Œÿ√„« ›Ì „”«— «· ﬁ—Ì— "
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If

            Set crep = New ClsReportProp
            crep.OpenFile = StrPath
            crep.LoadFile StrPath, FrmPreview
            crep.InvoID = XPTxtBillID.Text
            crep.ShowReport
            FrmPreview.show vbModal
            Set crep = Nothing
        End If

    ElseIf ShowType = 5 Then

        If XPTxtBillID.Text <> "" Then
            Set SaleReport = New ClsSaleReport
            SaleReport.ShowSallingData val(XPTxtBillID.Text), 1, val(Me.DBCboClientName.BoundText)

     
        End If

    ElseIf ShowType = 6 Then

        If XPTxtBillID.Text <> "" Then
            Set SaleReport = New ClsSaleReport
            SaleReport.ShowSallingData val(XPTxtBillID.Text), 2, val(Me.DBCboClientName.BoundText)
        
            SaleReport.PrintInvoiceReceipt val(XPTxtBillID.Text), P_Target
       
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPTxtDiscountVal_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtDiscountVal.Text, 0)
End Sub

Private Sub XPTxtSum_Change()
    On Error GoTo ErrTrap

    If CboPayMentType.ListIndex = 0 Then
        XPChkPayType(0).value = Checked
        XPTxtValue(0).Text = XPTxtSum.Text
    End If

    Me.LblTotal.Caption = XPTxtSum.Text
    CalculateInvPrecent
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
               ' Unload customer_screen

            Case vbCancel
                Cancel = True
              '  Unload customer_screen
        End Select

      '  Unload customer_screen
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
    Dim Fullcode As String
 
    GetCustomersDetail val(DBCboClientName.BoundText), , Fullcode, 1
    TxtSearchCode.Text = Fullcode

    If val(DBCboClientName.BoundText) <> 0 Then
        If (DBCboClientName.BoundText = 1 Or DBCboClientName.BoundText = 2) And Me.TxtModFlg.Text <> "R" Then
            CboPayMentType.locked = True
            '        CboPaymentType.ListIndex = 0
            Me.TxtCashCustomerName.Enabled = True
            Me.CmdCash(0).Enabled = True
            Me.CmdCash(1).Enabled = True
        Else
            CboPayMentType.locked = False
            Me.TxtCashCustomerName.Enabled = False
            Me.CmdCash(0).Enabled = False
            Me.CmdCash(1).Enabled = False
        End If
    
        If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
            Dim DefaultSalesPersonId As Integer
            GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId

            If Not DefaultSalesPersonId = 0 Then

                Me.DcboEmp.BoundText = DefaultSalesPersonId
            End If

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
                        '                 mina   Me.XPCboDiscountType.ListIndex = 0
                        '                 mina   Me.XPTxtDiscountVal.text = 0
                    ElseIf RsTemp("Trans_DiscountType").value = 1 Then
                        Me.XPCboDiscountType.ListIndex = 1
                        Me.XPTxtDiscountVal.Text = IIf(IsNull(RsTemp("Trans_Discount").value), "", RsTemp("Trans_Discount").value)
                    ElseIf RsTemp("Trans_DiscountType").value = 2 Then
                        Me.XPCboDiscountType.ListIndex = 2
                        Me.XPTxtDiscountVal.Text = IIf(IsNull(RsTemp("Trans_Discount").value), "", RsTemp("Trans_Discount").value)
                    End If

                Else
                    Me.XPCboDiscountType.ListIndex = 0
                    Me.XPTxtDiscountVal.Text = 0
                End If

            Else
                Me.CboSaleType.ListIndex = -1
                Me.XPCboDiscountType.ListIndex = 0
                Me.XPTxtDiscountVal.Text = 0
            End If

            RsTemp.Close
            Set RsTemp = Nothing
        End If
    End If
'    DcCustmer.BoundText = DBCboClientName.BoundText
'    If DcCustmer.BoundText = "" Then txtCustCode = ""
    FillVoucherGrid
    'FillOrderGrid
    Exit Sub
ErrTrap:
    Msg = Err.description & CHR(13) & ""
    Msg = Msg & Err.Source & CHR(13) & ""
    Msg = Msg & Me.Name & " DBCboClientName_Change:" & CHR(13) & ""
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub DBCboClientName_Click(Area As Integer)
    DBCboClientName_Change
End Sub

Private Sub XPTxtValue_Change(Index As Integer)
    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        If XPTxtValue(1).Text <> "" Then
            If val(Me.XPTxtValue(1).Text) > 0 Then
                ChkInstall.Enabled = True
            End If

        End If
    End If

    'If XPChkPayType(1).Value = 1 Then
    '            XPTxtValue(0).text = Val(LblTotal.Caption) - Val(XPTxtValue(1).text)
    'End If
    'If XPChkPayType(0).Value = 1 Then
    '    XPTxtValue(1).text = Val(LblTotal.Caption) - Val(XPTxtValue(0).text)
    'End If
    Exit Sub
ErrTrap:
End Sub

Public Sub ReplacementData()
    Dim Msg As String
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsReplace As ADODB.Recordset

    If Me.TxtModFlg.Text <> "R" Then Exit Sub

    '«·»ÕÀ ⁄‰ ⁄„·Ì«  «·«” »œ«· «·Œ«’… »«·›« Ê—…
    If FG.TextMatrix(FG.Row, FG.ColIndex("Code")) <> "" And FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) <> "" Then
        StrSQL = "select * From ReplacedItems where ReturnID=" & XPTxtBillID.Text
        StrSQL = StrSQL + " and ItemID=" & FG.TextMatrix(FG.Row, FG.ColIndex("Code"))
        StrSQL = StrSQL + " and ItemSerial='" & FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) & "'"
        Set RsReplace = New ADODB.Recordset
        RsReplace.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsReplace.EOF Or RsReplace.BOF) Then
            Msg = "·ﬁœ  „ «” »œ«· «·ﬁÿ⁄… : " & FG.Cell(flexcpTextDisplay, FG.Row, FG.ColIndex("Name")) & CHR(13)
            Msg = Msg + "–«  «·”Ì—Ì«· : " & FG.TextMatrix(FG.Row, FG.ColIndex("Serial")) & CHR(13)
            Msg = Msg + " »«·ﬁÿ⁄… –«  «·”Ì—Ì«· : " & RsReplace("newSerial").value & CHR(13)
            Msg = Msg + "›Ì ⁄„·Ì… ’Ì«‰…"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, "ﬁÿ⁄…  „ «” »œ«·Â«"
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Function AvailableDeal() As Boolean
    'On Error GoTo ErrTrap
    Dim RowNum As Integer
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    Dim RsSalle As ADODB.Recordset
    Dim LngItemID As Long

    For RowNum = 1 To FG.Rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            StrSQL = "select * From QryDelPurchase where Transaction_Date >=" & SQLDate(XPDtbBill.value, True) & ""
            StrSQL = StrSQL + " and Item_ID=" & FG.TextMatrix(RowNum, FG.ColIndex("Code"))
            StrSQL = StrSQL + " and Transaction_Type=9"

            If FG.Cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then
                If FG.TextMatrix(RowNum, FG.ColIndex("Serial")) <> "" Then
                    StrSQL = StrSQL + " and ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                End If
            End If

            Set RsSalle = New ADODB.Recordset
            RsSalle.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsSalle.EOF Or RsSalle.BOF) Then
                If FG.Cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then

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
                    Set RsTemp = GetItemQuantityStock(LngItemID, Me.DCboStoreName.BoundText, Me.XPDtbBill.value, val(Me.XPTxtBillID.Text))

                    If Not (RsTemp.EOF Or RsTemp.BOF) Then
                        If val(RsTemp("totalqty").value) < val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))) Then

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
    On Error Resume Next
    Dim StrTemp As String
    Dim RsTemp As ADODB.Recordset

    CboPayMentType.ListIndex = 1


    
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
            TxtTransSerial.Text = StrTemp
        Else
            TxtTransSerial.Text = 1
        End If

    Else
        TxtTransSerial.Text = 1
    End If

    DCPaymentNet.BoundText = 1

End Sub

Private Sub CalculateInvPrecent()
    Dim DblInvTotal As Double
    Dim DblInvProfit As Double
    Dim DblRes As Double

    DblInvProfit = val(Me.LblInvProfit.Caption)
    DblInvTotal = val(Me.XPTxtSum.Text)

    If DblInvProfit = 0 Or DblInvTotal = 0 Then
        DblRes = 0
    Else
        DblRes = 100 * (DblInvProfit / DblInvTotal)
    End If

    Me.lblInvPrecent.Caption = "%" & CStr(Int(DblRes)) 'Format(DblRes, SystemOptions.SysDefCurrencyForamt)
End Sub

Private Sub dcCar_Change()

    GetDriverInformation (val(DCCar.BoundText))

End Sub

Private Sub dcCar_Click(Area As Integer)
    GetDriverInformation (val(DCCar.BoundText))

End Sub

Function GetDriverInformation(ID As Integer)

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        Dim sql As String
        Dim rs As New ADODB.Recordset
 
        sql = " SELECT    * "
        sql = sql & " from dbo.TblCarsData"
        sql = sql & " Where (id = " & ID & ") "

        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If rs.RecordCount > 0 Then
            DCDriver.BoundText = IIf(IsNull(rs("Emp_id").value), 0, rs("Emp_id").value)
                  
        Else
            DCDriver = 0
               
        End If

    End If

End Function

Private Sub LoadCombosData()
    Dim StrSQL As String
    Dcombos.GetPaymentType Me.DCPaymentNet
    Dcombos.GetSalesRepData Me.DcboEmp
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetCustomersSuppliers 2, Me.DcbSuppler, True
    Dcombos.GetCustomersSuppliers 2, Me.DcbFarm, True
    Dcombos.GetDocTypebyid Me.DCDocTypes, 21, val(Me.dcBranch.BoundText)
    Dcombos.GetPrefix2 Me.DCPreFix, 7, 0
    Dcombos.GetCars Me.DCCar
    Dcombos.GetEmployees Me.DCDriver, , True
    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName
    Dcombos.GetStores Me.DCboStoreName
    Dcombos.GetCustomersSuppliers 1, Me.DcCustmer, False, branch_id

    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DBCboClientName
    cSearchDcbo(0).SetBuddyText Me.TxtCusID

    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DCboStoreName
    cSearchDcbo(1).SetBuddyText Me.TxtStoreID

    Set cSearchDcbo(3) = New clsDCboSearch
    Set cSearchDcbo(3).Client = Me.DcboEmp
    cSearchDcbo(3).SetBuddyText Me.TxtEmployeeID

    StrSQL = "  select  BankID,BankName  from BanksData   "
    fill_combo Dcbanks, StrSQL
      
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    lbl(34).Caption = "Comm T."
    lbl(35).Caption = "Comm%"
    lbl(36).Caption = "Total"
    lbl(72).Caption = "Adv.P"

    lbl(38).Caption = "No Of Ins"
    lbl(40).Caption = "First Ins "
    lbl(42).Caption = "Period"
    lbl(37).Caption = "Fixed Value"
    lbl(70).Caption = "Ins Disc."
    lbl(57).Caption = "Cash.visa"
    Label3.Caption = "Branch"
    Frame1.Caption = "Info"
    lbl(56).Caption = "Order No."
    lbl(58).Caption = " Total"
    lbl(59).Caption = " Payed"
    lbl(60).Caption = " Changed"
    lbl(63).Caption = " Total Qty"
    Frame2.Caption = "Color Map"
    lbl(68).Caption = " Profit"
    lbl(69).Caption = "Comm."
    lbl(71).Caption = "Net"
    Cmd1.Caption = "Attachments"

  '  Label1.Caption = "Doc Type"
    lbl(65).Caption = "Curr"
    lbl(66).Caption = "Rec No"
    lbl(67).Caption = "Manf No"
    Label6.Caption = "Price<cost"
    Label8.Caption = "Price=cost"
    Me.XPTab301.TabCaption(3) = "Attachments"
    Me.LblShortcutKeys.Caption = "(New F12 OR Enter) ,(Edit F11),(Save F10),(Undo F9),(Delete F8),(Search F7)"
    'Command2.Caption = "Convert to I. Voucher"
    Me.Caption = "Sales Invoice"
    Ele(9).Caption = Me.Caption
    lbl(5).Caption = "Invoice ID"
    lbl(6).Caption = "Invoice Date"
    lbl(7).Caption = "Customer Name"
    lbl(24).Caption = "Store Name"
    lbl(25).Caption = "Employee"
    lbl(9).Caption = "Cash/Credit"
    lbl(10).Caption = "Discount Type"
    lbl(8).Caption = "Value"
    lbl(22).Caption = "Profit Value"
    lbl(23).Caption = "Profit Perce"

    lbl(3).Caption = "Total:"
    lbl(49).Caption = "Net "
    lbl(50).Caption = "Disc"
    lbl(1).Caption = "By:"
    lbl(2).Caption = "Rec. Count:"

    lbl(31).Caption = "Item Code"
    lbl(30).Caption = "Item Name"
    lbl(29).Caption = "Item Case"
    lbl(28).Caption = "Item Serial"
    lbl(27).Caption = "Quantity"
    lbl(26).Caption = "Price"
    lbl(32).Caption = "Sales Type"
    lbl(33).Caption = "Customer Name"
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
    XPChkPayType(1).Caption = "Due"
    XPChkPayType(0).Caption = "Check"
    lbl(13).Caption = "Value"
    lbl(15).Caption = "Value"
    lbl(16).Caption = "Value"
    lbl(12).Caption = "Serial"
    lbl(14).Caption = "Serial"
    lbl(11).Caption = "Box Name"
    lbl(21).Caption = "Due Date"
    
    lbl(18).Caption = "Check NO."
    lbl(17).Caption = "Bank Name"
    lbl(19).Caption = "Due Date"
    CmdINSTALLMENT.Caption = "INSTALLMENT"
    Me.XPTab301.TabCaption(2) = "Comment On Invoice"
    Me.Ele(15).Caption = "Write any Comments about this Invoice"
    
    lbl(44).Caption = "Comment"
    XPChkPayType(0).Caption = "Cash"
    lbl(13).Caption = "Value"
    lbl(12).Caption = "ID"
    lbl(2).Caption = "Box"
    lbl(20).Caption = "Currency"
    XPChkPayType(1).Caption = "Credit"
    lbl(15).Caption = "Value"
    lbl(14).Caption = "ID"
    'Label1.Caption = "Due Date"
    ChkInstall.Caption = "Installment"
    CmdINSTALLMENT.Caption = "Calc"
    Label2.Caption = "Bank"
    CmdCheque.Caption = "Register"

    With FgInstallments
        .TextMatrix(0, .ColIndex("QestID")) = "ID"
        .TextMatrix(0, .ColIndex("value")) = "value"
        .TextMatrix(0, .ColIndex("Due_Date")) = "Due_Date"
 
    End With

    With FG
        .TextMatrix(0, .ColIndex("order_no")) = "ORD/INV NO."
 
    End With

    With FgCheques
 
        .TextMatrix(0, .ColIndex("CheckValue")) = "Value"
        .TextMatrix(0, .ColIndex("CheckNumber")) = "Cheque Number"
        .TextMatrix(0, .ColIndex("BankName")) = "Bank Name"
        .TextMatrix(0, .ColIndex("DueDate")) = "Due Date"
        .TextMatrix(0, .ColIndex("ReleaseDate")) = "Release Date"
 
    End With

    XPChkPayType(2).Caption = "Cheques"
    '«·ÊﬁÊ› ⁄‰œ «·«Ê—«ﬁ «·„«·ÌÂ
    lbl(61).Caption = "Bill type"
    BillBasedOn(0).Caption = "Direct Sales Invoices"
    BillBasedOn(1).Caption = "Based On Issue Vouchere"
    BillBasedOn(2).Caption = "Based On Purchase Orders"

    With Me.GRID1
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("noteserial1")) = "Voucher NO"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Voucher Date"
        .TextMatrix(0, .ColIndex("NoteSerial")) = "JE Voucher NO"
    End With
    With Me.GRID2
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("order_no")) = "Order No"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Voucher Date"
        .TextMatrix(0, .ColIndex("CusName")) = "Customer Name"
    End With

    Frame3.Caption = "JE Voucher NO"
    lbl(62).Caption = "JE Voucher NO"
    Cmd(10).Caption = "Print JE"

End Sub

Private Sub XPTxtValue_KeyPress(Index As Integer, _
                                KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtValue(Index).Text, 0)
End Sub

Private Function CheckCashCustomer() As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    If Trim$(Me.TxtCashCustomerName.Text) = "" Then
        CheckCashCustomer = True
    Else
        StrSQL = "Select * From Transactions Where CashCustomerName='" & Trim$(Me.TxtCashCustomerName.Text) & "'"
    
    End If

End Function

Private Sub XPTxtValue_MouseMove(Index As Integer, _
                                 Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)

    If val(Me.XPTxtValue(Index).Text) <> 0 Then
        Me.XPTxtValue(Index).ToolTipText = WriteNo(Me.XPTxtValue(Index).Text, 1, True)
    Else
        Me.XPTxtValue(Index).ToolTipText = ""
    End If

End Sub

Private Sub SumChecks()

    With Me.FgCheques

        If .Rows > 1 Then
            Me.lbl(19).Caption = .Aggregate(flexSTCount, .FixedRows, .ColIndex("CheckNumber"), .Rows - 1, .ColIndex("CheckNumber"))
            Me.lbl(18).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CheckValue"), .Rows - 1, .ColIndex("CheckValue"))
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
    Me.LblDiscount.Caption = 0
    Me.LblAdvPayment.Caption = 0
    lbl(19).Caption = ""
    lbl(18).Caption = ""
End Sub



 Sub GetInformationCustomer(Optional Cus_ID As Double, Optional TelNo As String = "")
Dim Rs6 As ADODB.Recordset
Set Rs6 = New ADODB.Recordset
Dim sql As String
If Cus_ID <> 0 Or TelNo <> "" Then
    If Cus_ID <> 0 Then
        sql = "select CusID,CusName ,CusNamee,Cus_mobile,Address from TblCustemers where CusID =" & Cus_ID & " "
        
    End If
Rs6.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
'TxtAddress.Text = IIf(IsNull(Rs6("Address").value), "", Rs6("Address").value)

DcCustmer.BoundText = IIf(IsNull(Rs6("CusID").value), "", Rs6("CusID").value)
'TxtCusID.Text = IIf(IsNull(Rs6("CusID").Value), "", Rs6("CusID").Value)
Else
'TxtCusID = ""
'TxtRecordNo = ""
'DcCustomerType.BoundText = ""
End If
End If
End Sub

Public Function generalSearch(StrSQL As String)
rs.Close
Set rs = Nothing


    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
          If Not (rs.BOF Or rs.EOF) Then
                rs.MoveLast
            End If
   Me.TxtModFlg.Text = "R"
            Retrive
          
            Me.TxtModFlg.Text = "R"
End Function
