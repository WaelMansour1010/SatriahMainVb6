VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmManAddNew1 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«· ÕÊÌ· „‰ Ê—‘… «·Ï Ê—‘…"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11700
   Icon            =   "FrmManAddNew1.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7395
   ScaleWidth      =   11700
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   9330
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   6690
      Visible         =   0   'False
      Width           =   645
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   3030
      Index           =   0
      Left            =   30
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   0
      Width           =   11640
      _cx             =   20532
      _cy             =   5345
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
      Appearance      =   5
      MousePointer    =   0
      Version         =   801
      BackColor       =   14871017
      ForeColor       =   192
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "»Ì«‰«  ≈Ì’«· «·’Ì«‰…"
      Align           =   0
      AutoSizeChildren=   0
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
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   765
         Left            =   0
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   62
         Top             =   2160
         Width           =   4185
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   8460
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   2520
         Width           =   1755
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   6180
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   8460
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   2160
         Width           =   1755
      End
      Begin VB.TextBox TxtNoteSerial1 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   6240
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TxtCashCustomerMobile 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   6180
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   1770
         Width           =   1695
      End
      Begin VB.TextBox TxtCashCustomerPhone 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   8460
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   1740
         Width           =   1755
      End
      Begin VB.CheckBox ChkInv 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   255
         Left            =   11250
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   690
         Width           =   255
      End
      Begin VB.TextBox TxtMaintanenceID 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   9090
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   270
         Width           =   1185
      End
      Begin VB.TextBox TxtTransSerial 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   5490
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   30
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox CboMaintenanceType 
         Height          =   315
         Left            =   1890
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   330
         Width           =   2295
      End
      Begin VB.TextBox TxtTransID 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   4650
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.TextBox TxtCashCustomerName 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   6180
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1380
         Width           =   4035
      End
      Begin VB.TextBox TxtReciptNumber 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   6240
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   270
         Width           =   1425
      End
      Begin ImpulseButton.ISButton XPBtnNewClients 
         Height          =   315
         Left            =   5940
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1020
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
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
         ButtonImage     =   "FrmManAddNew1.frx":038A
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
      End
      Begin MSDataListLib.DataCombo DBCboClientName 
         Height          =   315
         Left            =   6180
         TabIndex        =   2
         Top             =   1050
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker XPDtbGoInDtae 
         Height          =   345
         Left            =   2490
         TabIndex        =   5
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         _Version        =   393216
         Format          =   96337920
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker XPDtbGoOutDtae 
         Height          =   345
         Left            =   2490
         TabIndex        =   6
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         _Version        =   393216
         Format          =   96337921
         CurrentDate     =   38784
      End
      Begin MSDataListLib.DataCombo DcboEmp 
         Height          =   315
         Left            =   90
         TabIndex        =   8
         Top             =   660
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "7"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCboStoreName 
         Height          =   315
         Left            =   2610
         TabIndex        =   9
         Top             =   990
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "7"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton CmdSearchTrans 
         Height          =   345
         Left            =   5850
         TabIndex        =   12
         Top             =   630
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
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
         ButtonImage     =   "FrmManAddNew1.frx":0724
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton CmdOpenTrans 
         Height          =   345
         Left            =   11430
         TabIndex        =   13
         Top             =   420
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   609
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
         ButtonImage     =   "FrmManAddNew1.frx":0ABE
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin MSComCtl2.DTPicker ShfitFrom 
         Height          =   345
         Left            =   120
         TabIndex        =   64
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "'Time: 'hh:mm tt"
         Format          =   96337923
         UpDown          =   -1  'True
         CurrentDate     =   40909
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   120
         TabIndex        =   68
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "'Time: 'hh:mm tt"
         Format          =   96337923
         UpDown          =   -1  'True
         CurrentDate     =   40909
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Height          =   315
         Left            =   120
         TabIndex        =   69
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "7"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Ï Ê—‘…"
         Height          =   315
         Index           =   21
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Êﬁ  «·Œ—ÊÃ"
         Height          =   315
         Index           =   20
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   1680
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Êﬁ  «· ÕÊÌ·"
         Height          =   315
         Index           =   19
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„·«ÕŸ«  ⁄·Ì «·„⁄œÂ/«·”Ì«—…"
         Height          =   315
         Index           =   17
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Tag             =   "„·«ÕŸ«  ⁄·Ì «·„⁄œÂ/«·”Ì«—…/"
         Top             =   2160
         Width           =   1635
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «··ÊÕ…"
         Height          =   315
         Index           =   16
         Left            =   10560
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   2640
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„ÊœÌ·"
         Height          =   315
         Index           =   15
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   2160
         Width           =   555
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·„⁄œÂ/«·”Ì«—…"
         Height          =   315
         Index           =   14
         Left            =   10560
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   2160
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·›« Ê—…"
         Height          =   315
         Index           =   13
         Left            =   7680
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   720
         Width           =   945
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÃÊ«·"
         Height          =   315
         Index           =   2
         Left            =   7860
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   1800
         Width           =   435
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·Â« ›"
         Height          =   315
         Index           =   1
         Left            =   10620
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   1800
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «· ÕÊÌ·"
         Height          =   315
         Index           =   3
         Left            =   4590
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1350
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄„Ì·"
         Height          =   315
         Index           =   6
         Left            =   10350
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   1050
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·Œ—ÊÃ «·„ Êﬁ⁄"
         Height          =   435
         Index           =   7
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   1680
         Width           =   1665
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·⁄„·Ì…"
         Height          =   315
         Index           =   8
         Left            =   10530
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   300
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·’Ì«‰…"
         Height          =   315
         Index           =   10
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   330
         Width           =   1200
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·›‰Ì"
         Height          =   315
         Index           =   25
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   660
         Width           =   1080
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰ Ê—‘…"
         Height          =   315
         Index           =   24
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   990
         Width           =   960
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "’Ì«‰… ·›« Ê—…  „»Ì⁄« "
         Height          =   255
         Index           =   9
         Left            =   8520
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   690
         Width           =   2610
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·⁄„Ì· «·‰ﬁœÌ"
         Height          =   315
         Index           =   12
         Left            =   10290
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1410
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·≈Ì’«·"
         Height          =   315
         Index           =   18
         Left            =   7740
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   300
         Width           =   945
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   2
      Left            =   5070
      TabIndex        =   34
      Top             =   6660
      Width           =   885
      _ExtentX        =   1561
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
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   3
      Left            =   4170
      TabIndex        =   35
      Top             =   6660
      Width           =   840
      _ExtentX        =   1482
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
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   630
      TabIndex        =   36
      Top             =   6660
      Width           =   840
      _ExtentX        =   1482
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
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   3330
      TabIndex        =   37
      Top             =   6660
      Width           =   765
      _ExtentX        =   1349
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
      ColorTextShadow =   -2147483637
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   3255
      Index           =   5
      Left            =   30
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   3090
      Width           =   11640
      _cx             =   20532
      _cy             =   5741
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
      Appearance      =   5
      MousePointer    =   0
      Version         =   801
      BackColor       =   14871017
      ForeColor       =   192
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "»Ì«‰«  «·’‰› Ê«·ﬁÿ⁄… «·œ«Œ·… ··’Ì«‰…"
      Align           =   0
      AutoSizeChildren=   0
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
      Begin VB.TextBox TxtTicketNO 
         Height          =   360
         Left            =   8760
         TabIndex        =   14
         Top             =   3405
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.TextBox TxtCustomerNotes 
         Alignment       =   1  'Right Justify
         Height          =   1485
         Left            =   5850
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   1320
         Width           =   5625
      End
      Begin VB.TextBox TxtSerial 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   360
         Left            =   600
         Locked          =   -1  'True
         MaxLength       =   20
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   405
         Width           =   1665
      End
      Begin VB.TextBox TxtEmpNotes 
         Alignment       =   1  'Right Justify
         Height          =   1485
         Left            =   210
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   1335
         Width           =   5505
      End
      Begin VB.TextBox TxtQuantity 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   30
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   405
         Width           =   555
      End
      Begin MSDataListLib.DataCombo DCboItemsName 
         Height          =   315
         Left            =   2760
         TabIndex        =   17
         Top             =   405
         Width           =   4260
         _ExtentX        =   7514
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCboItemsCode 
         Height          =   315
         Left            =   7170
         TabIndex        =   16
         Top             =   405
         Width           =   2850
         _ExtentX        =   5027
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton CmdShowTransItems 
         Height          =   360
         Left            =   9930
         TabIndex        =   15
         Top             =   405
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ".... ÕœÌœ «’‰«›"
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
         ButtonImage     =   "FrmManAddNew1.frx":0E58
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton CmdSearch 
         Height          =   285
         Index           =   2
         Left            =   2340
         TabIndex        =   53
         Top             =   390
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
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
         ButtonImage     =   "FrmManAddNew1.frx":11F2
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„«  „ "
         Height          =   210
         Index           =   5
         Left            =   7620
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   1110
         Width           =   3855
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «· ﬂ "
         Height          =   240
         Index           =   11
         Left            =   7710
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   780
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·’‰›"
         Height          =   225
         Index           =   31
         Left            =   7290
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   180
         Width           =   2370
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈”„ «·’‰›"
         Height          =   210
         Index           =   30
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   180
         Width           =   2940
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·”Ì—Ì«·"
         Height          =   225
         Index           =   28
         Left            =   660
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   180
         Width           =   1575
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„·«ÕŸ… «·„»œ∆Ì… ··›‰Ì"
         Height          =   210
         Index           =   23
         Left            =   1650
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1140
         Width           =   3825
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬂ„Ì…"
         Height          =   210
         Index           =   0
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   180
         Width           =   495
      End
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   8280
      TabIndex        =   46
      Top             =   6690
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   4
      Left            =   1680
      TabIndex        =   66
      Top             =   6660
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ﬂ—ÊﬂÌ Ê’› «·„⁄œÂ/«·”Ì«—…"
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
      ColorTextShadow =   -2147483637
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   300
      Index           =   4
      Left            =   10440
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   6690
      Width           =   945
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   11540
      X2              =   240
      Y1              =   6540
      Y2              =   6555
   End
End
Attribute VB_Name = "FrmManAddNew1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cSearchDcbo(9) As clsDCboSearch
Dim TTP As clstooltip
Dim TTD As clstooltipdemand

Private Sub ChangeLang()
    CmdShowTransItems.Caption = "Select Items"
    Me.Caption = "Receive Maintenance"
    Ele(0).Caption = "Maintenance Data"
    lbl(13).Caption = "Bill#"
    Cmd(2).Caption = "save"
    Cmd(3).Caption = "Undo"
    Cmd(0).Caption = "Print"
    Cmd(1).Caption = "Exit"
    lbl(8).Caption = "ID"
    lbl(18).Caption = "Ticket NO."
    lbl(9).Caption = "For Bill NO"
    lbl(6).Caption = "Cust. Name"
    lbl(12).Caption = "Cash Cust."
    lbl(1).Caption = "Tel ."
    lbl(2).Caption = "Mob ."
    lbl(10).Caption = "Mant. Type"
    lbl(24).Caption = "Departement"
    lbl(25).Caption = "Technical"
    lbl(3).Caption = "Recived Date"
    lbl(7).Caption = "Expected Out Date"
    Ele(5).Caption = "Item Data"
    lbl(11).Caption = "Recipt NO."
    lbl(31).Caption = "Item Code"
    lbl(30).Caption = "Item Name"
    lbl(28).Caption = "Serial"
    lbl(0).Caption = "Qty"
    lbl(5).Caption = "Customer Complaint"
    lbl(23).Caption = " Technical  Notes"
    lbl(4).Caption = "BY"
End Sub

Private Sub ChkInv_Click()
    On Error GoTo ErrTrap

    If ChkInv.value = vbUnchecked Then
        TxtTransSerial.Enabled = False
        lbl(9).Enabled = False
        'CmdSearch.Enabled = False
        CmdSearchTrans.Enabled = False
        CmdOpenTrans.Enabled = False
    ElseIf ChkInv.value = vbChecked Then
        TxtTransSerial.Enabled = True
        lbl(9).Enabled = True
        'CmdSearch.Enabled = True
        CmdSearchTrans.Enabled = True
        CmdOpenTrans.Enabled = True
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 1
            Unload Me

        Case 2
            SaveData
    End Select

End Sub

Private Sub CmdSearch_Click(Index As Integer)
    Dim StrMSG As String
    Dim LngItemID As Long, LngStoreID As Long

    Select Case Index

        Case 0

            '        Load FrmItemSearch
            '        FrmItemSearch.RetrunType = 1
            '        Set FrmItemSearch.DcboItems = Me.DcboReItemName
            '        FrmItemSearch.Show vbModal
        Case 2
            Load FrmItemSearch
            FrmItemSearch.RetrunType = 1
            Set FrmItemSearch.DcboItems = Me.DCboItemsName
            FrmItemSearch.show vbModal

        Case 1
        
    End Select

End Sub

Private Sub CmdSearchTrans_Click()
    ' ›« Ê—… „»Ì⁄« 
    Load FrmBuySearch
    FrmBuySearch.DealingForm = InvoiceTransaction
    Set FrmBuySearch.ExtraRetrunObject = Me.TxtTransID
    Set FrmBuySearch.ExtraRetrunObject1 = Me.TxtNoteSerial1

    FrmBuySearch.CboPayMentType.Enabled = False

    If SystemOptions.UserInterface = ArabicInterface Then
        FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ⁄„·Ì… »Ì⁄"
    Else
        FrmBuySearch.Caption = "Search About Sales Invoices"
    End If

    FrmBuySearch.DCboClientsName.BoundText = Me.DBCboClientName.BoundText
    FrmBuySearch.show
        
End Sub

Private Sub CmdShowTransItems_Click()
    Dim Msg As String

    If val(Me.TxtTransID.text) = 0 Then
        Msg = "ÌÃ» «‰  ﬁÊ„ » ÕœÌœ ›« Ê—… »Ì⁄ .. Õ Ï ÌﬁÊ„ «·»—‰«„Ã »⁄—÷ √’‰«› Â–Â «·›« Ê—… ."
        Msg = Msg & Chr(13) & "· Œ «— „‰Â« «·√’‰«› «· Ï ”Ê› Ì „ ⁄„· ·Â« «·’Ì«‰…."
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Load FrmManChooseItems
    Set FrmManChooseItems.MyForm = Me
    FrmManChooseItems.LoadTrans val(Me.TxtTransID.text), InvoiceTransaction    '
    FrmManChooseItems.show vbModal
End Sub

Private Sub DBCboClientName_Change()
    Dim Msg As String

    On Error GoTo ErrTrap

    If val(Me.DBCboClientName.BoundText) = 2 Then
        Me.TxtCashCustomerName.Enabled = True
        Me.TxtCashCustomerPhone.Enabled = True
        Me.TxtCashCustomerMobile.Enabled = True
        Me.lbl(1).Enabled = True
        Me.lbl(2).Enabled = True
        Me.lbl(12).Enabled = True
    Else
        Me.TxtCashCustomerName.Enabled = False
        Me.TxtCashCustomerPhone.Enabled = False
        Me.TxtCashCustomerMobile.Enabled = False
        Me.lbl(1).Enabled = False
        Me.lbl(2).Enabled = False
        Me.lbl(12).Enabled = False
    
    End If

    Exit Sub
ErrTrap:

    If Err.Number = 7 Then
        Msg = "Ì⁄«‰Ï «·»—‰«„Ã „‰ ‰ﬁ’ ›Ï –«ﬂ—… «·ÃÂ«“"
        Msg = Msg & Chr(13) & "ÌÃ» €·ﬁ «·»—‰«„Ã Ê≈⁄«œ…  ‘€Ì· «·ÃÂ«“"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If

End Sub

Private Sub DBCboClientName_Click(Area As Integer)
    DBCboClientName_Change
End Sub

Private Sub DCboItemsCode_Change()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset

    If DCboItemsCode.BoundText = "" Then
        Exit Sub
    End If

    StrSQL = "select * From TblItems where ItemID=" & DCboItemsCode.BoundText
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        If RsTemp("ItemID").value <> Me.DCboItemsName.BoundText Then
            Me.DCboItemsName.BoundText = RsTemp("ItemID").value
        End If

        If RsTemp("HaveSerial").value = True Then
            TxtSerial.Enabled = True
            TxtQuantity.Enabled = False
            TxtQuantity.text = "1"
        ElseIf RsTemp("HaveSerial").value = False Then
            TxtSerial.Enabled = False
            TxtQuantity.Enabled = True
            TxtQuantity.text = ""
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub DCboItemsCode_Click(Area As Integer)
    DCboItemsCode_Change
End Sub

Private Sub DCboItemsName_Change()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset

    If DCboItemsName.BoundText = "" Then
        Exit Sub
    End If

    StrSQL = "select * From TblItems where ItemID=" & DCboItemsName.BoundText
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        If RsTemp("ItemID").value <> Me.DCboItemsCode.BoundText Then
            Me.DCboItemsCode.BoundText = RsTemp("ItemID").value
        End If

        If RsTemp("HaveSerial").value = True Then
            TxtSerial.Enabled = True
            TxtQuantity.Enabled = False
            TxtQuantity.text = "1"
        ElseIf RsTemp("HaveSerial").value = False Then
            TxtSerial.Enabled = False
            TxtQuantity.Enabled = True
            TxtQuantity.text = ""
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub DCboItemsName_Click(Area As Integer)
    DCboItemsName_Change
End Sub

Private Sub Form_Load()

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    Dim StrSQL As String
    Dim BGround As ClsBackGroundPic
    Dim RsItems As New ADODB.Recordset
    Dim StrList As String
    Dim Dcombos As ClsDataCombos
    On Error GoTo ErrTrap
    '--------------------------
 
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    CenterForm Me

    FormPostion Me, GetPostion
    '-------------------------
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    '-------------------------
    'AddTip
    SetDtpickerDate Me.XPDtbGoInDtae
    SetDtpickerDate XPDtbGoOutDtae
    Set Dcombos = New ClsDataCombos

    Dcombos.GetEmployees Me.DcboEmp
    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DcboEmp

    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DBCboClientName

    Dcombos.GetStores Me.DCboStoreName
    Set cSearchDcbo(2) = New clsDCboSearch
    Set cSearchDcbo(2).Client = Me.DCboStoreName

    Dcombos.GetItemsCodes Me.DCboItemsCode
    Set cSearchDcbo(3) = New clsDCboSearch
    Set cSearchDcbo(3).Client = Me.DCboItemsCode

    Dcombos.GetItemsNames Me.DCboItemsName
    Set cSearchDcbo(4) = New clsDCboSearch
    Set cSearchDcbo(4).Client = Me.DCboItemsName

    Dcombos.GetUsers Me.DCboUserName
    Me.DCboUserName.BoundText = user_id
    Set Dcombos = Nothing

    If SystemOptions.UserInterface = ArabicInterface Then

        With CboMaintenanceType
            .Clear
            .AddItem "œ«Œ· «·÷„«‰"
            .AddItem "Œ«—Ã «·÷„«‰"
            .ListIndex = 0
        End With

    Else

        With CboMaintenanceType
            .Clear
            .AddItem "Within warranty"
            .AddItem "WithOut warranty"
            .ListIndex = 0
        End With

    End If

    Me.ChkInv.value = vbUnchecked
    ChkInv_Click

    Me.TxtMaintanenceID.text = new_id("TblMainteneceNew", "MaintananceID", "")
    Me.TxtReciptNumber.text = Me.TxtMaintanenceID.text
    'DBCboClientName.BoundText = 2
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer

    FormPostion Me, SavePostion
    For i = LBound(cSearchDcbo) To UBound(cSearchDcbo)
        Set cSearchDcbo(i) = Nothing
    Next i

End Sub

Public Sub RetriveOrder(Optional Transaction_ID As Integer)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    On Error GoTo ErrTrap

    StrSQL = "Select * from transactions where Transaction_ID=" & Transaction_ID
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount < 1 Then
 
        Exit Sub
    Else
        DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
        Me.DCboStoreName.BoundText = IIf(IsNull(rs("storeid").value), "", rs("storeid").value)
        'Me.Dcbranch.BoundText = IIf(IsNull(rs("Branchid").value), "", rs("Branchid").value)

        'txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
    End If
 
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Private Sub TxtNoteSerial1_Change()

    If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
        RetriveOrder val(Me.TxtTransID)
        DCboItemsCode.text = ""
        DCboItemsName.text = ""
        TxtSerial.text = ""
        TxtQuantity.text = ""

    End If

End Sub

Private Sub TxtTransID_Change()
    Exit Sub
    Dim StrTemp As String
    Dim LngCuID As Long

    If Trim(Me.TxtTransID.text) = "" Then
        Me.TxtTransSerial.text = ""
    Else
        Dim RsOpt As ADODB.Recordset
        Set RsOpt = New ADODB.Recordset
        RsOpt.Open "select CheckSal from TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdText
        Dim myvar As String

        If RsOpt("CheckSal") Then myvar = 2 Else myvar = 21

        StrTemp = GetTransIDSerial(1, val(Me.TxtTransID.text), , myvar, LngCuID)
    
        If Trim$(Me.TxtTransSerial.text) <> StrTemp Then
            Me.TxtTransSerial.text = StrTemp
        End If

        If val(Me.DBCboClientName.BoundText) <> LngCuID Then
            Me.DBCboClientName.BoundText = LngCuID
        End If
    End If

End Sub

Private Sub TxtTransSerial_Change()
    Exit Sub

    If Trim$(Me.TxtTransSerial.text) = "" Then
        Me.TxtTransID.text = ""
    Else
        Dim RsOpt As ADODB.Recordset
        Set RsOpt = New ADODB.Recordset
        RsOpt.Open "select CheckSal from TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdText
        Dim myvar As String

        If RsOpt("CheckSal") Then myvar = 2 Else myvar = 21

        Me.TxtTransID.text = GetTransIDSerial(0, , Trim$(Me.TxtTransSerial.text), myvar)
    End If

End Sub

Private Sub XPBtnNewClients_Click()
    On Error GoTo ErrTrap
    'Load FrmAddNewCustemer

    'With FrmAddNewCustemer
    '    .DealingForm = Maintenance
    '    .Caption = "≈÷«›… ⁄„Ì· ÃœÌœ"
    '    .lbl(1).Caption = "ﬂÊœ «·⁄„Ì·"
    '    .lbl(0).Caption = "«”„ «·⁄„Ì·"
    '    .show vbModal
    '    cSearchDcbo(1).Refresh
    'End With

    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim rs As ADODB.Recordset
    Dim RsNotes As ADODB.Recordset
    Dim BolBegine As Boolean

    On Error GoTo ErrTrap

    If val(Me.TxtReciptNumber.text) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ≈Ì’«· «·œŒÊ· ··’Ì«‰…...!!"
        Else
            Msg = "Please Enter Recipt Numbe ...!!"
        End If

        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtReciptNumber.SetFocus
        Exit Sub
    End If

    If CboMaintenanceType.ListIndex = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄ «·’Ì«‰…"
        Else
            Msg = "Please Enter Maintenance type"
        End If
    
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CboMaintenanceType.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If DBCboClientName.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        
            Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·⁄„Ì·"
        Else
            Msg = "Please Select Customer Name"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DBCboClientName.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If Me.DcboEmp.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·›‰Ì...!!!"
        Else
            Msg = "Please  Enter Technical Name"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DcboEmp.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If ChkInv.value = vbChecked Then
        If TxtNoteSerial1.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  ÕœÌœ —ﬁ„ ›« Ê—… «·»Ì⁄ "
            Else
                Msg = "Please Select  Sales Inv."
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtNoteSerial1.SetFocus
            Exit Sub
        ElseIf PutTrans = False Then
            '        Exit Sub
        End If
    End If

    If Me.DCboStoreName.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then

            Msg = "ÌÃ» ≈Œ Ì«— «·„Œ“‰....!!! " & Chr(13)
        Else
            Msg = "Please Select Store....!!! " & Chr(13)
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCboStoreName.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    '-----------------------------------------------------------------------------
    If DCboItemsCode.text = "" Then
        Msg = "ÌÃ»  ÕœÌœ ﬂÊœ «·’‰›"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCboItemsCode.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If DCboItemsName.text = "" Then
        Msg = "ÌÃ»  ÕœÌœ «”„ «·’‰›"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCboItemsName.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If val(TxtQuantity.text) = 0 Then
        Msg = "ÌÃ»  ÕœÌœ «·ﬂ„Ì…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtQuantity.SetFocus
        Exit Sub
    End If

    LngQty = val(Me.TxtQuantity.text)

    If Me.TxtSerial.Enabled = True And Trim(Me.TxtSerial.text) = "" Then
        Msg = "»—Ã«¡ ≈œŒ«· «·”Ì—»«· «·Œ«’ »«·’‰›...!!"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtSerial.SetFocus
        Exit Sub
    End If

    If Me.ChkInv.value = vbChecked Then
        If Trim(Me.TxtNoteSerial1.text) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "»—Ã«¡ ﬂ «»… —ﬁ„ ›« Ê—… «·»Ì⁄...!!"
            Else
                Msg = "please Enter Sales Invoice No...!!"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtTransSerial.SetFocus
            Exit Sub
        End If

        If val(Me.TxtTransID.text) = 0 Then
            If PutTrans = False Then
                '   Exit Sub
            End If
        End If

        If CheckItemInv(Me.DCboItemsCode.BoundText, Trim(Me.TxtSerial.text), val(Me.TxtTransID.text)) = False Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "«·ﬁÿ⁄… „‰ «·’‰› : " & Me.DCboItemsName.text
                Msg = Msg & Chr(13) & "—ﬁ„ : " & Me.TxtSerial.text
                Msg = Msg & Chr(13) & "€Ì— „”Ã·… ›Ï «·›« Ê—… —ﬁ„ : " & Me.TxtNoteSerial1.text
            Else
                Msg = "Item : " & Me.DCboItemsName.text
                Msg = Msg & Chr(13) & "With Serial  : " & Me.TxtSerial.text
                Msg = Msg & Chr(13) & "Not Included In Invoice NO:" & Me.TxtNoteSerial1.text
            End If
        
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
    End If

    If Trim$(Me.TxtSerial.text) <> "" Then
        If Me.TxtModFlg.text = "N" Then
            LngTransID = 0
        ElseIf Me.TxtModFlg.text = "E" Then
            LngTransID = val(Me.TxtMaintanenceID.text)
        End If

        StrSQL = "SELECT QryManStockComplete.* "
        StrSQL = StrSQL + " FROM dbo.QryManStockComplete(" & LngTransID & ") QryManStockComplete"
        StrSQL = StrSQL + " Where ItemID=" & Me.DCboItemsCode.BoundText

        If Trim$(Me.TxtSerial.text) <> "" Then
            StrSQL = StrSQL + " AND ItemSerial='" & Trim$(Me.TxtSerial.text) & "'"
        End If

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Â–Â «·ﬁÿ⁄… „ÊÃÊœ… ›Ï «·„Œ“‰ ›⁄·«,,,"
            Else
                Msg = "This Item Already Exist In Store"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    End If

    Cn.BeginTrans
    BolBegine = True

    Set rs = New ADODB.Recordset
    rs.Open "TblMainteneceNew", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    rs.AddNew
    Me.TxtMaintanenceID.text = new_id("TblMainteneceNew", "MaintananceID", "")
    Me.TxtTicketNO.text = new_id("TblMainteneceNew", "TicketNO", "")
    rs("MaintananceID").value = val(Me.TxtMaintanenceID.text)
    rs("ReciptNumber").value = Trim$(Me.TxtReciptNumber.text)

    If Me.ChkInv.value = vbChecked Then
        rs("Transaction_ID").value = val(Me.TxtTransID.text)
    End If

    rs("CusID").value = val(Me.DBCboClientName.BoundText)

    If val(Me.DBCboClientName.BoundText) = 2 Then
        rs("CashCustomerName").value = Trim$(Me.TxtCashCustomerName.text)
        rs("CashCustomerPhone").value = Trim$(Me.TxtCashCustomerPhone.text)
        rs("CashCustomerMobile").value = Trim$(Me.TxtCashCustomerMobile.text)
        rs("CashCustomerEmail").value = Null
        rs("CashCustomerAddress").value = Null
    Else
        rs("CashCustomerName").value = Null
        rs("CashCustomerPhone").value = Null
        rs("CashCustomerMobile").value = Null
        rs("CashCustomerEmail").value = Null
        rs("CashCustomerAddress").value = Null
    End If

    If Me.CboMaintenanceType.ListIndex = 0 Then
        rs("MType").value = 0
    Else
        rs("MType").value = 1
    End If
    
    rs("EmpID").value = val(Me.DcboEmp.BoundText)
    rs("StoreID").value = val(Me.DCboStoreName.BoundText)
    rs("DateGoIN").value = Me.XPDtbGoInDtae.value
    rs("DateGoOUT").value = Me.XPDtbGoOutDtae.value
    
    rs("Remarks").value = Null
    rs("UserID").value = Me.DCboUserName.BoundText
    rs("PaymentType").value = 0
    rs("ManOperationTypeID").value = 1 'œŒÊ· ··’Ì«‰…
    
    rs("TicketNO").value = val(Me.TxtTicketNO.text)
    rs("ItemID").value = Me.DCboItemsName.BoundText

    If TxtSerial.Enabled = True Then
        rs("ItemSerial").value = Trim$(Me.TxtSerial.text)
    Else
        rs("ItemSerial").value = Null
    End If

    rs("Quantity").value = val(Me.TxtQuantity.text)
    rs("CustomerNotes").value = Trim$(Me.TxtCustomerNotes.text)
    rs("EmpNotes").value = Trim$(Me.TxtEmpNotes.text)
    rs("Cost").value = Null
    rs("SupDeci").value = Null
    rs("RetrunOrgID").value = Null
    rs.update

    Cn.CommitTrans
    BolBegine = False

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = " „  ⁄„·Ì… «·Õ›Ÿ...!!!"

    Else
        Msg = "Saved Successfully...!!!"
    End If

    MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    FrmManStore.LoadManStore
    Exit Sub
ErrTrap:

    If BolBegine = True Then
        Cn.RollbackTrans
        BolBegine = False
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ÕœÀ Œÿ« «À‰«¡ Õ›Ÿ «·»Ì«‰« ...!!!"
        Msg = Msg & Chr(13) & Err.description
        Msg = Msg & Chr(13) & Err.Number
        Msg = Msg & Chr(13) & Err.Source
    Else
        Msg = "An Error During Saving...!!!"
        Msg = Msg & Chr(13) & Err.description
        Msg = Msg & Chr(13) & Err.Number
        Msg = Msg & Chr(13) & Err.Source

    End If

    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Function PutTrans() As Boolean
    Dim StrTemp As String
    Dim Msg As String

    If Trim(Me.TxtTransSerial.text) = "" Then
        '    Me.TxtTransID.text = ""
    Else
        Dim RsOpt As ADODB.Recordset
        Set RsOpt = New ADODB.Recordset
        RsOpt.Open "select CheckSal from TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdText
        Dim myvar As String

        If RsOpt("CheckSal") Then myvar = 2 Else myvar = 21

        StrTemp = GetTransIDSerial(0, , Trim(Me.TxtTransSerial.text), myvar)

        If StrTemp = "" Then
            Msg = "·« ÊÃœ ›« Ê—… »Â–« «·—ﬁ„ ... ≈” Œœ„ ‘«‘… «·»ÕÀ.!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            PutTrans = False
        Else

            If Trim$(Me.TxtTransID.text) <> StrTemp Then
                Me.TxtTransID.text = StrTemp
            End If

            PutTrans = True
        End If
    End If

End Function
