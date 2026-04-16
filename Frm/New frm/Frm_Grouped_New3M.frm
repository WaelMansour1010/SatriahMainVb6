VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmPay_Garanty_Shipment3M 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8175
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   8325
   Icon            =   "Frm_Grouped_New3M.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   8325
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   8250
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8145
      _cx             =   14367
      _cy             =   14552
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
      Caption         =   "0|1|2|3|4|5|New Tab"
      Align           =   0
      CurrTab         =   2
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
      Flags(6)        =   2
      Begin VB.Frame frame6 
         Caption         =   "«·⁄ﬁÊœ"
         Height          =   7830
         Index           =   2
         Left            =   9690
         RightToLeft     =   -1  'True
         TabIndex        =   185
         Top             =   45
         Width           =   8055
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7515
            RightToLeft     =   -1  'True
            TabIndex        =   238
            Top             =   525
            Width           =   1605
         End
         Begin VB.CommandButton Command10 
            Caption         =   "ÿ»«⁄…"
            Height          =   525
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   237
            Top             =   0
            Width           =   3660
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3480
            RightToLeft     =   -1  'True
            TabIndex        =   219
            Top             =   330
            Width           =   1515
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H0080FFFF&
            Caption         =   "»‰“Ì‰ 80"
            Enabled         =   0   'False
            Height          =   315
            Left            =   9930
            MaskColor       =   &H0080FFFF&
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   218
            Top             =   1200
            Width           =   1545
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8250
            RightToLeft     =   -1  'True
            TabIndex        =   217
            Top             =   1200
            Width           =   1515
         End
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6300
            RightToLeft     =   -1  'True
            TabIndex        =   216
            Top             =   1200
            Width           =   1515
         End
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            Height          =   315
            Left            =   4470
            RightToLeft     =   -1  'True
            TabIndex        =   215
            Top             =   1200
            Width           =   1515
         End
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            Height          =   315
            Left            =   2700
            RightToLeft     =   -1  'True
            TabIndex        =   214
            Top             =   1200
            Width           =   1515
         End
         Begin VB.CommandButton Command12 
            BackColor       =   &H00FF8080&
            Caption         =   "»‰“Ì‰ 95"
            Enabled         =   0   'False
            Height          =   315
            Left            =   9900
            MaskColor       =   &H0080FFFF&
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   213
            Top             =   1590
            Width           =   1545
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8250
            RightToLeft     =   -1  'True
            TabIndex        =   212
            Top             =   1620
            Width           =   1515
         End
         Begin VB.TextBox Text10 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6300
            RightToLeft     =   -1  'True
            TabIndex        =   211
            Top             =   1620
            Width           =   1515
         End
         Begin VB.TextBox Text11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            Height          =   315
            Left            =   4470
            RightToLeft     =   -1  'True
            TabIndex        =   210
            Top             =   1620
            Width           =   1515
         End
         Begin VB.TextBox Text12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            Height          =   315
            Left            =   2700
            RightToLeft     =   -1  'True
            TabIndex        =   209
            Top             =   1620
            Width           =   1515
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H00404080&
            Caption         =   "”Ê·«—"
            Enabled         =   0   'False
            Height          =   315
            Left            =   9930
            MaskColor       =   &H0080FFFF&
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   208
            Top             =   2010
            Width           =   1545
         End
         Begin VB.TextBox Text13 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8250
            RightToLeft     =   -1  'True
            TabIndex        =   207
            Top             =   2010
            Width           =   1515
         End
         Begin VB.TextBox Text14 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6300
            RightToLeft     =   -1  'True
            TabIndex        =   206
            Top             =   2010
            Width           =   1515
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            Height          =   315
            Left            =   4470
            RightToLeft     =   -1  'True
            TabIndex        =   205
            Top             =   2010
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.TextBox Text16 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            Height          =   315
            Left            =   2700
            RightToLeft     =   -1  'True
            TabIndex        =   204
            Top             =   2010
            Width           =   1515
         End
         Begin VB.CommandButton Command14 
            BackColor       =   &H000040C0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   9930
            MaskColor       =   &H0080FFFF&
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   203
            Top             =   2430
            Width           =   1545
         End
         Begin VB.TextBox Text17 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8250
            RightToLeft     =   -1  'True
            TabIndex        =   202
            Top             =   2430
            Width           =   1515
         End
         Begin VB.TextBox Text18 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6300
            RightToLeft     =   -1  'True
            TabIndex        =   201
            Top             =   2430
            Width           =   1515
         End
         Begin VB.TextBox Text19 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            Height          =   315
            Left            =   4470
            RightToLeft     =   -1  'True
            TabIndex        =   200
            Top             =   2430
            Width           =   1515
         End
         Begin VB.TextBox Text20 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            Height          =   315
            Left            =   2700
            RightToLeft     =   -1  'True
            TabIndex        =   199
            Top             =   2430
            Width           =   1515
         End
         Begin VB.CommandButton Command15 
            BackColor       =   &H0080FFFF&
            Enabled         =   0   'False
            Height          =   315
            Left            =   9900
            MaskColor       =   &H0080FFFF&
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   198
            Top             =   2910
            Width           =   1545
         End
         Begin VB.TextBox Text21 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8220
            RightToLeft     =   -1  'True
            TabIndex        =   197
            Top             =   2910
            Width           =   1515
         End
         Begin VB.TextBox Text22 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6270
            RightToLeft     =   -1  'True
            TabIndex        =   196
            Top             =   2910
            Width           =   1515
         End
         Begin VB.TextBox Text23 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            Height          =   315
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   195
            Top             =   2910
            Width           =   1515
         End
         Begin VB.TextBox Text24 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            Height          =   315
            Left            =   2670
            RightToLeft     =   -1  'True
            TabIndex        =   194
            Top             =   2910
            Width           =   1515
         End
         Begin VB.TextBox Text25 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8280
            RightToLeft     =   -1  'True
            TabIndex        =   193
            Top             =   3750
            Width           =   1515
         End
         Begin VB.TextBox Text26 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6330
            RightToLeft     =   -1  'True
            TabIndex        =   192
            Top             =   3750
            Width           =   1515
         End
         Begin VB.TextBox Text27 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            Height          =   315
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   191
            Top             =   3750
            Width           =   1515
         End
         Begin VB.TextBox Text28 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000A&
            Height          =   315
            Left            =   2730
            RightToLeft     =   -1  'True
            TabIndex        =   190
            Top             =   3750
            Width           =   1515
         End
         Begin VB.TextBox Text29 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9960
            RightToLeft     =   -1  'True
            TabIndex        =   189
            Top             =   3720
            Width           =   1515
         End
         Begin VB.TextBox Text30 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Height          =   315
            Left            =   4650
            RightToLeft     =   -1  'True
            TabIndex        =   188
            Top             =   4680
            Width           =   1515
         End
         Begin VB.TextBox Text31 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            Height          =   315
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   187
            Top             =   4680
            Width           =   1515
         End
         Begin VB.CommandButton Command9 
            Caption         =   "«÷«›… «·„»Ì⁄«  «·‰ﬁœÌ… ··Œ“Ì‰…"
            Height          =   525
            Left            =   8010
            RightToLeft     =   -1  'True
            TabIndex        =   186
            Top             =   4740
            Width           =   3555
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   285
            Left            =   8760
            TabIndex        =   220
            TabStop         =   0   'False
            Top             =   390
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   503
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            Format          =   183697411
            CurrentDate     =   41640
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "Frm_Grouped_New3M.frx":000C
            Height          =   315
            Index           =   1
            Left            =   8670
            TabIndex        =   221
            Top             =   4380
            Width           =   2880
            _ExtentX        =   5080
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
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
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   285
            Left            =   4575
            TabIndex        =   239
            TabStop         =   0   'False
            Top             =   540
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   503
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            Format          =   183697411
            CurrentDate     =   41640
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   375
            Left            =   7095
            TabIndex        =   240
            TabStop         =   0   'False
            ToolTipText     =   "«÷€ÿ ·«÷«›… ⁄„Ì· ÃœÌœ"
            Top             =   525
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   661
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
            ButtonImage     =   "Frm_Grouped_New3M.frx":0021
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»ÕÀ »—ﬁ„ «·Ê—œÌ…"
            Height          =   255
            Index           =   58
            Left            =   8925
            RightToLeft     =   -1  'True
            TabIndex        =   243
            Top             =   525
            Width           =   1605
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·„÷Œ…"
            Height          =   210
            Index           =   28
            Left            =   7095
            RightToLeft     =   -1  'True
            TabIndex        =   242
            Top             =   60
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»ÕÀ » «—ÌŒ «·ÌÊ„"
            Height          =   255
            Index           =   62
            Left            =   5685
            RightToLeft     =   -1  'True
            TabIndex        =   241
            Top             =   570
            Width           =   1380
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "» «—ÌŒ «·ÌÊ„"
            Height          =   255
            Index           =   63
            Left            =   9900
            RightToLeft     =   -1  'True
            TabIndex        =   236
            Top             =   420
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·ÊﬁÊœ"
            Height          =   270
            Index           =   61
            Left            =   10410
            TabIndex        =   235
            Top             =   840
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ê—œÌ… —ﬁ„"
            Height          =   255
            Index           =   64
            Left            =   4890
            RightToLeft     =   -1  'True
            TabIndex        =   234
            Top             =   330
            Width           =   1515
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "’«›Ï ﬂ„Ì… „»«⁄  ﬁœÌ"
            Height          =   270
            Index           =   67
            Left            =   8040
            TabIndex        =   233
            Top             =   810
            Width           =   1680
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì ﬁÌ„… „»«⁄ ‰ﬁœÌ"
            Height          =   270
            Index           =   68
            Left            =   6150
            TabIndex        =   232
            Top             =   810
            Width           =   1680
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "’«›Ï ﬂ„Ì… „»«⁄ ¬Ã·"
            Height          =   270
            Index           =   69
            Left            =   4290
            TabIndex        =   231
            Top             =   810
            Width           =   1680
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì ﬁÌ„… „»«⁄ ¬Ã·"
            Height          =   270
            Index           =   70
            Left            =   2520
            TabIndex        =   230
            Top             =   810
            Width           =   1680
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì ﬂ„Ì… „»«⁄  ﬁœÌ"
            Height          =   270
            Index           =   71
            Left            =   8070
            TabIndex        =   229
            Top             =   3360
            Width           =   1680
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì ﬁÌ„… „»«⁄ ‰ﬁœÌ"
            Height          =   270
            Index           =   72
            Left            =   6180
            TabIndex        =   228
            Top             =   3360
            Width           =   1680
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì ﬂ„Ì… „»«⁄ ¬Ã·"
            Height          =   270
            Index           =   73
            Left            =   4320
            TabIndex        =   227
            Top             =   3360
            Width           =   1680
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì ﬁÌ„… „»«⁄ ¬Ã·"
            Height          =   270
            Index           =   74
            Left            =   2550
            TabIndex        =   226
            Top             =   3360
            Width           =   1680
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‘—› «·Ê—œÌ…"
            Height          =   270
            Index           =   75
            Left            =   9750
            TabIndex        =   225
            Top             =   3330
            Width           =   1680
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Œ — «·Œ“Ì‰…"
            Height          =   270
            Index           =   76
            Left            =   9810
            TabIndex        =   224
            Top             =   4110
            Width           =   1680
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì ﬂ„Ì… «·„»«⁄ "
            Height          =   270
            Index           =   79
            Left            =   4470
            TabIndex        =   223
            Top             =   4290
            Width           =   1680
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì ﬁÌ„… «·„»Ì⁄« "
            Height          =   270
            Index           =   80
            Left            =   2640
            TabIndex        =   222
            Top             =   4290
            Width           =   1680
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic100 
         Height          =   7830
         Left            =   -9000
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   45
         Width           =   8055
         _cx             =   14208
         _cy             =   13811
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
         Begin VB.TextBox txtQRCODE 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4560
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   183
            Top             =   5520
            Width           =   1515
         End
         Begin VB.PictureBox Picture2 
            Height          =   855
            Left            =   3720
            Picture         =   "Frm_Grouped_New3M.frx":041E
            RightToLeft     =   -1  'True
            ScaleHeight     =   795
            ScaleWidth      =   795
            TabIndex        =   182
            Top             =   5400
            Width           =   855
         End
         Begin C1SizerLibCtl.C1Elastic EleHeader 
            Height          =   675
            Left            =   0
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Width           =   7980
            _cx             =   14076
            _cy             =   1191
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
            Begin VB.TextBox TxtModFlg 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   2700
               RightToLeft     =   -1  'True
               TabIndex        =   11
               Top             =   120
               Visible         =   0   'False
               Width           =   945
            End
            Begin ImpulseButton.ISButton btnLast 
               Height          =   345
               Left            =   390
               TabIndex        =   12
               Top             =   150
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               ButtonStyle     =   1
               ButtonPositionImage=   4
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":3020
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext 
               Height          =   345
               Left            =   855
               TabIndex        =   13
               Top             =   150
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               ButtonStyle     =   1
               ButtonPositionImage=   4
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":33BA
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious 
               Height          =   345
               Left            =   1305
               TabIndex        =   14
               Top             =   150
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               ButtonStyle     =   1
               ButtonPositionImage=   4
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":3754
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst 
               Height          =   345
               Left            =   1770
               TabIndex        =   15
               Top             =   150
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               ButtonStyle     =   1
               ButtonPositionImage=   4
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":3AEE
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "ÊÕœ«  «·√’‰«›"
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
               Height          =   375
               Index           =   2
               Left            =   4680
               TabIndex        =   35
               Top             =   120
               Width           =   3015
            End
         End
         Begin C1SizerLibCtl.C1Elastic EltCont 
            Height          =   990
            Left            =   240
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   6480
            Width           =   6870
            _cx             =   12118
            _cy             =   1746
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
            Begin ImpulseButton.ISButton btnNew 
               Height          =   420
               Left            =   5805
               TabIndex        =   17
               Top             =   495
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":3E88
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave 
               Height          =   420
               Left            =   4290
               TabIndex        =   18
               Top             =   495
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":4222
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify 
               Height          =   420
               Left            =   4935
               TabIndex        =   19
               Top             =   495
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":45BC
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo 
               Height          =   420
               Left            =   3405
               TabIndex        =   20
               Top             =   495
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":4956
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete 
               Height          =   420
               Left            =   990
               TabIndex        =   21
               Top             =   495
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":4CF0
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnQuery 
               Height          =   330
               Left            =   2640
               TabIndex        =   22
               TabStop         =   0   'False
               ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
               Top             =   570
               Width           =   720
               _ExtentX        =   1270
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":528A
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUpdate 
               Height          =   330
               Left            =   5925
               TabIndex        =   23
               TabStop         =   0   'False
               ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
               Top             =   105
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":5624
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnCancel 
               Height          =   420
               Left            =   75
               TabIndex        =   24
               Top             =   495
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":59BE
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton ISButton1 
               Height          =   285
               Left            =   1800
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   600
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":5D58
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   0
               Left            =   3555
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   135
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   1
               Left            =   960
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   135
               Width           =   975
            End
            Begin VB.Label LabCurrRec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   2745
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   165
               Width           =   675
            End
            Begin VB.Label LabCountRec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   270
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   135
               Width           =   540
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   3270
            Left            =   120
            TabIndex        =   30
            Top             =   795
            Width           =   7650
            _cx             =   13494
            _cy             =   5768
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
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"Frm_Grouped_New3M.frx":60F2
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
         Begin VB.Frame Frm2 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1470
            Left            =   270
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   4080
            Width           =   7590
            Begin VB.TextBox TxtVacName 
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
               Height          =   345
               Left            =   195
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   6
               Top             =   465
               Width           =   5580
            End
            Begin VB.TextBox TxtUnitID 
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
               Height          =   345
               Left            =   4470
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   5
               Top             =   105
               Width           =   1305
            End
            Begin VB.ComboBox CmbType 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "Frm_Grouped_New3M.frx":61A0
               Left            =   2280
               List            =   "Frm_Grouped_New3M.frx":61B0
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   2550
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.TextBox TxtVacNamee 
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
               Height          =   345
               Left            =   195
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   3
               Top             =   840
               Width           =   5580
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·ÊÕœ… ⁄—»Ì"
               Height          =   255
               Index           =   0
               Left            =   6030
               RightToLeft     =   -1  'True
               TabIndex        =   9
               Top             =   450
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„”·”·"
               Height          =   285
               Index           =   3
               Left            =   6285
               RightToLeft     =   -1  'True
               TabIndex        =   8
               Top             =   90
               Width           =   810
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·ÊÕœ… «‰Ã·Ì“Ì"
               Height          =   255
               Index           =   1
               Left            =   6000
               RightToLeft     =   -1  'True
               TabIndex        =   7
               Top             =   840
               Width           =   1530
            End
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "«· ÊﬁÌ⁄ «·«·ﬂ —Ê‰Ì"
            Height          =   375
            Index           =   15
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   184
            Top             =   5520
            Width           =   1455
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   7830
         Left            =   -8700
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   45
         Width           =   8055
         _cx             =   14208
         _cy             =   13811
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
         Begin VB.TextBox XPMTxtRemark 
            Alignment       =   1  'Right Justify
            Height          =   3555
            Left            =   465
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   44
            Top             =   2490
            Width           =   6180
         End
         Begin VB.TextBox XPTxtBoxName 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   465
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   1755
            Width           =   6180
         End
         Begin VB.TextBox XPTxtBoxID 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   4590
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   1155
            Width           =   2055
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic5 
            Height          =   855
            Left            =   0
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   0
            Width           =   8055
            _cx             =   14208
            _cy             =   1508
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial (Arabic)"
               Size            =   18
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
            Begin VB.TextBox TxtModFlg1 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   2610
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   180
               Visible         =   0   'False
               Width           =   855
            End
            Begin ImpulseButton.ISButton XPBtnMove 
               Height          =   345
               Index           =   0
               Left            =   1155
               TabIndex        =   38
               Top             =   120
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   609
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":61C9
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
               Height          =   345
               Index           =   2
               Left            =   90
               TabIndex        =   39
               Top             =   120
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   609
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":6563
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
               Height          =   345
               Index           =   1
               Left            =   1680
               TabIndex        =   40
               Top             =   120
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   609
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":68FD
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
               Height          =   345
               Index           =   3
               Left            =   615
               TabIndex        =   41
               Top             =   120
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   609
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":6C97
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
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰«  √·Ê«‰ «·√’‰«›"
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
               Height          =   375
               Index           =   4
               Left            =   4080
               TabIndex        =   59
               Top             =   120
               Width           =   3735
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic6 
            Height          =   1185
            Left            =   465
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   6495
            Width           =   7020
            _cx             =   12383
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
            Begin ImpulseButton.ISButton Cmd 
               Height          =   420
               Index           =   0
               Left            =   5805
               TabIndex        =   49
               Top             =   495
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":7031
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   420
               Index           =   2
               Left            =   3930
               TabIndex        =   50
               Top             =   495
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":73CB
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   420
               Index           =   1
               Left            =   4815
               TabIndex        =   51
               Top             =   495
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":7765
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   420
               Index           =   3
               Left            =   2805
               TabIndex        =   52
               Top             =   495
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":7AFF
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   420
               Index           =   4
               Left            =   1830
               TabIndex        =   53
               Top             =   495
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":7E99
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   420
               Index           =   6
               Left            =   75
               TabIndex        =   54
               Top             =   495
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":8433
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label XPTxtCount 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   270
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   135
               Width           =   540
            End
            Begin VB.Label XPTxtCurrent 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   2745
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   165
               Width           =   675
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   3
               Left            =   960
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   135
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   2
               Left            =   3555
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   135
               Width           =   975
            End
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «··Ê‰"
            Height          =   345
            Index           =   0
            Left            =   6840
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   1170
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   375
            Index           =   1
            Left            =   6840
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   3210
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «··Ê‰"
            Height          =   375
            Index           =   3
            Left            =   6840
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   1755
            Width           =   1020
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   7830
         Left            =   45
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   45
         Width           =   8055
         _cx             =   14208
         _cy             =   13811
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
         Begin VB.TextBox XPMTxtRemark2 
            Alignment       =   1  'Right Justify
            Height          =   3405
            Left            =   465
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   62
            Top             =   2670
            Width           =   6090
         End
         Begin VB.TextBox XPTxtBoxName2 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   465
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   1935
            Width           =   6090
         End
         Begin VB.TextBox XPTxtBoxID2 
            Alignment       =   1  'Right Justify
            Height          =   390
            Left            =   3840
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   1185
            Width           =   2715
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic7 
            Height          =   855
            Left            =   0
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   0
            Width           =   8055
            _cx             =   14208
            _cy             =   1508
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial (Arabic)"
               Size            =   18
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
            Begin VB.TextBox TxtModFlg2 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   3090
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   180
               Visible         =   0   'False
               Width           =   855
            End
            Begin ImpulseButton.ISButton XPBtnMove2 
               Height          =   345
               Index           =   0
               Left            =   1155
               TabIndex        =   65
               Top             =   120
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   609
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":87CD
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
            Begin ImpulseButton.ISButton XPBtnMove2 
               Height          =   345
               Index           =   2
               Left            =   90
               TabIndex        =   66
               Top             =   120
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   609
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":8B67
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
            Begin ImpulseButton.ISButton XPBtnMove2 
               Height          =   345
               Index           =   1
               Left            =   1680
               TabIndex        =   67
               Top             =   120
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   609
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":8F01
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
            Begin ImpulseButton.ISButton XPBtnMove2 
               Height          =   345
               Index           =   3
               Left            =   615
               TabIndex        =   68
               Top             =   120
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   609
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":929B
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
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰«  „ﬁ«”«  «·√’‰«›"
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
               Height          =   375
               Index           =   5
               Left            =   4080
               TabIndex        =   83
               Top             =   120
               Width           =   3735
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic8 
            Height          =   1185
            Left            =   465
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   6495
            Width           =   7020
            _cx             =   12383
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
            Begin ImpulseButton.ISButton Cmd2 
               Height          =   420
               Index           =   0
               Left            =   5805
               TabIndex        =   73
               Top             =   495
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":9635
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd2 
               Height          =   420
               Index           =   2
               Left            =   3930
               TabIndex        =   74
               Top             =   495
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":99CF
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd2 
               Height          =   420
               Index           =   1
               Left            =   4815
               TabIndex        =   75
               Top             =   495
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":9D69
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd2 
               Height          =   420
               Index           =   3
               Left            =   2805
               TabIndex        =   76
               Top             =   495
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":A103
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd2 
               Height          =   420
               Index           =   4
               Left            =   1830
               TabIndex        =   77
               Top             =   495
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":A49D
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd2 
               Height          =   420
               Index           =   6
               Left            =   75
               TabIndex        =   78
               Top             =   495
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":AA37
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   5
               Left            =   3555
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   135
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   4
               Left            =   960
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   135
               Width           =   975
            End
            Begin VB.Label XPTxtCurrent2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   2745
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   165
               Width           =   675
            End
            Begin VB.Label XPTxtCount2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   270
               RightToLeft     =   -1  'True
               TabIndex        =   79
               Top             =   120
               Width           =   540
            End
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·„ﬁ«”"
            Height          =   330
            Index           =   5
            Left            =   6555
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   1215
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   240
            Index           =   4
            Left            =   6555
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   3240
            Width           =   1410
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„ﬁ«”"
            Height          =   375
            Index           =   2
            Left            =   6645
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   1935
            Width           =   1215
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   7830
         Left            =   8790
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   45
         Width           =   8055
         _cx             =   14208
         _cy             =   13811
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
         Begin VB.TextBox XPMTxtRemark3 
            Alignment       =   1  'Right Justify
            Height          =   2400
            Left            =   660
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   104
            Top             =   3570
            Width           =   5715
         End
         Begin VB.TextBox XPTxtBoxName3 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   660
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   2040
            Width           =   5715
         End
         Begin VB.TextBox XPTxtBoxID3 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   4410
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   1440
            Width           =   1965
         End
         Begin VB.TextBox XPTxtBoxNamee 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   660
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   101
            Top             =   2565
            Width           =   5715
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic9 
            Height          =   1185
            Left            =   465
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   6495
            Width           =   7020
            _cx             =   12383
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
            Begin ImpulseButton.ISButton Cmd3 
               Height          =   420
               Index           =   0
               Left            =   5805
               TabIndex        =   85
               Top             =   495
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":ADD1
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd3 
               Height          =   420
               Index           =   2
               Left            =   3930
               TabIndex        =   86
               Top             =   495
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":B16B
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd3 
               Height          =   420
               Index           =   1
               Left            =   4815
               TabIndex        =   87
               Top             =   495
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":B505
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd3 
               Height          =   420
               Index           =   3
               Left            =   2805
               TabIndex        =   88
               Top             =   495
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":B89F
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd3 
               Height          =   420
               Index           =   4
               Left            =   1830
               TabIndex        =   89
               Top             =   495
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":BC39
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd3 
               Height          =   420
               Index           =   6
               Left            =   75
               TabIndex        =   90
               Top             =   495
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   741
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":C1D3
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label XPTxtCount3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   270
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   120
               Width           =   540
            End
            Begin VB.Label XPTxtCurrent3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   2745
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   165
               Width           =   675
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   7
               Left            =   960
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   135
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   6
               Left            =   3555
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   135
               Width           =   975
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic10 
            Height          =   855
            Left            =   0
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   0
            Width           =   8055
            _cx             =   14208
            _cy             =   1508
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial (Arabic)"
               Size            =   18
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
            Begin VB.TextBox TxtModFlg3 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   2730
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   180
               Visible         =   0   'False
               Width           =   855
            End
            Begin ImpulseButton.ISButton XPBtnMove3 
               Height          =   345
               Index           =   0
               Left            =   1155
               TabIndex        =   97
               Top             =   120
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   609
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":C56D
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
            Begin ImpulseButton.ISButton XPBtnMove3 
               Height          =   345
               Index           =   2
               Left            =   90
               TabIndex        =   98
               Top             =   120
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   609
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":C907
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
            Begin ImpulseButton.ISButton XPBtnMove3 
               Height          =   345
               Index           =   1
               Left            =   1680
               TabIndex        =   99
               Top             =   120
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   609
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":CCA1
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
            Begin ImpulseButton.ISButton XPBtnMove3 
               Height          =   345
               Index           =   3
               Left            =   615
               TabIndex        =   100
               Top             =   120
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   609
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":D03B
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
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«„«ﬂ‰ «· Œ“Ì‰"
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
               Height          =   375
               Index           =   6
               Left            =   4080
               TabIndex        =   109
               Top             =   120
               Width           =   3735
            End
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   570
            TabIndex        =   180
            Top             =   3015
            Width           =   5805
            _ExtentX        =   10239
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Œ“‰"
            Height          =   240
            Index           =   24
            Left            =   7035
            RightToLeft     =   -1  'True
            TabIndex        =   181
            Top             =   3030
            Width           =   450
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ "
            Height          =   345
            Index           =   9
            Left            =   6555
            RightToLeft     =   -1  'True
            TabIndex        =   108
            Top             =   1440
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   240
            Index           =   8
            Left            =   6555
            RightToLeft     =   -1  'True
            TabIndex        =   107
            Top             =   4005
            Width           =   1230
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„ ⁄—»Ì"
            Height          =   375
            Index           =   7
            Left            =   6555
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Top             =   2040
            Width           =   1230
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„ «‰Ã·Ì“Ì"
            Height          =   375
            Index           =   6
            Left            =   6555
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   2565
            Width           =   1230
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   7830
         Left            =   9090
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   45
         Width           =   8055
         _cx             =   14208
         _cy             =   13811
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
         Begin VB.Frame Frm24 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1500
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   121
            Top             =   5145
            Width           =   7335
            Begin VB.TextBox TxtSerial 
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
               Left            =   4290
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   124
               Top             =   45
               Width           =   1785
            End
            Begin VB.TextBox TxtVacName4 
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
               Left            =   315
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   123
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„  «·‰Ê⁄"
               Top             =   405
               Width           =   5760
            End
            Begin VB.TextBox TxtVacNamee4 
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
               Left            =   315
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   122
               Top             =   765
               Width           =   5760
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ  "
               Height          =   195
               Index           =   10
               Left            =   5985
               RightToLeft     =   -1  'True
               TabIndex        =   127
               Top             =   30
               Width           =   1110
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄—»Ì"
               Height          =   285
               Index           =   9
               Left            =   6060
               RightToLeft     =   -1  'True
               TabIndex        =   126
               Top             =   360
               Width           =   1035
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«‰Ã·Ì“Ì"
               Height          =   285
               Index           =   8
               Left            =   5985
               RightToLeft     =   -1  'True
               TabIndex        =   125
               Top             =   720
               Width           =   1170
            End
         End
         Begin VB.Frame FraHeader 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   645
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   110
            Top             =   0
            Width           =   8055
            Begin VB.Frame Frmo2 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   113
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DCUser 
                  CausesValidation=   0   'False
                  Height          =   315
                  Left            =   -255
                  TabIndex        =   114
                  Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
                  Top             =   15
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
                  TabIndex        =   115
                  Top             =   45
                  Width           =   855
               End
            End
            Begin VB.TextBox TxtModFlg4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   2100
               RightToLeft     =   -1  'True
               TabIndex        =   112
               Text            =   "modflag"
               Top             =   -150
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.TextBox TxtVac_ID 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Left            =   2790
               RightToLeft     =   -1  'True
               TabIndex        =   111
               Top             =   -90
               Visible         =   0   'False
               Width           =   945
            End
            Begin MSComctlLib.ImageList GrdImageList 
               Left            =   3840
               Top             =   480
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
                     Picture         =   "Frm_Grouped_New3M.frx":D3D5
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New3M.frx":D76F
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New3M.frx":DB09
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New3M.frx":DEA3
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New3M.frx":E23D
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New3M.frx":E5D7
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New3M.frx":E971
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New3M.frx":EF0B
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btnLast4 
               Height          =   315
               Left            =   90
               TabIndex        =   116
               Top             =   150
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":F2A5
               ColorButton     =   16777215
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext4 
               Height          =   315
               Left            =   555
               TabIndex        =   117
               Top             =   150
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":F63F
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious4 
               Height          =   315
               Left            =   1155
               TabIndex        =   118
               Top             =   150
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":F9D9
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst4 
               Height          =   315
               Left            =   1620
               TabIndex        =   119
               Top             =   150
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":FD73
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   " ⁄—Ì› «·„Ê«’›« "
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
               Height          =   375
               Index           =   7
               Left            =   4095
               RightToLeft     =   -1  'True
               TabIndex        =   120
               Top             =   90
               Width           =   3720
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic11 
            Height          =   1080
            Left            =   930
            TabIndex        =   128
            TabStop         =   0   'False
            Top             =   6660
            Width           =   6270
            _cx             =   11060
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
            Begin ImpulseButton.ISButton btnNew4 
               Height          =   330
               Left            =   4575
               TabIndex        =   129
               Top             =   555
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":1010D
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave4 
               Height          =   330
               Left            =   3030
               TabIndex        =   130
               Top             =   555
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":104A7
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify4 
               Height          =   330
               Left            =   3795
               TabIndex        =   131
               Top             =   555
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":10841
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo4 
               Height          =   330
               Left            =   2265
               TabIndex        =   132
               Top             =   555
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":10BDB
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete4 
               Height          =   330
               Left            =   1500
               TabIndex        =   133
               Top             =   555
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":10F75
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnQuery4 
               Height          =   330
               Left            =   5040
               TabIndex        =   134
               TabStop         =   0   'False
               ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
               Top             =   90
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":1150F
               ColorButton     =   14737632
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUpdate4 
               Height          =   330
               Left            =   3765
               TabIndex        =   135
               TabStop         =   0   'False
               ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
               Top             =   105
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":118A9
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnPrint 
               Height          =   285
               Left            =   4725
               TabIndex        =   136
               TabStop         =   0   'False
               Top             =   150
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               ButtonStyle     =   1
               ButtonPositionImage=   2
               Caption         =   ""
               BackColor       =   14871017
               FontSize        =   14.25
               FontName        =   "Arial"
               FontBold        =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "Frm_Grouped_New3M.frx":11C43
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnCancel4 
               Height          =   330
               Left            =   705
               TabIndex        =   137
               Top             =   555
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":11FDD
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label LabCountRec4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   141
               Top             =   225
               Width           =   540
            End
            Begin VB.Label LabCurrRec4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   140
               Top             =   240
               Width           =   675
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   9
               Left            =   810
               RightToLeft     =   -1  'True
               TabIndex        =   139
               Top             =   225
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   8
               Left            =   2505
               RightToLeft     =   -1  'True
               TabIndex        =   138
               Top             =   225
               Width           =   975
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid4 
            Height          =   4275
            Left            =   0
            TabIndex        =   142
            Top             =   690
            Width           =   8055
            _cx             =   14208
            _cy             =   7541
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
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"Frm_Grouped_New3M.frx":12377
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic12 
         Height          =   7830
         Left            =   9390
         TabIndex        =   143
         TabStop         =   0   'False
         Top             =   45
         Width           =   8055
         _cx             =   14208
         _cy             =   13811
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
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   645
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   154
            Top             =   0
            Width           =   8055
            Begin VB.TextBox TxtVac_ID5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   285
               Left            =   2430
               RightToLeft     =   -1  'True
               TabIndex        =   159
               Top             =   -210
               Visible         =   0   'False
               Width           =   1065
            End
            Begin VB.TextBox TxtModFlg5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   3780
               RightToLeft     =   -1  'True
               TabIndex        =   158
               Text            =   "modflag"
               Top             =   450
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   155
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DCUser5 
                  CausesValidation=   0   'False
                  Height          =   315
                  Left            =   -255
                  TabIndex        =   156
                  Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
                  Top             =   15
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
                  Index           =   16
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   157
                  Top             =   45
                  Width           =   855
               End
            End
            Begin MSComctlLib.ImageList GrdImageList5 
               Left            =   3720
               Top             =   -480
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
                     Picture         =   "Frm_Grouped_New3M.frx":123FF
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New3M.frx":12799
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New3M.frx":12B33
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New3M.frx":12ECD
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New3M.frx":13267
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New3M.frx":13601
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New3M.frx":1399B
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "Frm_Grouped_New3M.frx":13F35
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btnLast5 
               Height          =   315
               Left            =   90
               TabIndex        =   160
               Top             =   150
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":142CF
               ColorButton     =   16777215
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext5 
               Height          =   315
               Left            =   555
               TabIndex        =   161
               Top             =   150
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":14669
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious5 
               Height          =   315
               Left            =   1155
               TabIndex        =   162
               Top             =   150
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":14A03
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst5 
               Height          =   315
               Left            =   1620
               TabIndex        =   163
               Top             =   150
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":14D9D
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "⁄‰«’— «· ﬂ«·Ì› «·’‰«⁄ÌÂ"
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
               Height          =   375
               Index           =   17
               Left            =   4215
               RightToLeft     =   -1  'True
               TabIndex        =   164
               Top             =   90
               Width           =   3630
            End
         End
         Begin VB.Frame Frm25 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Caption         =   "”"
            Enabled         =   0   'False
            Height          =   1785
            Left            =   465
            RightToLeft     =   -1  'True
            TabIndex        =   144
            Top             =   4815
            Width           =   7230
            Begin VB.TextBox TxtVacName5 
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
               Left            =   240
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   149
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·„Õ«›Ÿ…"
               Top             =   390
               Width           =   4890
            End
            Begin VB.TextBox TxtSerial5 
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
               Left            =   4065
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   148
               Top             =   30
               Width           =   1065
            End
            Begin VB.ComboBox Combo1 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "Frm_Grouped_New3M.frx":15137
               Left            =   2280
               List            =   "Frm_Grouped_New3M.frx":15147
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   147
               Top             =   3150
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.TextBox TxtVacNamee5 
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
               Left            =   240
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   145
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·„Õ«›Ÿ…"
               Top             =   720
               Width           =   4890
            End
            Begin MSDataListLib.DataCombo DcboExpensesID 
               Height          =   315
               Left            =   240
               TabIndex        =   146
               Tag             =   "«Œ — «·œÊ·… „‰ ›÷·ﬂ"
               Top             =   1110
               Width           =   4890
               _ExtentX        =   8625
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·⁄‰’— ⁄—»Ì"
               Height          =   285
               Index           =   15
               Left            =   5220
               RightToLeft     =   -1  'True
               TabIndex        =   153
               Top             =   390
               Width           =   1650
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·⁄‰’—"
               Height          =   195
               Index           =   14
               Left            =   5565
               RightToLeft     =   -1  'True
               TabIndex        =   152
               Top             =   30
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„’—Ê›"
               Height          =   285
               Index           =   12
               Left            =   5400
               RightToLeft     =   -1  'True
               TabIndex        =   151
               Top             =   1110
               Width           =   1290
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·⁄‰’— «‰Ã·Ì“Ì"
               Height          =   285
               Index           =   11
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   150
               Top             =   720
               Width           =   1530
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic13 
            Height          =   1095
            Left            =   855
            TabIndex        =   165
            TabStop         =   0   'False
            Top             =   6615
            Width           =   5700
            _cx             =   10054
            _cy             =   1931
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
            Begin ImpulseButton.ISButton btnNew5 
               Height          =   330
               Left            =   4575
               TabIndex        =   166
               Top             =   555
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":15160
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave5 
               Height          =   330
               Left            =   3030
               TabIndex        =   167
               Top             =   555
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":154FA
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify5 
               Height          =   330
               Left            =   3795
               TabIndex        =   168
               Top             =   555
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":15894
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo5 
               Height          =   330
               Left            =   2265
               TabIndex        =   169
               Top             =   555
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":15C2E
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete5 
               Height          =   330
               Left            =   1500
               TabIndex        =   170
               Top             =   555
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":15FC8
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnQuery5 
               Height          =   330
               Left            =   5160
               TabIndex        =   171
               TabStop         =   0   'False
               ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
               Top             =   90
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":16562
               ColorButton     =   14737632
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUpdate5 
               Height          =   330
               Left            =   3765
               TabIndex        =   172
               TabStop         =   0   'False
               ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
               Top             =   105
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":168FC
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton21 
               Height          =   285
               Left            =   4725
               TabIndex        =   173
               TabStop         =   0   'False
               Top             =   150
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               ButtonStyle     =   1
               ButtonPositionImage=   2
               Caption         =   ""
               BackColor       =   14871017
               FontSize        =   14.25
               FontName        =   "Arial"
               FontBold        =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "Frm_Grouped_New3M.frx":16C96
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnCancel5 
               Height          =   330
               Left            =   705
               TabIndex        =   174
               Top             =   555
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "Frm_Grouped_New3M.frx":17030
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   11
               Left            =   2505
               RightToLeft     =   -1  'True
               TabIndex        =   178
               Top             =   225
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   10
               Left            =   810
               RightToLeft     =   -1  'True
               TabIndex        =   177
               Top             =   225
               Width           =   975
            End
            Begin VB.Label LabCurrRec5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   176
               Top             =   240
               Width           =   675
            End
            Begin VB.Label LabCountRec5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   175
               Top             =   225
               Width           =   540
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid Grid5 
            Height          =   3960
            Left            =   0
            TabIndex        =   179
            Top             =   720
            Width           =   8055
            _cx             =   14208
            _cy             =   6985
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
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"Frm_Grouped_New3M.frx":173CA
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
Attribute VB_Name = "FrmPay_Garanty_Shipment3M"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SendForm As Integer

Dim RsSavRec As ADODB.Recordset
Dim RecId As String
Dim ii As Long

Dim rs As ADODB.Recordset
Dim TTP As clstooltip

Dim rs2 As ADODB.Recordset
Dim TTP2 As clstooltip

Dim Rs3 As ADODB.Recordset
Dim TTP3 As clstooltip

Dim RsSavRec4 As ADODB.Recordset
Dim BKGrndPic4 As ClsBackGroundPic
Dim RecId4 As String
'Dim II4 As Long
Public mIsQrCode As Boolean
Dim RsSavRec5 As ADODB.Recordset
Dim BKGrndPic5 As ClsBackGroundPic
Dim RecId5 As String
'Dim II5 As Long
Dim cSearch  As clsDCboSearch



Private Sub DCboStoreName_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetStores Me.DCboStoreName
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

Private Sub TxtStoreID_KeyDown(KeyCode As Integer, Shift As Integer)
' Dim StoreId As Integer

'    If KeyCode = vbKeyReturn Then
'    StoreId = getStoreInformatin(TxtStoreID)
'        DCboStoreName.BoundText = StoreId
'    End If
End Sub

Private Sub ChangeLang()

    Label1(3).Caption = "Code"
    Label1(0).Caption = "Name Ar"
    Label1(1).Caption = "Name Eng"
    ISButton1.Caption = "Prient"
    btnQuery.Caption = "Search"
    With Grid
        .TextMatrix(0, .ColIndex("UnitID")) = "Unit Code"
        .TextMatrix(0, .ColIndex("UnitName")) = "Unit Name AR"
        .TextMatrix(0, .ColIndex("UnitNameE")) = "Unit Name Eng"
        Label1(2).Caption = "Unit  Data"
        btnNew.Caption = "New"
        btnModify.Caption = "Modify"
        btnSave.Caption = "Save"
        BtnUndo.Caption = "Undo"
        btnDelete.Caption = "Delete"
        btnCancel.Caption = "Exit"
        Label2(0).Caption = "Current Record"
        Label2(1).Caption = "NO Of Record"
    End With

    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic
'#####################################################################################################################################################
    Dim XPic2 As IPictureDisp
    Set XPic2 = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic2
    Set XPic2 = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic2

    Label1(4).Caption = "Color Data"
    lbl(0).Caption = "Color Code"
    lbl(3).Caption = "color  Name"
    lbl(1).Caption = "Remarks"
    Label2(2).Caption = "Current Record"
    Label2(3).Caption = "NO. Recordes"
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(6).Caption = "Exit"
'#####################################################################################################################################################
    Set XPic = Me.XPBtnMove2(1).ButtonImage
    Set Me.XPBtnMove2(1).ButtonImage = Me.XPBtnMove2(2).ButtonImage
    Set Me.XPBtnMove2(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove2(0).ButtonImage
    Set Me.XPBtnMove2(0).ButtonImage = Me.XPBtnMove2(3).ButtonImage
    Set Me.XPBtnMove2(3).ButtonImage = XPic
    
    Label1(5).Caption = "Size Data"
    lbl(5).Caption = "Size Code"
    lbl(2).Caption = "Size  Name"
    Label2(4).Caption = "Remarks"
    Label2(5).Caption = "Current Record"
    lbl(4).Caption = "NO. Recordes"
    Me.Cmd2(0).Caption = "New"
    Me.Cmd2(1).Caption = "Edit"
    Me.Cmd2(2).Caption = "Save"
    Me.Cmd2(3).Caption = "Undo"
    Me.Cmd2(4).Caption = "Delete"
    Me.Cmd2(6).Caption = "Exit"
'#####################################################################################################################################################
    Set XPic = Me.XPBtnMove3(1).ButtonImage
    Set Me.XPBtnMove3(1).ButtonImage = Me.XPBtnMove3(2).ButtonImage
    Set Me.XPBtnMove3(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove3(0).ButtonImage
    Set Me.XPBtnMove3(0).ButtonImage = Me.XPBtnMove3(3).ButtonImage
    Set Me.XPBtnMove3(3).ButtonImage = XPic

    Label1(6).Caption = "Stores Locations"
    lbl(9).Caption = " Code"
    lbl(7).Caption = " Name Ar"
    lbl(6).Caption = " Name Eng"
    lbl(8).Caption = "Remarks"
    Label2(6).Caption = "Current Record"
    Label2(7).Caption = "NO. Recordes"
    Me.Cmd3(0).Caption = "New"
    Me.Cmd3(1).Caption = "Edit"
    Me.Cmd3(2).Caption = "Save"
    Me.Cmd3(3).Caption = "Undo"
    Me.Cmd3(4).Caption = "Delete"
    Me.Cmd3(6).Caption = "Exit"
'#####################################################################################################################################################
    Set XPic = Me.btnFirst4.ButtonImage
    Set Me.btnFirst4.ButtonImage = Me.btnLast4.ButtonImage
    Set Me.btnLast4.ButtonImage = XPic
    Set XPic = Me.btnPrevious4.ButtonImage
    Set Me.btnPrevious4.ButtonImage = Me.btnNext4.ButtonImage
    Set Me.btnNext4.ButtonImage = XPic

    Label1(7).Caption = "Items Specifications"
    Label1(10).Caption = "Code"
    Label1(9).Caption = "Name AR"
    Label1(8).Caption = "Name ENG"

    Label2(8).Caption = "Current Record"
    Label2(9).Caption = "NO. Recordes"

    btnNew4.Caption = "New"
    btnModify4.Caption = "Modify"
    btnSave4.Caption = "Save"
    BtnUndo4.Caption = "Undo"
    btnDelete4.Caption = "Delete"
    btnCancel4.Caption = "Exit"

    With Me.Grid4
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("id")) = "Id"
        .TextMatrix(0, .ColIndex("name")) = "Name AR"
        .TextMatrix(0, .ColIndex("namee")) = "Name ENG"
    End With
'######################################################################################################################################################
    Set XPic = Me.btnFirst5.ButtonImage
    Set Me.btnFirst5.ButtonImage = Me.btnLast5.ButtonImage
    Set Me.btnLast5.ButtonImage = XPic
    Set XPic = Me.btnPrevious5.ButtonImage
    Set Me.btnPrevious5.ButtonImage = Me.btnNext5.ButtonImage
    Set Me.btnNext5.ButtonImage = XPic

    Label1(17).Caption = "Production Cost component "
    With Me.Grid5
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("TBLProductionElementsId")) = "Element ID"
        .TextMatrix(0, .ColIndex("Name")) = " Name A"
        .TextMatrix(0, .ColIndex("NameE")) = " Name E"
        .TextMatrix(0, .ColIndex("ExpensesID")) = "Expenses Name"
    End With
    
    Label1(14).Caption = "ID"
    Label1(15).Caption = "Name AR"
    Label1(11).Caption = "Name En"
    Label1(12).Caption = "Expenses Name"
    Label2(11).Caption = "Curr. Rec."
    Label2(10).Caption = "Rec. Count."
    btnNew5.Caption = "New"
    btnModify5.Caption = "Modify"
    btnSave5.Caption = "Save"
    BtnUndo5.Caption = "Undo"
    btnDelete5.Caption = "Delete"
    btnCancel5.Caption = "Exit"
End Sub
Private Sub Form_Load()

    On Error GoTo ErrTrap
    
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    If SendForm = 0 Then
        '«·ÊÕœ« 
        frame6(2).left = 10290
    End If
    Dim Dcombos As ClsDataCombos
    If SendForm = 0 Then
        Dim cGrdBack As New ClsBackGroundPic
        Set Me.Grid.WallPaper = cGrdBack.Picture
        Dim i      As Integer
        Dim My_SQL As String
    
        ScreenNameArabic = "  «·ÊÕœ«  «·„” Œœ„… ›Ï «·»—‰«„Ã "
        ScreenNameEnglish = " Units Data  "
        RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
        Set RsSavRec = New ADODB.Recordset
        If mIsQrCode Then
            My_SQL = "Select * from TblUnites Where IsNull(QRCODE,'') <> '' "
            
            RsSavRec.CursorLocation = adUseClient
            RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic

        Else
            My_SQL = "TblUnites"
            RsSavRec.CursorLocation = adUseClient
            RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect

        End If
        Me.TxtModFlg.text = "R"
        Resize_Form Me
        FillGridWithData
        With Me.Grid
            For i = 0 To .Cols - 1
                .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
            Next
            .ExtendLastCol = True
        End With
        BtnFirst_Click
        '#####################################################################################################################################################
    ElseIf SendForm = 1 Then
        ScreenNameArabic = "‘«‘… «·Ê«‰ «·«’‰«›  "
        ScreenNameEnglish = "  Items Color Data "
        RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
    
        Resize_Form Me
        Set rs = New ADODB.Recordset
        rs.Open "TblItemsColors", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        Me.TxtModFlg1.text = "R"
        Retrive
        '#####################################################################################################################################################
    ElseIf SendForm = 2 Then
        ScreenNameArabic = "»Ì«‰«  „ﬁ«”«  «·√’‰«›   "
        ScreenNameEnglish = "  Items Size"
        RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"

        Set Cmd2(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
        Set Cmd2(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
        Set Cmd2(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
        Set Cmd2(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
        Set Cmd2(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
        Set Cmd2(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
        Resize_Form Me
        Set rs2 = New ADODB.Recordset
        rs2.Open "TblItemsSizes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        Me.TxtModFlg2.text = "R"
        Retrive2
        '#####################################################################################################################################################
    ElseIf SendForm = 3 Then
        ScreenNameArabic = " «‰Ê«⁄ ›—“ «·«’‰«›  "
        ScreenNameEnglish = " Items Class "
        Set Dcombos = New ClsDataCombos
        Dcombos.GetStores Me.DCboStoreName
        RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"

        Set Cmd3(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
        Set Cmd3(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
        Set Cmd3(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
        Set Cmd3(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
        Set Cmd3(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
        Set Cmd3(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
        Resize_Form Me
        Set Rs3 = New ADODB.Recordset
        Rs3.Open "TblstoresLocations", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        Me.TxtModFlg3.text = "R"
        Retrive3
        '#####################################################################################################################################################
    ElseIf SendForm = 4 Then
        My_SQL = "TblSpecification"
        Set BKGrndPic4 = New ClsBackGroundPic
        Set RsSavRec4 = New ADODB.Recordset
        RsSavRec4.CursorLocation = adUseClient
        RsSavRec4.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        Me.TxtModFlg4.text = "R"
        Resize_Form Me
        My_SQL = "select UserID,UserName From tblUsers "
        fill_combo DCUser, My_SQL

        FillGrid4WithData

        With Me.Grid4
            .Cell(flexcpPicture, 0, .ColIndex("name")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
            .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon
            For i = 0 To .Cols - 1
                .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
            Next
            .ExtendLastCol = True
            .WallPaper = BKGrndPic4.Picture
            .RowHeight(-1) = 300
        End With
        btnFirst4_Click
        '######################################################################################################################################################
    ElseIf SendForm = 5 Then
        ScreenNameArabic = " ⁄‰«’— «· ﬂ«·Ì› «·’‰«⁄ÌÂ "
        ScreenNameEnglish = "  Production Cost Elements "
        RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
        My_SQL = "TBLProductionElements"
        Set BKGrndPic5 = New ClsBackGroundPic
        Set RsSavRec5 = New ADODB.Recordset
        RsSavRec5.CursorLocation = adUseClient
        RsSavRec5.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        Me.TxtModFlg5.text = "R"
        Resize_Form Me
        My_SQL = "select UserID,UserName From tblUsers "
        fill_combo DCUser5, My_SQL
        Set Dcombos = New ClsDataCombos
        Dcombos.GetExpensesNames Me.DcboExpensesID
        Set cSearch = New clsDCboSearch
        Set cSearch.Client = Me.DcboExpensesID
        ModFgLib.LinkFgColWithDataCombo Grid5, Grid5.ColIndex("ExpensesID"), Me.DcboExpensesID
        FillGrid5WithData
        With Me.Grid5
            .Cell(flexcpPicture, 0, .ColIndex("Name")) = Me.GrdImageList5.ListImages("Vac_Name").ExtractIcon
            .Cell(flexcpPicture, 0, .ColIndex("NameE")) = Me.GrdImageList5.ListImages("Vac_Name").ExtractIcon
            .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList5.ListImages("Ser").ExtractIcon
            For i = 0 To .Cols - 1
                .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
            Next
            .ExtendLastCol = True
            .WallPaper = BKGrndPic5.Picture
            .RowHeight(-1) = 300
        End With
        btnFirst5_Click
        
        '######################################################################################################################################################
    End If
    C1Tab1.TabVisible(SendForm) = True
    C1Tab1.CurrTab = SendForm
  
ErrTrap:
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim IntResult As String
    Dim StrMSG As String
    
    On Error GoTo ErrTrap

    If SendForm = 0 Then
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
        End If
    ElseIf SendForm = 1 Then
        If Me.TxtModFlg1.text <> "R" Then
            Select Case Me.TxtModFlg1.text
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
        End If
    ElseIf SendForm = 2 Then
        If Me.TxtModFlg2.text <> "R" Then
            Select Case Me.TxtModFlg2.text
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
        End If
    ElseIf SendForm = 3 Then
        If Me.TxtModFlg3.text <> "R" Then
            Select Case Me.TxtModFlg3.text
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
        End If
    ElseIf SendForm = 4 Then
        If Me.TxtModFlg4.text <> "R" Then
            Select Case Me.TxtModFlg4.text
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
        End If
    ElseIf SendForm = 5 Then
        If Me.TxtModFlg5.text <> "R" Then
            Select Case Me.TxtModFlg5.text
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
        End If
    'ElseIf SendForm = 6 Then
    '    If Me.TxtModFlg6.Text <> "R" Then
    '        Select Case Me.TxtModFlg6.Text
    '            Case "N"
    '                If SystemOptions.UserInterface = EnglishInterface Then
    '                    StrMSG = "You will close this screen before save " & Chr(13)
    '                    StrMSG = StrMSG & " the new data  " & Chr(13)
    '                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
    '                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
    '                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
    '                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    '                Else
    '                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
    '                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
    '                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
    '                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
    '                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
    '                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
    '                End If
    '            Case "E"
    '                If SystemOptions.UserInterface = EnglishInterface Then
    '                    StrMSG = "You will close this screen before save  " & Chr(13)
    '                    StrMSG = StrMSG & " the Modifications  " & Chr(13)
    '                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
    '                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
    '                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
    '                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    '                Else
    '                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
    '                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
    '                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
    '                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
    '                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
    '                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
    '                End If
    '        End Select
    '    End If
    'ElseIf SendForm = 7 Then
    '    If Me.TxtModFlg7.Text <> "R" Then
    '        Select Case Me.TxtModFlg7.Text
    '            Case "N"
    '                If SystemOptions.UserInterface = EnglishInterface Then
    '                    StrMSG = "You will close this screen before save " & Chr(13)
    '                    StrMSG = StrMSG & " the new data  " & Chr(13)
    '                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
    '                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
    '                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
    '                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    '                Else
    '                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
    '                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
    '                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
    '                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
    '                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
    '                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
    '                End If
    '            Case "E"
    '                If SystemOptions.UserInterface = EnglishInterface Then
    '                    StrMSG = "You will close this screen before save  " & Chr(13)
    '                    StrMSG = StrMSG & " the Modifications  " & Chr(13)
    '                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
    '                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
    '                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
    '                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    '                Else
    '                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
    '                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
    '                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
    '                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
    '                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
    '                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
    '                End If
    '        End Select
    '    End If
    End If
    
    If StrMSG <> "" Then
        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)
        Select Case IntResult
            Case vbYes
                Cancel = True
                Select Case SendForm
                    Case 0
                        btnSave_Click
                    Case 1
                        SaveData
                    Case 2
                        SaveData2
                    Case 3
                        SaveData3
                    Case 4
                        btnSave4_Click
                    Case 5
                        btnSave5_Click
                    Case 6
                        'btnSave6_Click
                    Case 7
                        'btnSave7_Click
                End Select
            Case vbCancel
                Cancel = True
        End Select
    End If
    
    
    Exit Sub
ErrTrap:
End Sub
Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo ErrTrap
    
'#####################################################################################################################################################
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish
    If RsSavRec.State = adStateOpen Then
        If Not (RsSavRec.EOF Or RsSavRec.BOF) Then
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
        End If
        RsSavRec.Close
        Set RsSavRec = Nothing
    End If
'#####################################################################################################################################################
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
'#####################################################################################################################################################
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish
    If rs2.State = adStateOpen Then
        If Not (rs2.EOF Or rs2.BOF) Then
            If rs2.EditMode <> adEditNone Then
                rs2.CancelUpdate
            End If
        End If
        rs2.Close
    End If
    Set rs2 = Nothing
    Set TTP2 = Nothing
'#####################################################################################################################################################
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish
    If Rs3.State = adStateOpen Then
        If Not (Rs3.EOF Or Rs3.BOF) Then
            If Rs3.EditMode <> adEditNone Then
                Rs3.CancelUpdate
            End If
        End If
        Rs3.Close
    End If
    Set Rs3 = Nothing
    Set TTP3 = Nothing
'#####################################################################################################################################################
    If RsSavRec4.State = adStateOpen Then
        If Not (RsSavRec4.EOF Or RsSavRec4.BOF) Then
            If RsSavRec4.EditMode <> adEditNone Then
                RsSavRec4.CancelUpdate
            End If
        End If
        RsSavRec4.Close
        Set RsSavRec4 = Nothing
    End If
'#####################################################################################################################################################
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish
    If RsSavRec5.State = adStateOpen Then
        If Not (RsSavRec5.EOF Or RsSavRec5.BOF) Then
            If RsSavRec5.EditMode <> adEditNone Then
                RsSavRec5.CancelUpdate
            End If
        End If
        RsSavRec5.Close
        Set RsSavRec5 = Nothing
    End If
    Set cSearch = Nothing
'######################################################################################################################################################
Exit Sub
ErrTrap:
End Sub
'#####################################################################################################################################################
'#####################################################################################################################################################
'#####################################################################################################################################################
Private Sub btnQuery_Click()
    Load FrmSearchUnit
    FrmSearchUnit.show
End Sub
Function print_report(Optional NoteSerial As String, Optional X As Integer)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
   
    MySQL = "  SELECT     UnitID, UnitName, UnitNamee"
    MySQL = MySQL & " From dbo.TblUnites"

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repUnit.rpt"
    Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repUnit.rpt"
    End If
    
    If Dir(StrFileName) = "" Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "There's no data to show"
        End If
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    
    Dim total As String
    Dim dif As String
    Dim totl As Double

    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Private Sub BtnCancel_Click()
    Unload Me
End Sub
Private Sub btnDelete_Click()

    On Error GoTo ErrTrap

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    
    If TxtUnitID.text <> "" Then
        If UnitsHaveTransactions(val(TxtUnitID.text)) = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = " ·« Ì„ﬂ‰ Õ–› Â–… «·ÊÕœ… ·ÊÃÊœ ⁄„·Ì«  „— »ÿÂ »Â« "
            Else
                Msg = " Can't Modify Unit - Unit Have Transaction "
            End If

            MsgBox Msg, vbCritical
            Exit Sub
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
            MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbYesNo + vbMsgBoxRight, App.Title)
        Else
            MSGType = MsgBox("Delete This Record", vbYesNo + vbMsgBoxRight, App.Title)
        End If
        If MSGType = vbYes Then
            RsSavRec.Find "UnitID=" & val(TxtUnitID.text), , adSearchForward, 1
            If Not (RsSavRec.BOF Or RsSavRec.EOF) Then
                CuurentLogdata ("D")
                RsSavRec.delete
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbOKOnly + vbMsgBoxRight, App.Title
                Else
                    MsgBox "Delete Success...", vbOKOnly + vbMsgBoxRight, App.Title
                End If
                FillGridWithData
                BtnNext_Click
            End If
        End If
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217873, -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
                StrMSG = "Sorry .. can't Delete this record , Reason : Data integrity"
            End If
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.Title
            Cn.Errors.Clear
    End Select
End Sub
Private Sub BtnFirst_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(Me.TxtUnitID.text)
        Me.TxtModFlg.text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveFirst
    FiLLTXT
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry.. this record Already Deleted" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnLast_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(Me.TxtUnitID.text)
        Me.TxtModFlg.text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveLast
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry.. this record Already Deleted" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If TxtUnitID.text <> "" Then
        TxtModFlg = "E"
        Frm2.Enabled = True
        Me.TxtVacName.SetFocus
        CuurentLogdata
    End If

    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«" & CHR(13)
                Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
                Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
                Msg = "Sorry" & CHR(13)
                Msg = Msg & " Can't Edit this record now - Another user work with it now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
    End Select
End Sub
Private Sub btnNew_Click()

    On Error GoTo ErrTrap
    
    Dim My_SQL As String

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    clear_all Me
    TxtModFlg.text = "N"
    My_SQL = "TblUnites"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtUnitID.text = rs.RecordCount + 1
    Else
        TxtUnitID.text = 1
    End If
    rs.Close
    CmbType.ListIndex = 0
    TxtVacName.SetFocus
ErrTrap:
End Sub
Private Sub BtnNext_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(Me.TxtUnitID.text)
        Me.TxtModFlg.text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    If RsSavRec.EOF Then
        RsSavRec.MoveLast
    Else
        RsSavRec.MoveNext
        If RsSavRec.EOF Then
            RsSavRec.MoveLast
        End If
    End If
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnPrevious_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    
    If Me.TxtModFlg.text = "N" Then
        FindRec val(Me.TxtUnitID.text)
        Me.TxtModFlg.text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MovePrevious
    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnSave_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control

    If Trim(Me.TxtVacName.text) = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ﬂ «»… «”„ «·ÊÕœ… ...!!!"
        Else
            Msg = "Please Enter The name"
        End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtVacName.SetFocus
        Exit Sub
    End If
    StrVacName = IsRecExist("TblUnites", "UnitName", Trim(TxtVacName.text), "UnitName", "UnitID<>'" & Trim(TxtUnitID.text) & "'")

    If StrVacName <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–Â «·ÊÕœ… „‰ ﬁ»·"
        Else
            Msg = "this Unit Already Exist"
        End If

        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.Title
        TxtVacName.SetFocus
        Exit Sub
    End If
    Select Case Me.TxtModFlg.text
        Case "N"
            AddNewRec
            BtnLast_Click

        Case "E"
            FiLLRec
    End Select
    Exit Sub
ErrTrap:
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.Title
    Else
        MsgBox "Error in Enterd data", vbOKOnly + vbMsgBoxRight, App.Title
    End If
End Sub
Private Sub BtnUndo_Click()
    FindRec val(TxtUnitID.text)
    Me.TxtModFlg.text = "R"
End Sub
Private Sub BtnUpdate_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim FristCount As Long
    Dim LastCount As Long
    
    FristCount = RsSavRec.RecordCount
    RsSavRec.Requery
    LastCount = RsSavRec.RecordCount
    BtnUndo_Click
    If SystemOptions.UserInterface = ArabicInterface Then
        If FristCount = LastCount Then
            Msg = "·«  ÊÃœ »Ì«‰«  ÃœÌœ…"
        Else
            Msg = "⁄œœ «·”Ã·«  ﬁ»· «· ÕœÌÀ" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
            If LastCount > FristCount Then
                Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·ÃœÌœ…" & vbCrLf & LastCount - FristCount
            Else
                Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·„Õ–Ê›…" & vbCrLf & FristCount - LastCount
            End If
        End If
    Else
    If FristCount = LastCount Then
            Msg = "No new data"
        Else
            Msg = "No. of records before the update" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
        
            If LastCount > FristCount Then
                Msg = Msg + vbCrLf & "No. of new records" & vbCrLf & LastCount - FristCount
            Else
                Msg = Msg + vbCrLf & "No. of deleted records" & vbCrLf & FristCount - LastCount
            End If
        End If
    End If
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.Title
ErrTrap:
End Sub
Public Sub AddNewRec()

    On Error GoTo ErrTrap
    
    Dim StrRecID As String
    
    StrRecID = new_id("TblUnites", "UnitID", "")
    RsSavRec.AddNew
    RsSavRec.Fields("UnitID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
Public Sub FiLLRec()

    On Error GoTo ErrTrap

    RsSavRec.Fields("UnitName").value = IIf(TxtVacName.text <> "", Trim(TxtVacName.text), Null)
    RsSavRec.Fields("UnitNamee").value = IIf(TxtVacNamee.text <> "", Trim(TxtVacNamee.text), Null)
    RsSavRec("QRCODE").value = Trim(txtQRCODE.text)
    RsSavRec.update
    CuurentLogdata
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbOKOnly + vbMsgBoxRight, App.Title
    Else
        MsgBox "Saved Successfully", vbOKOnly + vbMsgBoxRight, App.Title
    End If
    FillGridWithData
    TxtModFlg = "R"
    Exit Sub
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If
End Sub
Public Sub FiLLTXT()

    On Error GoTo ErrTrap
    
    Dim i As Integer
    
    Frm2.Enabled = False
    TxtUnitID.text = IIf(IsNull(RsSavRec.Fields("UnitID").value), "", RsSavRec.Fields("UnitID").value)
    TxtVacName.text = IIf(IsNull(RsSavRec.Fields("UnitName").value), "", RsSavRec.Fields("UnitName").value)
    TxtVacNamee.text = IIf(IsNull(RsSavRec.Fields("UnitNamee").value), "", RsSavRec.Fields("UnitNamee").value)
    txtQRCODE.text = IIf(IsNull(RsSavRec.Fields("QRCODE").value), "", RsSavRec.Fields("QRCODE").value)
    
    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount
    With Grid
        For i = 1 To .rows - 1
            If Trim(TxtUnitID.text) = .TextMatrix(i, .ColIndex("UnitID")) Then
                .row = i
                Exit Sub
            End If
        Next
    End With
ErrTrap:
End Sub
Public Sub EditRec(StrTable As String, RecId As String)
    FiLLRec
End Sub
Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid.TextMatrix(Me.Grid.row, Me.Grid.ColIndex("UnitID")))
ErrTrap:
End Sub
Private Sub TxtDis_Count_KeyPress(KeyAscii As Integer)
    KeyAscii = DataFormat(CurOnly, KeyAscii)
End Sub
Private Sub ISButton1_Click()
    print_report
End Sub

Private Sub TxtUnitID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.text
    TxtModFlg.text = ""
    TxtModFlg = TxtMod
End Sub
Public Function FindRec(ByVal RecId As Long)

    On Error GoTo ErrTrap
    
    RsSavRec.Find "UnitID=" & RecId, , adSearchForward, 1

    If Not (RsSavRec.EOF) Then
        FiLLTXT
    End If
    Exit Function
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        BtnUndo_Click
    End If
End Function
Private Sub TxtModFlg_Change()
    If TxtModFlg.text = "N" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
    ElseIf TxtModFlg.text = "R" Then
        Frm2.Enabled = False
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False
        If TxtUnitID.text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
        End If
        BtnUpdate.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
    ElseIf TxtModFlg.text = "E" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        Grid.Enabled = False
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    End If

End Sub
Public Sub FillGridWithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    If mIsQrCode Then
        My_SQL = "select * From TblUnites  Where IsNull(QRCODE ,'') <> '' order by UnitID"
    Else
        My_SQL = "select * From TblUnites order by UnitID"
    End If
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With Me.Grid
        .rows = 2
        .Clear flexClearScrollable
        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 1
            rs.MoveFirst
            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(rs.Fields("UnitName").value), "", rs.Fields("UnitName").value)
                .TextMatrix(i, .ColIndex("UnitNamee")) = IIf(IsNull(rs.Fields("UnitNamee").value), "", rs.Fields("UnitNamee").value)
                .TextMatrix(i, .ColIndex("UnitID")) = IIf(IsNull(rs.Fields("UnitID").value), "", rs.Fields("UnitID").value)
                rs.MoveNext
            Next
            rs.Close
        End If
        .RowHeight(-1) = 300
    End With
ErrTrap:
End Sub
Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·ÊÕœ…   " & TxtUnitID.text & CHR(13) & "  «”„ «·ÊÕœ… " & TxtVacName.text
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Unit No   " & TxtUnitID.text & CHR(13) & " Unit Name" & TxtVacNamee.text
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D"
    End If
End Function
Private Sub TxtVacName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub TxtVacNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub
'#####################################################################################################################################################
'#####################################################################################################################################################
'#####################################################################################################################################################

Private Sub Cmd_Click(index As Integer)

    On Error GoTo ErrTrap

    Select Case index
        Case 0
            TxtModFlg1.text = "N"
            clear_all Me
            XPTxtBoxID.text = CStr(new_id("TblItemsColors", "ColorID", "", True))
            XPTxtBoxName.SetFocus
        Case 1
            TxtModFlg1.text = "E"
            CuurentLogdata1
        Case 2
            SaveData
        Case 3
            Call Undo
        Case 4
            Del_Company
        Case 5
        Case 6
            Unload Me
    End Select
    Exit Sub
ErrTrap:
End Sub
Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Function CuurentLogdata1(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «··Ê‰   " & XPTxtBoxID.text & CHR(13) & "  «”„ «··Ê‰ " & XPTxtBoxName.text & CHR(13) & "  „·«ÕŸ«  " & XPMTxtRemark.text
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Color No   " & XPTxtBoxID.text & CHR(13) & " Color Name" & XPTxtBoxName.text & CHR(13) & "  Remarks " & XPMTxtRemark.text
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg1
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D"
    End If
End Function
Private Sub TxtModFlg1_Change()

    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg1.text
        Case "R"
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
            Me.XPTxtBoxID.locked = True
            Me.XPTxtBoxName.locked = True
            Me.XPMTxtRemark.locked = True
            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If
        Case "N"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.XPTxtBoxID.locked = True
            Me.XPTxtBoxName.locked = False
            Me.XPMTxtRemark.locked = False
        Case "E"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            Me.XPTxtBoxID.locked = True
            Me.XPTxtBoxName.locked = False
            Me.XPMTxtRemark.locked = False
    End Select
    Exit Sub
ErrTrap:
End Sub
Public Sub Retrive(Optional Lngid As Long = 0)

    On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If
    XPTxtBoxID.text = IIf(IsNull(rs("ColorID").value), "", val(rs("ColorID").value))
    XPTxtBoxName.text = IIf(IsNull(rs("ColorName").value), "", Trim(rs("ColorName").value))
    XPMTxtRemark.text = IIf(IsNull(rs("ColorComment").value), "", Trim(rs("ColorComment").value))
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub
Private Sub XPBtnMove_Click(index As Integer)

    On Error GoTo ErrTrap

    If Me.TxtModFlg1.text = "N" Then
        clear_all Me
        Me.TxtModFlg1.text = "R"
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
Private Sub SaveData()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
    
    On Error GoTo ErrTrap
    
    If Me.TxtModFlg1.text <> "R" Then
        If XPTxtBoxName.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "„‰ ›÷·ﬂ √œŒ· «”„ «··Ê‰ ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Else
                MsgBox "Please enter the name", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            End If
            XPTxtBoxName.SetFocus
            Exit Sub
        End If
        Select Case Me.TxtModFlg1.text
            Case "N"
                StrSQL = "select * from  TblItemsColors where ColorName ='" & Trim(XPTxtBoxName.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If RsTemp.RecordCount > 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "Â‰«ﬂ ·Ê‰ „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «··Ê‰"
                    Else
                        Msg = "This record already exists"
                    End If
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    XPTxtBoxName.SetFocus
                    Exit Sub
                End If
            Case "E"
                StrSQL = "select * from  TblItemsColors where ColorName='" & Trim(XPTxtBoxName.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If RsTemp.RecordCount > 0 Then
                    If RsTemp("ColorID").value <> val(XPTxtBoxID.text) Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "Â‰«ﬂ ·Ê‰  „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                            Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                            Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «··Ê‰"
                        Else
                            Msg = "This record already exists"
                        End If
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        XPTxtBoxName.SetFocus
                        Exit Sub
                    End If
                End If
        End Select
        Cn.BeginTrans
        BeginTrans = True
        Select Case Me.TxtModFlg1.text
            Case "N"
                rs.AddNew
                rs("ColorID").value = val(XPTxtBoxID.text)
            Case "E"
                If rs("ColorID").value <> val(Me.XPTxtBoxID.text) Then
                    rs.Find "ColorID=" & val(Me.XPTxtBoxID.text), , adSearchForward, 1

                    If rs.EOF Or rs.EOF Then
                        Exit Sub
                    End If
                End If
        End Select
        rs("ColorName").value = Trim(XPTxtBoxName.text)
        rs("ColorComment").value = IIf(XPMTxtRemark.text = "", Null, Trim(XPMTxtRemark.text))
        
        rs.update
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        CuurentLogdata1
        Select Case Me.TxtModFlg1.text
            Case "N"
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «··Ê‰" & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Recored saved successfully , do you want to add another recored"
                End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            Case "E"
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Else
                    MsgBox "Record edited successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                End If
        End Select
        TxtModFlg1.text = "R"
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & CHR(13)
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub
Private Sub Undo()

    On Error GoTo ErrTrap

    Select Case TxtModFlg1.text
        Case "N"
            clear_all Me
            Me.TxtModFlg1.text = "R"
            XPBtnMove_Click (1)
        Case "E"
            rs.Find "BoxID='" & val(XPTxtBoxID.text) & "'", , adSearchForward, adBookmarkFirst
            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg1.text = "R"
                Exit Sub
            End If
            Retrive
            Me.TxtModFlg1.text = "R"
    End Select
    Exit Sub
ErrTrap:
End Sub
Private Sub Del_Company()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    On Error GoTo ErrTrap
    If XPTxtBoxID.text <> "" Then
        If val(Me.XPTxtBoxID.text) = 1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã·...!!!"
            Else
                Msg = "sorry, this record cannot be deleted "
            End If
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If
        StrSQL = "select * from Transaction_Details where ColorID=" & Trim(XPTxtBoxID.text)
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «··Ê‰" & CHR(13)
                Msg = Msg + "Â‰«ﬂ »⁄÷ «·⁄„·Ì«  „— »ÿ… »Â–« «··Ê‰"
            Else
                Msg = "sorry, this record cannot be deleted due to data integration"
            End If
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ì „ Õ–› »Ì«‰«  «··Ê‰ —ﬁ„ " & CHR(13)
            Msg = Msg + (XPTxtBoxID.text) & CHR(13)
            Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
            Msg = "Are you sure you want to delete this record"
        End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                CuurentLogdata1 ("D")
                rs.delete
                rs.MoveFirst
                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg1_Change
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
            Msg = "This operation is not available due to lack of records"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg1_Change
        Exit Sub
    End If
    TxtModFlg1_Change
    Exit Sub
ErrTrap:
    If Err.Number = -2147217887 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «··Ê‰ "
            Msg = Msg & CHR(13) & Err.Description
        Else
            Msg = "sorry, this record cannot be deleted due to data integration"
        End If
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
        rs.CancelUpdate
    End If

End Sub
'2
'#####################################################################################################################################################
'#####################################################################################################################################################
'#####################################################################################################################################################
Private Sub Cmd2_Click(index As Integer)

    'On Error GoTo ErrTrap

    Select Case index
        Case 0
            TxtModFlg2.text = "N"
            clear_all Me
            XPTxtBoxID2.text = CStr(new_id("TblItemsSizes", "SizeId", "", True))
            XPTxtBoxName2.SetFocus
        Case 1
            TxtModFlg2.text = "E"
            CuurentLogdata2
        Case 2
            SaveData2
        Case 3
            Call Undo2
        Case 4
            Del_Company2
        Case 5
        Case 6
            Unload Me
    End Select
    Exit Sub
ErrTrap:
End Sub
Private Sub CmdHelp2_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub
Function CuurentLogdata2(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·„ﬁ«”   " & XPTxtBoxID2.text & CHR(13) & "  «”„ «·„ﬁ«” " & XPTxtBoxName2.text & CHR(13) & "  „·«ÕŸ«  " & XPMTxtRemark2.text
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Size No   " & XPTxtBoxID2.text & CHR(13) & " Size Name" & XPTxtBoxName2.text & CHR(13) & "  Remarks " & XPMTxtRemark2.text
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg2
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D"
    End If
End Function

Private Sub TxtModFlg2_Change()

    On Error GoTo ErrTrap
    
    Select Case Me.TxtModFlg2.text
        Case "R"
            Me.Cmd2(2).Enabled = False
            Me.Cmd2(3).Enabled = False
            Me.Cmd2(0).Enabled = True
            Me.Cmd2(1).Enabled = True
            Me.Cmd2(4).Enabled = True
            Me.XPBtnMove2(0).Enabled = True
            Me.XPBtnMove2(1).Enabled = True
            Me.XPBtnMove2(2).Enabled = True
            Me.XPBtnMove2(3).Enabled = True
            Me.XPTxtBoxID2.locked = True
            Me.XPTxtBoxName2.locked = True
            Me.XPMTxtRemark2.locked = True
            If rs2.RecordCount < 1 Then
                Me.XPBtnMove2(0).Enabled = False
                Me.XPBtnMove2(1).Enabled = False
                Me.XPBtnMove2(2).Enabled = False
                Me.XPBtnMove2(3).Enabled = False
                Me.Cmd2(1).Enabled = False
                Me.Cmd2(4).Enabled = False
            End If
        Case "N"
            Me.Cmd2(2).Enabled = True
            Me.Cmd2(3).Enabled = True
            Me.Cmd2(0).Enabled = False
            Me.Cmd2(1).Enabled = False
            Me.Cmd2(4).Enabled = False
            Me.XPTxtBoxID2.locked = True
            Me.XPTxtBoxName2.locked = False
            Me.XPMTxtRemark2.locked = False
        Case "E"
            Me.Cmd2(2).Enabled = True
            Me.Cmd2(3).Enabled = True
            Me.Cmd2(0).Enabled = False
            Me.Cmd2(1).Enabled = False
            Me.Cmd2(4).Enabled = False
            Me.XPBtnMove2(0).Enabled = False
            Me.XPBtnMove2(1).Enabled = False
            Me.XPBtnMove2(2).Enabled = False
            Me.XPBtnMove2(3).Enabled = False
            Me.XPTxtBoxID2.locked = True
            Me.XPTxtBoxName2.locked = False
            Me.XPMTxtRemark2.locked = False
    End Select

    Exit Sub
ErrTrap:
End Sub
Public Sub Retrive2(Optional Lngid As Long = 0)

    On Error GoTo ErrTrap

    If rs2.RecordCount < 1 Then
        XPTxtCurrent2.Caption = 0
        XPTxtCount2.Caption = 0
        Exit Sub
    End If
    XPTxtBoxID2.text = IIf(IsNull(rs2("SizeId").value), "", val(rs2("SizeId").value))
    XPTxtBoxName2.text = IIf(IsNull(rs2("SizeName").value), "", Trim(rs2("SizeName").value))
    XPMTxtRemark2.text = IIf(IsNull(rs2("SizeComment").value), "", Trim(rs2("SizeComment").value))
    XPTxtCurrent2.Caption = rs2.AbsolutePosition
    XPTxtCount2.Caption = rs2.RecordCount
    Exit Sub
ErrTrap:
End Sub
Private Sub XPBtnMove2_Click(index As Integer)

    On Error GoTo ErrTrap

    If Me.TxtModFlg2.text = "N" Then
        clear_all Me
        Me.TxtModFlg2.text = "R"
        XPBtnMove_Click (1)
    End If
    Select Case index
        Case 0
            If Not (rs2.EOF Or rs2.BOF) Then
                rs2.MovePrevious

                If rs2.BOF Then rs2.MoveFirst
            End If
        Case 1
            If Not (rs2.EOF Or rs2.BOF) Then
                rs2.MoveFirst
            End If
        Case 2
            If Not (rs2.EOF Or rs2.BOF) Then
                rs2.MoveLast
            End If
        Case 3
            If Not (rs2.EOF Or rs2.BOF) Then
                rs2.MoveNext
                If rs2.EOF Then rs2.MoveLast
            End If
    End Select
    Retrive2
    Exit Sub
ErrTrap:
End Sub
Private Sub SaveData2()

    Dim Msg As String
    Dim Strs2QL As String
    Dim rs2Temp As New ADODB.Recordset
    Dim rs2TempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
    
    On Error GoTo ErrTrap

    If Me.TxtModFlg2.text <> "R" Then
        If XPTxtBoxName2.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "„‰ ›÷·ﬂ √œŒ· «”„ «·„ﬁ«” ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Else
                MsgBox "Please enter the name", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            End If
            XPTxtBoxName2.SetFocus
            Exit Sub
        End If
        Select Case Me.TxtModFlg2.text
            Case "N"
                Strs2QL = "select * from  TblItemsSizes where SizeName ='" & Trim(XPTxtBoxName2.text) & "'"
                rs2Temp.Open Strs2QL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If rs2Temp.RecordCount > 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "Â‰«ﬂ „ﬁ«” „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„ﬁ«”"
                    Else
                        Msg = "This record already exists"
                    End If
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    XPTxtBoxName2.SetFocus
                    Exit Sub
                End If
                rs2Temp.Close
                Strs2QL = "select * from  TblItemsSizes where SizeID=" & val(XPTxtBoxID2.text)
                rs2Temp.Open Strs2QL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If rs2Temp.RecordCount > 0 Then
                    If rs2Temp("SizeId").value <> val(XPTxtBoxID2.text) Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "Â‰«ﬂ „ﬁ«”  „”Ã· „”»ﬁ« »Â–« «·—ﬁ„" & CHR(13)
                            Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·—ﬁ„ «·’ÕÌÕ " & CHR(13)
                            Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ —ﬁ„ «·„ﬁ«”"
                        Else
                            Msg = "This record already exists"
                        End If
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        XPTxtBoxID2.SetFocus
                        Exit Sub
                    End If
                End If
            Case "E"
                Strs2QL = "select * from  TblItemsSizes where SizeName='" & Trim(XPTxtBoxName2.text) & "'"
                rs2Temp.Open Strs2QL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If rs2Temp.RecordCount > 0 Then
                    If rs2Temp("SizeId").value <> val(XPTxtBoxID2.text) Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "Â‰«ﬂ „ﬁ«”  „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                            Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                            Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„ﬁ«”"
                        Else
                            Msg = "This record already exists"
                        End If
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        XPTxtBoxName2.SetFocus
                        Exit Sub
                    End If
                End If
                
                rs2Temp.Close
                
                Strs2QL = "select * from  TblItemsSizes where SizeID=" & val(XPTxtBoxID2.text)
                rs2Temp.Open Strs2QL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If rs2Temp.RecordCount > 0 Then
                    If rs2Temp("SizeId").value <> val(XPTxtBoxID2.text) Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "Â‰«ﬂ „ﬁ«”  „”Ã· „”»ﬁ« »Â–« «·—ﬁ„" & CHR(13)
                            Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·—ﬁ„ «·’ÕÌÕ " & CHR(13)
                            Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ —ﬁ„ «·„ﬁ«”"
                        Else
                            Msg = "This record already exists"
                        End If
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        XPTxtBoxID2.SetFocus
                        Exit Sub
                    End If
                End If
        End Select
        Cn.BeginTrans
        BeginTrans = True
        Select Case Me.TxtModFlg2.text
            Case "N"
                rs2.AddNew
                rs2("SizeId").value = val(XPTxtBoxID2.text)
            Case "E"
                If rs2("SizeId").value <> val(Me.XPTxtBoxID2.text) Then
                    rs2.Find "SizeId=" & val(Me.XPTxtBoxID2.text), , adSearchForward, 1
                    If rs2.EOF Or rs2.EOF Then
                        Exit Sub
                    End If
                End If
        End Select
        rs2("SizeName").value = Trim(XPTxtBoxName2.text)
        rs2("SizeComment").value = IIf(XPMTxtRemark2.text = "", Null, Trim(XPMTxtRemark2.text))
        rs2.update
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent2.Caption = rs2.AbsolutePosition
        XPTxtCount2.Caption = rs2.RecordCount
        CuurentLogdata2
        Select Case Me.TxtModFlg2.text
            Case "N"
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·„ﬁ«”" & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Recored saved successfully , do you want to add another recored"
                End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            Case "E"
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Else
                    MsgBox "Record Edited successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                End If
        End Select
        TxtModFlg2.text = "R"
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & CHR(13)
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub
Private Sub Undo2()

    On Error GoTo ErrTrap

    Select Case TxtModFlg2.text
        Case "N"
            clear_all Me
            Me.TxtModFlg2.text = "R"
            XPBtnMove2_Click (1)
        Case "E"
            rs2.Find "SizeId='" & val(XPTxtBoxID2.text) & "'", , adSearchForward, adBookmarkFirst
            If rs2.EOF Or rs2.BOF Then
                Me.TxtModFlg2.text = "R"
                Exit Sub
            End If
            Retrive2
            Me.TxtModFlg2.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub
Private Sub Del_Company2()

    Dim Msg As String
    Dim Strs2QL As String
    Dim rs2Temp As New ADODB.Recordset
    
    On Error GoTo ErrTrap

    If XPTxtBoxID2.text <> "" Then
        If val(Me.XPTxtBoxID2.text) = 1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã·...!!!"
            Else
                Msg = "sorry, this record cannot be deleted "
            End If
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If
        Strs2QL = "select * from Transaction_Details where ItemSize=" & Trim(XPTxtBoxID2.text)
        rs2Temp.Open Strs2QL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If Not (rs2Temp.EOF Or rs2Temp.BOF) Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «··Ê‰" & CHR(13)
                Msg = Msg + "Â‰«ﬂ »⁄÷ «·⁄„·Ì«  „— »ÿ… »Â–« «··Ê‰"
            Else
                Msg = "sorry, this record cannot be deleted due to data integration"
            End If
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ì „ Õ–› »Ì«‰«  «··Ê‰ —ﬁ„ " & CHR(13)
            Msg = Msg + (XPTxtBoxID2.text) & CHR(13)
            Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
            Msg = "Are you sure you want to delete this record"
        End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs2.RecordCount < 1 Then
                rs2.delete
                CuurentLogdata2 ("D")
                rs2.MoveFirst
                If rs2.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg2_Change
                    XPTxtCurrent2.Caption = 0
                    XPTxtCount2.Caption = 0
                Else
                    Retrive2
                End If
            End If
        End If
    Else
        clear_all Me
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
            Msg = "This operation is not available due to lack of records"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg2_Change
        Exit Sub
    End If
    TxtModFlg2_Change
    Exit Sub
ErrTrap:
    If Err.Number = -2147217887 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «··Ê‰ "
            Msg = Msg & CHR(13) & Err.Description
        Else
            Msg = "sorry, this record cannot be deleted due to data integration"
        End If
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
        rs2.CancelUpdate
    End If
End Sub
'3
'#####################################################################################################################################################
'#####################################################################################################################################################
'#####################################################################################################################################################
Private Sub Cmd3_Click(index As Integer)

    On Error GoTo ErrTrap

    Select Case index
        Case 0
            TxtModFlg3.text = "N"
            clear_all Me
            XPTxtBoxName3.SetFocus
        Case 1
            TxtModFlg3.text = "E"
            CuurentLogdata3
        Case 2
            SaveData3
        Case 3
            Call Undo3
        Case 4
            Del_Company3
        Case 5
        Case 6
            Unload Me
    End Select
    Exit Sub
ErrTrap:
End Sub
Function CuurentLogdata3(Optional Currentmode As String)

    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·›—“   " & XPTxtBoxID3.text & CHR(13) & "  «”„ «·›—“ " & XPTxtBoxName3.text & CHR(13) & "  „·«ÕŸ«  " & XPMTxtRemark3.text
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Class No   " & XPTxtBoxID3.text & CHR(13) & " Class Name" & XPTxtBoxName3.text & CHR(13) & "  Remarks " & XPMTxtRemark3.text
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg3
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D"
    End If
End Function

Private Sub TxtModFlg3_Change()

    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg3.text
        Case "R"
            Me.Cmd3(2).Enabled = False
            Me.Cmd3(3).Enabled = False
            Me.Cmd3(0).Enabled = True
            Me.Cmd3(1).Enabled = True
            Me.Cmd3(4).Enabled = True
            Me.XPBtnMove3(0).Enabled = True
            Me.XPBtnMove3(1).Enabled = True
            Me.XPBtnMove3(2).Enabled = True
            Me.XPBtnMove3(3).Enabled = True
            Me.XPTxtBoxID3.locked = True
            Me.XPTxtBoxName3.locked = True
            Me.XPMTxtRemark3.locked = True
            If Rs3.RecordCount < 1 Then
                Me.XPBtnMove3(0).Enabled = False
                Me.XPBtnMove3(1).Enabled = False
                Me.XPBtnMove3(2).Enabled = False
                Me.XPBtnMove3(3).Enabled = False
                Me.Cmd3(1).Enabled = False
                Me.Cmd3(4).Enabled = False
            End If
        Case "N"
            Me.Cmd3(2).Enabled = True
            Me.Cmd3(3).Enabled = True
            Me.Cmd3(0).Enabled = False
            Me.Cmd3(1).Enabled = False
            Me.Cmd3(4).Enabled = False
            Me.XPTxtBoxID3.locked = True
            Me.XPTxtBoxName3.locked = False
            Me.XPMTxtRemark3.locked = False
        Case "E"
            Me.Cmd3(2).Enabled = True
            Me.Cmd3(3).Enabled = True
            Me.Cmd3(0).Enabled = False
            Me.Cmd3(1).Enabled = False
            Me.Cmd3(4).Enabled = False
            Me.XPBtnMove3(0).Enabled = False
            Me.XPBtnMove3(1).Enabled = False
            Me.XPBtnMove3(2).Enabled = False
            Me.XPBtnMove3(3).Enabled = False
            Me.XPTxtBoxID3.locked = True
            Me.XPTxtBoxName3.locked = False
            Me.XPMTxtRemark3.locked = False
    End Select
    Exit Sub
ErrTrap:
End Sub
Public Sub Retrive3(Optional Lngid As Long = 0)

    On Error GoTo ErrTrap

    If Rs3.RecordCount < 1 Then
        XPTxtCurrent3.Caption = 0
        XPTxtCount3.Caption = 0
        Exit Sub
    End If

    XPTxtBoxID3.text = IIf(IsNull(Rs3("Locid").value), "", val(Rs3("Locid").value))
    XPTxtBoxName3.text = IIf(IsNull(Rs3("name").value), "", Trim(Rs3("name").value))
    XPTxtBoxNamee.text = IIf(IsNull(Rs3("namee").value), "", Trim(Rs3("namee").value))
    XPMTxtRemark3.text = IIf(IsNull(Rs3("Comment").value), "", Trim(Rs3("Comment").value))
    Me.DCboStoreName.BoundText = IIf(IsNull(Rs3("StoreID").value), "", Rs3("StoreID").value)
    XPTxtCurrent3.Caption = Rs3.AbsolutePosition
    XPTxtCount3.Caption = Rs3.RecordCount
    Exit Sub
ErrTrap:
End Sub
Private Sub XPBtnMove3_Click(index As Integer)

    On Error GoTo ErrTrap

    If Me.TxtModFlg3.text = "N" Then
        clear_all Me
        Me.TxtModFlg3.text = "R"
        XPBtnMove3_Click (1)
    End If

    Select Case index
        Case 0
            If Not (Rs3.EOF Or Rs3.BOF) Then
                Rs3.MovePrevious
                If Rs3.BOF Then Rs3.MoveFirst
            End If
        Case 1
            If Not (Rs3.EOF Or Rs3.BOF) Then
                Rs3.MoveFirst
            End If
        Case 2
            If Not (Rs3.EOF Or Rs3.BOF) Then
                Rs3.MoveLast
            End If
        Case 3
            If Not (Rs3.EOF Or Rs3.BOF) Then
                Rs3.MoveNext
                If Rs3.EOF Then Rs3.MoveLast
            End If
    End Select
    Retrive3
    Exit Sub
ErrTrap:
End Sub
Private Sub SaveData3()

    Dim Msg As String
    Dim Strs3QL As String
    Dim rs3Temp As New ADODB.Recordset
    Dim rs3TempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
    
    On Error GoTo ErrTrap

    If Me.TxtModFlg3.text <> "R" Then
        If XPTxtBoxName3.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "„‰ ›÷·ﬂ √œŒ· «”„ «·„ﬂ«‰ ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Else
                MsgBox "Please enter the name", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            End If
            XPTxtBoxName3.SetFocus
            Exit Sub
        End If
        Select Case Me.TxtModFlg3.text
            Case "N"
                Strs3QL = "select * from  TblstoresLocations where name ='" & Trim(XPTxtBoxName3.text) & "'"
                rs3Temp.Open Strs3QL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If rs3Temp.RecordCount > 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "Â‰«ﬂ „ﬂ«‰   „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„ﬁ«”"
                    Else
                        Msg = "This record already exists"
                    End If
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    XPTxtBoxName3.SetFocus
                    Exit Sub
                End If
            Case "E"
                Strs3QL = "select * from  TblstoresLocations where name='" & Trim(XPTxtBoxName3.text) & "'"
                rs3Temp.Open Strs3QL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If rs3Temp.RecordCount > 0 Then
                    If rs3Temp("Locid").value <> val(XPTxtBoxID3.text) Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "Â‰«ﬂ „ﬂ«‰  „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
                            Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                            Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„ﬁ«”"
                        Else
                            Msg = "This record already exists"
                        End If
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        XPTxtBoxName3.SetFocus
                        Exit Sub
                    End If
                End If
        End Select
        Cn.BeginTrans
        BeginTrans = True
        Select Case Me.TxtModFlg3.text
            Case "N"
                Rs3.AddNew
                XPTxtBoxID3.text = CStr(new_id("TblstoresLocations", "Locid", "", True))
                Rs3("Locid").value = val(XPTxtBoxID3.text)
            Case "E"
                If Rs3("Locid").value <> val(Me.XPTxtBoxID3.text) Then
                    Rs3.Find "Locid=" & val(Me.XPTxtBoxID3.text), , adSearchForward, 1
                    If Rs3.EOF Or Rs3.EOF Then
                        Exit Sub
                    End If
                End If
        End Select
        Rs3("name").value = Trim(XPTxtBoxName3.text)
        Rs3("namee").value = Trim(XPTxtBoxNamee.text)
        Rs3("Comment").value = IIf(XPMTxtRemark3.text = "", Null, Trim(XPMTxtRemark3.text))
        Rs3("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, val(DCboStoreName.BoundText))
        Rs3.update
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent3.Caption = Rs3.AbsolutePosition
        XPTxtCount3.Caption = Rs3.RecordCount
        CuurentLogdata3
        Select Case Me.TxtModFlg3.text
            Case "N"
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–« «·„ﬁ«”" & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Recored saved successfully , do you want to add another recored"
                End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd3_Click (0)
                    Exit Sub
                End If
            Case "E"
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Else
                    MsgBox "Record saved successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                End If
        End Select
        TxtModFlg3.text = "R"
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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry , somthing went wrong while saving data" & CHR(13)
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub
Private Sub Undo3()

    On Error GoTo ErrTrap

    Select Case TxtModFlg3.text
        Case "N"
            clear_all Me
            Me.TxtModFlg3.text = "R"
            XPBtnMove3_Click (1)
        Case "E"
            Rs3.Find "BoxID='" & val(XPTxtBoxID3.text) & "'", , adSearchForward, adBookmarkFirst
            If Rs3.EOF Or Rs3.BOF Then
                Me.TxtModFlg3.text = "R"
                Exit Sub
            End If
            Retrive3
            Me.TxtModFlg3.text = "R"
    End Select
    Exit Sub
ErrTrap:
End Sub
Private Sub Del_Company3()

    Dim Msg As String
    Dim Strs3QL As String
    Dim rs3Temp As New ADODB.Recordset
    
    On Error GoTo ErrTrap
    
    If XPTxtBoxID3.text <> "" Then
        'If val(Me.XPTxtBoxID3.Text) = 1 Then
            'If SystemOptions.UserInterface = ArabicInterface Then
                'Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã·...!!!"
            'Else
                'Msg = "sorry, this record cannot be deleted"
            'End If
            'MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            'Exit Sub
        'End If
        
        'Strs3QL = "select * from Transaction_Details where Locid=" & Trim(XPTxtBoxID3.Text)
        'rs3Temp.Open Strs3QL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        'If Not (rs3Temp.EOF Or rs3Temp.BOF) Then
        '    Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «··Ê‰" & Chr(13)
        '    Msg = Msg + "Â‰«ﬂ »⁄÷ «·⁄„·Ì«  „— »ÿ… »Â–« «··Ê‰"
        '    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        '    Exit Sub
        'End If
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ì „ Õ–› »Ì«‰«  «··Ê‰ —ﬁ„ " & CHR(13)
            Msg = Msg + (XPTxtBoxID3.text) & CHR(13)
            Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
            Msg = "Are you sure you want to delete this record"
        End If
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not Rs3.RecordCount < 1 Then
                CuurentLogdata3 ("D")
                Rs3.delete
                Rs3.MoveFirst

                If Rs3.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg3_Change
                    XPTxtCurrent3.Caption = 0
                    XPTxtCount3.Caption = 0
                Else
                    Retrive3
                End If
            End If
        End If
    Else
        clear_all Me
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
            Msg = "This operation is not available due to lack of records"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg3_Change
        Exit Sub
    End If
    TxtModFlg3_Change
    Exit Sub
ErrTrap:
    If Err.Number = -2147217887 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «··Ê‰ "
            Msg = Msg & CHR(13) & Err.Description
        Else
            Msg = "sorry, this record cannot be deleted due to data integration"
        End If
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
        Rs3.CancelUpdate
    End If
End Sub
'4
'######################################################################################################################################################
'######################################################################################################################################################
'######################################################################################################################################################

Private Sub btnCancel4_Click()
    Unload Me
End Sub
Private Sub btnDelete4_Click()

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    
    On Error GoTo ErrTrap

    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.Title)
    Else
        MSGType = MsgBox("Do you want to delete this record", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.Title)
    End If
    If MSGType = vbYes Then
        RsSavRec4.Find "id=" & val(TxtVac_ID.text), , adSearchForward, 1
        RsSavRec4.delete
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Else
            MsgBox "Record deleted successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        End If
        FillGrid4WithData
        btnNext4_Click
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217873, -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
                StrMSG = "sorry, this record cannot be deleted due to data integration"
            End If
            RsSavRec4.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.Title
            Cn.Errors.Clear
    End Select
End Sub
Private Sub btnFirst4_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg4.text = "N" Then
        FindRec4 val(TxtVac_ID.text)
        Me.TxtModFlg4.text = "R"
    End If

    TxtModFlg4 = "R"

    If RsSavRec4.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec4.MoveFirst
    FiLLTXT4
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec4.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnLast4_Click()

    On Error GoTo ErrTrap

    Dim Msg As String
    
    If Me.TxtModFlg4.text = "N" Then
        FindRec4 val(TxtVac_ID.text)
        Me.TxtModFlg4.text = "R"
    End If
    TxtModFlg4 = "R"
    If RsSavRec4.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec4.MoveLast
    FiLLTXT4
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec4.Requery
            Resume BegnieWork
    End Select

End Sub
Private Sub btnModify4_Click()

    Dim Msg As String

    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap

    If TxtVac_ID.text <> "" Then
        TxtModFlg4 = "E"
        Frm24.Enabled = True
        Me.TxtVacName4.SetFocus
    End If

    Exit Sub
ErrTrap:

    Select Case Err.Number
        Case -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«" & CHR(13)
                Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
                Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
        
            If RsSavRec4.EditMode <> adEditNone Then
                RsSavRec4.CancelUpdate
            End If
    End Select
End Sub
Private Sub btnNew4_Click()

    Dim My_SQL As String
    Dim rs As ADODB.Recordset

    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm24.Enabled = True
    clear_all Me
    TxtModFlg4.text = "N"

    My_SQL = "TblSpecification"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial.text = rs.RecordCount + 1
    Else
        TxtSerial.text = 1
    End If

    rs.Close
   
    TxtVacName4.SetFocus
ErrTrap:
End Sub
Private Sub btnNext4_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String

    If Me.TxtModFlg4.text = "N" Then
        FindRec4 val(TxtVac_ID.text)
        Me.TxtModFlg4.text = "R"
    End If
    TxtModFlg4 = "R"
    If RsSavRec4.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    If RsSavRec4.EOF Then
        RsSavRec4.MoveLast
    Else
        RsSavRec4.MoveNext
        If RsSavRec4.EOF Then
            RsSavRec4.MoveLast
        End If
    End If
    FiLLTXT4
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec4.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnPrevious4_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    If Me.TxtModFlg4.text = "N" Then
        FindRec4 val(TxtVac_ID.text)
        Me.TxtModFlg4.text = "R"
    End If
    TxtModFlg4 = "R"
    If RsSavRec4.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:
    RsSavRec4.MovePrevious
    If RsSavRec4.BOF Then
        RsSavRec4.MoveFirst
    End If
    FiLLTXT4
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec4.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnSave4_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    For Each CtrlTxt In Me.Controls

        If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
            If CtrlTxt.text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.Title
                CtrlTxt.SetFocus
                Exit Sub
            End If
        End If
    Next
    StrVacName = IsRecExist("TblSpecification", "nameˆ", Trim(TxtVacName4.text), "name", "ID<>'" & Trim(TxtVac_ID.text) & "'")

    If StrVacName <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
        Else
            Msg = "This record already exists"
        End If
        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.Title
        TxtVacName4.SetFocus
    
        Exit Sub

    End If
    Select Case Me.TxtModFlg4.text
        Case "N"
            AddNewRec4
            btnLast4_Click
        Case "E"
            FiLLRec4
    End Select
    Exit Sub
ErrTrap:
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.Title
    Else
        MsgBox "Something went wrong while inserting data", vbOKOnly + vbMsgBoxRight, App.Title
    End If
End Sub
 
Private Sub BtnUndo4_Click()
    FindRec4 val(TxtVac_ID.text)
    Me.TxtModFlg4.text = "R"
End Sub

Private Sub BtnUpdate4_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim FristCount As Long
    Dim LastCount As Long
    
    FristCount = RsSavRec4.RecordCount
    RsSavRec4.Requery
    LastCount = RsSavRec4.RecordCount
    BtnUndo4_Click
    If SystemOptions.UserInterface = ArabicInterface Then
        If FristCount = LastCount Then
            Msg = "·«  ÊÃœ »Ì«‰«  ÃœÌœ…"
        Else
            Msg = "⁄œœ «·”Ã·«  ﬁ»· «· ÕœÌÀ" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
            If LastCount > FristCount Then
                Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·ÃœÌœ…" & vbCrLf & LastCount - FristCount
            Else
                Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·„Õ–Ê›…" & vbCrLf & FristCount - LastCount
            End If
        End If
    Else
        If FristCount = LastCount Then
            Msg = "No new data"
        Else
            Msg = "No. of records before the update" & vbCrLf & FristCount & vbCrLf & "No. of records after update" & vbCrLf & LastCount
        
            If LastCount > FristCount Then
                Msg = Msg + vbCrLf & "No. of new records" & vbCrLf & LastCount - FristCount
            Else
                Msg = Msg + vbCrLf & "No. of deleted records" & vbCrLf & FristCount - LastCount
            End If
        End If
    End If
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.Title
ErrTrap:
End Sub

Public Sub AddNewRec4()

    On Error GoTo ErrTrap
    
    Dim StrRecId4 As String
    
    StrRecId4 = new_id("TblSpecification", "id", "")
    RsSavRec4.AddNew
    RsSavRec4.Fields("id").value = IIf(StrRecId4 <> "", StrRecId4, Null)
    FiLLRec4
ErrTrap:
End Sub
Public Sub FiLLRec4()

    On Error GoTo ErrTrap

    RsSavRec4.Fields("name").value = IIf(TxtVacName4.text <> "", Trim(TxtVacName4.text), Null)
    RsSavRec4.Fields("namee").value = IIf(TxtVacNamee4.text <> "", Trim(TxtVacNamee4.text), Null)
    RsSavRec4.update
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Else
        MsgBox "Record saved successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    End If
    FillGrid4WithData
    TxtModFlg4 = "R"
    Exit Sub
ErrTrap:
    If RsSavRec4.EditMode <> adEditNone Then
        RsSavRec4.CancelUpdate
    End If
End Sub
Public Sub FiLLTXT4()

    On Error GoTo ErrTrap
    
    Dim i As Integer
    
    Frm24.Enabled = False
    TxtVac_ID.text = IIf(IsNull(RsSavRec4.Fields("id").value), "", RsSavRec4.Fields("id").value)
    TxtVacName4.text = IIf(IsNull(RsSavRec4.Fields("name").value), "", RsSavRec4.Fields("name").value)
    TxtVacNamee4.text = IIf(IsNull(RsSavRec4.Fields("namee").value), "", RsSavRec4.Fields("namee").value)
    LabCurrRec4.Caption = RsSavRec4.AbsolutePosition
    LabCountRec4.Caption = RsSavRec4.RecordCount
    With Grid4
        For i = 1 To .rows - 1
            If Trim(TxtVac_ID.text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial.text = .TextMatrix(i, .ColIndex("Ser"))
                .row = i
                Exit Sub
            End If
        Next
    End With
ErrTrap:
End Sub

Public Sub EditRec4(StrTable As String, RecId4 As String)
    FiLLRec4
End Sub
Private Sub Grid4_EnterCell()

    On Error GoTo ErrTrap
    
    FindRec4 val(Me.Grid4.TextMatrix(Me.Grid4.row, Me.Grid4.ColIndex("id")))
ErrTrap:
End Sub
Private Sub TxtVac_ID_Change()

    Dim TxtMod As String
    
    TxtMod = TxtModFlg4.text
    TxtModFlg4.text = ""
    TxtModFlg4 = TxtMod
End Sub
Public Function FindRec4(ByVal RecId4 As Long)

    On Error GoTo ErrTrap
    
    RsSavRec4.Find "id=" & RecId4, , adSearchForward, 1
    If Not (RsSavRec4.EOF) Then
        FiLLTXT4
    End If
    Exit Function
ErrTrap:
    If RsSavRec4.EditMode <> adEditNone Then
        RsSavRec4.CancelUpdate
        BtnUndo4_Click
    End If
End Function
Private Sub TxtModFlg4_Change()
    If TxtModFlg4.text = "N" Then
        Frm24.Enabled = True
        Me.btnNew4.Enabled = False
        btnModify4.Enabled = False
        btnDelete4.Enabled = False
        Me.btnQuery4.Enabled = False
        Grid4.Enabled = False
        BtnUndo4.Enabled = True
        Me.btnSave4.Enabled = True
        BtnUpdate4.Enabled = False
    ElseIf TxtModFlg4.text = "R" Then
        Frm24.Enabled = False
        Grid4.Enabled = True
        btnModify4.Enabled = False
        btnDelete4.Enabled = False
        If TxtVac_ID.text <> "" Then
            btnModify4.Enabled = True
            btnDelete4.Enabled = True
        End If
        BtnUpdate4.Enabled = True
        Me.btnQuery4.Enabled = True
        Me.btnNew4.Enabled = True
        BtnUndo4.Enabled = False
        Me.btnSave4.Enabled = False
        btnNext4.Enabled = True
        btnPrevious4.Enabled = True
        btnFirst4.Enabled = True
        btnLast4.Enabled = True
    ElseIf TxtModFlg4.text = "E" Then
        Frm24.Enabled = True
        Me.btnNew4.Enabled = False
        btnModify4.Enabled = False
        btnDelete4.Enabled = False
        Me.btnQuery4.Enabled = False
        BtnUpdate4.Enabled = False
        BtnUndo4.Enabled = True
        Me.btnSave4.Enabled = True
        Grid4.Enabled = False
        btnNext4.Enabled = False
        btnPrevious4.Enabled = False
        btnFirst4.Enabled = False
        btnLast4.Enabled = False
    End If
End Sub
Public Sub FillGrid4WithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From TblSpecification order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid4
        .rows = 2
        .Clear flexClearScrollable
        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 1
            rs.MoveFirst
            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)
                .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs.Fields("namee").value), "", rs.Fields("namee").value)
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                rs.MoveNext
            Next
            rs.Close
        End If
        .RowHeight(-1) = 300
    End With
ErrTrap:
End Sub
Private Function CheckDelCountry(Lngid As Long) As Boolean

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    
    StrSQL = "Select * From TblEmployee Where id=" & Lngid & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rs.BOF Or rs.EOF) Then
        CheckDelCountry = False
    Else
        CheckDelCountry = True
    End If
    rs.Close
    Set rs = Nothing
End Function
Private Sub TxtVacName4_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub
Private Sub TxtVacNamee4_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub
'5
'######################################################################################################################################################
'######################################################################################################################################################
'######################################################################################################################################################

 

Private Sub btnCancel5_Click()
    Unload Me
End Sub

Private Sub btnDelete5_Click()

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    
    On Error GoTo ErrTrap
    
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    If TxtVac_ID5.text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.Title)
        Else
            MSGType = MsgBox("Do you want to delete this record", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.Title)
        End If
        If MSGType = vbYes Then
            RsSavRec5.Find "TBLProductionElementsId=" & val(TxtVac_ID5.text), , adSearchForward, 1
            CuurentLogdata5 ("D")
            RsSavRec5.delete
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
                MsgBox "Record deleted successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            End If
            FillGrid5WithData
            btnNext5_Click
        End If
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217873, -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
                StrMSG = "sorry, this record cannot be deleted due to data integration"
            End If
            RsSavRec5.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.Title
            Cn.Errors.Clear
    End Select
End Sub
Private Sub btnFirst5_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg5.text = "N" Then
        FindRec5 val(TxtVac_ID5.text)
        Me.TxtModFlg5.text = "R"
    End If
    TxtModFlg5 = "R"
    If RsSavRec5.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec5.MoveFirst
    FiLLTXT5
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec5.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnLast5_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg5.text = "N" Then
        FindRec5 val(TxtVac_ID5.text)
        Me.TxtModFlg5.text = "R"
    End If
    TxtModFlg5 = "R"
    If RsSavRec5.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec5.MoveLast
    FiLLTXT5
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec5.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify5_Click()

    Dim Msg As String

    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap

    If TxtVac_ID5.text <> "" Then
        TxtModFlg5 = "E"
        Frm25.Enabled = True
        Me.TxtVacName5.SetFocus
        CuurentLogdata5
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«" & CHR(13)
                Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
                Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            If RsSavRec5.EditMode <> adEditNone Then
                RsSavRec5.CancelUpdate
            End If
    End Select
End Sub
Private Sub btnNew5_Click()

    Dim My_SQL As String
    Dim rs As ADODB.Recordset

    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap
    
    Set rs = New ADODB.Recordset
    Frm25.Enabled = True
    Me.TxtVac_ID5.text = ""
    Me.TxtVacName5.text = ""
    Me.TxtVacNamee5.text = ""
    Me.DcboExpensesID.BoundText = ""
    TxtModFlg5.text = "N"

    My_SQL = "TBLProductionElements"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial5.text = rs.RecordCount + 1
    Else
        TxtSerial5.text = 1
    End If
    rs.Close
    CmbType.ListIndex = 0
    TxtVacName5.SetFocus
ErrTrap:
End Sub
Private Sub btnNext5_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String

    If Me.TxtModFlg5.text = "N" Then
        FindRec5 val(TxtVac_ID5.text)
        Me.TxtModFlg5.text = "R"
    End If
    TxtModFlg5 = "R"
    If RsSavRec5.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    If RsSavRec5.EOF Then
        RsSavRec5.MoveLast
    Else
        RsSavRec5.MoveNext
        If RsSavRec5.EOF Then
            RsSavRec5.MoveLast
        End If
    End If
    FiLLTXT5
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec5.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnPrevious5_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    
    If Me.TxtModFlg5.text = "N" Then
        FindRec5 val(TxtVac_ID5.text)
        Me.TxtModFlg5.text = "R"
    End If
    TxtModFlg5 = "R"
    If RsSavRec5.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec5.MovePrevious
    If RsSavRec5.BOF Then
        RsSavRec5.MoveFirst
    End If
    FiLLTXT5
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec5.Requery
            Resume BegnieWork
    End Select
End Sub

Private Sub btnSave5_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control

    For Each CtrlTxt In Me.Controls
        If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
            If CtrlTxt.text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.Title
                CtrlTxt.SetFocus
                Exit Sub
            End If
        End If
    Next
    StrVacName = IsRecExist("TBLProductionElements", "Name", Trim(TxtVacName5.text), "Name", "Vac_ID<>'" & Trim(TxtVac_ID5.text) & "'")
    If StrVacName <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
        Else
            Msg = "This record already exists"
        End If
        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.Title
        TxtVacName5.SetFocus
        Exit Sub
    End If
    Select Case Me.TxtModFlg5.text
        Case "N"
            AddNewRec5
            btnLast5_Click

        Case "E"
            FiLLRec5
    End Select
    Exit Sub
ErrTrap:
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.Title
    Else
        MsgBox "Record saved successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    End If
End Sub
 
Private Sub BtnUndo5_Click()
    FindRec5 val(TxtVac_ID5.text)
    Me.TxtModFlg5.text = "R"
End Sub
Private Sub BtnUpdate5_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim FristCount As Long
    Dim LastCount As Long
    
    FristCount = RsSavRec5.RecordCount
    RsSavRec5.Requery
    LastCount = RsSavRec5.RecordCount
    BtnUndo5_Click
    If SystemOptions.UserInterface = ArabicInterface Then
        If FristCount = LastCount Then
            Msg = "·«  ÊÃœ »Ì«‰«  ÃœÌœ…"
        Else
            Msg = "⁄œœ «·”Ã·«  ﬁ»· «· ÕœÌÀ" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
            If LastCount > FristCount Then
                Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·ÃœÌœ…" & vbCrLf & LastCount - FristCount
            Else
                Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·„Õ–Ê›…" & vbCrLf & FristCount - LastCount
            End If
        End If
    Else
        If FristCount = LastCount Then
            Msg = "No new data"
        Else
            Msg = "No. of records before the update" & vbCrLf & FristCount & vbCrLf & "No. of records after update" & vbCrLf & LastCount
        
            If LastCount > FristCount Then
                Msg = Msg + vbCrLf & "No. of new records" & vbCrLf & LastCount - FristCount
            Else
                Msg = Msg + vbCrLf & "No. of deleted records" & vbCrLf & FristCount - LastCount
            End If
        End If
    End If
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.Title
ErrTrap:
End Sub
Private Sub DcboExpensesID_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        Set Dcombos = New ClsDataCombos
        Dcombos.GetExpensesNames Me.DcboExpensesID
    End If
End Sub
Public Sub AddNewRec5()

    On Error GoTo ErrTrap
    
    Dim StrRecId5 As String
    
    StrRecId5 = new_id("TBLProductionElements", "TBLProductionElementsId", "")
    RsSavRec5.AddNew
    RsSavRec5.Fields("TBLProductionElementsId").value = IIf(StrRecId5 <> "", StrRecId5, Null)
    FiLLRec5
ErrTrap:
End Sub
Function CuurentLogdata5(Optional Currentmode As String)
   
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " ﬂÊœ   " & TxtSerial5.text & CHR(13) & "  «”„ «·⁄‰’— ⁄—»Ì " & TxtVacName5.text & CHR(13) & "  «”„ «·„’—Ê› " & DcboExpensesID
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Code   " & TxtSerial5.text & CHR(13) & "Element English Name" & TxtVacNamee5.text & CHR(13) & " Expenses Namae " & DcboExpensesID
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg5
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D"
    End If
End Function
Public Sub FiLLRec5()

    On Error GoTo ErrTrap

    RsSavRec5.Fields("Name").value = IIf(TxtVacName5.text <> "", Trim(TxtVacName5.text), Null)
    RsSavRec5.Fields("Namee").value = IIf(TxtVacNamee5.text <> "", Trim(TxtVacNamee5.text), Null)
    RsSavRec5.Fields("ExpensesID").value = IIf(DcboExpensesID.BoundText <> 0, val(DcboExpensesID.BoundText), Null)
    RsSavRec5.update
    CuurentLogdata5
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Else
        MsgBox "Saves Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

    End If
    FillGrid5WithData
    TxtModFlg5 = "R"
    Exit Sub
ErrTrap:
    If RsSavRec5.EditMode <> adEditNone Then
        RsSavRec5.CancelUpdate
    End If
End Sub
Public Sub FiLLTXT5()

    On Error GoTo ErrTrap
    
    Dim i As Integer
    
    Frm25.Enabled = False
    TxtVac_ID5.text = IIf(IsNull(RsSavRec5.Fields("TBLProductionElementsId").value), "", RsSavRec5.Fields("TBLProductionElementsId").value)
    TxtVacName5.text = IIf(IsNull(RsSavRec5.Fields("Name").value), "", RsSavRec5.Fields("Name").value)
    TxtVacNamee5.text = IIf(IsNull(RsSavRec5.Fields("Namee").value), "", RsSavRec5.Fields("Namee").value)
    Me.DcboExpensesID.BoundText = IIf(IsNull(RsSavRec5.Fields("ExpensesID").value), "", RsSavRec5.Fields("ExpensesID").value)
    LabCurrRec5.Caption = RsSavRec5.AbsolutePosition
    LabCountRec5.Caption = RsSavRec5.RecordCount
    With Grid5
        For i = 1 To .rows - 1
            If Trim(TxtVac_ID5.text) = .TextMatrix(i, .ColIndex("TBLProductionElementsId")) Then
                TxtSerial5.text = .TextMatrix(i, .ColIndex("Ser"))
                .row = i
                Exit Sub
            End If
        Next
    End With
ErrTrap:
End Sub
Public Sub EditRec5(StrTable As String, RecId5 As String)
    FiLLRec5
End Sub
Private Sub Grid5_EnterCell()

    On Error GoTo ErrTrap
    
    FindRec5 val(Me.Grid5.TextMatrix(Me.Grid5.row, Me.Grid5.ColIndex("TBLProductionElementsId")))
ErrTrap:
End Sub
Private Sub TxtVac_ID5_Change()

    Dim TxtMod As String
    
    TxtMod = TxtModFlg5.text
    TxtModFlg5.text = ""
    TxtModFlg5 = TxtMod
End Sub
Public Function FindRec5(ByVal RecId5 As Long)

    On Error GoTo ErrTrap
    
    RsSavRec5.Find "TBLProductionElementsId=" & RecId5, , adSearchForward, 1
    If Not (RsSavRec5.EOF) Then
        FiLLTXT5
    End If
    Exit Function
ErrTrap:
    If RsSavRec5.EditMode <> adEditNone Then
        RsSavRec5.CancelUpdate
        BtnUndo5_Click
    End If
End Function
Private Sub TxtModFlg5_Change()
    If TxtModFlg5.text = "N" Then
        Frm25.Enabled = True
        Me.btnNew5.Enabled = False
        btnModify5.Enabled = False
        btnDelete5.Enabled = False
        Me.btnQuery5.Enabled = False
        Grid5.Enabled = False
        BtnUndo5.Enabled = True
        Me.btnSave5.Enabled = True
        BtnUpdate5.Enabled = False
    ElseIf TxtModFlg5.text = "R" Then
        Frm25.Enabled = False
        Grid5.Enabled = True
        btnModify5.Enabled = False
        btnDelete5.Enabled = False
        If TxtVac_ID5.text <> "" Then
            btnModify5.Enabled = True
            btnDelete5.Enabled = True
        End If
        BtnUpdate5.Enabled = True
        Me.btnQuery5.Enabled = True
        Me.btnNew5.Enabled = True
        BtnUndo5.Enabled = False
        Me.btnSave5.Enabled = False
        btnNext5.Enabled = True
        btnPrevious5.Enabled = True
        btnFirst5.Enabled = True
        btnLast5.Enabled = True
    ElseIf TxtModFlg5.text = "E" Then
        Frm25.Enabled = True
        Me.btnNew5.Enabled = False
        btnModify5.Enabled = False
        btnDelete5.Enabled = False
        Me.btnQuery5.Enabled = False
        BtnUpdate5.Enabled = False
        BtnUndo5.Enabled = True
        Me.btnSave5.Enabled = True
        Grid5.Enabled = False
        btnNext5.Enabled = False
        btnPrevious5.Enabled = False
        btnFirst5.Enabled = False
        btnLast5.Enabled = False
    End If
End Sub
Public Sub FillGrid5WithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String
    Set rs = New ADODB.Recordset
    
    My_SQL = "select * From TBLProductionElements order by TBLProductionElementsId"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    With Me.Grid5
        .rows = 2
        .Clear flexClearScrollable
        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 1
            rs.MoveFirst
            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs.Fields("Name").value), "", rs.Fields("Name").value)
                .TextMatrix(i, .ColIndex("Namee")) = IIf(IsNull(rs.Fields("Namee").value), "", rs.Fields("Namee").value)
                .TextMatrix(i, .ColIndex("TBLProductionElementsId")) = IIf(IsNull(rs.Fields("TBLProductionElementsId").value), "", rs.Fields("TBLProductionElementsId").value)
                .TextMatrix(i, .ColIndex("ExpensesID")) = IIf(IsNull(rs.Fields("ExpensesID").value), "", rs.Fields("ExpensesID").value)
                rs.MoveNext
            Next
            rs.Close
        End If
        .RowHeight(-1) = 300
    End With
ErrTrap:
End Sub
Private Function CheckDelCountry5(LngExpensesID As Long) As Boolean

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    
    StrSQL = "Select * From TblEmployee Where TBLProductionElementsId=" & LngExpensesID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rs.BOF Or rs.EOF) Then
        CheckDelCountry5 = False
    Else
        CheckDelCountry5 = True
    End If
    rs.Close
    Set rs = Nothing
End Function
Private Sub TxtVacName5_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub
Private Sub TxtVacNamee5_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub

