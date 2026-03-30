VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmDestruction 
   Caption         =   "«–‰ ’—› „Ê«œ ⁄·Ï «·„‘«—Ì⁄"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18090
   HelpContextID   =   370
   Icon            =   "FrmDestruction.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8670
   ScaleWidth      =   18090
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8670
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   18090
      _cx             =   31909
      _cy             =   15293
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   465
         Index           =   3
         Left            =   2250
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   7545
         Width           =   12720
         _cx             =   22437
         _cy             =   820
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
            Height          =   225
            Left            =   7500
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   45
            Width           =   1200
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   4740
            TabIndex        =   13
            Top             =   45
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
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
            Height          =   270
            Left            =   0
            TabIndex        =   55
            Top             =   0
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   225
            Index           =   1
            Left            =   6330
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   45
            Width           =   1140
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   90
            Width           =   825
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   195
            Left            =   2205
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   75
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   195
            Index           =   2
            Left            =   930
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   75
            Width           =   1155
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   195
            Index           =   0
            Left            =   3210
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   75
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "≈Ã„«·Ì «·”‰œ"
            Height          =   180
            Index           =   3
            Left            =   8715
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   45
            Width           =   1425
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1710
         Index           =   0
         Left            =   15
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   660
         Width           =   18015
         _cx             =   31776
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
         Begin VB.TextBox txtNoteId 
            Alignment       =   1  'Right Justify
            Height          =   240
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Top             =   0
            Visible         =   0   'False
            Width           =   1980
         End
         Begin VB.TextBox TXTOverProject 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   975
            Left            =   0
            TabIndex        =   126
            Text            =   "«·”‰œ „ ŒÿÌ "
            Top             =   720
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.TextBox txtManualNO 
            Alignment       =   1  'Right Justify
            Height          =   225
            Left            =   13695
            RightToLeft     =   -1  'True
            TabIndex        =   111
            Top             =   30
            Width           =   1005
         End
         Begin VB.TextBox TxtNoteSerial2 
            Alignment       =   1  'Right Justify
            Height          =   210
            Left            =   15960
            RightToLeft     =   -1  'True
            TabIndex        =   110
            Top             =   45
            Width           =   1230
         End
         Begin VB.ComboBox CBoBasedON 
            Height          =   315
            ItemData        =   "FrmDestruction.frx":038A
            Left            =   1470
            List            =   "FrmDestruction.frx":038C
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   77
            Top             =   120
            Width           =   1980
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   210
            Left            =   15960
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   45
            Width           =   1230
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Height          =   225
            Left            =   7170
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   120
            Width           =   1005
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   210
            Left            =   105
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox TXT_order_no 
            Alignment       =   1  'Right Justify
            Height          =   210
            Left            =   105
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox TxtBillComment 
            Alignment       =   1  'Right Justify
            Height          =   810
            Left            =   3210
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   67
            Top             =   720
            Width           =   13980
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   240
            Left            =   15960
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   405
            Width           =   1230
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   210
            Left            =   12750
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   -1320
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   210
            Left            =   12750
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   -1410
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Frame Frame2 
            Height          =   645
            Left            =   20490
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   75
            Width           =   7305
            Begin MSDataListLib.DataCombo Dcemp 
               Height          =   315
               Left            =   7920
               TabIndex        =   53
               Top             =   120
               Width           =   6315
               _ExtentX        =   11139
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Text            =   ""
               RightToLeft     =   -1  'True
            End
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Height          =   240
            Left            =   1470
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   375
            Width           =   1980
         End
         Begin VB.Frame Frame1 
            Enabled         =   0   'False
            Height          =   600
            Left            =   18975
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   75
            Visible         =   0   'False
            Width           =   4365
            Begin VB.TextBox ItemMakingCost 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   120
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Top             =   480
               Width           =   825
            End
            Begin VB.TextBox ItemMakingQty 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   1920
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   480
               Width           =   1185
            End
            Begin MSDataListLib.DataCombo DcboItemMaking 
               Height          =   315
               Left            =   120
               TabIndex        =   44
               Top             =   120
               Width           =   2955
               _ExtentX        =   5212
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl«· ﬂ·›… 
               Alignment       =   1  'Right Justify
               Caption         =   "«· ﬂ·›…"
               Height          =   255
               Left            =   960
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   480
               Width           =   615
            End
            Begin VB.Label lbl’‰›„Ã„⁄ 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·’‰› «·„’‰⁄"
               Height          =   315
               Index           =   13
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   120
               Width           =   1455
            End
            Begin VB.Label lbl«·ﬂ„Ì… 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂ„Ì…"
               Height          =   315
               Index           =   0
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   480
               Width           =   495
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   750
            Index           =   5
            Left            =   18465
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   810
            Visible         =   0   'False
            Width           =   6510
            _cx             =   11483
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
            ForeColor       =   128
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   "’—› »‰«¡ ⁄·Ï √„— ‘€·"
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
            Begin MSDataListLib.DataCombo DcboWorkOrders 
               Height          =   315
               Left            =   3540
               TabIndex        =   34
               Top             =   210
               Width           =   2205
               _ExtentX        =   3889
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               ListField       =   "7"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboItems 
               Height          =   315
               Left            =   3540
               TabIndex        =   37
               Top             =   540
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo Dcwarsah 
               Height          =   315
               Left            =   120
               TabIndex        =   48
               Top             =   210
               Width           =   2235
               _ExtentX        =   3942
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Ê—‘…"
               Height          =   195
               Index           =   13
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’‰› „Ã„⁄"
               Enabled         =   0   'False
               Height          =   315
               Index           =   12
               Left            =   6000
               RightToLeft     =   -1  'True
               TabIndex        =   36
               Top             =   570
               Width           =   1095
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "√„— ‘€·"
               Enabled         =   0   'False
               Height          =   315
               Index           =   11
               Left            =   5730
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   330
               Width           =   1335
            End
         End
         Begin VB.ComboBox CboType 
            Height          =   315
            Left            =   60
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   1200
            Visible         =   0   'False
            Width           =   1245
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   570
            Index           =   4
            Left            =   60
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1275
            Visible         =   0   'False
            Width           =   1305
            _cx             =   2302
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
            ForeColor       =   192
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   "«·ﬁÌœ «·„Õ«”»Ì"
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
            Begin MSDataListLib.DataCombo DcboDebitSide 
               Height          =   315
               Left            =   60
               TabIndex        =   25
               Top             =   75
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboCreditSide 
               Height          =   315
               Left            =   60
               TabIndex        =   26
               Top             =   360
               Width           =   480
               _ExtentX        =   847
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—› „œÌ‰"
               Height          =   210
               Index           =   32
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   150
               Width           =   345
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—› œ«∆‰"
               Height          =   195
               Index           =   10
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   360
               Width           =   345
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·ﬁÌœ:"
               Height          =   210
               Index           =   9
               Left            =   1110
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   150
               Width           =   135
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·› —… :"
               Height          =   195
               Index           =   8
               Left            =   1110
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   360
               Width           =   135
            End
            Begin VB.Label LblDevID 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   885
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   150
               Width           =   225
            End
            Begin VB.Label lblAccountInterval 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   195
               Left            =   885
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   360
               Width           =   225
            End
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   12345
            TabIndex        =   0
            Top             =   405
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcproject 
            Height          =   315
            Left            =   9135
            TabIndex        =   56
            Top             =   -1410
            Visible         =   0   'False
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcopr 
            Height          =   315
            Left            =   9135
            TabIndex        =   59
            Top             =   -1320
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   285
            Index           =   9
            Left            =   210
            TabIndex        =   63
            Top             =   420
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
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
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   4500
            TabIndex        =   64
            Top             =   120
            Width           =   2610
            _ExtentX        =   4604
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbProcess 
            Height          =   315
            Left            =   4500
            TabIndex        =   68
            Top             =   -1410
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   210
            Left            =   1065
            TabIndex        =   72
            TabStop         =   0   'False
            ToolTipText     =   "«÷€ÿ ·«÷«›… ⁄„Ì· ÃœÌœ"
            Top             =   120
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   370
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
            ButtonImage     =   "FrmDestruction.frx":038E
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   330
            Left            =   9255
            TabIndex        =   76
            Top             =   45
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   582
            _Version        =   393216
            Format          =   220594177
            CurrentDate     =   38784
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·—ﬁ„ «·ÌœÊÌ"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   14790
            TabIndex        =   112
            Top             =   45
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï "
            Height          =   210
            Index           =   7
            Left            =   3000
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   120
            Width           =   1260
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·⁄„·ÌÂ"
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   8250
            TabIndex        =   69
            Top             =   -1410
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„·«ÕŸ« "
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   17385
            TabIndex        =   66
            Top             =   645
            Width           =   600
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   8235
            TabIndex        =   65
            Top             =   120
            Width           =   615
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·»‰œ "
            Height          =   240
            Index           =   16
            Left            =   14010
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   -1320
            Width           =   1305
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„‘—Ê⁄"
            Height          =   210
            Index           =   14
            Left            =   14010
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   -1410
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   195
            Index           =   15
            Left            =   3210
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   420
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   225
            Index           =   4
            Left            =   16680
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   405
            Width           =   1305
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·⁄„·Ì…"
            Height          =   225
            Index           =   6
            Left            =   11325
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   45
            Width           =   1305
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·⁄„·Ì…"
            Height          =   285
            Index           =   5
            Left            =   16680
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   45
            Width           =   1305
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   615
         Left            =   15
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   18015
         _cx             =   31776
         _cy             =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   20.25
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
         Caption         =   " «–‰ ’—› „Ê«œ ⁄·Ï «·„‘«—Ì⁄ "
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
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008080FF&
            Height          =   375
            Left            =   7710
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   120
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008080FF&
            Height          =   360
            Left            =   10440
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   120
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008080FF&
            Height          =   360
            Left            =   9240
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   120
            Visible         =   0   'False
            Width           =   1155
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   2805
            TabIndex        =   7
            Top             =   120
            Width           =   1125
            _ExtentX        =   1984
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
            ButtonImage     =   "FrmDestruction.frx":078B
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
            Left            =   1500
            TabIndex        =   8
            Top             =   120
            Width           =   1140
            _ExtentX        =   2011
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
            ButtonImage     =   "FrmDestruction.frx":0B25
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
            Left            =   4035
            TabIndex        =   9
            Top             =   120
            Width           =   1170
            _ExtentX        =   2064
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
            ButtonImage     =   "FrmDestruction.frx":0EBF
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
            Left            =   180
            TabIndex        =   10
            Top             =   120
            Width           =   1230
            _ExtentX        =   2170
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
            ButtonImage     =   "FrmDestruction.frx":1259
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
            Height          =   405
            Left            =   6165
            Picture         =   "FrmDestruction.frx":15F3
            Stretch         =   -1  'True
            Top             =   0
            Width           =   555
         End
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   5040
         Left            =   0
         TabIndex        =   78
         Top             =   2325
         Width           =   18105
         _cx             =   31935
         _cy             =   8890
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
         Caption         =   "«·»Ì«‰«  «·«”«”Ì…|Õ«·… «·«⁄ „«œ"
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   4620
            Left            =   45
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   45
            Width           =   18015
            _cx             =   31776
            _cy             =   8149
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
            Begin VB.Timer Timer1 
               Interval        =   500
               Left            =   120
               Top             =   960
            End
            Begin VB.TextBox Text7 
               Alignment       =   1  'Right Justify
               Height          =   240
               Left            =   15105
               RightToLeft     =   -1  'True
               TabIndex        =   124
               Top             =   750
               Width           =   1530
            End
            Begin VB.TextBox Text6 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   15105
               RightToLeft     =   -1  'True
               TabIndex        =   123
               Top             =   465
               Width           =   1530
            End
            Begin VB.TextBox Text5 
               Alignment       =   1  'Right Justify
               Height          =   270
               Left            =   15105
               RightToLeft     =   -1  'True
               TabIndex        =   113
               Top             =   180
               Width           =   1530
            End
            Begin XtremeSuiteControls.CheckBox ChAuto 
               Height          =   225
               Left            =   3105
               TabIndex        =   80
               Top             =   180
               Width           =   1200
               _Version        =   786432
               _ExtentX        =   2117
               _ExtentY        =   397
               _StockProps     =   79
               Caption         =   " Õ„Ì· «·„Ê«œ «·Ì«"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   555
               Index           =   2
               Left            =   30
               TabIndex        =   81
               TabStop         =   0   'False
               Top             =   1035
               Width           =   17955
               _cx             =   31671
               _cy             =   979
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
               Begin VB.ComboBox CboItemCase 
                  Height          =   315
                  Left            =   8340
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   85
                  Top             =   240
                  Width           =   2325
               End
               Begin VB.TextBox TxtQuantity 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   3060
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   240
                  Width           =   2145
               End
               Begin VB.TextBox TxtSerial 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   5205
                  MaxLength       =   20
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   240
                  Width           =   3045
               End
               Begin VB.TextBox TxtPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   255
                  Left            =   870
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   82
                  Top             =   240
                  Width           =   2115
               End
               Begin MSDataListLib.DataCombo DCboItemsName 
                  Height          =   315
                  Left            =   10695
                  TabIndex        =   86
                  Top             =   240
                  Width           =   3405
                  _ExtentX        =   6006
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCboItemsCode 
                  Height          =   315
                  Left            =   13935
                  TabIndex        =   87
                  Top             =   240
                  Width           =   3660
                  _ExtentX        =   6456
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdAdd 
                  Height          =   300
                  Left            =   750
                  TabIndex        =   88
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1050
                  _ExtentX        =   1852
                  _ExtentY        =   529
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
                  ButtonImage     =   "FrmDestruction.frx":525B
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
               Begin ImpulseButton.ISButton ISButton4 
                  Height          =   345
                  Left            =   -720
                  TabIndex        =   127
                  Top             =   120
                  Width           =   1635
                  _ExtentX        =   2884
                  _ExtentY        =   609
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈œ—«Ã"
                  BackColor       =   14871017
                  FontSize        =   13.5
                  FontBold        =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   13.5
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ButtonImage     =   "FrmDestruction.frx":55F5
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   -2147483637
                  ColorOutline    =   0
                  Alignment       =   1
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   -2147483637
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·’‰›"
                  Height          =   195
                  Index           =   31
                  Left            =   14445
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   0
                  Width           =   3135
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈”„ «·’‰›"
                  Height          =   195
                  Index           =   30
                  Left            =   11160
                  RightToLeft     =   -1  'True
                  TabIndex        =   93
                  Top             =   0
                  Width           =   3105
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·… «·’‰›"
                  Height          =   195
                  Index           =   29
                  Left            =   8610
                  RightToLeft     =   -1  'True
                  TabIndex        =   92
                  Top             =   0
                  Width           =   2190
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ì—Ì«·"
                  Height          =   195
                  Index           =   28
                  Left            =   5400
                  RightToLeft     =   -1  'True
                  TabIndex        =   91
                  Top             =   0
                  Width           =   3000
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂ„Ì…"
                  Height          =   195
                  Index           =   27
                  Left            =   3690
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   0
                  Width           =   1560
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”⁄—"
                  Height          =   195
                  Index           =   26
                  Left            =   1395
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   0
                  Width           =   1695
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid FG 
               Height          =   2340
               Left            =   0
               TabIndex        =   95
               Top             =   1680
               Width           =   17970
               _cx             =   31697
               _cy             =   4128
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
               Cols            =   19
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmDestruction.frx":598F
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
            Begin MSDataListLib.DataCombo dcopr2 
               Height          =   315
               Left            =   4485
               TabIndex        =   96
               Top             =   480
               Width           =   10545
               _ExtentX        =   18600
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcproject1 
               Height          =   315
               Left            =   4485
               TabIndex        =   97
               Top             =   180
               Width           =   10545
               _ExtentX        =   18600
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbProcess1 
               Height          =   315
               Left            =   4485
               TabIndex        =   98
               Top             =   750
               Width           =   10545
               _ExtentX        =   18600
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSComctlLib.Toolbar TBar 
               Height          =   630
               Left            =   525
               TabIndex        =   99
               Top             =   4035
               Width           =   7035
               _ExtentX        =   12409
               _ExtentY        =   1111
               ButtonWidth     =   609
               ButtonHeight    =   1005
               Appearance      =   1
               _Version        =   393216
            End
            Begin ImpulseButton.ISButton ISButton2 
               Height          =   300
               Left            =   525
               TabIndex        =   100
               Top             =   0
               Visible         =   0   'False
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   529
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
               ButtonImage     =   "FrmDestruction.frx":5C5E
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
            Begin ImpulseButton.ISButton ISButton1 
               Height          =   345
               Left            =   2640
               TabIndex        =   125
               Top             =   360
               Visible         =   0   'False
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   609
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "≈œ—«Ã"
               BackColor       =   14871017
               FontSize        =   13.5
               FontBold        =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmDestruction.frx":5FF8
               ColorButton     =   14871017
               ColorHighlight  =   16777215
               ColorHoverText  =   16711680
               ColorShadow     =   -2147483637
               ColorOutline    =   0
               Alignment       =   1
               DrawFocusRectangle=   0   'False
               ColorToggledHoverText=   16711680
               LowerToggledContent=   0   'False
               ColorTextShadow =   -2147483637
            End
            Begin VB.Label LblItemsCount 
               Alignment       =   1  'Right Justify
               Height          =   270
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   4245
               Width           =   390
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·»‰œ"
               Height          =   165
               Index           =   17
               Left            =   16290
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   480
               Width           =   1305
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„‘—Ê⁄"
               Height          =   270
               Index           =   18
               Left            =   16395
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   180
               Width           =   1305
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄„·ÌÂ"
               Height          =   255
               Index           =   19
               Left            =   16395
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   765
               Width           =   1305
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   4620
            Left            =   18750
            TabIndex        =   105
            TabStop         =   0   'False
            Top             =   45
            Width           =   18015
            _cx             =   31776
            _cy             =   8149
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
            Begin VSFlex8UCtl.VSFlexGrid GRID2 
               Height          =   3795
               Left            =   120
               TabIndex        =   106
               Tag             =   "1"
               Top             =   135
               Width           =   17640
               _cx             =   31115
               _cy             =   6694
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
               FormatString    =   $"FrmDestruction.frx":6392
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
            Begin VB.Label Label1100 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
               Height          =   225
               Left            =   11670
               RightToLeft     =   -1  'True
               TabIndex        =   108
               Top             =   4815
               Width           =   3570
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
               Height          =   240
               Left            =   6855
               RightToLeft     =   -1  'True
               TabIndex        =   107
               Top             =   4170
               Width           =   3555
            End
         End
      End
      Begin ImpulseButton.ISButton Accredit 
         Height          =   345
         Left            =   120
         TabIndex        =   109
         Top             =   7380
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   609
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   540
         Index           =   0
         Left            =   16050
         TabIndex        =   114
         Top             =   8070
         Width           =   1860
         _ExtentX        =   3281
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
         Left            =   14055
         TabIndex        =   115
         Top             =   8070
         Width           =   1890
         _ExtentX        =   3334
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
         Left            =   12030
         TabIndex        =   116
         Top             =   8070
         Width           =   1860
         _ExtentX        =   3281
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
         Left            =   10230
         TabIndex        =   117
         Top             =   8070
         Width           =   1620
         _ExtentX        =   2858
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
         Left            =   8115
         TabIndex        =   118
         Top             =   8070
         Width           =   1755
         _ExtentX        =   3096
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
         Left            =   6060
         TabIndex        =   119
         Top             =   8070
         Width           =   1695
         _ExtentX        =   2990
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
         Left            =   0
         TabIndex        =   120
         Top             =   8070
         Width           =   1800
         _ExtentX        =   3175
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
      Begin ImpulseButton.ISButton CmdHelp 
         Height          =   540
         Left            =   2010
         TabIndex        =   121
         Top             =   8070
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   953
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
         Height          =   540
         Index           =   7
         Left            =   3975
         TabIndex        =   122
         Top             =   8070
         Width           =   1845
         _ExtentX        =   3254
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
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   0
      TabIndex        =   38
      Top             =   -720
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   315
      Left            =   0
      TabIndex        =   40
      Top             =   1560
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label lbl√„—‘€· 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "√„— ‘€·"
      Enabled         =   0   'False
      Height          =   315
      Index           =   13
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lbl«·’‰›«·„’‰⁄ 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·’‰› «·„’‰⁄"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   -720
      Width           =   975
   End
End
Attribute VB_Name = "FrmDestruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim NewGrid As New ClsGrid
Dim SaleReport As ClsSaleReport
Dim cSearchDcbo(3)   As clsDCboSearch
Public project_id1 As Integer
Public BolPrint As Boolean
Dim materilaAcc2 As String


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
     
     
    TxtFillData.Text = "F"
    TxtFillData_Change
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Function GetTransectionID(Optional Noseri As String) As Long
Dim StrSQL As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
StrSQL = "SELECT Transaction_ID FROM Transactions WHERE NoteSerial1=N'" & Noseri & "'"
If CBoBasedON.ListIndex = 3 Then
    StrSQL = StrSQL & " and (Transaction_Type=42)"
ElseIf CBoBasedON.ListIndex = 2 Then
    StrSQL = StrSQL & " and (Transaction_Type=26 )"
End If
rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetTransectionID = IIf(IsNull(rs2("Transaction_ID").value), 0, rs2("Transaction_ID").value)
Else
GetTransectionID = 0
End If
End Function

Public Sub RetriveOrderProd(Optional order_no As String = "", Optional ByVal mType As Long = 2)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim mTransID As Long
    
    mTransID = GetTransectionID(order_no)
    Dim Num As Long
    On Error GoTo ErrTrap
        FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 2
    StrSQL = " Select * from transactions "
    If mType = 2 Then
        StrSQL = StrSQL & "  WHERE     (dbo.Transactions.Transaction_Type = 26) "
        
    ElseIf mType = 3 Then
         StrSQL = StrSQL & "  WHERE     (dbo.Transactions.Transaction_Type = 42) "
    End If
    StrSQL = StrSQL & "  AND (dbo.Transactions.Transaction_ID = " & mTransID & ")"
    
    
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

 
    If rs.RecordCount < 1 Then
 
 

        Exit Sub
    Else
      '  DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)

           
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
 
    FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.Text = ""
    Dim mCostPrice As Double
    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))
            
            mCostPrice = 0
            If val(TxtNoteSerial1) <> 0 And SystemOptions.CostByProduction Then
           
            
                    s = "SELECT T2.* "
                    s = s & " from  Transactions AS t "
                    s = s & " Inner Join Transaction_Details T2 On T2.Transaction_ID = t.Transaction_ID"
                    s = s & " WHERE t.Transaction_Type = 26 and t.OrderID =  " & val(mTransID)
                    s = s & " and  T2.Item_ID = " & val(RsDetails!Item_ID & "")
                    s = s & " and T2.UnitId= " & val(RsDetails!UnitID & "")
                    Set rsDummy = New ADODB.Recordset
    
    '
                    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    If rsDummy.EOF Then
                        mCostPrice = 0
                    Else
                        mCostPrice = val(rsDummy!ShowPrice & "")
                    End If
    
   
    
                If mCostPrice <> 0 Then
                    FG.TextMatrix(Num, FG.ColIndex("Price")) = mCostPrice
                
                End If
            End If
            If mCostPrice = 0 Then
                FG.TextMatrix(Num, FG.ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(val(FG.TextMatrix(Num, FG.ColIndex("Code"))), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, , , val(Me.DCboStoreName.BoundText))
            End If
            'ModItemCostPrice.GetCostItemPrice(Fg.TextMatrix(RowNum, Fg.ColIndex("Code")), 0, Fg.TextMatrix(RowNum, Fg.ColIndex("Serial")), , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Me.Text1.Text), RsDetails("UnitID").value, Me.DCboStoreName.BoundText)
           
           ' Fg.TextMatrix(Num, Fg.ColIndex("ShowQty")) = IIf(IsNull(RsDetails("ShowQty")), "", (RsDetails("ShowQty").value))
           ' Fg.TextMatrix(Num, Fg.ColIndex("showPrice")) = IIf(IsNull(RsDetails("showPrice")), "", (RsDetails("showPrice").value))
            FG.TextMatrix(Num, FG.ColIndex("Valu")) = val(FG.TextMatrix(Num, FG.ColIndex("Count"))) * val(FG.TextMatrix(Num, FG.ColIndex("Price")))
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
            FG.TextMatrix(Num, FG.ColIndex("L")) = IIf(IsNull(RsDetails("L")), "", (RsDetails("L").value))
             FG.TextMatrix(Num, FG.ColIndex("W")) = IIf(IsNull(RsDetails("W")), "", (RsDetails("W").value))
             FG.TextMatrix(Num, FG.ColIndex("H1")) = IIf(IsNull(RsDetails("H1")), "", (RsDetails("H1").value))
             FG.TextMatrix(Num, FG.ColIndex("H2")) = IIf(IsNull(RsDetails("H2")), "", (RsDetails("H2").value))
             FG.TextMatrix(Num, FG.ColIndex("NoCount")) = IIf(IsNull(RsDetails("NoCount")), "", (RsDetails("NoCount").value))
             FG.TextMatrix(Num, FG.ColIndex("Area")) = IIf(IsNull(RsDetails("Area")), "", (RsDetails("Area").value))
             FG.TextMatrix(Num, FG.ColIndex("Height")) = IIf(IsNull(RsDetails("Height")), "", (RsDetails("Height").value))
             FG.TextMatrix(Num, FG.ColIndex("length")) = IIf(IsNull(RsDetails("length")), "", (RsDetails("length").value))
             FG.TextMatrix(Num, FG.ColIndex("Width")) = IIf(IsNull(RsDetails("Width")), "", (RsDetails("Width").value))
         
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        
            RsDetails.MoveNext
            Debug.Print Num

            If FG.rows > 10 Then
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
Public Sub RetriveOrderProduct(Optional order_no As String = "")
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
    StrSQLMain = " SELECT     dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ShowQty"
    StrSQLMain = StrSQLMain & " FROM         dbo.Transactions INNER JOIN"
    StrSQLMain = StrSQLMain & " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
    StrSQLMain = StrSQLMain & "  WHERE     (dbo.Transactions.Transaction_Type = 26) AND (dbo.Transactions.Transaction_Serial = N'" & order_no & "')"
    RsMainData.Open StrSQLMain, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsMainData.RecordCount < 1 Then
 
        Exit Sub
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
               FG.TextMatrix(currentrow, FG.ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(LngItemID2, 0, "", , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(XPTxtBillID.Text), val(FG.TextMatrix(currentrow, FG.ColIndex("UnitID"))), val(Me.DCboStoreName.BoundText))
                 
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

    TxtFillData.Text = "F"
    Screen.MousePointer = vbDefault
    ' XPDtbBill_Change

    'XPTxtCurrent.Caption = rs.AbsolutePosition
    'XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub
Private Sub Accredit_Click()
    Dim BeginTrans As Boolean
 
If val(XPTxtBillID.Text) = 0 Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "«Õ›Ÿ «·”‰œ «Ê·«", vbCritical
     Else
     MsgBox "Save Doc First", vbCritical
     End If
      
      Exit Sub
      End If
 
 
    SendTopost Me.Name, "Transactions", "Transaction_ID", 0, val(dcBranch.BoundText), val(XPTxtBillID.Text), TxtNoteSerial2.Text, , , IIf(Me.TXTOverProject.Visible, 1, 0)
    If TxtModFlg.Text <> "N" And TxtModFlg.Text <> "E" Then
    rs.Resync
    End If
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If
    fillapprovData
End Sub

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

Private Sub CboType_Change()

    If Me.TxtModFlg.Text = "R" Then Exit Sub
    WriteDev

    If Me.CboType.ListIndex = 2 Then
        Me.DcboWorkOrders.Enabled = True
        Me.Frame1.Enabled = True
        Me.lbl(11).Enabled = True
        dcproject.Visible = True
        DcEmp.Visible = False

        If SystemOptions.UserInterface = EnglishInterface Then
            lbl(14).Caption = "Project"
        Else
            lbl(14).Caption = "«·„‘—Ê⁄"
        End If
    
    Else

        If Me.CboType.ListIndex = 1 Then
            DcEmp.Visible = True
            dcproject.Visible = False

            If SystemOptions.UserInterface <> EnglishInterface Then
                lbl(14).Caption = "«·„ÊŸ›"
            Else
                lbl(14).Caption = "Employee"
            End If

        Else

            Me.DcboWorkOrders.Enabled = False
            Me.lbl(11).Enabled = False
            Me.DcboItems.BoundText = ""
            NewGrid.RestrictedAssbliedItemID = 0
            NewGrid.WorkOrderID = 0
            '   Me.Frame1.Enabled = False
            dcproject.Visible = False
            DcEmp.Visible = False
            '   DCboStoreName.BoundText = ""
        End If
    End If

End Sub

Private Sub CboType_Click()
    CboType_Change
End Sub

Private Sub ChAuto_Click()
If Me.ChAuto.value = vbChecked Then
ISButton4.Visible = False
ISButton1.Visible = True
Else
ISButton1.Visible = False
ISButton4.Visible = True
End If
End Sub

Private Sub Cmd_Click(Index As Integer)
    Dim AskOption As Boolean
    Dim intDef As Integer
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTest As ADODB.Recordset
Me.CboType.ListIndex = 2
    'On Error GoTo ErrTrap
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            clear_all Me
            
            TxtModFlg.Text = "N"
            NewGrid.GridDefaultValue 1
            Accredit.Caption = ""
               Grid2.Clear flexClearScrollable, flexClearEverything
              Grid2.rows = 1
              Timer1.Enabled = False
              
            Me.DCboUserName.BoundText = user_id
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSaleStore", 1)
            DCboStoreName.BoundText = intDef
'            FG.SetFocus
            FG.Col = FG.ColIndex("Code")
            FG.Row = FG.rows - 2

            If SystemOptions.UserInterface = EnglishInterface Then
                CboItemCase.Clear
                CboItemCase.AddItem "New"
                CboItemCase.AddItem "Used"
            End If
dcBranch.BoundText = Current_branch
CboType.ListIndex = 2
        Case 1
        
                    If ScreenAproved(val(Me.XPTxtBillID.Text), Me.Name) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "«·Õ—ﬂÂ „— »ÿÂ »«·«⁄ „«œ« "
                Else
                    MsgBox "Can not edit.This process associated with approvals"
                End If
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
              
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.DcboWorkOrders.BoundText) <> 0 Then
                If CheckOrderState(val(Me.DcboWorkOrders.BoundText)) = False Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·«Ì„ﬂ‰ «· ⁄œÌ· .. ›ﬁœ  „ ≈€·«ﬁ «„— «·‘€·...!!!"
                   Else
                   Msg = "You Can not Edit "
                   End If
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If

            TxtModFlg.Text = "E"
            Me.DCboUserName.BoundText = user_id

        Case 2
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
                    Msg = "Õœœ «·›—⁄ «Ê·«  "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                dcBranch.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
  my_branch = Me.dcBranch.BoundText
  
  
            If Me.TxtModFlg.Text = "N" Then
         
                If TxtNoteSerial.Text = "" Then
                    If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
                        MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
                    Else
                       
                        If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
                            MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                        Else
                            TxtNoteSerial.Text = Notes_coding(val(my_branch), XPDtbBill.value)
                        End If
                    End If
                End If
 
            End If
              Dim TxtNoteSerial1str As String
my_branch = val(Me.dcBranch.BoundText)
    If TxtNoteSerial2.Text = "" Then
     TxtNoteSerial1str = Voucher_coding(val(my_branch), XPDtbBill.value, 66, 66)
    
                If TxtNoteSerial1str = "error" Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›…  Õ—ﬂ…  ÃœÌœ…  ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                Else
                               
                    If TxtNoteSerial1str = "" Then
                        MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„  «·Õ—ﬂ… ÃœÌœ     ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                    End If
                End If
    End If
            SaveData

        Case 3
            Call Undo

        Case 4
                        If ScreenAproved(val(Me.XPTxtBillID.Text), Me.Name) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "«·Õ—ﬂÂ „— »ÿÂ »«·«⁄ „«œ« "
                Else
                    MsgBox "Can not edit.This process associated with approvals"
                End If
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

            Del_TransAction

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            FrmBuySearch.DealingForm = Destruction
            FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ”‰œ ’—› «·„‘«—Ì⁄"

            FrmBuySearch.FG.ColHidden(FrmBuySearch.FG.ColIndex("ClientNmae")) = True
            FrmBuySearch.show vbModal

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            PrintReport

        Case 6
            rs.Requery
            Unload Me
         Case 9
           ShowGL_cc Me.TxtNoteSerial.Text, , 200, TXTNoteID
    End Select
      
    Exit Sub
ErrTrap:
End Sub

Sub FillGrid()
Dim i As Integer
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Dim sql As String
Dim current_row As Integer
sql = " SELECT     dbo.TblMatrials.Pand, dbo.TblMatrials.ProjectID, dbo.TblMatrials.monthly, dbo.TblMatrials.catalogID, dbo.TblMatrials.OperCode, dbo.TblMatrials.priceapro, "
sql = sql & "                       dbo.TblMatrials.Quntapro, dbo.TblMatrials.Price, dbo.TblMatrials.[Count], dbo.TblMatrials.ItemID, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,"
sql = sql & "                       dbo.TblItems.fullcode , dbo.terms_operations.OPRIDD"
sql = sql & "  FROM         dbo.TblMatrials RIGHT OUTER JOIN"
sql = sql & "                       dbo.terms_operations ON dbo.TblMatrials.Opr = dbo.terms_operations.id LEFT OUTER JOIN"
sql = sql & "                       dbo.TblItems ON dbo.TblMatrials.ItemID = dbo.TblItems.ItemID"
sql = sql & "  Where (dbo.TblMatrials.ProjectID =" & project_id1 & ") "
If dcopr2.Text <> "" And val(dcopr2.BoundText) <> 0 Then
'sql = sql & " And dbo.TblMatrials.Pand = " & val(dcopr2.BoundText) & ""
sql = sql & " and dbo.terms_operations.ProjectDes_ID =" & val(dcopr2.BoundText) & ""

End If
If DcbProcess1.Text <> "" And val(DcbProcess1.BoundText) <> 0 Then
sql = sql & " And dbo.terms_operations.OPRIDD  =" & val(DcbProcess1.BoundText) & ""
End If
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then

    

    With FG
    rs2.MoveFirst
 For i = 1 To rs2.RecordCount
    FG.rows = FG.rows + 1
        current_row = FG.rows - 1
    .TextMatrix(current_row, .ColIndex("operaid")) = DcbProcess1.BoundText
     .TextMatrix(current_row, .ColIndex("pandid")) = Me.dcopr2.BoundText
     .TextMatrix(current_row, .ColIndex("projectid")) = project_id1
     .TextMatrix(current_row, .ColIndex("project")) = Me.dcproject1.Text
     .TextMatrix(current_row, .ColIndex("pand")) = Me.dcopr2.Text
     .TextMatrix(current_row, .ColIndex("opera")) = DcbProcess1.Text
    ' .TextMatrix(i, .ColIndex("Count")) = IIf(IsNull(Rs2("Quntapro").value), 0, Rs2("Quntapro").value)
    ' .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(Rs2("priceapro").value), 0, Rs2("priceapro").value)
    ' .TextMatrix(i, .ColIndex("Valu")) = val(.TextMatrix(i, .ColIndex("Count"))) * val(.TextMatrix(i, .ColIndex("Price")))
    ' .TextMatrix(i, .ColIndex("opera")) = IIf(IsNull(Rs2("ItemID").value), 0, Rs2("ItemID").value)
    ' .TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(Rs2("ItemID").value), "", Rs2("ItemID").value)
    ' .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs2("ItemName").value), "", Rs2("ItemName").value)
    
     
     ' DCboItemsCode.BoundText = IIf(IsNull(Rs2("ItemID").value), 0, Rs2("ItemID").value)
      DCboItemsName.BoundText = IIf(IsNull(rs2("ItemID").value), "", rs2("ItemID").value)
       TxtQuantity.Text = IIf(IsNull(rs2("Quntapro").value), 0, rs2("Quntapro").value)
      TxtPrice.Text = IIf(IsNull(rs2("priceapro").value), 0, rs2("priceapro").value)
    NewGrid.CmdAddData_Click
    rs2.MoveNext
  Next i
  
    End With
 End If
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub DCboItemsCode_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF3 Then
        
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 7
        FrmItemSearch.show vbModal
    End If

End Sub

Private Sub DCboStoreName_Click(Area As Integer)

    If DCboStoreName.Text = "" Then Exit Sub
    If TxtModFlg.Text <> "R" Then
        WriteDev
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

Private Sub DcboWorkOrders_Change()
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    If val(Me.DcboWorkOrders.BoundText) = 0 Then
        Me.DcboItems.BoundText = ""
    Else
        StrSQL = "Select * From TblWorkOrdersData Where WorkOrderID=" & val(Me.DcboWorkOrders.BoundText)
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            If IsNull(rs("AssbliedItemID").value) = True Then
                Me.DcboItems.BoundText = ""
                NewGrid.RestrictedAssbliedItemID = 0
                NewGrid.WorkOrderID = 0
            Else
                Me.DcboItems.BoundText = rs("AssbliedItemID").value
                NewGrid.WorkOrderID = val(Me.DcboWorkOrders.BoundText)
                NewGrid.RestrictedAssbliedItemID = rs("AssbliedItemID").value

                If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
                    'NewGrid.AddItemInGrid 0, Rs("AssbliedItemID").Value, 1
                End If
            End If
        End If

        rs.Close
        Set rs = Nothing
    End If

End Sub

Private Sub DcboWorkOrders_Click(Area As Integer)
    DcboWorkOrders_Change
End Sub

Private Sub Dcbranch_Change()
Dcbranch_Click (0)
End Sub

Private Sub Dcbranch_Click(Area As Integer)
    
   If Me.TxtModFlg.Text <> "R" Then
   TxtNoteSerial2.Text = ""
      If ChekSanNumber(Current_branch, 66) = True Then
          TxtNoteSerial2.Text = ""
      End If
      TxtNoteSerial2.Text = ""
   End If
End Sub

Private Sub Dcemp_Click(Area As Integer)
    On Error Resume Next

    If DcEmp.BoundText <> "" Then
        If TxtModFlg.Text <> "R" Then
            WriteDev
        End If
    End If

End Sub

Private Sub dcopr_Click(Area As Integer)
 Dim Dcombos As ClsDataCombos
 Dim project_id As Integer
       Set Dcombos = New ClsDataCombos
  If dcproject.BoundText <> "" Then
         'project_id = get_project_id(DCproject.BoundText, "Material_account")
         project_id = val(dcproject.BoundText)
         
         If Me.dcopr.BoundText <> "" Then
         Dcombos.GetProcessOfProjedt DcbProcess, project_id, , dcopr.BoundText, 2
         End If
       
    End If
End Sub

Private Sub dcopr_KeyUp(KeyCode As Integer, _
                        Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim project_id As Integer
        On Error Resume Next

        If dcproject.BoundText <> "" Then
            WriteDev
            'project_id = get_project_id(DCproject.BoundText, "Material_account")
            project_id = val(dcproject.BoundText)
            'Material_account = materilaAcc2
            fillopr (project_id)
        End If

    End If

End Sub

Private Sub dcopr2_Click(Area As Integer)
 Dim Dcombos As ClsDataCombos
 Dim project_id As Integer
       Set Dcombos = New ClsDataCombos
  If dcproject1.BoundText <> "" Then
         'project_id = get_project_id(dcproject1.BoundText, "Material_account")
         project_id = val(dcproject.BoundText)
         'Material_account = materilaAcc2
         If Me.dcopr2.BoundText <> "" Then
         Dcombos.GetProcessOfProjedt DcbProcess1, project_id, , dcopr2.BoundText, 2
         End If
       
    End If
End Sub

Private Sub dcproject_Click(Area As Integer)
    Dim project_id As Integer
    On Error Resume Next

    If dcproject.BoundText <> "" Then
        WriteDev
        'project_id = get_project_id(DCproject.BoundText, "Material_account")
        project_id = val(dcproject.BoundText)
         
        fillterms (project_id)
    End If

End Sub

  



Function fillterms(project_id As Integer)
    Dim My_SQL As String
 
    My_SQL = " select oprid,des from dbo.projects_des where project_id=" & project_id

    fill_combo Me.dcopr, My_SQL
       
        
    dcopr.ReFill
End Function
Function fillterms1(project_id As Integer)
    Dim My_SQL As String
 
    My_SQL = " select oprid,des from dbo.projects_des where project_id=" & project_id

  
        fill_combo Me.dcopr2, My_SQL
        
    dcopr.ReFill
End Function

Function fillopr(project_id As Integer)
    Dim My_SQL As String
 
    My_SQL = "  select fullcode,name from terms_operations where project_id=" & project_id

    fill_combo Me.dcopr, My_SQL
    dcopr.ReFill
End Function

Private Sub dcproject_KeyUp(KeyCode As Integer, _
                            Shift As Integer)

        If KeyCode = vbKeyF3 Then
         FrmProjectSearch.lblSearchtype.Caption = 1010
             FrmProjectSearch.show vbModal
           
        End If
        
        

    If KeyCode = vbKeyF5 Then
        Dim My_SQL As String
        My_SQL = " select id,Project_name from projects"
 
        fill_combo dcproject, My_SQL

    End If

End Sub

Private Sub dcproject1_Click(Area As Integer)
    
    On Error Resume Next

    If dcproject1.BoundText <> "" Then
       
    'project_id1 = get_project_id(dcproject1.BoundText, "Material_account")
        project_id1 = val(dcproject1.BoundText)
        If project_id1 = 0 Then project_id1 = val(dcproject1.BoundText)
        fillterms1 (project_id1)
    End If
End Sub

Private Sub dcproject1_KeyUp(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF3 Then
         FrmProjectSearch.lblSearchtype.Caption = 11
              
              FrmProjectSearch.show vbModal
           
        End If

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
StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtBillID.Text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If RsDetails.RecordCount > 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If
Accredit.Enabled = False
Else
Accredit.Enabled = True
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If
End If
 If Not (RsDetails.EOF Or RsDetails.BOF) Then
        Grid2.rows = RsDetails.RecordCount + 1
 

        For Num = 1 To RsDetails.RecordCount
        
       Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
    If Grid2.TextMatrix(Num, Grid2.ColIndex("Currcursor")) = "1" Then
   Grid2.cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
   Else
    Grid2.cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
    End If
    
        Grid2.TextMatrix(Num, Grid2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
           If SystemOptions.UserInterface = ArabicInterface Then
            Grid2.TextMatrix(Num, Grid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
          Else
             Grid2.TextMatrix(Num, Grid2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
          End If
            If SystemOptions.UserInterface = ArabicInterface Then
            Grid2.TextMatrix(Num, Grid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            Else
            Grid2.TextMatrix(Num, Grid2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            End If
            Grid2.TextMatrix(Num, Grid2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
          Grid2.TextMatrix(Num, Grid2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
 
 
RsDetails.MoveNext
If Num = RsDetails.RecordCount Then

        If Grid2.TextMatrix(Num, Grid2.ColIndex("Approved")) <> "" Then
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
 Grid2.rows = 1
    End If
RsDetails.Close

End Function

Private Sub FG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
   Dim Rs1 As ADODB.Recordset
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
Dim sql As String
Set Rs1 = New ADODB.Recordset
    With FG

        Select Case .ColKey(Col)
               Case "project"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("projectid"), False, True)
                .TextMatrix(Row, .ColIndex("projectid")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("operaid")) = 0
                .TextMatrix(Row, .ColIndex("pandid")) = 0
                .TextMatrix(Row, .ColIndex("pand")) = ""
                .TextMatrix(Row, .ColIndex("opera")) = ""
               Case "pand"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("pandid"), False, True)
                .TextMatrix(Row, .ColIndex("pandid")) = StrAccountCode
                 .TextMatrix(Row, .ColIndex("operaid")) = 0
                .TextMatrix(Row, .ColIndex("opera")) = ""
              Case "opera"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("operaid"), False, True)
                .TextMatrix(Row, .ColIndex("operaid")) = StrAccountCode
 End Select
 End With
End Sub

Private Sub FG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With FG
Select Case .ColKey(Col)
Case "Code"
.ComboList = ""
Case "Serial"
.ComboList = ""
Case "Count"
.ComboList = ""
Case "Price"
.ComboList = ""
Case "Valu"
Cancel = True
End Select
End With
End Sub

Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
 
    With Me.FG

        Select Case .ColKey(Col)
      Case "project"
               StrSQL = "  SELECT     id, Project_name, Project_nameE"
               StrSQL = StrSQL & "     From dbo.Projects"
             If SystemOptions.UserInterface = ArabicInterface Then
              StrSQL = StrSQL & " Where (Not (Project_name Is Null)) and Project_name<>N'""'"
              Else
              StrSQL = StrSQL & " Where (Not (Project_nameE Is Null))and Project_nameE <>N'""'"
             End If
       Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "Project_name", "id")
                Else
                    StrComboList = .BuildComboList(rs, "Project_nameE", "id")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
         ''//////////////
              Case "pand"
               StrSQL = " SELECT     oprid, des"
               StrSQL = StrSQL & "          From dbo.projects_des"
               StrSQL = StrSQL & "          Where (project_id = " & val(.TextMatrix(Row, .ColIndex("ProjectID"))) & " and project_id<>0)"
        
       Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
               StrComboList = .BuildComboList(rs, "des", "oprid")
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
                

          Case "opera"
               StrSQL = "       SELECT     dbo.terms_operations.OPRIDD, dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.ProcessNameE"
               StrSQL = StrSQL & "       FROM         dbo.terms_operations LEFT OUTER JOIN"
               StrSQL = StrSQL & "       dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEFID"
               StrSQL = StrSQL & "   Where (dbo.terms_operations.project_id = " & val(.TextMatrix(Row, .ColIndex("ProjectID"))) & ") And (dbo.terms_operations.ProjectDes_ID = " & val(.TextMatrix(Row, .ColIndex("PandID"))) & ")"
       Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                   StrComboList = .BuildComboList(rs, "ProcessName", "OPRIDD")
                Else
                   StrComboList = .BuildComboList(rs, "ProcessNameE", "OPRIDD")
                End If
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
   End Select
   End With
End Sub

Private Sub ImgFavorites_Click()
    AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub ISButton1_Click()
If ChAuto.value = vbChecked Then
FillGrid
ChAuto.value = vbUnchecked
Exit Sub
End If
End Sub

Private Sub ISButton3_Click()
If val(CBoBasedON.ListIndex) = 0 Then
FrmBuySearch.DealingForm = GridTransType.internalorder
  FrmBuySearch.Index = 1
            FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ÿ·»«   œ«Œ·Ì…"
            FrmBuySearch.show vbModal
ElseIf val(CBoBasedON.ListIndex) = 1 Then
FrmBuySearch.DealingForm = PurchaseTransaction
  FrmBuySearch.Index = 11
            FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ›Ê« Ì— «·„‘ —Ì«  "
            FrmBuySearch.show vbModal
  ElseIf val(CBoBasedON.ListIndex) = 2 Then
        Order_no_search2.show
        Order_no_search2.RetrunType = 12
ElseIf val(CBoBasedON.ListIndex) = 3 Then
FrmBuySearch.DealingForm = salespricelist
  FrmBuySearch.Index = 11
            FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ⁄—Ê÷ «·«”⁄«— "
            FrmBuySearch.show vbModal
    
End If
End Sub
Sub RetriveoOrder(Optional TransID As Integer = 0, Optional Transaction_Type As Integer)
Dim StrSQL As String
Dim RsDetails As ADODB.Recordset
Dim Rs1 As ADODB.Recordset
Set RsDetails = New ADODB.Recordset
StrSQL = "SELECT * FROM Transactions WHERE Transaction_ID=" & TransID & " "
    Set Rs1 = New ADODB.Recordset
    Rs1.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
    TxtBillComment.Text = IIf(IsNull(Rs1("TransactionComment")), "", (Rs1("TransactionComment").value))
    DCboStoreName.BoundText = IIf(IsNull(Rs1("StoreID")), 0, (Rs1("StoreID").value))
    End If
   FG.Clear flexClearScrollable, flexClearEverything
    FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
    StrSQL = " SELECT   dbo.Transaction_Details.unitid as itmemunitid,   dbo.TblItems.HaveSerial AS Expr1, *"
    StrSQL = StrSQL + "  FROM         dbo.TblItems INNER JOIN"
    StrSQL = StrSQL + "                   dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID INNER JOIN"
    StrSQL = StrSQL + "                  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL + "                   dbo.TblProcessDEF ON dbo.Transaction_Details.Oper_ID = dbo.TblProcessDEF.TblProcessDEFID LEFT OUTER JOIN"
     StrSQL = StrSQL + "                  dbo.projects_des ON dbo.Transaction_Details.Pand_ID = dbo.projects_des.oprid LEFT OUTER JOIN"
     StrSQL = StrSQL + "                  dbo.projects ON dbo.Transaction_Details.project_ID1 = dbo.projects.id"

    StrSQL = StrSQL + " where Transaction_ID=" & TransID

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
        ''//
         FG.TextMatrix(Num, FG.ColIndex("projectid")) = IIf(IsNull(RsDetails("project_ID1")), "", (RsDetails("project_ID1").value))
         FG.TextMatrix(Num, FG.ColIndex("pandid")) = IIf(IsNull(RsDetails("Pand_ID")), "", (RsDetails("Pand_ID").value))
         FG.TextMatrix(Num, FG.ColIndex("operaid")) = IIf(IsNull(RsDetails("Oper_ID")), "", (RsDetails("Oper_ID").value))
         FG.TextMatrix(Num, FG.ColIndex("pand")) = IIf(IsNull(RsDetails("des")), "", (RsDetails("des").value))
         If SystemOptions.UserInterface = ArabicInterface Then
         FG.TextMatrix(Num, FG.ColIndex("project")) = IIf(IsNull(RsDetails("Project_name")), "", (RsDetails("Project_name").value))
         FG.TextMatrix(Num, FG.ColIndex("opera")) = IIf(IsNull(RsDetails("ProcessName")), "", (RsDetails("ProcessName").value))
         Else
         FG.TextMatrix(Num, FG.ColIndex("project")) = IIf(IsNull(RsDetails("Project_nameE")), "", (RsDetails("Project_nameE").value))
         FG.TextMatrix(Num, FG.ColIndex("opera")) = IIf(IsNull(RsDetails("ProcessNameE")), "", (RsDetails("ProcessNameE").value))
         End If
        ''//
    
        
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))
           ' FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice")), "", (RsDetails("showPrice").value))
        
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
        
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("itmemunitid")), "", (RsDetails("itmemunitid").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        

  FG.TextMatrix(Num, FG.ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(Num, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Me.XPTxtBillID), val(FG.cell(flexcpData, Num, FG.ColIndex("UnitID"))))
        



'        FG.TextMatrix(Num, FG.ColIndex("RequestLimit")) = IIf(IsNull(RsDetails("RequestLimit")), 0, (RsDetails("RequestLimit").value))
'        FG.TextMatrix(Num, FG.ColIndex("LastPurchaseDate")) = IIf(IsNull(RsDetails("LastPurchaseDate")), "", (RsDetails("LastPurchaseDate").value))
'        FG.TextMatrix(Num, FG.ColIndex("LastPurchasePrice")) = IIf(IsNull(RsDetails("LastPurchasePrice")), 0, (RsDetails("LastPurchasePrice").value))
'        FG.TextMatrix(Num, FG.ColIndex("LastPurchaseqty")) = IIf(IsNull(RsDetails("LastPurchaseqty")), 0, (RsDetails("LastPurchaseqty").value))
'        FG.TextMatrix(Num, FG.ColIndex("AverageIssue")) = IIf(IsNull(RsDetails("AverageIssue")), 0, (RsDetails("AverageIssue").value))
        
            RsDetails.MoveNext
            Debug.Print Num

            If FG.rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num

    End If
End Sub

Private Sub ISButton4_Click()
If Me.dcproject1.Text = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„‘—Ê⁄"
Else
MsgBox "Please Select Project"
End If
dcproject1.SetFocus
Exit Sub
End If
If DCboItemsName.Text = "" Or val(DCboItemsName.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·’‰› "
Else
MsgBox "Please Select Item"
End If
DCboItemsName.SetFocus
Exit Sub
End If
If val(TxtQuantity.Text) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ  ÕœÌœ «·ﬂ„Ì… "
Else
MsgBox "Please Eneter Qty"
End If
TxtQuantity.SetFocus
Exit Sub
End If
    Dim current_row As Integer

    If FG.rows = 1 Then
        FG.rows = FG.rows + 1
        current_row = 1
    Else
        
        FG.rows = FG.rows + 1
        current_row = FG.rows - 1
    End If

    With FG
    .TextMatrix(current_row, .ColIndex("operaid")) = DcbProcess1.BoundText
     .TextMatrix(current_row, .ColIndex("pandid")) = Me.dcopr2.BoundText
    .TextMatrix(current_row, .ColIndex("projectid")) = project_id1
    .TextMatrix(current_row, .ColIndex("project")) = Me.dcproject1.Text
        .TextMatrix(current_row, .ColIndex("pand")) = Me.dcopr2.Text
    .TextMatrix(current_row, .ColIndex("opera")) = DcbProcess1.Text
    End With
 NewGrid.CmdAddData_Click
End Sub

Private Sub ItemMakingQty_Change()

    If val(ItemMakingQty.Text) <> 0 Then ItemMakingCost = val(XPTxtSum.Text) / val(ItemMakingQty.Text)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
Dim ID As Double
Dim materilaAcc As String
If Text5.Text <> "" Then
GetCodeIDProject ID, Text5.Text, 1, materilaAcc
Me.dcproject1.BoundText = materilaAcc
End If
    On Error Resume Next

    If dcproject1.BoundText <> "" Then
       
        'project_id1 = get_project_id(dcproject1.BoundText, "Material_account")
        project_id1 = val(dcproject1.BoundText)
        fillterms1 (project_id1)
    End If
    
End Sub

Private Sub Timer1_Timer()
 
TXTOverProject.Text = "«·”‰œ „ ŒÿÌ"
    If TXTOverProject.ForeColor = &HFF& Then
        TXTOverProject.ForeColor = &HC0FFFF
    Else
       TXTOverProject.ForeColor = &HFF&
    End If
    
    
End Sub

Private Sub Txt_order_no_Change()
If Me.TxtModFlg.Text <> "R" Then
RetriveoOrder val(TXT_order_no.Text)
End If
End Sub

Private Sub TxtFillData_Change()

    If TxtFillData.Text = "F" Then
        NewGrid.Calculate 1
    End If

End Sub

Private Sub XPBtnAdd_Click()

    If FG.TextMatrix(FG.rows - 1, FG.ColIndex("Code")) <> "" Then
        FG.rows = FG.rows + 1
        NewGrid.GridDefaultValue FG.rows - 1
        FG.Row = FG.rows - 1
        FG.Col = FG.ColIndex("Code")
        FG.ShowCell FG.rows - 1, FG.ColIndex("Code")
        FG.SetFocus
    End If

End Sub
Private Sub TxtNoteSerial1_Change()
If val(CBoBasedON.ListIndex) = 2 Then
'RetriveOrderProduct TxtNoteSerial1.Text
    RetriveOrderProd TxtNoteSerial1.Text, 2
ElseIf val(CBoBasedON.ListIndex) = 3 Then
    RetriveOrderProd TxtNoteSerial1.Text, 3
End If
End Sub

Private Sub XPBtnMove_Click(Index As Integer)
'    On Error GoTo ErrTrap

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

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            Cmd_Click (0)
        Else
            Sendkeys "{TAB}"
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
        If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
            XPBtnAdd_Click
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
            XPBtnRemove_Click
        End If
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

Private Sub Form_Load()
    Dim StrSQL As String
    Dim My_SQL As String
    Dim Num As Integer
    Dim StrList As String
    Dim BGround As New ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
   ' On Error GoTo ErrTrap

    My_SQL = "  select id,name from warsha"

    fill_combo Dcwarsah, My_SQL

    My_SQL = "    select oprid,des from dbo.projects_des"

    fill_combo Me.dcopr, My_SQL


    fill_combo Me.dcopr2, My_SQL
    If SystemOptions.UserInterface = ArabicInterface Then
    With CBoBasedON
    .Clear
    .AddItem "ÿ·» œ«Œ·Ì"
    .AddItem "›« Ê—… „‘ —Ì« "
    .AddItem "«„— «‰ «Ã"
    .AddItem "⁄—÷ ”⁄—"
    End With
    Else
    With CBoBasedON
    .Clear
    .AddItem "Internal Order"
    .AddItem "Purchase Invoice"
    .AddItem "Production Order"
    CBoBasedON.AddItem "Offer price"
    End With
    End If


 Dim s As String
        Dim rsDummyAcc As New ADODB.Recordset
        
        s = "Select A14 from branches "
        rsDummyAcc.Open s, Cn, adOpenStatic, adLockReadOnly
        DcboCreditSide.BoundText = rsDummyAcc!A14 & ""
        materilaAcc2 = rsDummyAcc!A14 & ""
        rsDummyAcc.Close


    'My_SQL = "  select Account_code,Account_Name from ACCOUNTS  where last_account=1 and Account_code like'a3a17a%'"
 If SystemOptions.UserInterface = ArabicInterface Then
    My_SQL = " select id,Project_name from projects where Project_name<>N'""' And Not (Project_name Is Null)"
Else
My_SQL = " select id,Project_nameE from projects where Project_nameE<>N'""' And Not (Project_nameE Is Null)"
End If



   fill_combo dcproject, My_SQL
    fill_combo dcproject1, My_SQL



    My_SQL = " select Emp_ID,Emp_Name from TblEmployee"
    ' Project_name,expanses_account from projects
    fill_combo DcEmp, My_SQL

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture

    With Me.CboType
        .Clear
        .AddItem " ·›Ì« "
        .AddItem "„”ÕÊ»«  ‘Œ’Ì…"
        '   .AddItem "„”ÕÊ»«  ·“Ê„ «·⁄„· ›Ï «·‘—ﬂ…"
        .AddItem "’—› ··„‘«—Ì⁄ Ê «·Ê—‘"
        .AddItem "Âœ«Ì« Ê⁄Ì‰« "
    End With

    Set NewGrid.Grid = FG
    NewGrid.GridTrans = Destruction
    Set NewGrid.TxtModFlag = TxtModFlg
    Set NewGrid.txtTotal = XPTxtSum
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.DtpBillDate = Me.XPDtbBill
    Set NewGrid.StoreName = Me.DCboStoreName
    Set NewGrid.GrdTBar = Me.TBar
    Set NewGrid.LblItemsCount = Me.LblItemsCount
    Set NewGrid.TxtInvID = Me.XPTxtBillID
    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    Set NewGrid.DCboItemName = DCboItemsName

    Set NewGrid.DCboItemCode = DCboItemsCode
    Set NewGrid.CboItemCase = CboItemCase
    Set NewGrid.CmdAddData = CmdAdd
    Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.TxtQuantity = TxtQuantity
    Set NewGrid.TxtPrice = TxtPrice
    Set NewGrid.LblTotalQty = Me.LblTotalQty
    NewGrid.FillGrid
    FG.WallPaper = BGround.Picture
    AddTip
    SetDtpickerDate XPDtbBill
    Set Dcombos = New ClsDataCombos
     
    Dcombos.GetStores Me.DCboStoreName
    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DCboStoreName
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetAccountingCodes Me.DcboDebitSide
    Dcombos.GetAccountingCodes Me.DcboCreditSide
    Dcombos.GetWorkOrders Me.DcboWorkOrders
    Dcombos.GetItemsNames Me.DcboItems, 0, 1
    Dcombos.GetItemsMaking Me.DcboItemMaking, 0, 1

Dcombos.GetProcessOfProjedt Me.DcbProcess
Dcombos.GetProcessOfProjedt Me.DcbProcess1
Dcombos.GetBranches Me.dcBranch
Dim cl
    StrSQL = "SELECT * FROM Transactions WHERE (Transaction_Type=8 or Transaction_Type=990 OR Transaction_Type=17 OR Transaction_Type=18) "
StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"
StrSQL = StrSQL & " Order By Transaction_ID "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
  
  
  Me.TxtModFlg.Text = "R"
    XPBtnMove_Click 2
    
    'Resize_Form Me, TransactionSize

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
    If SystemOptions.SysAppAccoutingType = SimpleAccoutning Then
        Me.Ele(4).Visible = False
    End If


    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
ChAuto.Caption = "Auto"
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
Label3.Caption = "Branch"
Label4.Caption = "Manual No. "
    Me.Caption = "Projects Issue Vouchers"
    Label1.Caption = "Remarks"
    Cmd(9).Caption = "Print Ge"
    lbl(18).Caption = "Project"
    lbl(17).Caption = "Item"
     lbl(19).Caption = "Operation"
     
     
    C1Elastic6.Caption = Me.Caption
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"
    lbl(3).Caption = "Total"
    lbl(1).Caption = "BY"
    lbl(0).Caption = "Current rec"
    lbl(2).Caption = "Rec count"
    lbl(31).Caption = "Item Code"
    lbl(30).Caption = "Item Name"
    lbl(29).Caption = "Status"
    lbl(28).Caption = "Serial"
    lbl(27).Caption = "Qty"
    lbl(26).Caption = "Price"
    lbl(12).Caption = "Composit Item"
    lbl(11).Caption = "work order"
    lbl(13).Caption = "Workshop"
    lbl(14).Caption = "Project"
    lbl(4).Caption = "Inventory"
    lbl(6).Caption = "Date"
    lbl(5).Caption = "OPR#"
    lbl(7).Caption = "Type"
    lbl’‰›„Ã„⁄(13).Caption = "Composit Item"
    lbl«·ﬂ„Ì…(0).Caption = "QTY"
    lbl«· ﬂ·›….Caption = "Cost"
    Me.CboType.Clear

    With Me.CboType
        .Clear
        .AddItem "Damage"
        .AddItem "personal withdrawal"
        '   .AddItem "„”ÕÊ»«  ·“Ê„ «·⁄„· ›Ï «·‘—ﬂ…"
        .AddItem "Projects"
    End With

    Ele(5).Caption = "Based ON"

Label2.Caption = "Operation"
'lbl(19).Caption = "Operation"


    With FG
        .TextMatrix(0, .ColIndex("Code")) = "Item Code"
        .TextMatrix(0, .ColIndex("Name")) = " Item Name"
        .TextMatrix(0, .ColIndex("ItemCase")) = "Item Case "
        .TextMatrix(0, .ColIndex("HaveSerial")) = "Have Serial"
        .TextMatrix(0, .ColIndex("count")) = "Item Qty"
        .TextMatrix(0, .ColIndex("Price")) = "Price"
'        .TextMatrix(0, .ColIndex("opr_fullcode")) = "Operation"

.TextMatrix(0, .ColIndex("pand")) = "Item"
.TextMatrix(0, .ColIndex("project")) = "Project"
.TextMatrix(0, .ColIndex("opera")) = "Operation"

    End With

    CboItemCase.Clear
    CboItemCase.AddItem "New"
    CboItemCase.AddItem "Used"

    lbl(14).Caption = "Project"
    lbl(15).Caption = "GL#"

    lbl(16).Caption = "Operation"
    lbl(17).Caption = "Operation"
    
    Accredit.Caption = "Send For Approval"
    Me.C1Tab1.TabCaption(0) = "Basic Data"
    Me.C1Tab1.TabCaption(1) = "Approval Status"
    Label11.Caption = "Approval Requested By"
    
    With Grid2
        .TextMatrix(0, .ColIndex("Approved")) = "Approved"
        .TextMatrix(0, .ColIndex("levelName")) = "Level"
        .TextMatrix(0, .ColIndex("EmpName")) = "Employee"
        .TextMatrix(0, .ColIndex("ApprovDate")) = "Approve Date"
        .TextMatrix(0, .ColIndex("Remarks")) = "Notes"
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
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
    Set SaleReport = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap
    Dim RsTest As ADODB.Recordset
    Dim StrSQL As String

    Select Case Me.TxtModFlg.Text

        Case "R"
            '     Me.Caption = "≈–‰ ’—› »÷«⁄…"
            Me.CboType.Enabled = False
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
        
         '   Me.XPDtbBill.Enabled = False
           ' Me.DCboStoreName.locked = True
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

'            Ele(2).Enabled = False

        Case "N"
            '     Me.Caption = "≈–‰ ’—› »÷«⁄…( ÃœÌœ )"
            Me.CboType.Enabled = True
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
            '   Me.XPBtnMove(0).Enabled = False
            '   Me.XPBtnMove(1).Enabled = False
            '   Me.XPBtnMove(2).Enabled = False
            '   Me.XPBtnMove(3).Enabled = False
   
            FG.Enabled = True
            FG.rows = 1
            Me.XPDtbBill.Enabled = True
            XPDtbBill.value = Date
            Me.DCboStoreName.locked = False
        
            FG.Editable = flexEDKbdMouse
            Ele(2).Enabled = True
            CboItemCase.ListIndex = 0

        Case "E"
            '     Me.Caption = "≈–‰ ’—› »÷«⁄…(  ⁄œÌ· )"
            Me.CboType.Enabled = True
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
            XPBtnAdd.Enabled = True
            XPBtnRemove.Enabled = True
            XPBtnRemove.Enabled = True
            FG.Enabled = True
            Me.XPDtbBill.Enabled = True
            Me.DCboStoreName.locked = False
            Ele(2).Enabled = True
            FG.Editable = flexEDKbdMouse
        
    End Select

    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0, Optional NoteID As Long = 0)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim RsTest  As ADODB.Recordset
    Dim LngNoteID As Long
    Dim RsDev As ADODB.Recordset
    Dim Num As Long

   ' On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

If IsNull(rs("OverProject").value) Then
 TXTOverProject.Visible = False
 
 Else
 If rs("OverProject").value = 1 Then
 TXTOverProject.Visible = True
 Else
 TXTOverProject.Visible = False
 End If
 
 
End If

     

 If NoteID <> 0 Then
        rs.Find "NoteId=" & NoteID, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
        GoTo ll
    End If
    
    
    
    If Lngid <> 0 Then
        rs.Find "Transaction_ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If
    
ll:
 If rs.EOF Or rs.BOF Then
        Exit Sub
    End If
    TxtFillData.Text = "T"
    Screen.MousePointer = vbArrowHourglass
    XPTxtBillID.Text = IIf(IsNull(rs("Transaction_ID").value), "", val(rs("Transaction_ID").value))
       DcbProcess.BoundText = IIf(IsNull(rs("Tms_Oper_ID").value), "", rs("Tms_Oper_ID").value)
    dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
     TxtBillComment.Text = IIf(IsNull(rs("TransactionComment").value), "", (rs("TransactionComment").value))
    TxtTransSerial.Text = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
    XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", (rs("Transaction_Date").value))
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    Me.TxtNoteSerial2.Text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
    Me.Dcwarsah.Text = IIf(IsNull(rs("warsha").value), "", rs("warsha").value)
    TxtManualNO.Text = IIf(IsNull(rs("ManualNO").value), "", (rs("ManualNO").value))
       TXTNoteID = val(rs!NoteID & "")
''//
Me.TXT_order_no.Text = IIf(IsNull(rs("OrderID").value), "", rs("OrderID").value)
Me.TxtNoteSerial1.Text = IIf(IsNull(rs("NoteSerial2").value), "", rs("NoteSerial2").value)
''//
    Me.dcproject = IIf(IsNull(rs("project").value), "", rs("project").value)

    Me.dcopr.BoundText = IIf(IsNull(rs("opr_fullcode").value), "", rs("opr_fullcode").value)
    Me.DcEmp.BoundText = IIf(IsNull(rs("empid").value), "", rs("empid").value)

    FG.Clear flexClearScrollable, flexClearEverything
    Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)

    If rs("Transaction_Type").value = 8 Then
        Me.CboType.ListIndex = 0
    ElseIf rs("Transaction_Type").value = 17 Then
        Me.CboType.ListIndex = 1
    ElseIf rs("Transaction_Type").value = 18 Then
        Me.CboType.ListIndex = 2
    End If
    CBoBasedON.ListIndex = IIf(IsNull(rs("basedOn").value), 0, rs("basedOn").value)
    If Not (IsNull(rs("WorkOrderID").value)) Then
        Me.DcboWorkOrders.BoundText = rs("WorkOrderID").value
    Else
        Me.DcboWorkOrders.BoundText = ""
    End If

    Me.DcboItemMaking.BoundText = IIf(IsNull(rs("ItemMaking").value), "", rs("ItemMaking").value)
           
    Me.ItemMakingQty.Text = IIf(IsNull(rs("ItemMakingQty").value), "", rs("ItemMakingQty").value) ' rs("ItemMakingQty").value
    Me.ItemMakingCost.Text = IIf(IsNull(rs("ItemMakingCost").value), "", rs("ItemMakingCost").value) ' rs("ItemMakingCost").value

    FG.rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh

  StrSQL = " SELECT    Transaction_Details.OverProject ,dbo.Transaction_Details.unitid as itemunitid , dbo.TblItems.HaveSerial,  dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transaction_Details.project_ID1 AS project_ID11, dbo.projects.Project_name, "
  StrSQL = StrSQL + "                    dbo.Transaction_Details.Pand_ID AS Pand_ID1, dbo.projects_des.des, dbo.Transaction_Details.Oper_ID AS Oper_ID1, dbo.TblProcessDEF.ProcessName,"
  StrSQL = StrSQL + "                    dbo.TblProcessDEF.ProcessNameE, dbo.Transaction_Details.*"
  StrSQL = StrSQL + " FROM         dbo.TblProcessDEF RIGHT OUTER JOIN"
  StrSQL = StrSQL + "                    dbo.Transaction_Details ON dbo.TblProcessDEF.TblProcessDEFID = dbo.Transaction_Details.Oper_ID LEFT OUTER JOIN"
  StrSQL = StrSQL + "                    dbo.projects_des ON dbo.Transaction_Details.Pand_ID = dbo.projects_des.oprid and dbo.projects_des.oprid <> 0 LEFT OUTER JOIN"
  StrSQL = StrSQL + "                    dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
  StrSQL = StrSQL + "                    dbo.projects ON dbo.Transaction_Details.project_ID1 = dbo.projects.id LEFT OUTER JOIN"
  StrSQL = StrSQL + "                    dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
  StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)
  StrSQL = StrSQL + " order by Transaction_Details.id "
    Set RsDetails = New ADODB.Recordset
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
        
        
        FG.TextMatrix(Num, FG.ColIndex("OverProject")) = IIf(IsNull(RsDetails("OverProject")), "", (RsDetails("OverProject").value))
        
        
                If val(FG.TextMatrix(Num, FG.ColIndex("OverProject"))) = 1 Then
                FG.cell(flexcpBackColor, Num, 1, Num, FG.Cols - 1) = &H8080FF
            Else
                    FG.cell(flexcpBackColor, Num, 1, Num, FG.Cols - 1) = vbWhite
            End If


         FG.TextMatrix(Num, FG.ColIndex("projectid")) = IIf(IsNull(RsDetails("project_ID11")), "", (RsDetails("project_ID11").value))
            FG.TextMatrix(Num, FG.ColIndex("project")) = IIf(IsNull(RsDetails("Project_name")), "", Trim(RsDetails("Project_name").value))
            FG.TextMatrix(Num, FG.ColIndex("pandid")) = IIf(IsNull(RsDetails("Pand_ID1")), "", Trim(RsDetails("Pand_ID1").value))
             FG.TextMatrix(Num, FG.ColIndex("pand")) = IIf(IsNull(RsDetails("des")), "", (RsDetails("des").value))
            FG.TextMatrix(Num, FG.ColIndex("operaid")) = IIf(IsNull(RsDetails("Oper_ID1")), "", Trim(RsDetails("Oper_ID1").value))
            
            
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial")), "", Trim(RsDetails("ItemSerial").value))
             If SystemOptions.UserInterface = ArabicInterface Then
             FG.TextMatrix(Num, FG.ColIndex("opera")) = IIf(IsNull(RsDetails("ProcessName")), "", (RsDetails("ProcessName").value))
             Else
             FG.TextMatrix(Num, FG.ColIndex("opera")) = IIf(IsNull(RsDetails("ProcessNameE")), "", (RsDetails("ProcessNameE").value))
             End If

            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If

            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice")), "", (RsDetails("showPrice").value))
            
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("guaranteeTime")) = IIf(IsNull(RsDetails("guaranteeTime")), "", (RsDetails("guaranteeTime").value))
            FG.TextMatrix(Num, FG.ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks")), "", (RsDetails("Remarks").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 0, (RsDetails("ItemSize").value))
     
            FG.cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("itemunitid")), "", (RsDetails("itemunitid").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
            
'            FG.TextMatrix(Num, FG.ColIndex("opr_fullcode")) = IIf(IsNull(RsDetails("opr_fullcode")), "", (RsDetails("opr_fullcode").value))
        
            RsDetails.MoveNext
        Next Num

    End If

    StrSQL = "Select * From NOTES Where Transaction_ID=" & val(Me.XPTxtBillID.Text)
    Set RsNotes = New ADODB.Recordset
    RsNotes.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsNotes.BOF Or RsNotes.EOF) Then
        Me.TxtNoteSerial.Text = IIf(IsNull(RsNotes("NoteSerial").value), "", (RsNotes("NoteSerial").value))
        LngNoteID = RsNotes("NoteID").value
        StrSQL = "Select * From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & LngNoteID & ""
        StrSQL = StrSQL + " Order BY DEV_ID_Line_No"
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or RsDev.EOF) Then
            RsDev.MoveFirst

            For i = 1 To RsDev.RecordCount

                If RsDev("Credit_Or_Debit").value = 0 Then
                    Me.DcboDebitSide.BoundText = RsDev("Account_Code").value
                ElseIf RsDev("Credit_Or_Debit").value = 1 Then
                    Me.DcboCreditSide.BoundText = RsDev("Account_Code").value
                End If

                RsDev.MoveNext
            Next i

        End If
    End If

    TxtFillData.Text = "F"
    Screen.MousePointer = vbDefault
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    fillapprovData
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Private Sub Undo()
    Dim Msg As String

    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                clear_all Me
                Me.TxtModFlg.Text = "R"
                XPBtnMove_Click (1)
            End If

        Case "E"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ⁄œÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                rs.Find "Transaction_ID='" & val(XPTxtBillID.Text) & "'", , adSearchForward, adBookmarkFirst

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

Private Sub Del_TransAction()
''    On Error GoTo ErrTrap

    If XPTxtBillID.Text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (XPTxtBillID.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
       Else
       Msg = "Confirm Delete"
    End If

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            If Not rs.RecordCount < 1 Then
   Deletepost Me.Name, "Transactions", "Transaction_ID", 0, val(dcBranch.BoundText), val(XPTxtBillID.Text), TxtNoteSerial2.Text
    
          '  CuurentLogdata ("D")
             '   Cn.BeginTrans
               'Cn.Execute "delete from Transactions where Transaction_ID =" & val(XPTxtBillID.Text)
                Cn.Execute "delete from Transaction_Details where Transaction_ID =" & val(XPTxtBillID.Text)
                Cn.Execute "delete from notes where NoteSerial ='" & Me.TxtNoteSerial.Text & "'"
             rs.delete
             '   Cn.CommitTrans
            
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

    Else
        clear_all Me
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
        Msg = "This is process not available .no record"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change

    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then
       ' Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
       If SystemOptions.UserInterface = EnglishInterface Then
       Msg = "Can not delete this is record"
       Else
       Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰« "
       End If
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
        rs.CancelUpdate
    End If

End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hWnd, "≈–‰ ’—› »÷«⁄…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  √’‰«›  «·›…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "≈–‰ ’—› »÷«⁄…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
    End With

    With TTP
        .Create Me.hWnd, "≈–‰ ’—› »÷«⁄…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "≈–‰ ’—› »÷«⁄…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ «·»Ì«‰«  «·Õ«·Ì…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "≈–‰ ’—› »÷«⁄…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·≈÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "≈–‰ ’—› »÷«⁄…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "≈–‰ ’—› »÷«⁄…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄„·Ì… ≈Â·«ﬂ« " & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "≈–‰ ’—› »÷«⁄…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "≈–‰ ’—› »÷«⁄…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "≈–‰ ’—› »÷«⁄…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "≈–‰ ’—› »÷«⁄…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "≈–‰ ’—› »÷«⁄…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "≈–‰ ’—› »÷«⁄…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RowNum As Integer
    Dim RSTransDetails As ADODB.Recordset
    Dim RsNotes As ADODB.Recordset
    Dim RsTemp  As New ADODB.Recordset
    Dim RsTest As New ADODB.Recordset
    Dim RsRepeat As ADODB.Recordset
    Dim RsDetalis As ADODB.Recordset
    Dim StrSQL As String
    Dim StrSqlDel As String
    Dim note_id As Integer
    Dim BeginTrans As Boolean
    Dim LngItemID As Long
    Dim LngNoteID As Long
    Dim LngDev As Long
    Dim LngLineNO As Long
    Dim StrDes As String
    Dim project_id As Integer
    'On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass
     Dim Posted As Integer
            If CheckAprroveScreen(Me.Name) = True Then
            Posted = 1
            Else
            Posted = 0
            End If
            
    If Me.TxtModFlg.Text <> "R" Then
        If DCboStoreName.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ «·„Œ“‰"
            Else
            Msg = "Please Select Store"
            End If
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCboStoreName.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If XPDtbBill.value = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ  «—ÌŒ  ”ÃÌ· Â–Â «·⁄„·Ì…"
            Else
            Msg = "Please Select Date"
            End If
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPDtbBill.SetFocus
            Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If Me.CboType.ListIndex = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄ ≈–‰ «·’—›"
           Else
           Msg = "Please Select Type"
           End If
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
             CboType.SetFocus
           Sendkeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If Me.CboType.ListIndex = 2 Then
    
            If dcproject.BoundText = "" Then
             '   Msg = "ÌÃ» ≈Œ Ì«—   «·„‘—Ê⁄ «·„‰’—› ·Â «·»÷«⁄… ...!!!"
             '   MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
             '   DCPROJECT.SetFocus
             '  SendKeys "{F4}"
             '   Screen.MousePointer = vbDefault
             '   Exit Sub
        
            End If

            '   If Val(Me.DcboWorkOrders.BoundText) = 0 Then
            '       Msg = "ÌÃ» ≈Œ Ì«— √„— «·‘€· «·„‰’—›… ·Â «·»÷«⁄… ...!!!"
            '       MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '       DcboWorkOrders.SetFocus
            '       SendKeys "{F4}"
            '       Screen.MousePointer = vbDefault
            '       Exit Sub
            '   End If
            '   If Val(Me.DcboWorkOrders.BoundText) <> 0 Then
            '       If CheckOrderState(Val(Me.DcboWorkOrders.BoundText)) = False Then
            '           Msg = "·«Ì„ﬂ‰ «·Õ›Ÿ .. ›ﬁœ  „ ≈€·«ﬁ «„— «·‘€·...!!!"
            '           MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '           Exit Sub
            '       End If
            '   End If
        End If
    
        '     If Me.CboType.ListIndex = 2 Then
        '     If Val(Me.DcboWorkOrders.BoundText) = 0 Then
        '         Msg = "ÌÃ» ≈Œ Ì«— √„— «·‘€· «·„‰’—›… ·Â «·»÷«⁄… ...!!!"
        '         MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '         DcboWorkOrders.SetFocus
        '         SendKeys "{F4}"
        '         Screen.MousePointer = vbDefault
        '         Exit Sub
        '     End If
  
        ' End If
    
        'Check the Items Grid
    
    'check all have projects
    
    '**************************
    If 1 = 1 Then
            For RowNum = 1 To FG.rows - 1

      If FG.TextMatrix(RowNum, FG.ColIndex("projectid")) = "" Or val(FG.TextMatrix(RowNum, FG.ColIndex("projectid"))) = 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "·„ Ì „ «Œ Ì«— „‘—Ê⁄ ›Ì «·”ÿ— —ﬁ„ " & RowNum & "Ê·« Ì„ﬂ‰ «·Õ›Ÿ", vbCritical
                    Else
                            MsgBox "Please select Project in row " & RowNum & "Can not Save", vbCritical
                    End If
         Exit Sub
         End If
            
        Next RowNum
        
        
        
     End If
     
     
     
         For RowNum = 1 To FG.rows - 1
     
                      If FG.TextMatrix(RowNum, FG.ColIndex("projectid")) <> "" And val(FG.TextMatrix(RowNum, FG.ColIndex("projectid"))) <> 0 Then
                                       If FG.TextMatrix(RowNum, FG.ColIndex("pandid")) = "" Or val(FG.TextMatrix(RowNum, FG.ColIndex("pandid"))) = 0 Then
                                            
                                                         If SystemOptions.UserInterface = ArabicInterface Then
                                                                 MsgBox "·„ Ì „ «Œ Ì«— «·»‰œ  ··„‘—Ê⁄ ›Ì «·”ÿ— —ﬁ„ " & RowNum & "Ê·« Ì„ﬂ‰ «·Õ›Ÿ", vbCritical
                                                         Else
                                                                 MsgBox "Please select Project Pand in row " & RowNum & "Can not Save", vbCritical
                                                         End If
                                                         Exit Sub
                                    End If
                    
                    
                    End If
            
        Next RowNum
        
        
        
        
        
     
         For RowNum = 1 To FG.rows - 1
     
                      If FG.TextMatrix(RowNum, FG.ColIndex("pandid")) <> "" And val(FG.TextMatrix(RowNum, FG.ColIndex("pandid"))) <> 0 Then
                                       If (FG.TextMatrix(RowNum, FG.ColIndex("operaid")) = "" Or val(FG.TextMatrix(RowNum, FG.ColIndex("operaid"))) = 0) And ProjectItemsCheck(val(FG.TextMatrix(RowNum, FG.ColIndex("projectid")))) > 0 Then
                                            
                                                         If SystemOptions.UserInterface = ArabicInterface Then
                                                                 MsgBox "·„ Ì „ «Œ Ì«— «·⁄„·ÌÂ ··»‰œ ›Ì   ··„‘—Ê⁄ ›Ì «·”ÿ— —ﬁ„ " & RowNum & "Ê·« Ì„ﬂ‰ «·Õ›Ÿ", vbCritical
                                                         Else
                                                                 MsgBox "Please select Project Pand in row " & RowNum & "Can not Save", vbCritical
                                                         End If
                                                         Exit Sub
                                                         
                                    End If
                    
                    
                    End If
            
        Next RowNum
        
        
        
                 For RowNum = 1 To FG.rows - 1
     
                      If FG.TextMatrix(RowNum, FG.ColIndex("operaid")) <> "" And val(FG.TextMatrix(RowNum, FG.ColIndex("operaid"))) <> 0 Then
                                       If (FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Or val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))) <> 0) Then
                                            
                                            
                                              If ProjectItemsCheck(val(FG.TextMatrix(RowNum, FG.ColIndex("projectid"))), val(FG.TextMatrix(RowNum, FG.ColIndex("pandid"))), val(FG.TextMatrix(RowNum, FG.ColIndex("operaid"))), val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))) <= 0 Then
                                             
                                                         If SystemOptions.UserInterface = ArabicInterface Then
                                                                 MsgBox "«·’‰› «·„Œ’’ ··⁄„·Ì…   €Ì— „ÊÃÊœ „⁄ ⁄„·ÌÂ «·„‘—Ê⁄ «·”ÿ— —ﬁ„ " & RowNum & "Ê·« Ì„ﬂ‰ «·Õ›Ÿ", vbCritical
                                                         Else
                                                                 MsgBox "Please select Project Pand in row " & RowNum & "Can not Save", vbCritical
                                                         End If
                                                         Exit Sub
                                                End If
                                                
                                    
                                    End If
                    
                    
                    End If
            
        Next RowNum
        
         Dim OPRQTY As Double
     Dim IssuedOprQty As Double
       TXTOverProject.Visible = False
                 For RowNum = 1 To FG.rows - 1
    FG.TextMatrix(RowNum, FG.ColIndex("OverProject")) = 0
    FG.cell(flexcpBackColor, RowNum, 1, RowNum, FG.Cols - 1) = vbWhite
                      If FG.TextMatrix(RowNum, FG.ColIndex("operaid")) <> "" And val(FG.TextMatrix(RowNum, FG.ColIndex("operaid"))) <> 0 Then
                                       If (FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Or val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))) <> 0) Then
                                            
                                            If (FG.TextMatrix(RowNum, FG.ColIndex("Count")) <> "" Or val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))) <> 0) Then
                                            
                                                            
                                                            
                                                              If CheckRemainsetyforprojectopr(val(FG.TextMatrix(RowNum, FG.ColIndex("projectid"))), val(FG.TextMatrix(RowNum, FG.ColIndex("pandid"))), val(FG.TextMatrix(RowNum, FG.ColIndex("operaid"))), val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))), val(Me.XPTxtBillID.Text), val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))), Me.TxtModFlg, OPRQTY, IssuedOprQty) < 0 Then
                                                             TXTOverProject.Visible = True
                                                             Timer1.Enabled = True
                                                             FG.TextMatrix(RowNum, FG.ColIndex("OverProject")) = 1
                                                                      
                                                            FG.cell(flexcpBackColor, RowNum, 1, RowNum, FG.Cols - 1) = &H8080FF
         
                                                                         If SystemOptions.UserInterface = ArabicInterface Then
                                                                                 MsgBox "«·ﬂ„ÌÂ «·„Œ’’Â ·Â–« «·»‰œ ”  ŒÿÌ «·„ Êﬁ⁄ ›Ì «·”ÿ— —ﬁ„ " & RowNum & CHR(13) & "«·’‰›" & (FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "") & CHR(13) & "   ”Ì „ «·Õ›Ÿ Ê·ﬂ‰ »„·«ÕŸ… ·· ⁄„Ìœ " & CHR(13) & "ﬂ„ÌÂ «·»‰œ „‰ «·’‰› " & OPRQTY & CHR(13) & " «·„‰’—› Õ Ì  «—ÌŒÂ " & Abs(IssuedOprQty) & CHR(13) & " «·„—«œ ’—›Â »«·”‰œ  " & val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))) & CHR(13) & " «·„ »ﬁÌ" & OPRQTY - IssuedOprQty
                                                                         Else
                                                                                 MsgBox "Please select Project Pand in row " & RowNum & "Can not Save", vbCritical
                                                                         End If
                                                                       '  Exit Sub
                                                                End If
                                                                
                                                                
                                              End If
                                    
                                    End If
                    
                    
                    End If
            
        Next RowNum
                
        
        
      
     
    '********************
        If NewGrid.CheckDataEntered = False Then
            Exit Sub
        End If

        If NewGrid.Calculate(1) = False Then
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        Cn.BeginTrans
        BeginTrans = True
        Screen.MousePointer = vbArrowHourglass

        If Me.TxtModFlg.Text = "N" Then
            rs.AddNew
                        XPTxtBillID.Text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            TxtTransSerial.Text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "(Transaction_Type=8 OR Transaction_Type=17 OR Transaction_Type=18)"))

            rs("Transaction_ID").value = val(XPTxtBillID.Text)
        ElseIf Me.TxtModFlg.Text = "E" Then
        
        End If

        Set RSTransDetails = New ADODB.Recordset
     '   RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
        Set RsNotes = New ADODB.Recordset
  '      RsNotes.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
           If TxtNoteSerial2.Text = "" Then
              TxtNoteSerial2.Text = Voucher_coding(val(my_branch), XPDtbBill.value, 66, 66)
          End If
          If TXTOverProject.Visible = True Then
          rs("OverProject").value = 1
          Else
          rs("OverProject").value = 0
          
          End If
          
          rs("NoteSerial1").value = IIf(Me.TxtNoteSerial2 <> "", Trim(TxtNoteSerial2.Text), Null)
       rs("TransactionComment").value = IIf(Trim$(TxtBillComment.Text) = "", Null, Trim$(TxtBillComment.Text))
        rs("Transaction_Serial").value = IIf(Trim(Me.TxtTransSerial.Text) = "", "", Trim(Me.TxtTransSerial.Text))
        rs("Transaction_Date").value = XPDtbBill.value
        rs("warsha").value = IIf(Trim(Me.Dcwarsah.Text) = "", "", Trim(Me.Dcwarsah.Text))
        rs("project").value = IIf(Trim(Me.dcproject.Text) = "", "", Trim(Me.dcproject.Text))
        rs("Tms_Oper_ID").value = IIf(Me.DcbProcess.BoundText = "", 0, val(DcbProcess.BoundText))
      rs("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
        If dcproject.BoundText <> "" Then
            'project_id = get_project_id(DCproject.BoundText, "Material_account")
            project_id = val(dcproject.BoundText)
            rs("project_id").value = project_id
        End If
    
        If Me.dcopr.BoundText <> "" Then
            rs("opr_fullcode").value = Me.dcopr.BoundText
    
        End If

        rs("empid").value = IIf(Me.DcEmp.BoundText = "", 0, Me.DcEmp.BoundText)
      If Posted = 1 Then
      rs("Transaction_Type").value = 990
      Else
        If Me.CboType.ListIndex = 0 Then
            rs("Transaction_Type").value = 8
        ElseIf Me.CboType.ListIndex = 1 Then
            rs("Transaction_Type").value = 17
        ElseIf Me.CboType.ListIndex = 2 Then
            rs("Transaction_Type").value = 18
        End If
End If
        rs("ManualNO").value = IIf(TxtManualNO.Text = "", Null, (TxtManualNO.Text))
        rs("UserID").value = user_id
        rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, val(DCboStoreName.BoundText))

        If Me.CboType.ListIndex = 2 And val(Me.DcboWorkOrders.BoundText) <> 0 Then
            rs("WorkOrderID").value = val(Me.DcboWorkOrders.BoundText)
        Else
            rs("WorkOrderID").value = Null
        End If
 rs("OrderID").value = val(Me.TXT_order_no.Text)
        rs("NoteSerial2").value = Me.TxtNoteSerial1.Text
        
        rs("ItemMaking").value = IIf(DcboItemMaking.BoundText = "", Null, val(DcboItemMaking.BoundText))
        rs("ItemMakingQty").value = val(ItemMakingQty.Text)
        rs("basedOn").value = val(CBoBasedON.ListIndex)

        If val(ItemMakingQty.Text) <> 0 Then ItemMakingCost = val(XPTxtSum.Text) / val(ItemMakingQty.Text)
        rs("ItemMakingCost").value = val(ItemMakingCost.Text)
       
        rs.update

        If DcboItemMaking.Text <> "" And val(ItemMakingQty.Text) <> 0 And val(ItemMakingCost.Text) <> 0 Then
            ' «·’‰› «·„’‰⁄
            rs.AddNew
            rs("Transaction_ID").value = val(XPTxtBillID.Text) + 1
            rs("Transaction_Type").value = 3
            rs("Transaction_Serial").value = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=3"))
        
            'XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            ' Me.TxtTransSerial.text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=3"))
            rs("UserID").value = user_id
            rs("Transaction_Date").value = XPDtbBill.value
            rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, val(DCboStoreName.BoundText))
       
            rs.update
        End If

        If Me.TxtModFlg.Text = "E" Then
            StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
        End If

        For RowNum = 1 To FG.rows - 1

            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then

                'Check Repeat Serial
                If FG.TextMatrix(RowNum, FG.ColIndex("Serial")) <> "" Then
                    StrSQL = "select * From Transaction_Details where ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                    StrSQL = StrSQL + " and Transaction_ID =" & XPTxtBillID.Text
                    Set RsTemp = New ADODB.Recordset
                    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (RsTemp.EOF Or RsTemp.BOF) Then
                        Msg = "«·”Ì—Ì«· «·Œ«’ »«·’‰›" & CHR(13)
                        Msg = Msg + FG.cell(flexcpTextDisplay, RowNum, FG.ColIndex("name")) & CHR(13)
                        Msg = Msg + " „ √œŒ«·Â ·ﬁÿ⁄… √Œ—Ï ›Ì Â–Â «·›« Ê—…"
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        RsTemp.Close
                        FG.Row = RowNum
                        FG.Col = FG.ColIndex("name")
                        FG.ShowCell RowNum, FG.ColIndex("name")
                        FG.SetFocus
                        Screen.MousePointer = vbDefault
                        BeginTrans = False
                        Cn.RollbackTrans
                        Exit Sub
                    End If

                    RsTemp.Close
                End If

                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").value = val(XPTxtBillID.Text)
                RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
         
                '   RSTransDetails("Quantity").Value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("Count")) = ""), Null, Val(Fg.TextMatrix(RowNum, Fg.ColIndex("Count"))))
                '            RSTransDetails("ItemName").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Name")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Name"))))
            
                If dcproject.BoundText <> "" Then
        
                    'project_id = get_project_id(DCproject.BoundText, "Material_account")
                    project_id = val(dcproject.BoundText)
                    RSTransDetails("project_id").value = project_id
            
                End If
        
                If Me.dcopr.BoundText <> "" Then
        
'                    RSTransDetails("opr_fullcode").value = Me.dcopr.BoundText
        
                Else
'                    RSTransDetails("opr_fullcode").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("opr_fullcode")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("opr_fullcode")))
                End If

                RSTransDetails("payed").value = 1
                RSTransDetails("Transaction_Date").value = Me.XPDtbBill
            
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
                

          RSTransDetails("OverProject").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("OverProject")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("OverProject"))))
          
RSTransDetails("Oper_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("operaid")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("operaid"))))
RSTransDetails("Pand_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("pandid")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("pandid"))))
RSTransDetails("project_ID1").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("projectid")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("projectid"))))
project_id = IIf(IsNull(RSTransDetails("project_ID1").value), 0, RSTransDetails("project_ID1").value)

                RSTransDetails("showPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
                RSTransDetails("ItemDiscountType").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType"))))
                RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
                RSTransDetails("ItemDiscount").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal"))))
                RSTransDetails("guaranteeTime").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime"))))
            
                RSTransDetails("UnitID").value = IIf(FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
                RSTransDetails("ShowQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
            
                RSTransDetails("Remarks").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Remarks")) = ""), Null, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("Remarks"))))
                RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
                RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), "", Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
            RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
            
                Dim RsUnitData As ADODB.Recordset
                Dim LngCurItemID As Long
                Dim LngUnitID As Long
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
                End If
 RSTransDetails("Price").value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
             
             Dim OldQty As Double
             Dim OldCost As Double
              Dim NewQty As Double
               Dim NewCost As Double
               
getItemCostData XPDtbBill.value, RSTransDetails("Item_ID").value, val(DCboStoreName.BoundText), val(Me.XPTxtBillID.Text), OldQty, OldCost, NewQty, NewCost, , LngUnitID
       RSTransDetails("OldQty").value = NewQty
       RSTransDetails("OldCost").value = NewCost
       
      RSTransDetails("NewQty").value = RSTransDetails("OldQty").value - RSTransDetails("Quantity").value
       RSTransDetails("NewCost").value = RSTransDetails("OldCost").value ' ((RSTransDetails("OldQty").value * RSTransDetails("OldCost").value) + (RSTransDetails("Quantity").value * RSTransDetails("Price").value)) / (RSTransDetails("Quantity").value + RSTransDetails("OldQty").value)
       



                RSTransDetails.update
            End If

        Next RowNum

        If DcboItemMaking.Text <> "" And val(ItemMakingQty.Text) <> 0 And val(ItemMakingCost.Text) <> 0 Then

            ' «·’‰› «·„’‰⁄
            RSTransDetails.AddNew
            RSTransDetails("Transaction_ID").value = val(XPTxtBillID.Text) + 1
            RSTransDetails("Item_ID").value = IIf(DcboItemMaking.BoundText = "", Null, val(DcboItemMaking.BoundText))
            RSTransDetails("ItemCase").value = 1
        
            Set rsItm = New ADODB.Recordset
            itmcode_new = IIf(DcboItemMaking.BoundText = "", Null, val(DcboItemMaking.BoundText))
            rsItm.Open "select ItemCode from TblItems where ItemID=" & itmcode_new, Cn, adOpenStatic, adLockOptimistic, adCmdText
            '        RSTransDetails("ItemSerial").Value = ""
            RSTransDetails("Quantity").value = val(ItemMakingQty.Text)
            RSTransDetails("Price").value = val(ItemMakingCost.Text)
            RSTransDetails("ColorID").value = 1

            If val(ItemMakingQty.Text) <> 0 Then ItemMakingCost = val(XPTxtSum.Text) / val(ItemMakingQty.Text)
            RSTransDetails("ItemSize").value = ""
            RSTransDetails.update

        End If
        
        If val(TXTNoteID) <> 0 Then
            Cn.Execute "delete from notes where NoteID =" & val(TXTNoteID)
        Else
            Cn.Execute "delete from notes where NoteSerial =" & val(Me.TxtNoteSerial.Text)
        End If
        Set RsNotes = New ADODB.Recordset
      '  RsNotes.Open "NOTES", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      
        RsNotes.AddNew
        LngNoteID = new_id("Notes", "NoteID", "")
        rs!NoteID = LngNoteID
        TXTNoteID = LngNoteID
        rs.update
        RsNotes("NoteID").value = LngNoteID
        RsNotes("branch_no").value = val(Me.dcBranch.BoundText)
        RsNotes("NoteDate").value = Me.XPDtbBill.value
        RsNotes("NoteType").value = 100
        ''//
       
        ''//
        
        RsNotes("NoteSerial").value = Me.TxtNoteSerial.Text
        RsNotes("NoteSerial1").value = Me.TxtNoteSerial2.Text
        
        
        RsNotes("numbering_type").value = sand_numbering_type(0) '„”·”· «·ﬁÌœ
        RsNotes("sanad_year").value = year(XPDtbBill.value)
        RsNotes("sanad_month").value = Month(XPDtbBill.value)
        RsNotes("note_value_by_characters").value = WriteNo(Format(Me.XPTxtSum.Text, "0.00"), 0, True, ".")
            
        'RsNotes("NoteSerial").value = new_id("Notes", "NoteSerial", "")
        RsNotes("Note_Value").value = val(Me.XPTxtSum.Text)
        RsNotes("Transaction_ID").value = val(Me.XPTxtBillID.Text)
        RsNotes.update
        LngDev = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
        
'        If DCPROJECT.BoundText <> "" Then
         
     'Dim project_id As Integer
     Dim linevalue As Double
     Dim Material_account As String
     Dim lineno As Integer
     lineno = 1
        For RowNum = 1 To FG.rows - 1

            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
         'project_id = RSTransDetails("project_ID1").value
       project_id = IIf(IsNull(RSTransDetails("project_ID1").value), 0, RSTransDetails("project_ID1").value)
     '     project_id = project_id1
                            If project_id = 0 Then
                            'project_id = get_project_id(DCproject.BoundText, "Material_account")
                            project_id = val(dcproject.BoundText)
                            End If
                            'projectid
               Material_account = get_project_Account(val(FG.TextMatrix(RowNum, FG.ColIndex("projectid"))), "Material_account")
               If Material_account = "" Then
            Material_account = get_project_Account(project_id, "Material_account")
            Material_account = get_project_Account(val(FG.TextMatrix(RowNum, FG.ColIndex("projectid"))), "AccountUnderImp")
            
            End If
               If Material_account = "" Then
                Material_account = materilaAcc2
               End If
         Dim pandid As Integer
          Dim operaID As Integer
          
     
pandid = val((FG.TextMatrix(RowNum, FG.ColIndex("pandid"))))
operaID = val((FG.TextMatrix(RowNum, FG.ColIndex("operaid"))))
 
linevalue = Round((FG.TextMatrix(RowNum, FG.ColIndex("Valu"))), 2)

Dim des As String
   des = "»‰«¡ ⁄·Ï ”‰œ ’—› „Ê«œ „‘«—Ì⁄ —ﬁ„"
   des = des & TxtNoteSerial2.Text
   des = des & "«·—ﬁ„ «·ÌœÊÌ"
   des = des & TxtManualNO.Text
                        If ModAccounts.AddNewDev(LngDev, lineno, Material_account, linevalue, 0, des, LngNoteID, , , CInt(SystemOptions.SysCurrentAccountIntervalID), XPDtbBill.value, , val(XPTxtBillID.Text), , , , , , , , , , project_id, , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , pandid, operaID, , , , , , , , , Posted) = False Then
                            GoTo ErrTrap
                        End If
                        lineno = lineno + 1
            End If
            
        Next RowNum
            lineno = lineno + 1


          
            If ModAccounts.AddNewDev(LngDev, lineno, Me.DcboCreditSide.BoundText, val(Me.XPTxtSum.Text), 1, des, LngNoteID, , , CInt(SystemOptions.SysCurrentAccountIntervalID), XPDtbBill.value, , val(XPTxtBillID.Text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText), , , , , , , , , , , , , , , , , , , , , , , , , Posted) = False Then
                GoTo ErrTrap
            End If
        
'        Else

       '     If ModAccounts.AddNewDev(LngDev, 1, Me.DcboDebitSide.BoundText, val(Me.XPTxtSum.text), 0, "", LngNoteID, , , CInt(SystemOptions.SysCurrentAccountIntervalID), Date, , , , , , , , , , , , , , , , , , , val(Me.DcBranch.BoundText)) = False Then
       '         GoTo ErrTrap
       '     End If
'
'            If ModAccounts.AddNewDev(LngDev, 2, Me.DcboCreditSide.BoundText, val(Me.XPTxtSum.text), 1, "", LngNoteID, , , CInt(SystemOptions.SysCurrentAccountIntervalID), Date, , , , , , , , , , , , , , , , , , , val(Me.DcBranch.BoundText)) = False Then
'                GoTo ErrTrap
'            End If
'        End If
        
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount

        Select Case Me.TxtModFlg.Text

            Case "N"
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
              Else
                    Msg = "Operation data was saved " & CHR(13)
                    Msg = Msg + "need another operation""Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
              
             End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                    Cmd_Click (0)
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            
            Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
               Else
                MsgBox "Update Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
               End If
        End Select

        TxtModFlg.Text = "R"
    End If
fillapprovData
    Screen.MousePointer = vbDefault

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    Screen.MousePointer = vbDefault
    If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
    Msg = "Sorry...error douring save data " & CHR(13)
    End If
    Msg = Msg & CHR(13) & Err.Description
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & Err.Source
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub XPBtnRemove_Click()
    On Error GoTo ErrTrap

    If FG.rows > 1 Then
        If FG.rows = 2 Then
            FG.Clear flexClearScrollable, flexClearEverything
            NewGrid.Calculate 1, True
        Else

            If FG.rows > 1 Then
                If FG.Row <> FG.FixedRows - 1 Then
                    FG.RemoveItem (FG.Row)
                End If
            End If

            NewGrid.Calculate 1
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub PrintReport()
    On Error GoTo ErrTrap
    Dim ShowType As Integer
    'Dim clrep As ClsReportProp
    Dim StrPath As String
    Dim Msg As String

    If XPTxtBillID.Text <> "" Then
        Set SaleReport = New ClsSaleReport
        SaleReport.DestructionReport XPTxtBillID.Text
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
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
    NewGrid.Calculate 1
End Sub

Private Sub WriteDev()

    If TxtModFlg.Text = "R" Then Exit Sub
    If DCboStoreName.Text = "" Then Exit Sub

    On Error Resume Next
    Me.DcboCreditSide.BoundText = ""
    Me.DcboDebitSide.BoundText = ""
 
    Dim Account_Code_dynamic As String

    Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")

    If Account_Code_dynamic = "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
        Else
        MsgBox "Please Select Account     ", vbCritical
        End If
        
        Exit Sub
    End If

    Me.DcboCreditSide.BoundText = Account_Code_dynamic '  Õ”«» «·„Œ“Ê‰
    'Me.DcboCreditSide.BoundText = "a1a2a5" '  Õ”«» «·„Œ“Ê‰

    If Me.CboType.ListIndex = 0 Then ' ·›Ì« 

        Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code1")

        If Account_Code_dynamic = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»     '›—Êﬁ Ê ·›Ì«  „Œ“Ê‰   ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                    Else
        MsgBox "Please Select Account     ", vbCritical
        End If
        
            Exit Sub
        End If
        
        Me.DcboDebitSide.BoundText = Account_Code_dynamic '›—Êﬁ Ê ·›Ì«  „Œ“Ê‰
        'Me.DcboDebitSide.BoundText = "a3a7"    '›—Êﬁ Ê ·›Ì«  „Œ“Ê‰
    ElseIf Me.CboType.ListIndex = 1 Then '„”ÕÊ»«  ‘Œ’Ì…
       
        Account_Code_dynamic = get_EMPLOYEE_Account(DcEmp.BoundText, "Account_Code")

        If Account_Code_dynamic = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«» –„„ ·Â–« «·„ÊŸﬁ ·Â–… «·⁄„·Ì…", vbCritical
                    Else
        MsgBox "Please Select Account     ", vbCritical
        End If
        
            Exit Sub
        End If
            
        Me.DcboDebitSide.BoundText = Account_Code_dynamic ' Õ”«»   –„„ «·„ÊŸ›
        'Me.DcboDebitSide.BoundText = "a2a1a2" ' Õ”«» Ê”Ìÿ «›  «ÕÌ
    ElseIf Me.CboType.ListIndex = 2 Then '«·„‘—Ê⁄
            
        If dcproject.Text <> "" Then
            Me.DcboDebitSide.BoundText = dcproject.BoundText
        Else
         
            Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code2")

            If Account_Code_dynamic = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»       ”ÊÌ«  Ã—œÌ… ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                        Else
        MsgBox "Please Select Account     ", vbCritical
        End If
        
                Exit Sub
            End If

            Me.DcboDebitSide.BoundText = Account_Code_dynamic ' ”ÊÌ«  Ã—œÌ…
            ' Me.DcboDebitSide.BoundText = "a3a8" '›—Êﬁ«  Ê“Ì«œ… ›Ì «·„Œ“Ê‰
        End If
    
    ElseIf Me.CboType.ListIndex = 3 Then

        Account_Code_dynamic = get_account_code_branch(17, my_branch)
        
        If Account_Code_dynamic = "NO branch" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
           Else
            MsgBox " Please Create Branch", vbCritical
           End If
            Exit Sub
        Else

            If Account_Code_dynamic = "NO account" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»    Âœ«Ì« Ê⁄Ì‰«  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        Else
        MsgBox "Please Select Account     ", vbCritical
        End If
        
                Exit Sub
         
            End If
        End If

        Me.DcboDebitSide.BoundText = Account_Code_dynamic 'Âœ«Ì« Ê⁄Ì‰« 
      
    End If

End Sub

Private Function CheckOrderState(LngOrderID As Long) As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    StrSQL = "Select * From TblWorkOrdersData Where WorkOrderID=" & LngOrderID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        If rs("WorkOrderState").value = 0 Then
            CheckOrderState = True
        ElseIf rs("WorkOrderState").value = 1 Then
            CheckOrderState = False
        End If

    Else
        CheckOrderState = False
    End If

    rs.Close
    Set rs = Nothing
End Function

Private Sub XPDtbBill_Change()
   If Me.TxtModFlg.Text <> "R" Then
   TxtNoteSerial2.Text = ""
      If ChekSanNumber(Current_branch, 66) = True Then
          TxtNoteSerial2.Text = ""
      End If
      TxtNoteSerial2.Text = ""
   End If
End Sub
