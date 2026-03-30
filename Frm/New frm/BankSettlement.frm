VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form BankSettlementt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ”ÊÌ«  »‰ﬂÌ…"
   ClientHeight    =   10065
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16995
   ForeColor       =   &H00000000&
   Icon            =   "BankSettlement.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10065
   ScaleWidth      =   16995
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   10065
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   16995
      _cx             =   29977
      _cy             =   17754
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic8 
         Height          =   825
         Left            =   0
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   9240
         Width           =   16965
         _cx             =   29924
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
         Begin ImpulseButton.ISButton btnNew 
            Height          =   450
            Left            =   15375
            TabIndex        =   53
            Top             =   285
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   794
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
            ButtonImage     =   "BankSettlement.frx":6852
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   450
            Left            =   11160
            TabIndex        =   54
            Top             =   285
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   794
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
            ButtonImage     =   "BankSettlement.frx":D0B4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   450
            Left            =   13335
            TabIndex        =   55
            Top             =   285
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   794
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
            ButtonImage     =   "BankSettlement.frx":D44E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   450
            Left            =   9120
            TabIndex        =   56
            Top             =   285
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   794
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
            ButtonImage     =   "BankSettlement.frx":13CB0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   450
            Left            =   2790
            TabIndex        =   57
            Top             =   285
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   794
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
            ButtonImage     =   "BankSettlement.frx":1404A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   450
            Left            =   495
            TabIndex        =   58
            Top             =   285
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   794
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
            ButtonImage     =   "BankSettlement.frx":145E4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   450
            Left            =   7110
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   285
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   794
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… "
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
            ButtonImage     =   "BankSettlement.frx":1497E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   450
            Left            =   5055
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   285
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   794
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
            ButtonImage     =   "BankSettlement.frx":1B1E0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   810
         Left            =   0
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   8520
         Width           =   16965
         _cx             =   29924
         _cy             =   1429
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic7 
            Height          =   600
            Left            =   0
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   120
            Width           =   6120
            _cx             =   10795
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
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   270
               Index           =   0
               Left            =   4515
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   270
               Index           =   1
               Left            =   1815
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   120
               Width           =   1605
            End
            Begin VB.Label LabCurrRec 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00800000&
               Height          =   270
               Left            =   3660
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   120
               Width           =   705
            End
            Begin VB.Label LabCountRec 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   630
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   120
               Width           =   795
            End
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   12555
            TabIndex        =   43
            Top             =   225
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   300
            Left            =   10440
            TabIndex        =   44
            ToolTipText     =   "Õ–› «·’› «·Õ«·Ì"
            Top             =   225
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   529
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–› «·’› «·Õ«·Ì"
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
            ButtonImage     =   "BankSettlement.frx":1B57A
            ButtonImageDisabled=   "BankSettlement.frx":21DDC
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   300
            Left            =   8820
            TabIndex        =   45
            ToolTipText     =   "Õ–› «·ﬂ·"
            Top             =   225
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   529
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–› «·ﬂ· "
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
            ButtonImage     =   "BankSettlement.frx":40FC6
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   300
            Index           =   8
            Left            =   15750
            TabIndex        =   46
            Top             =   225
            Width           =   885
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   660
         Left            =   0
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2355
         Width           =   16965
         _cx             =   29924
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
         Begin VB.TextBox txtto 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   2265
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   36
            Top             =   120
            Width           =   6135
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Height          =   330
            Left            =   11910
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   120
            Width           =   1680
         End
         Begin VB.TextBox oldTxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Height          =   330
            Left            =   9150
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   120
            Width           =   1680
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   330
            Left            =   14340
            TabIndex        =   33
            Top             =   120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            _Version        =   393216
            Format          =   94044161
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   405
            Left            =   120
            TabIndex        =   37
            Top             =   120
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   714
            ButtonPositionImage=   3
            Caption         =   "«÷«›…"
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
            ButtonImage     =   "BankSettlement.frx":47828
            ColorButton     =   14871017
            AlignmentVertical=   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‘—Õ"
            Height          =   270
            Index           =   0
            Left            =   8100
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   120
            Width           =   1260
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„…"
            Height          =   270
            Index           =   9
            Left            =   13515
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   120
            Width           =   915
         End
         Begin VB.Label Labelbnkrf 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„—Ã⁄ «·»‰ﬂ"
            Height          =   330
            Left            =   10845
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label lblmovment 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·Õ—ﬂ…"
            Height          =   210
            Left            =   15885
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   120
            Width           =   1095
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   645
         Left            =   0
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1755
         Width           =   17085
         _cx             =   30136
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
         Begin VB.TextBox txtFile 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   4260
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   195
            Visible         =   0   'False
            Width           =   1680
         End
         Begin VB.OptionButton check2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«” Ì—«œ „‰ „·›"
            Height          =   300
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   90
            Value           =   -1  'True
            Width           =   1875
         End
         Begin VB.OptionButton check1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÌœÊÌ"
            Height          =   300
            Left            =   8955
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   90
            Width           =   1140
         End
         Begin ImpulseButton.ISButton CmdImport 
            Height          =   390
            Left            =   1275
            TabIndex        =   30
            Top             =   90
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   688
            Caption         =   " Õ„Ì· «·„·›"
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
            ButtonImage     =   "BankSettlement.frx":4E08A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin ImpulseButton.ISButton CMDSelectFile 
            Height          =   390
            Left            =   3300
            TabIndex        =   74
            Top             =   90
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   688
            Caption         =   "Õœœ „”«— «·„·›"
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
            ButtonImage     =   "BankSettlement.frx":548EC
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin MSComDlg.CommonDialog CD1 
            Left            =   240
            Top             =   120
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «· ”ÊÌ…"
            Height          =   285
            Index           =   1
            Left            =   11310
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   195
            Width           =   2310
         End
      End
      Begin C1SizerLibCtl.C1Elastic Frm2 
         Height          =   1170
         Left            =   0
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   720
         Width           =   17085
         _cx             =   30136
         _cy             =   2064
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
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   330
            Left            =   14175
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   135
            Width           =   1425
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   525
            Left            =   105
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   495
            Width           =   8250
            _cx             =   14552
            _cy             =   926
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
            Begin MSComCtl2.DTPicker DtpDateFrom 
               Height          =   330
               Left            =   4170
               TabIndex        =   21
               Top             =   120
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   94044163
               CurrentDate     =   38887
            End
            Begin MSComCtl2.DTPicker DtpDateTo 
               Height          =   330
               Left            =   1440
               TabIndex        =   22
               Top             =   135
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   94044163
               CurrentDate     =   38887
            End
            Begin ImpulseButton.ISButton ISButton6 
               Height          =   255
               Left            =   -405
               TabIndex        =   23
               Top             =   135
               Visible         =   0   'False
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   450
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
               ButtonImage     =   "BankSettlement.frx":5B14E
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               LowerToggledContent=   0   'False
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄·Ï «·› —… "
               Height          =   255
               Index           =   13
               Left            =   6435
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   135
               Width           =   1305
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   255
               Index           =   3
               Left            =   5490
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   135
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   255
               Index           =   5
               Left            =   2940
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   135
               Width           =   1260
            End
         End
         Begin Dynamic_Byte.NourHijriCal Txt_DateHigri 
            Height          =   330
            Left            =   6990
            TabIndex        =   10
            Top             =   135
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   582
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   330
            Left            =   10905
            TabIndex        =   11
            Top             =   135
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   582
            _Version        =   393216
            Format          =   94044161
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker SettlementDT 
            Height          =   330
            Left            =   14175
            TabIndex        =   12
            Top             =   600
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   582
            _Version        =   393216
            Format          =   94044161
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "BankSettlement.frx":619B0
            Height          =   315
            Left            =   225
            TabIndex        =   13
            Top             =   135
            Width           =   4515
            _ExtentX        =   7964
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
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   8490
            TabIndex        =   14
            Top             =   600
            Width           =   4710
            _ExtentX        =   8308
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbldate 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   360
            Left            =   12885
            TabIndex        =   20
            Top             =   240
            Width           =   780
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «· ”ÊÌ…"
            Height          =   210
            Index           =   15
            Left            =   15705
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label lblhjdate 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ «·ÂÃ—Ì"
            Height          =   270
            Left            =   9240
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   135
            Width           =   1320
         End
         Begin VB.Label lblcode 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   270
            Left            =   15630
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   135
            Width           =   1275
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            Height          =   240
            Left            =   5265
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   135
            Width           =   1245
         End
         Begin VB.Label Labelbank 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «·»‰ﬂ"
            Height          =   240
            Left            =   13185
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   600
            Width           =   1065
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   735
         Left            =   0
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   0
         Width           =   17100
         _cx             =   30163
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
         BackColor       =   16777215
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
         Begin ImpulseButton.ISButton btnLast 
            Height          =   255
            Left            =   720
            TabIndex        =   62
            Top             =   210
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   450
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
            ButtonImage     =   "BankSettlement.frx":619C5
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   255
            Left            =   2370
            TabIndex        =   63
            Top             =   210
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   450
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
            ButtonImage     =   "BankSettlement.frx":61D5F
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   255
            Left            =   1230
            TabIndex        =   64
            Top             =   210
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   450
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
            ButtonImage     =   "BankSettlement.frx":620F9
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   255
            Left            =   1845
            TabIndex        =   65
            Top             =   210
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   450
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
            ButtonImage     =   "BankSettlement.frx":62493
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " ”ÊÌ«  »‰ﬂÌ…"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   420
            Index           =   2
            Left            =   10890
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   210
            Width           =   4200
         End
         Begin VB.Image Image1 
            Height          =   510
            Left            =   15480
            Picture         =   "BankSettlement.frx":6282D
            Stretch         =   -1  'True
            Top             =   120
            Width           =   1215
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid FG 
         Height          =   2235
         Left            =   7770
         TabIndex        =   67
         Top             =   6315
         Width           =   9195
         _cx             =   16219
         _cy             =   3942
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
         BackColor       =   12648447
         ForeColor       =   -2147483640
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   12648447
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
         Rows            =   1
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"BankSettlement.frx":63DD2
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
         Begin MSComctlLib.ProgressBar ProgressBar2 
            Height          =   615
            Left            =   2760
            TabIndex        =   68
            Top             =   1440
            Visible         =   0   'False
            Width           =   11655
            _ExtentX        =   20558
            _ExtentY        =   1085
            _Version        =   393216
            Appearance      =   0
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid Fg1 
         Height          =   2235
         Left            =   0
         TabIndex        =   69
         Top             =   6315
         Width           =   7710
         _cx             =   13600
         _cy             =   3942
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
         BackColor       =   12648447
         ForeColor       =   -2147483640
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   12648447
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
         Rows            =   1
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"BankSettlement.frx":63F4F
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
         Begin MSComctlLib.ProgressBar ProgressBar4 
            Height          =   615
            Left            =   2760
            TabIndex        =   70
            Top             =   1200
            Visible         =   0   'False
            Width           =   11655
            _ExtentX        =   20558
            _ExtentY        =   1085
            _Version        =   393216
            Appearance      =   0
         End
      End
      Begin ALLButtonS.ALLButton BtnShow 
         Height          =   315
         Left            =   6690
         TabIndex        =   73
         Top             =   5970
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   " ‰›Ì– «·„ﬁ«—‰…"
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
         BCOL            =   12583104
         BCOLO           =   12583104
         FCOL            =   16777088
         FCOLO           =   16777088
         MCOL            =   192
         MPTR            =   1
         MICON           =   "BankSettlement.frx":6402B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
         Height          =   2835
         Left            =   7770
         TabIndex        =   76
         Top             =   3045
         Width           =   9195
         _cx             =   16219
         _cy             =   5001
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
         Rows            =   1
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"BankSettlement.frx":64047
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
         Begin MSComctlLib.ProgressBar ProgressBar3 
            Height          =   615
            Left            =   2760
            TabIndex        =   77
            Top             =   1440
            Visible         =   0   'False
            Width           =   11655
            _ExtentX        =   20558
            _ExtentY        =   1085
            _Version        =   393216
            Appearance      =   0
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid Grid 
         Height          =   2835
         Left            =   0
         TabIndex        =   78
         Top             =   3045
         Width           =   7710
         _cx             =   13600
         _cy             =   5001
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
         Rows            =   1
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"BankSettlement.frx":641E1
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
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   615
            Left            =   2760
            TabIndex        =   79
            Top             =   1200
            Visible         =   0   'False
            Width           =   11655
            _ExtentX        =   20558
            _ExtentY        =   1085
            _Version        =   393216
            Appearance      =   0
         End
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "Õ—ﬂ«  „”Ã·… ⁄·Ï «·»‰ﬂ Ê€Ì— „ÊÃÊœ… ›Ì «·»—‰«„Ã"
         ForeColor       =   &H8000000B&
         Height          =   330
         Index           =   4
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   5970
         Width           =   5730
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         Caption         =   "Õ—ﬂ«  „”Ã·… ⁄·Ï «·»—‰«„Ã Ê€Ì— „ÊÃÊœ… ›Ì «·»‰ﬂ"
         ForeColor       =   &H8000000B&
         Height          =   330
         Index           =   2
         Left            =   9495
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   5970
         Width           =   6705
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   23640
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Text            =   "modflag"
      Top             =   1320
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "BankSettlement.frx":642DE
      Left            =   23640
      List            =   "BankSettlement.frx":642EE
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2040
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   23640
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   945
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   20280
      TabIndex        =   3
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   -360
      Width           =   2100
      _ExtentX        =   3704
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
   Begin MSDataListLib.DataCombo DCPreFix 
      Height          =   315
      Left            =   23640
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComctlLib.ImageList GrdImageList 
      Left            =   23640
      Top             =   2280
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
            Picture         =   "BankSettlement.frx":64307
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BankSettlement.frx":646A1
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BankSettlement.frx":64A3B
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BankSettlement.frx":64DD5
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BankSettlement.frx":6516F
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BankSettlement.frx":65509
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BankSettlement.frx":658A3
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BankSettlement.frx":65E3D
            Key             =   "BuyValue"
         EndProperty
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
      Left            =   19920
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   -960
      Width           =   855
   End
End
Attribute VB_Name = "BankSettlementt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Dim RsSavRec As ADODB.Recordset
 Dim StrSQL As String
 Dim RsDevsub As ADODB.Recordset
 Dim BKGrndPic As ClsBackGroundPic
 Dim RecId As String
 Dim II As Long
Private Sub btnQuery_Click()
Load BankSetmSearch
BankSetmSearch.show vbModal
End Sub
Sub fillgrid1()
Dim i As Integer
If Grid.Rows >= 2 Then
Dim count As Integer
With Fg1

If .Rows = 1 Then
.Rows = .Rows + 1
Else
.Rows = 2
End If
count = 1
For i = 1 To Grid.Rows - 1
If val(Grid.TextMatrix(i, Grid.ColIndex("RowG"))) = 0 Then
.TextMatrix(count, .ColIndex("Ser")) = count
.TextMatrix(count, .ColIndex("BankValue")) = Grid.TextMatrix(i, Grid.ColIndex("BankValue"))
.TextMatrix(count, .ColIndex("MoveDT")) = Grid.TextMatrix(i, Grid.ColIndex("MoveDT"))
.TextMatrix(count, .ColIndex("BankRF")) = Grid.TextMatrix(i, Grid.ColIndex("BankRF"))
.TextMatrix(count, .ColIndex("Explan")) = Grid.TextMatrix(i, Grid.ColIndex("Explan"))
.Rows = .Rows + 1
count = count + 1
End If
Next i
.Rows = .Rows - 1
End With
End If
End Sub
Sub fillgrid()
Dim i As Integer

Dim count As Integer
If VSFlexGrid2.Rows >= 2 Then
With Fg

If .Rows = 1 Then
.Rows = .Rows + 1
Else
.Rows = 2
End If
count = 1

For i = 1 To VSFlexGrid2.Rows - 1
If val(VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("RowG"))) = 0 Then
.TextMatrix(count, .ColIndex("Ser")) = count
.TextMatrix(count, .ColIndex("NoteSerial1")) = VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("NoteSerial1"))
.TextMatrix(count, .ColIndex("NotesTypeName")) = VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("NotesTypeName"))
.TextMatrix(count, .ColIndex("BankValue")) = VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("BankValue"))
.TextMatrix(count, .ColIndex("NoteSerial")) = VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("NoteSerial"))
.TextMatrix(count, .ColIndex("MoveDT")) = VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("MoveDT"))
.TextMatrix(count, .ColIndex("BankRF")) = VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("BankRF"))
.TextMatrix(count, .ColIndex("Explan")) = VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("Explan"))
.TextMatrix(count, .ColIndex("NoteType")) = VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("NoteType"))
.Rows = .Rows + 1
count = count + 1
End If
Next i
.Rows = .Rows - 1
End With
End If
End Sub
Sub ComparBetwenSystemAndBank1()
Dim bol As Boolean
Dim i As Integer
Dim j As Integer
For i = 1 To VSFlexGrid2.Rows - 1
bol = False
For j = 1 To Grid.Rows - 1

If val(Me.VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("BankRF"))) <> 0 And val(Me.Grid.TextMatrix(j, Grid.ColIndex("BankRF"))) <> 0 And (val(Me.Grid.TextMatrix(j, Grid.ColIndex("BankRF"))) = val(Me.VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("BankRF")))) And val(Me.Grid.TextMatrix(j, Grid.ColIndex("RowG"))) = 0 And val(Me.VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("RowG"))) = 0 Then
Me.Grid.TextMatrix(j, Grid.ColIndex("RowG")) = i
Me.VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("RowG")) = j
 VSFlexGrid2.Cell(flexcpBackColor, i, VSFlexGrid2.ColIndex("Ser"), i, VSFlexGrid2.ColIndex("Explan")) = &HFFFF80
 Grid.Cell(flexcpBackColor, j, Grid.ColIndex("Ser"), j, Grid.ColIndex("Explan")) = &HFFFF80
 bol = True
 End If
Next j
If bol = False Then
For j = 1 To Grid.Rows - 1

If val(Me.Grid.TextMatrix(j, Grid.ColIndex("BankValue"))) <> 0 And (Me.VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("MoveDT")) = Me.Grid.TextMatrix(j, Grid.ColIndex("MoveDT"))) And (val(Me.VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("BankValue"))) = val(Me.Grid.TextMatrix(j, Grid.ColIndex("BankValue")))) And val(Me.Grid.TextMatrix(j, Grid.ColIndex("RowG"))) = 0 And val(Me.VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("RowG"))) = 0 Then
Me.Grid.TextMatrix(j, Grid.ColIndex("RowG")) = i
Me.VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("RowG")) = j
 VSFlexGrid2.Cell(flexcpBackColor, i, VSFlexGrid2.ColIndex("Ser"), i, VSFlexGrid2.ColIndex("Explan")) = &HFFFF80
  Grid.Cell(flexcpBackColor, j, Grid.ColIndex("Ser"), j, Grid.ColIndex("Explan")) = &HFFFF80
 bol = True
 End If
Next j
End If

If bol = False Then
For j = 1 To Grid.Rows - 1

If val(Me.VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("RowG"))) = 0 Then
Me.VSFlexGrid2.TextMatrix(i, VSFlexGrid2.ColIndex("RowG")) = 0
 VSFlexGrid2.Cell(flexcpBackColor, i, VSFlexGrid2.ColIndex("Ser"), i, VSFlexGrid2.ColIndex("Explan")) = &HFF&

 bol = True
 End If
Next j
End If
Next i
End Sub
Sub ComparBetwenSystemAndBank()
Dim bol As Boolean
Dim i As Integer
Dim j As Integer
For i = 1 To Grid.Rows - 1
bol = False
For j = 1 To VSFlexGrid2.Rows - 1

If val(Me.VSFlexGrid2.TextMatrix(j, VSFlexGrid2.ColIndex("BankRF"))) <> 0 And val(Me.Grid.TextMatrix(i, Grid.ColIndex("BankRF"))) <> 0 And (val(Me.Grid.TextMatrix(i, Grid.ColIndex("BankRF"))) = val(Me.VSFlexGrid2.TextMatrix(j, VSFlexGrid2.ColIndex("BankRF")))) And val(Me.Grid.TextMatrix(i, Grid.ColIndex("RowG"))) = 0 And val(Me.VSFlexGrid2.TextMatrix(j, VSFlexGrid2.ColIndex("RowG"))) = 0 Then
Me.Grid.TextMatrix(i, Grid.ColIndex("RowG")) = j
Me.VSFlexGrid2.TextMatrix(j, VSFlexGrid2.ColIndex("RowG")) = i
 VSFlexGrid2.Cell(flexcpBackColor, j, VSFlexGrid2.ColIndex("Ser"), j, VSFlexGrid2.ColIndex("Explan")) = &HFFFF80
 Grid.Cell(flexcpBackColor, i, Grid.ColIndex("Ser"), i, Grid.ColIndex("Explan")) = &HFFFF80
 bol = True
 End If
Next j
If bol = False Then
For j = 1 To VSFlexGrid2.Rows - 1
'MsgBox val(Me.Grid.TextMatrix(i, Grid.ColIndex("BankValue")))
If val(Me.Grid.TextMatrix(i, Grid.ColIndex("BankValue"))) <> 0 And (Me.VSFlexGrid2.TextMatrix(j, VSFlexGrid2.ColIndex("MoveDT")) = Me.Grid.TextMatrix(i, Grid.ColIndex("MoveDT"))) And (val(Me.VSFlexGrid2.TextMatrix(j, VSFlexGrid2.ColIndex("BankValue"))) = val(Me.Grid.TextMatrix(i, Grid.ColIndex("BankValue")))) And val(Me.Grid.TextMatrix(i, Grid.ColIndex("RowG"))) = 0 And val(Me.VSFlexGrid2.TextMatrix(j, VSFlexGrid2.ColIndex("RowG"))) = 0 Then
Me.Grid.TextMatrix(i, Grid.ColIndex("RowG")) = j
Me.VSFlexGrid2.TextMatrix(j, VSFlexGrid2.ColIndex("RowG")) = i
 VSFlexGrid2.Cell(flexcpBackColor, j, VSFlexGrid2.ColIndex("Ser"), j, VSFlexGrid2.ColIndex("Explan")) = &HFFFF80
  Grid.Cell(flexcpBackColor, i, Grid.ColIndex("Ser"), i, Grid.ColIndex("Explan")) = &HFFFF80
 bol = True
 End If
Next j
End If

If bol = False Then
For j = 1 To VSFlexGrid2.Rows - 1
If val(Me.Grid.TextMatrix(i, Grid.ColIndex("RowG"))) = 0 Then
  Grid.Cell(flexcpBackColor, i, Grid.ColIndex("Ser"), i, Grid.ColIndex("Explan")) = &HFF&
 bol = True
 End If
Next j
End If
Next i
End Sub

Private Sub BtnShow_Click()
SearchData
If Me.Grid.Rows < 2 Or VSFlexGrid2.Rows < 2 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «œŒ«· «œŒ«· «·»Ì«‰«  »‘ﬂ· ’ÕÌÕ ﬁ»· «·„ﬁ«—‰Â"
Else
MsgBox "Please Enter Data Before Compare"
End If
Exit Sub
End If
ComparBetwenSystemAndBank
ComparBetwenSystemAndBank1
fillgrid
fillgrid1
End Sub

 Private Sub CmdImport_Click()
If txtFile.Text = "" Then MsgBox "Õœœ «·„·› «Ê·«": Exit Sub
Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Integer
Dim currentvalue As String

Dim BDate As Date
Dim val1 As Double
Dim des As String
Dim CheqNo As String

  

    Set ExcelObj = CreateObject("Excel.Application")
    Set ExcelSheet = CreateObject("Excel.Sheet")

    ExcelObj.Workbooks.Open txtFile.Text   ' App.Path & "\TrialBalance.xls"
DoEvents
    Set ExcelBook = ExcelObj.Workbooks(1)
    Set ExcelSheet = ExcelBook.Worksheets(1)
 
    With ExcelSheet
    i = 2
    Grid.Rows = 2
    Do Until .cells(i, 2) & "" = ""
 '       Set l = lvwList.ListItems.Add(, , .Cells(i, 1))
   BDate = .cells(i, 1)
    val1 = .cells(i, 2)
         des = .cells(i, 4)
        CheqNo = .cells(i, 3)
       
         
        
 With Grid

   '  MsgBox .Rows
  .TextMatrix(i - 1, .ColIndex("MoveDT")) = (BDate)
  .TextMatrix(i - 1, .ColIndex("BankValue")) = val(val1)
   .TextMatrix(i - 1, .ColIndex("BankRF")) = (CheqNo)
   .TextMatrix(i - 1, .ColIndex("Explan")) = (des)
   
  ' Grid_AfterEdit i, .ColIndex("account_serial")
   
     
    
          


'    Fg_Journal_AfterEdit i, .ColIndex("BranchId")
'
'   If Val(DebitValue) > 0 Then
'      .TextMatrix(i, .ColIndex("DebitValue")) = Val(DebitValue)
'         Fg_Journal_AfterEdit i, .ColIndex("DebitValue")
'
'    End If
'
'       If Val(CreditValue) > 0 Then
'     .TextMatrix(i, .ColIndex("CreditValue")) = Val(CreditValue)
'     Fg_Journal_AfterEdit i, .ColIndex("CreditValue")
'      End If
      
   
 End With
        i = i + 1
       Grid.Rows = Grid.Rows + 1
        
    Loop

    End With
Grid.Rows = Grid.Rows - 1
    ExcelObj.Workbooks.Close

    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelObj = Nothing
End Sub

   Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    conection = "select * from TBLBankSettlement WHERE 1=1 "
    StrSQL = StrSQL & "  AND      BranchID in(" & Current_branchSql & ")"
    
          If SystemOptions.usertype <> UserAdmin Then
      '  StrSQL = StrSQL & " AND   BranchID=" & Current_branch
    End If
    StrSQL = StrSQL & "  order by  IDBS  "
    
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
   'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetBanks Me.DcboBox
     
    BtnLast_Click
    ShowTip
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        SwitchKeyboardLang LANG_ENGLISH
        Else
        SwitchKeyboardLang LANG_ARABIC
    End If
    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If
       Me.Refresh
ErrTrap:
End Sub
' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Sub saveGrid()
Dim Rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim Rs3 As ADODB.Recordset
Dim Rs4 As ADODB.Recordset
    Set Rs1 = New ADODB.Recordset
    StrSQL = "SELECT  *  from TBLBankSettlementJoin Where (1 = -1)"
    Rs1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    With Grid
       For i = .FixedRows To .Rows - 1
     If .TextMatrix(i, .ColIndex("Ser")) <> "" Then
                Rs1.AddNew
                Rs1("IDBS").value = Me.TxtSerial1.Text
                Rs1("MoveDT").value = IIf((.TextMatrix(i, .ColIndex("MoveDT"))) = "", Null, .TextMatrix(i, .ColIndex("MoveDT")))
                Rs1("BankValue").value = IIf((.TextMatrix(i, .ColIndex("BankValue"))) = "", Null, .TextMatrix(i, .ColIndex("BankValue")))
                Rs1("BankRF").value = IIf((.TextMatrix(i, .ColIndex("BankRF"))) = "", Null, .TextMatrix(i, .ColIndex("BankRF")))
                Rs1("Explan").value = IIf((.TextMatrix(i, .ColIndex("Explan"))) = "", Null, .TextMatrix(i, .ColIndex("Explan")))
                Rs1("RowG").value = IIf((.TextMatrix(i, .ColIndex("RowG"))) = "", Null, .TextMatrix(i, .ColIndex("RowG")))
                Rs1("TypeTrans").value = 0
                Rs1.update
      End If
     Next i
     End With
   '''/////
    Set rs2 = New ADODB.Recordset
    StrSQL = "SELECT  *  from TBLBankSettlementJoin Where (1 = -1)"
    rs2.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      With Fg1
       For i = .FixedRows To .Rows - 1
     If .TextMatrix(i, .ColIndex("Ser")) <> "" Then
                rs2.AddNew
                rs2("IDBS").value = Me.TxtSerial1.Text
                rs2("MoveDT").value = IIf((.TextMatrix(i, .ColIndex("MoveDT"))) = "", Null, .TextMatrix(i, .ColIndex("MoveDT")))
                rs2("BankValue").value = IIf((.TextMatrix(i, .ColIndex("BankValue"))) = "", Null, .TextMatrix(i, .ColIndex("BankValue")))
                rs2("BankRF").value = IIf((.TextMatrix(i, .ColIndex("BankRF"))) = "", Null, .TextMatrix(i, .ColIndex("BankRF")))
                rs2("Explan").value = IIf((.TextMatrix(i, .ColIndex("Explan"))) = "", Null, .TextMatrix(i, .ColIndex("Explan")))
                rs2("TypeTrans").value = 1
                rs2.update
      End If
     Next i
     End With
 '''///
 Set Rs3 = New ADODB.Recordset
    StrSQL = "SELECT  *  from TBLBankSettlementJoin Where (1 = -1)"
    Rs3.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With VSFlexGrid2
       For i = .FixedRows To .Rows - 1
     If .TextMatrix(i, .ColIndex("Ser")) <> "" Then
                Rs3.AddNew
                Rs3("IDBS").value = Me.TxtSerial1.Text
                Rs3("MoveDT").value = IIf((.TextMatrix(i, .ColIndex("MoveDT"))) = "", Null, .TextMatrix(i, .ColIndex("MoveDT")))
                Rs3("BankValue").value = IIf((.TextMatrix(i, .ColIndex("BankValue"))) = "", Null, .TextMatrix(i, .ColIndex("BankValue")))
                Rs3("BankRF").value = IIf((.TextMatrix(i, .ColIndex("BankRF"))) = "", Null, .TextMatrix(i, .ColIndex("BankRF")))
                Rs3("Explan").value = IIf((.TextMatrix(i, .ColIndex("Explan"))) = "", Null, .TextMatrix(i, .ColIndex("Explan")))
                Rs3("TypeTrans").value = 2
                Rs3("NoteType").value = IIf((.TextMatrix(i, .ColIndex("NoteType"))) = "", Null, .TextMatrix(i, .ColIndex("NoteType")))
                Rs3("NoteSerial1").value = IIf((.TextMatrix(i, .ColIndex("NoteSerial1"))) = "", Null, .TextMatrix(i, .ColIndex("NoteSerial1")))
                Rs3("NoteSerial").value = IIf((.TextMatrix(i, .ColIndex("NoteSerial"))) = "", Null, .TextMatrix(i, .ColIndex("NoteSerial")))
                Rs3("RowG").value = IIf((.TextMatrix(i, .ColIndex("RowG"))) = "", Null, .TextMatrix(i, .ColIndex("RowG")))
                Rs3.update
      End If
     Next i
     End With
 '''///
  Set Rs4 = New ADODB.Recordset
    StrSQL = "SELECT  *  from TBLBankSettlementJoin Where (1 = -1)"
    Rs4.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     With Fg
       For i = .FixedRows To .Rows - 1
     If .TextMatrix(i, .ColIndex("Ser")) <> "" Then
                Rs4.AddNew
                Rs4("IDBS").value = Me.TxtSerial1.Text
                Rs4("MoveDT").value = IIf((.TextMatrix(i, .ColIndex("MoveDT"))) = "", Null, .TextMatrix(i, .ColIndex("MoveDT")))
                Rs4("BankValue").value = IIf((.TextMatrix(i, .ColIndex("BankValue"))) = "", Null, .TextMatrix(i, .ColIndex("BankValue")))
                Rs4("BankRF").value = IIf((.TextMatrix(i, .ColIndex("BankRF"))) = "", Null, .TextMatrix(i, .ColIndex("BankRF")))
                Rs4("Explan").value = IIf((.TextMatrix(i, .ColIndex("Explan"))) = "", Null, .TextMatrix(i, .ColIndex("Explan")))
                Rs4("TypeTrans").value = 3
                Rs4("NoteType").value = IIf((.TextMatrix(i, .ColIndex("NoteType"))) = "", Null, .TextMatrix(i, .ColIndex("NoteType")))
                Rs4("NoteSerial1").value = IIf((.TextMatrix(i, .ColIndex("NoteSerial1"))) = "", Null, .TextMatrix(i, .ColIndex("NoteSerial1")))
                Rs4("NoteSerial").value = IIf((.TextMatrix(i, .ColIndex("NoteSerial"))) = "", Null, .TextMatrix(i, .ColIndex("NoteSerial")))
                Rs4.update
      End If
     Next i
     End With
 
End Sub

Public Sub FiLLRec()
    On Error GoTo ErrTrap
    If TxtModFlg = "E" Then
    If check1.value = True Then
    StrSQL = "Delete From TBLBankSettlementJoin Where IDBS='" & val(TxtSerial1.Text) & "'"
    Cn.Execute StrSQL, , adExecuteNoRecords
    End If
    End If
    RsSavRec.Fields("DateM").value = XPDtbTrans.value
    RsSavRec.Fields("DateH").value = Me.Txt_DateHigri.value
    RsSavRec.Fields("BranchID").value = val(Me.dcBranch.BoundText)
    RsSavRec.Fields("SettlementDT").value = SettlementDT.value
    RsSavRec.Fields("BankID").value = val(Me.DcboBox.BoundText)
    RsSavRec.Fields("FromDT").value = DtpDateFrom.value
    RsSavRec.Fields("ToDT").value = DtpDateTo.value
       
    If check1.value = True Then
    RsSavRec.Fields("EXPCheck").value = 0
    Else
    RsSavRec.Fields("EXPCheck").value = 1
    End If
     If Not (IsNull(RsSavRec.Fields("LockedInterval").value)) Then
    If RsSavRec.Fields("LockedInterval").value = True Then
    btnDelete.Enabled = False
    btnModify.Enabled = False
    Else
    btnDelete.Enabled = True
    btnModify.Enabled = True
    End If
    End If
        
    RsSavRec.Fields("UserID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    RsSavRec.update
    saveGrid
    ' save grid

      Select Case Me.TxtModFlg.Text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " Saved... " & CHR(13)
                Msg = Msg + "Do you want to enter another operation?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                 If SystemOptions.UserInterface = ArabicInterface Then
             Else
               Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                TxtModFlg = "R"
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            End If
       End Select
  Exit Sub
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If
   End Sub
' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()
   On Error GoTo ErrTrap
    Dim i As Integer
    ProgressBar1.Visible = True
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("IDBS").value), "", RsSavRec.Fields("IDBS").value): ProgressBar1.value = 10
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("DateM").value), Date, RsSavRec.Fields("DateM").value): ProgressBar1.value = 20
    Txt_DateHigri.value = IIf(IsNull(RsSavRec.Fields("DateH").value), "", RsSavRec.Fields("DateH").value): ProgressBar1.value = 30
    dcBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value): ProgressBar1.value = 40
    SettlementDT.value = IIf(IsNull(RsSavRec.Fields("SettlementDT").value), "", RsSavRec.Fields("SettlementDT").value): ProgressBar1.value = 50
    DcboBox.BoundText = IIf(IsNull(RsSavRec.Fields("BankID").value), "", RsSavRec.Fields("BankID").value): ProgressBar1.value = 60
    DtpDateFrom.value = IIf(IsNull(RsSavRec.Fields("FromDT").value), Date, RsSavRec.Fields("FromDT").value): ProgressBar1.value = 70
    DtpDateTo.value = IIf(IsNull(RsSavRec.Fields("ToDT").value), Date, RsSavRec.Fields("ToDT").value): ProgressBar1.value = 80
    ''''''''''''''''
     If RsSavRec.Fields("EXPCheck").value = 0 Then
     check1.value = vbChecked
     Me.Grid.Clear flexClearScrollable, flexClearEverything
     Me.Grid.Enabled = True
     FillTextGridData
     Else
     check2.value = vbChecked
     Me.Grid.Clear flexClearScrollable, flexClearEverything
     Me.Grid.Enabled = False
     End If
    ''''''''''''''''''''''
    If Not (IsNull(RsSavRec.Fields("LockedInterval").value)) Then
    If RsSavRec.Fields("LockedInterval").value = True Then
    btnModify.Enabled = False
    btnDelete.Enabled = False
    Else
    btnModify.Enabled = True
    btnDelete.Enabled = True
    End If
    End If
     DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
     LabCurrRec.Caption = RsSavRec.AbsolutePosition
     LabCountRec.Caption = RsSavRec.RecordCount: ProgressBar1.value = 80
      ProgressBar1.Visible = False
 ProgressBar1.value = 0
ErrTrap:
  ProgressBar1.Visible = False
 ProgressBar1.value = 0
End Sub
  Sub FillTextGridData()
 If check1.value = True Then
  Dim Rs1 As ADODB.Recordset
  Dim i As Integer
  Dim sql As String
    sql = "SELECT     ID, IDBS AS IDBSJON, MoveDT, BankValue, BankRF, Explan, RowG , TypeTrans"
    sql = sql & " From dbo.TBLBankSettlementJoin"
    sql = sql & " Where (IDBS =" & val(TxtSerial1.Text) & ") And (TypeTrans = 0)"
 Set Rs1 = New ADODB.Recordset
    Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
        
     With Me.Grid
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("IDBS")) = IIf(IsNull(Rs1("IDBSJON").value), "", Rs1("IDBSJON").value)
                   .TextMatrix(i, .ColIndex("MoveDT")) = IIf(IsNull(Rs1("MoveDT").value), "", Rs1("MoveDT").value)
                   .TextMatrix(i, .ColIndex("BankValue")) = IIf(IsNull(Rs1("BankValue").value), "", Rs1("BankValue").value)
                   .TextMatrix(i, .ColIndex("BankRF")) = IIf(IsNull(Rs1("BankRF").value), "", Rs1("BankRF").value)
                   .TextMatrix(i, .ColIndex("Explan")) = IIf(IsNull(Rs1("Explan").value), "", Rs1("Explan").value)
                   .TextMatrix(i, .ColIndex("RowG")) = IIf(IsNull(Rs1("RowG").value), "", Rs1("RowG").value)
                   If val(.TextMatrix(i, .ColIndex("RowG"))) <> 0 Then
                    .Cell(flexcpBackColor, i, .ColIndex("Ser"), i, .ColIndex("Explan")) = &HFFFF80
                    Else
                    .Cell(flexcpBackColor, i, .ColIndex("Ser"), i, .ColIndex("Explan")) = &HFF&
                  End If
                   Rs1.MoveNext
             Next i
              .AutoSize 0, .Cols - 1, False
        End With
        
      End If
   '''/////////////////////////
     sql = "SELECT     ID, IDBS AS IDBSJON, MoveDT, BankValue, BankRF, Explan, RowG , TypeTrans"
    sql = sql & " From dbo.TBLBankSettlementJoin"
    sql = sql & " Where (IDBS =" & val(TxtSerial1.Text) & ") And (TypeTrans = 1)"
 Set Rs1 = New ADODB.Recordset
    Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     With Me.Fg1
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("IDBS")) = IIf(IsNull(Rs1("IDBSJON").value), "", Rs1("IDBSJON").value)
                   .TextMatrix(i, .ColIndex("MoveDT")) = IIf(IsNull(Rs1("MoveDT").value), "", Rs1("MoveDT").value)
                   .TextMatrix(i, .ColIndex("BankValue")) = IIf(IsNull(Rs1("BankValue").value), "", Rs1("BankValue").value)
                   .TextMatrix(i, .ColIndex("BankRF")) = IIf(IsNull(Rs1("BankRF").value), "", Rs1("BankRF").value)
                   .TextMatrix(i, .ColIndex("Explan")) = IIf(IsNull(Rs1("Explan").value), "", Rs1("Explan").value)
                   Rs1.MoveNext
             Next i
              .AutoSize 0, .Cols - 1, False
        End With
          End If
   '''//////////////
      '''/////////////////////////
     sql = "SELECT     dbo.TBLBankSettlementJoin.ID, dbo.TBLBankSettlementJoin.IDBS AS IDBSJON, dbo.TBLBankSettlementJoin.MoveDT, dbo.TBLBankSettlementJoin.BankValue, "
  sql = sql + "                    dbo.TBLBankSettlementJoin.BankRF, dbo.TBLBankSettlementJoin.Explan, dbo.TBLBankSettlementJoin.notes_id,"
  sql = sql + "                    dbo.TBLBankSettlementJoin.Double_Entry_Vouchers_ID, dbo.TBLBankSettlementJoin.RowG, dbo.TBLBankSettlementJoin.NoteType,"
  sql = sql + "                    dbo.TBLBankSettlementJoin.NoteSerial, dbo.TBLBankSettlementJoin.NoteSerial1, dbo.TBLBankSettlementJoin.TypeTrans, dbo.TblNotesTypes.NotesTypeName,"
  sql = sql + "                    dbo.TblNotesTypes.NotesTypeNameE"
  sql = sql + "  FROM         dbo.TBLBankSettlementJoin LEFT OUTER JOIN"
  sql = sql + "                    dbo.TblNotesTypes ON dbo.TBLBankSettlementJoin.NoteType = dbo.TblNotesTypes.NotesType"
  sql = sql + " Where (dbo.TBLBankSettlementJoin.IDBS = " & val(TxtSerial1.Text) & ") And (dbo.TBLBankSettlementJoin.TypeTrans = 2)"
 Set Rs1 = New ADODB.Recordset
    Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     With Me.VSFlexGrid2
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("IDBS")) = IIf(IsNull(Rs1("IDBSJON").value), "", Rs1("IDBSJON").value)
                   .TextMatrix(i, .ColIndex("MoveDT")) = IIf(IsNull(Rs1("MoveDT").value), "", Rs1("MoveDT").value)
                   .TextMatrix(i, .ColIndex("BankValue")) = IIf(IsNull(Rs1("BankValue").value), "", Rs1("BankValue").value)
                   .TextMatrix(i, .ColIndex("BankRF")) = IIf(IsNull(Rs1("BankRF").value), "", Rs1("BankRF").value)
                   .TextMatrix(i, .ColIndex("NoteType")) = IIf(IsNull(Rs1("NoteType").value), "", Rs1("NoteType").value)
                   .TextMatrix(i, .ColIndex("RowG")) = IIf(IsNull(Rs1("RowG").value), "", Rs1("RowG").value)
                   .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(Rs1("NoteSerial1").value), "", Rs1("NoteSerial1").value)
                   .TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(Rs1("NoteSerial").value), "", Rs1("NoteSerial").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("NotesTypeName")) = IIf(IsNull(Rs1("NotesTypeName").value), "", Rs1("NotesTypeName").value)
                   Else
                   .TextMatrix(i, .ColIndex("NotesTypeName")) = IIf(IsNull(Rs1("NotesTypeNamee").value), "", Rs1("NotesTypeNamee").value)
                   End If
                       If val(.TextMatrix(i, .ColIndex("RowG"))) <> 0 Then
                    .Cell(flexcpBackColor, i, .ColIndex("Ser"), i, .ColIndex("Explan")) = &HFFFF80
                    Else
                    .Cell(flexcpBackColor, i, .ColIndex("Ser"), i, .ColIndex("Explan")) = &HFF&
                  End If
                   Rs1.MoveNext
             Next i
              .AutoSize 0, .Cols - 1, False
        End With
          End If
            '''/////////////////////////
     sql = "SELECT     dbo.TBLBankSettlementJoin.ID, dbo.TBLBankSettlementJoin.IDBS AS IDBSJON, dbo.TBLBankSettlementJoin.MoveDT, dbo.TBLBankSettlementJoin.BankValue, "
  sql = sql + "                    dbo.TBLBankSettlementJoin.BankRF, dbo.TBLBankSettlementJoin.Explan, dbo.TBLBankSettlementJoin.notes_id,"
  sql = sql + "                    dbo.TBLBankSettlementJoin.Double_Entry_Vouchers_ID, dbo.TBLBankSettlementJoin.RowG, dbo.TBLBankSettlementJoin.NoteType,"
  sql = sql + "                    dbo.TBLBankSettlementJoin.NoteSerial, dbo.TBLBankSettlementJoin.NoteSerial1, dbo.TBLBankSettlementJoin.TypeTrans, dbo.TblNotesTypes.NotesTypeName,"
  sql = sql + "                    dbo.TblNotesTypes.NotesTypeNameE"
  sql = sql + "  FROM         dbo.TBLBankSettlementJoin LEFT OUTER JOIN"
  sql = sql + "                    dbo.TblNotesTypes ON dbo.TBLBankSettlementJoin.NoteType = dbo.TblNotesTypes.NotesType"
  sql = sql + " Where (dbo.TBLBankSettlementJoin.IDBS = " & val(TxtSerial1.Text) & ") And (dbo.TBLBankSettlementJoin.TypeTrans = 3)"
 Set Rs1 = New ADODB.Recordset
    Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     With Me.Fg
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("IDBS")) = IIf(IsNull(Rs1("IDBSJON").value), "", Rs1("IDBSJON").value)
                   .TextMatrix(i, .ColIndex("MoveDT")) = IIf(IsNull(Rs1("MoveDT").value), "", Rs1("MoveDT").value)
                   .TextMatrix(i, .ColIndex("BankValue")) = IIf(IsNull(Rs1("BankValue").value), "", Rs1("BankValue").value)
                   .TextMatrix(i, .ColIndex("BankRF")) = IIf(IsNull(Rs1("BankRF").value), "", Rs1("BankRF").value)
                   .TextMatrix(i, .ColIndex("NoteType")) = IIf(IsNull(Rs1("NoteType").value), "", Rs1("NoteType").value)
                   
                   .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(Rs1("NoteSerial1").value), "", Rs1("NoteSerial1").value)
                   .TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(Rs1("NoteSerial").value), "", Rs1("NoteSerial").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("NotesTypeName")) = IIf(IsNull(Rs1("NotesTypeName").value), "", Rs1("NotesTypeName").value)
                   Else
                   .TextMatrix(i, .ColIndex("NotesTypeName")) = IIf(IsNull(Rs1("NotesTypeNamee").value), "", Rs1("NotesTypeNamee").value)
                   End If
                   Rs1.MoveNext
             Next i
              .AutoSize 0, .Cols - 1, False
        End With
          End If
     End If
 End Sub
  Sub FillGridDataWithAdd()
   Dim i As Integer
  Grid.Rows = Grid.Rows + 1
  i = Grid.Rows
  i = i - 1
  With Grid
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("IDBS")) = TxtSerial1.Text
                .TextMatrix(i, .ColIndex("MoveDT")) = DTPicker2.value
                .TextMatrix(i, .ColIndex("BankValue")) = Text1.Text
                .TextMatrix(i, .ColIndex("BankRF")) = oldTxtSerial1.Text
                .TextMatrix(i, .ColIndex("Explan")) = txtto.Text
  End With
        Text1.Text = ""
        oldTxtSerial1.Text = ""
        txtto.Text = ""
 End Sub

Private Sub ISButton2_Click()
        '+++++++++++++++++++++++++++++++++++++++++++++++
      If Text1.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «·ﬁÌ„… ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Text1.SetFocus
             Exit Sub
                 Else
            MsgBox "Write Value ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Text1.SetFocus
            Exit Sub
            End If
     End If
      ' If oldTxtSerial1.text = "" Then
      '  If SystemOptions.UserInterface = ArabicInterface Then
      '      MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· „—Ã⁄ «·»‰ﬂ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
      '      oldTxtSerial1.SetFocus
      '       Exit Sub
      '           Else
      '      MsgBox "Write Bank Reference ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
      '      oldTxtSerial1.SetFocus
      '      Exit Sub
      '      End If
    ' End If
    '    If txtto.text = "" Then
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «·‘—Õ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
    '        txtto.SetFocus
    '         Exit Sub
    '             Else
    '        MsgBox "Write Explanation ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '        txtto.SetFocus
    '        Exit Sub
    '        End If
    ' End If
     FillGridDataWithAdd
   End Sub
Private Sub ISButton3_Click()
On Error Resume Next
    With Me.Grid
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
 End Sub
 Private Sub ISButton4_Click()
 On Error Resume Next
 Me.Grid.Clear flexClearScrollable, flexClearEverything
 cleargriid
 End Sub
 Private Sub ISButton6_Click()
 SearchData
 End Sub
 Public Sub SearchData()
 VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid2.Rows = VSFlexGrid2.FixedRows
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    sql = "SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.[Value], dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate, "
    sql = sql & "                  dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Descriptione,"
    sql = sql & "                  dbo.DOUBLE_ENTREY_VOUCHERS.RecordDateH, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.Notes.ChqueNum,"
    sql = sql & "                  dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID,"
    sql = sql & "                  dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No, dbo.TblNotesTypes.NotesTypeName, dbo.TblNotesTypes.NotesTypeNamee, dbo.Notes.NoteSerial,"
    sql = sql & "                  dbo.Notes.NoteSerial1, dbo.Notes.NoteType"
    sql = sql & " FROM         dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN"
    sql = sql & "                  dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID INNER JOIN"
    sql = sql & "                  dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType"
               
       BolBegine = False
       StrWhere = ""

     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      StrWhere = StrWhere & " AND dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = N'" & get_bank_Account(val(Me.DcboBox.BoundText), "Account_Code") & "'"
      
   '-----------------------------------
    sql = sql & StrWhere
    sql = sql & " order by dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate ,dbo.Notes.ChqueNum, dbo.DOUBLE_ENTREY_VOUCHERS.[Value]"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
          Exit Sub
    Else
        With Me.VSFlexGrid2
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
             rs.MoveFirst
                 For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("NoteType")) = IIf(IsNull(rs("NoteType").value), "", rs("NoteType").value)
                .TextMatrix(i, .ColIndex("MoveDT")) = IIf(IsNull(rs("RecordDate").value), "", rs("RecordDate").value)
                .TextMatrix(i, .ColIndex("BankValue")) = IIf(IsNull(rs("Value").value), "", rs("Value").value)
                .TextMatrix(i, .ColIndex("BankRF")) = IIf(IsNull(rs("ChqueNum").value), "", rs("ChqueNum").value)
                .TextMatrix(i, .ColIndex("Explan")) = IIf(IsNull(rs("Double_Entry_Vouchers_Descriptione").value), "", rs("Double_Entry_Vouchers_Descriptione").value)
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
                .TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("NotesTypeName")) = IIf(IsNull(rs("NotesTypeName").value), "", rs("NotesTypeName").value)
                Else
                .TextMatrix(i, .ColIndex("NotesTypeName")) = IIf(IsNull(rs("NotesTypeNamee").value), "", rs("NotesTypeNamee").value)
                End If
                 rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
         End With
    End If
End Sub

Private Sub CMDSelectFile_Click()
CD1.ShowOpen
txtFile.Text = CD1.filename
 End Sub

' change date to hj
  Private Sub XPDtbTrans_Change()
  If Me.TxtModFlg.Text <> "R" Then
              Txt_DateHigri.value = ToHijriDate(XPDtbTrans.value)
   End If
   End Sub
' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
        If ChekClodePeriod(SettlementDT.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
      If ChekClodePeriod(SettlementDT.value) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                Else
                MsgBox "Please Change Date Becouse This is Period is Closed"
                End If
             Exit Sub
    End If
    '---------------------- check if data Vaclete -----------------------
      If dcBranch.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            dcBranch.SetFocus
            Exit Sub
            Else
            MsgBox "Write Branch Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
            dcBranch.SetFocus
         End If
     End If
    '+++++++++++++++++++++++++++++++++++++++++++++++
     If DcboBox.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «”„ «·»‰ﬂ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            DcboBox.SetFocus
             Exit Sub
             Else
            MsgBox "Write Bank Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcboBox.SetFocus
            Exit Sub
            End If
     End If
    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.Text
            '------------------------------ new record ----------------------------
        Case "N"
                  '------------------------- save record -----------------------------
          AddNewRecored
          AddNewRec
           
        '  BtnLast_Click
        Case "E"
            '----------------------------- save edit -------------------------------
            FiLLRec
    End Select
    Exit Sub
ErrTrap:
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TBLBankSettlement", "IDBS", "")
    Me.TxtSerial1.Text = StrRecID
    RsSavRec.AddNew
    RsSavRec.Fields("IDBS").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
' change id search
Private Sub TxtSerial1_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub
' search for select id
Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.find "IDBS=" & RecId, , adSearchForward, 1
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
  ' cancel camnd sub
  '+++++++++++++++++++++++++++++++
  Private Sub BtnCancel_Click()
    Unload Me
End Sub
' undo sub
 Private Sub BtnUndo_Click()
    FindRec val(TxtSerial1.Text)
    Me.TxtModFlg.Text = "R"
    FiLLTXT
     BtnLast_Click
End Sub
' delet sub
Private Sub btnDelete_Click()
        If ChekClodePeriod(SettlementDT.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    Dim X As Integer
    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If
    If X = vbNo Then Exit Sub
     If TxtSerial1.Text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
       End If
               Else
                RsSavRec.find "IDBS=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.delete
               '''''''''''''''''''''''''''''''
                 StrSQL = "Delete From TBLBankSettlementJoin Where IDBS='" & val(TxtSerial1.Text) & "'"
                 Cn.Execute StrSQL, , adExecuteNoRecords
                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Deletion Process Success ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               End If
               cleargriid
     End If
                            '------------------------------ Move Next ---------------------------.
        Me.Refresh
       ' FillGridWithData
        BtnNext_Click
     Exit Sub
ErrTrap:
     Select Case Err.Number
        Case -2147217873, -2147467259
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
           Cn.Errors.Clear
    End Select

End Sub
' exit without save sub
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
               btnSave_Click
        Case vbCancel
              Cancel = True
        End Select
    End If
    Exit Sub
ErrTrap:
End Sub
Private Sub Form_Terminate()
     ' Set FrmVacancy = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    If RsSavRec.State = adStateOpen Then
        If Not (RsSavRec.EOF Or RsSavRec.BOF) Then
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
        End If
        RsSavRec.Close
        Set RsSavRec = Nothing
    End If
ErrTrap:
End Sub
Private Sub Form_Activate()
    Me.ZOrder 0
End Sub
Public Sub EditRec(StrTable As String, _
                   RecId As String)
     FiLLRec
End Sub
'Private Sub Grid_EnterCell()
 '   On Error GoTo ErrTrap
  '  FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("Ser")))
'ErrTrap:
'End Sub
Private Sub TxtModFlg_Change()
    If TxtModFlg.Text = "N" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        ISButton1.Enabled = False
     '   Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
              
       
    ElseIf TxtModFlg.Text = "R" Then
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False
        If TxtSerial1.Text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
    End If
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
        ISButton1.Enabled = True
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
   ElseIf TxtModFlg.Text = "E" Then
       Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
    '    Grid.Enabled = False
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    End If
End Sub

' move btowen recored
Private Sub BtnFirst_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        cleargriid
        Exit Sub
    End If
BegnieWork:
If Me.check1.value = True Then
ISButton3.Enabled = True
ISButton4.Enabled = True
Else
ISButton3.Enabled = False
ISButton4.Enabled = False
End If
    RsSavRec.MoveFirst
    cleargriid
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnLast_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        cleargriid
        Exit Sub
    End If
BegnieWork:
If Me.check1.value = True Then
ISButton3.Enabled = True
ISButton4.Enabled = True
Else
ISButton3.Enabled = False
ISButton4.Enabled = False
End If
    RsSavRec.MoveLast
    cleargriid
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify_Click()
        If ChekClodePeriod(SettlementDT.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
    Dim Msg As String
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtSerial1.Text <> "" Then
        TxtModFlg = "E"
        Me.DCboUserName.BoundText = user_id
      '  Me.Dcbranch.BoundText = branch_id
        Frm2.Enabled = True
        Me.dcBranch.SetFocus
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
                    If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If
    End Select
End Sub
Private Sub btnNew_Click()
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    clear_all Me
    cleargriid
    Me.VSFlexGrid2.Rows = 1
    TxtModFlg.Text = "N"
    CmbType.ListIndex = 0
    Me.DCboUserName.BoundText = user_id
    Me.dcBranch.BoundText = branch_id
    CmbType.ListIndex = 0
    dcBranch.SetFocus
    check1.value = True
    Me.Grid.Clear flexClearScrollable, flexClearEverything
    Me.VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
    Me.Fg.Clear flexClearScrollable, flexClearEverything
     Fg.Rows = 1
    Me.Fg1.Clear flexClearScrollable, flexClearEverything
        Fg1.Rows = 1
ErrTrap:
End Sub
Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
      clear_all Me
      cleargriid
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
    cleargriid
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        cleargriid
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MovePrevious
    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If
     cleargriid
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
'Information for camand
'++++++++++++++++++++++++++++++++++++++
Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = CHR(13) + CHR(10)
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
             .AddControl btnNew, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        .AddControl btnLast, Msg, True
    End With
ErrTrap:
End Sub
' short cut for keys
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrTrap
    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            btnNew_Click
        Else
            SendKeys "{TAB}"
        End If
    End If
    'New ---------------------------
    If KeyCode = vbKeyF12 Then
        If btnNew.Enabled = False Then Exit Sub
        btnNew_Click
    End If
    'Edit ------------------------
    If KeyCode = vbKeyF11 Then
        If btnModify.Enabled = False Then Exit Sub
        btnModify_Click
    End If
    'save --------------------------------------------------------------------------------
    If KeyCode = vbKeyF10 Then
        If btnSave.Enabled = False Then Exit Sub
        btnSave_Click
    End If
    'undo ------------------------------------------------------------------------------
    If KeyCode = vbKeyF9 Then
        If BtnUndo.Enabled = False Then Exit Sub
        BtnUndo_Click
    End If
    'Delete ---------------------------------------------------------------------------
    If KeyCode = vbKeyF8 Then
        If btnDelete.Enabled = False Then Exit Sub
        btnDelete_Click
    End If
    'Exit ----------------------------------------------------------------------
    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If btnCancel.Enabled = False Then Exit Sub
            BtnCancel_Click
        End If
    End If
    'Moveing through Records ---------------------------------------------------------------------------
    'If TxtModFlg.Text = "R" Then
    'Move first --------------------------------------------
    If KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
        If btnFirst.Enabled = False Then Exit Sub
        BtnFirst_Click
    End If
    'Move Previous---------------------------------------------------------
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
        If btnPrevious.Enabled = False Then Exit Sub
        BtnPrevious_Click
    End If
    'Move Next---------------------------------------------------------
    If KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
        If btnNext.Enabled = False Then Exit Sub
        BtnNext_Click
    End If
    'Move Last---------------------------------------------------------
    If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
        If btnLast.Enabled = False Then Exit Sub
        BtnLast_Click
    End If
    'End If
    Exit Sub
ErrTrap:
End Sub
' print Events
'++++++++++++++++++++++++++++++++++++++++++
Private Sub BtnPrint_Click()
On Error GoTo ErrTrap
  If val(Me.TxtSerial1.Text) <> 0 Then
      print_report
  End If
ErrTrap:
End Sub
Private Sub ISButton1_Click()
On Error GoTo ErrTrap
   If val(Me.TxtSerial1.Text) <> 0 Then
       print_report
   End If
ErrTrap:
End Sub
Function print_report(Optional NoteSerial As String)
On Error GoTo ErrTrap
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
   sql = "SELECT     dbo.TBLBankSettlement.IDBS, dbo.TBLBankSettlement.DateM, dbo.TBLBankSettlement.DateH, dbo.TBLBankSettlement.BranchID, "
   sql = sql & "                  dbo.TBLBankSettlement.SettlementDT, dbo.TBLBankSettlement.BankID, dbo.TBLBankSettlement.FromDT, dbo.TBLBankSettlement.ToDT,"
   sql = sql & "                  dbo.TBLBankSettlement.EXPCheck, dbo.TBLBankSettlement.UserID, dbo.TBLBankSettlementJoin.ID, dbo.TBLBankSettlementJoin.IDBS AS IDBSJON,"
   sql = sql & "                   dbo.TBLBankSettlementJoin.MoveDT, dbo.TBLBankSettlementJoin.BankValue, dbo.TBLBankSettlementJoin.BankRF, dbo.TBLBankSettlementJoin.Explan,"
   sql = sql & "                   dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.BanksData.BankName, dbo.BanksData.BankNamee,"
   sql = sql & "                   dbo.TBLBankSettlementJoin.notes_id, dbo.TBLBankSettlementJoin.Double_Entry_Vouchers_ID, dbo.TBLBankSettlementJoin.RowG,"
   sql = sql & "                   dbo.TBLBankSettlementJoin.NoteType, dbo.TBLBankSettlementJoin.TypeTrans, dbo.TBLBankSettlementJoin.NoteSerial, dbo.TBLBankSettlementJoin.NoteSerial1,"
   sql = sql & "                    dbo.TblNotesTypes.NotesTypeName , dbo.TblNotesTypes.NotesTypeNameE"
   sql = sql & " FROM         dbo.TblNotesTypes RIGHT OUTER JOIN"
   sql = sql & "                     dbo.TBLBankSettlementJoin ON dbo.TblNotesTypes.NotesType = dbo.TBLBankSettlementJoin.NoteType RIGHT OUTER JOIN"
   sql = sql & "                     dbo.TblBranchesData RIGHT OUTER JOIN"
   sql = sql & "                     dbo.TBLBankSettlement LEFT OUTER JOIN"
   sql = sql & "                     dbo.BanksData ON dbo.TBLBankSettlement.BankID = dbo.BanksData.BankID ON dbo.TblBranchesData.branch_id = dbo.TBLBankSettlement.BranchID ON"
   sql = sql & "                     dbo.TBLBankSettlementJoin.IDBS = dbo.TBLBankSettlement.IDBS"
    sql = sql & " Where (dbo.TBLBankSettlement.IDBS = " & val(TxtSerial1.Text) & ")"
                      
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "BankSettlementtRPT.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "BankSettlementtRPTEN.rpt"
        End If
   If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
   Set RsData = New ADODB.Recordset
    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
     If RsData.BOF Or RsData.EOF Then
       Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
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
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        StrReportTitle = "" '& StrAccountName
        Else
         xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
          xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
ErrTrap:
  End Function
' chang langeg Event
'++++++++++++++++++++++++++++++++++++
'Private Sub TxtVacName_GotFocus()
 '   SwitchKeyboardLang LANG_ARABIC
'End Sub
'Private Sub TxtVacNamee_GotFocus()
'SwitchKeyboardLang LANG_ENGLISH
'End Sub
Private Sub ChangeLang()
On Error GoTo ErrTrap
   ' form name
    Me.Caption = "Bank Settlement"
    ' labell name
    CmdImport.Caption = "Upload File"
    CMDSelectFile.Caption = "Select File"
    lbl(2).Caption = "Registered on the program and does not exist in the bank movements"
    lbl(4).Caption = "Registered on the bank and does not exist in the program movements"
    Me.Label1(2).Caption = Me.Caption
    Me.lblcode.Caption = "Code"
    Me.lbldate.Caption = "Date"
    Me.lblhjdate.Caption = "HJ Date"
    Me.Label3.Caption = "Branch"
    Me.lbl(15).Caption = "Settlement Date"
    Me.Labelbank.Caption = "Select Bank "
    '''''''''''''' next
    Me.lbl(13).Caption = "Date Between"
    Me.lbl(3).Caption = "From"
    Me.lbl(5).Caption = "To"
    '''''''''''''''''''''''' next
    Me.lbl(1).Caption = "Settlement Type"
    Me.check1.Caption = "Manual"
    Me.check2.Caption = "Import File"
   ' ISButton5.Caption = "Download File"
    Me.lblmovment.Caption = "Movement Date"
    Me.Labelbnkrf.Caption = "Bank Reference"
    Me.Label1(9).Caption = "Value"
    Me.lbl(0).Caption = "Explanation"
    Me.ISButton2.Caption = "Add"
    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "NO. Recordes"
    Me.lbl(8).Caption = "by"
    '''''''''''''''''''''''''''''''' next
    ISButton6.Caption = "Search"
    ISButton3.Caption = "Delet Select"
    ISButton4.Caption = "Delet All"
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    ISButton1.Caption = "Print"
    btnQuery.Caption = "Search"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
    BtnShow.Caption = "Comparison"
    With Me.Grid
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("IDBS")) = "Code"
        .TextMatrix(0, .ColIndex("MoveDT")) = "Movement Date"
        .TextMatrix(0, .ColIndex("BankValue")) = "Value"
        .TextMatrix(0, .ColIndex("BankRF")) = "Bank Reference"
        .TextMatrix(0, .ColIndex("Explan")) = "Explanation"
        '''''''''''''''''''''''
       End With
   With Me.VSFlexGrid2
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("IDBS")) = "Code"
        .TextMatrix(0, .ColIndex("MoveDT")) = "Date"
        .TextMatrix(0, .ColIndex("BankValue")) = "Value"
        .TextMatrix(0, .ColIndex("BankRF")) = "Check NO."
        .TextMatrix(0, .ColIndex("Explan")) = "Explanation"
           .TextMatrix(0, .ColIndex("NoteSerial1")) = "Trans NO."
        .TextMatrix(0, .ColIndex("NotesTypeName")) = "Trans Type"
        .TextMatrix(0, .ColIndex("NoteSerial")) = "No.Entry"
  End With
         With Me.Fg1
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("IDBS")) = "Code"
        .TextMatrix(0, .ColIndex("MoveDT")) = "Movement Date"
        .TextMatrix(0, .ColIndex("BankValue")) = "Value"
        .TextMatrix(0, .ColIndex("BankRF")) = "Bank Reference"
        .TextMatrix(0, .ColIndex("Explan")) = "Explanation"
        '''''''''''''''''''''''
       End With
   With Me.Fg
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("IDBS")) = "Code"
        .TextMatrix(0, .ColIndex("MoveDT")) = "Date"
        .TextMatrix(0, .ColIndex("BankValue")) = "Value"
        .TextMatrix(0, .ColIndex("BankRF")) = "Check NO."
        .TextMatrix(0, .ColIndex("Explan")) = "Explanation"
           .TextMatrix(0, .ColIndex("NoteSerial1")) = "Trans NO."
        .TextMatrix(0, .ColIndex("NotesTypeName")) = "Trans Type"
        .TextMatrix(0, .ColIndex("NoteSerial")) = "No.Entry"
  End With
ErrTrap:
End Sub
' check box event
'+++++++++++++++++++++++++++++++++++
  Private Sub Check1_Click()
  On Error GoTo ErrTrap
  If check1.value = True Then
  CmdImport.Enabled = False
  DTPicker2.Enabled = True
 CMDSelectFile.Enabled = False
  Text1.Enabled = True
  oldTxtSerial1.Enabled = True
  txtto.Enabled = True
  Text1.Text = ""
  oldTxtSerial1.Text = ""
  txtto.Text = ""
  ISButton3.Enabled = True
  ISButton4.Enabled = True
  ISButton2.Enabled = True
  Grid.Enabled = True
  Me.Grid.Clear flexClearScrollable, flexClearEverything
  End If
ErrTrap:
 End Sub
 Private Sub Check2_Click()
  On Error GoTo ErrTrap
  If check2.value = True Then
  DTPicker2.Enabled = False
CmdImport.Enabled = True
CMDSelectFile.Enabled = True
  Text1.Enabled = False
  oldTxtSerial1.Enabled = False
  txtto.Enabled = False
  Text1.Text = ""
  oldTxtSerial1.Text = ""
  txtto.Text = ""
  ISButton3.Enabled = False
  ISButton4.Enabled = False
  ISButton2.Enabled = False
  Grid.Enabled = False
  Me.Grid.Clear flexClearScrollable, flexClearEverything
  End If
ErrTrap:
End Sub
Private Sub Dcbranch_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  DcboBox.SetFocus
  End If
ErrTrap:
End Sub
Private Sub DcboBox_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  Text1.SetFocus
  End If
ErrTrap:
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrTrap
   If KeyAscii = 13 Then
   oldTxtSerial1.SetFocus
   Else
   If KeyAscii = 8 Then
     If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
   KeyAscii = 0
    End If
    End If
    End If
ErrTrap:
End Sub
Private Sub oldTxtSerial1_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  txtto.SetFocus
  End If
ErrTrap:
End Sub
Private Sub txtto_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  Call ISButton2_Click
  End If
ErrTrap:
End Sub
Private Sub cleargriid()
Me.Grid.Rows = 1
End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TBLBankSettlement"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub

'+++++++++++++++++++++++++++++++++ end






