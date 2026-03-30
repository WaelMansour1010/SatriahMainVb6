VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmFinaStatem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·ﬁÊ«∆„ «·„«·Ì…"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13545
   Icon            =   "FrmFinaStatem.frx":0000
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   9075
   ScaleWidth      =   13545
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   23640
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   960
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmFinaStatem.frx":6852
      Left            =   23640
      List            =   "FrmFinaStatem.frx":6862
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2040
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   23640
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Text            =   "modflag"
      Top             =   1320
      Visible         =   0   'False
      Width           =   465
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   9075
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13545
      _cx             =   23892
      _cy             =   16007
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
      Begin C1SizerLibCtl.C1Elastic Frm2 
         Height          =   585
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   2670
         Width           =   13500
         _cx             =   23813
         _cy             =   1032
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
         Begin VB.TextBox TxtRemarks 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   285
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   59
            Top             =   120
            Width           =   4725
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«»«  «·‘—ﬂ…"
            Height          =   975
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   1920
            Visible         =   0   'False
            Width           =   3375
            Begin MSDataListLib.DataCombo DcbAccount4 
               Height          =   315
               Left            =   120
               TabIndex        =   40
               Top             =   240
               Width           =   2805
               _ExtentX        =   4948
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbAccount3 
               Height          =   315
               Left            =   120
               TabIndex        =   41
               Top             =   600
               Width           =   2805
               _ExtentX        =   4948
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Õ”«» œ«∆‰"
               Height          =   315
               Index           =   13
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   600
               Width           =   1380
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Õ”«» „œÌ‰"
               Height          =   315
               Index           =   14
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   240
               Width           =   1380
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰”»… «·‘—ﬂ…"
            Height          =   975
            Left            =   10320
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   1920
            Visible         =   0   'False
            Width           =   1935
            Begin VB.TextBox TxtStay1 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   360
               TabIndex        =   34
               Top             =   600
               Width           =   810
            End
            Begin VB.TextBox TxtCivilin1 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   360
               TabIndex        =   33
               Top             =   240
               Width           =   810
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " %"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   315
               Index           =   12
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   600
               Width           =   300
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " %"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   315
               Index           =   11
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   240
               Width           =   300
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " ‰”»… «·„ﬁÌ„Ì‰"
               Height          =   315
               Index           =   10
               Left            =   1035
               RightToLeft     =   -1  'True
               TabIndex        =   36
               Top             =   600
               Width           =   1500
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " ‰”»… «·„Ê«ÿ‰Ì‰"
               Height          =   315
               Index           =   9
               Left            =   1035
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   240
               Width           =   1500
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«»«  «·„ÊŸ›"
            Height          =   975
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   1920
            Visible         =   0   'False
            Width           =   3495
            Begin MSDataListLib.DataCombo DcbAccount2 
               Height          =   315
               Left            =   120
               TabIndex        =   28
               Top             =   240
               Width           =   2805
               _ExtentX        =   4948
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbAccount1 
               Height          =   315
               Left            =   120
               TabIndex        =   29
               Top             =   600
               Width           =   2805
               _ExtentX        =   4948
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Õ”«» œ«∆‰"
               Height          =   315
               Index           =   7
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   600
               Width           =   1380
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Õ”«» „œÌ‰"
               Height          =   315
               Index           =   6
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   240
               Width           =   1380
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰”»… «·„ÊŸ›"
            Height          =   975
            Left            =   8160
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   1800
            Visible         =   0   'False
            Width           =   2175
            Begin VB.TextBox TxtCivilin 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   360
               TabIndex        =   23
               Top             =   240
               Width           =   810
            End
            Begin VB.TextBox TxtStay 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   360
               TabIndex        =   22
               Top             =   600
               Width           =   810
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " ‰”»… «·„Ê«ÿ‰Ì‰"
               Height          =   315
               Index           =   3
               Left            =   1275
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   240
               Width           =   1500
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " ‰”»… «·„ﬁÌ„Ì‰"
               Height          =   315
               Index           =   1
               Left            =   1275
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   600
               Width           =   1500
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " %"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   315
               Index           =   4
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   240
               Width           =   300
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " %"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   315
               Index           =   5
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   600
               Width           =   300
            End
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   11865
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   135
            Width           =   1125
         End
         Begin MSComCtl2.DTPicker RecordDate 
            Height          =   360
            Left            =   9480
            TabIndex        =   3
            Top             =   120
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   635
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   95879169
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   10920
            TabIndex        =   4
            Top             =   1800
            Visible         =   0   'False
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   5640
            TabIndex        =   44
            Top             =   120
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   270
            Left            =   4815
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   120
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„‰›–"
            Height          =   270
            Index           =   8
            Left            =   8415
            TabIndex        =   45
            Top             =   120
            Width           =   885
         End
         Begin VB.Label Labelbank 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·„›—œ"
            Height          =   255
            Left            =   12480
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   1800
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.Label lblcode 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„"
            Height          =   270
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   120
            Width           =   420
         End
         Begin VB.Label lbldate 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «· ‰›Ì–"
            Height          =   360
            Left            =   10710
            TabIndex        =   5
            Top             =   150
            Width           =   1260
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   540
         Left            =   -120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Width           =   13980
         _cx             =   24659
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
            Height          =   210
            Left            =   600
            TabIndex        =   9
            Top             =   150
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   370
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
            ButtonImage     =   "FrmFinaStatem.frx":687B
            ColorButton     =   16777215
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   210
            Left            =   1950
            TabIndex        =   10
            Top             =   150
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   370
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
            ButtonImage     =   "FrmFinaStatem.frx":6C15
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   210
            Left            =   1020
            TabIndex        =   11
            Top             =   150
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   370
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
            ButtonImage     =   "FrmFinaStatem.frx":6FAF
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   210
            Left            =   1515
            TabIndex        =   12
            Top             =   150
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   370
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
            ButtonImage     =   "FrmFinaStatem.frx":7349
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Image Image1 
            Height          =   390
            Left            =   11085
            Picture         =   "FrmFinaStatem.frx":76E3
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1005
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·ﬁÊ«∆„ «·„«·Ì…"
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
            Left            =   7350
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   75
            Width           =   3420
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic8 
         Height          =   870
         Left            =   120
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   8235
         Width           =   13500
         _cx             =   23813
         _cy             =   1535
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
            Height          =   435
            Left            =   12015
            TabIndex        =   47
            Top             =   285
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   767
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
            ButtonImage     =   "FrmFinaStatem.frx":CEB5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   435
            Left            =   9450
            TabIndex        =   48
            Top             =   285
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   767
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
            ButtonImage     =   "FrmFinaStatem.frx":13717
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   435
            Left            =   10845
            TabIndex        =   49
            Top             =   285
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   767
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
            ButtonImage     =   "FrmFinaStatem.frx":13AB1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   435
            Left            =   7995
            TabIndex        =   50
            Top             =   285
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   767
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
            ButtonImage     =   "FrmFinaStatem.frx":1A313
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   435
            Left            =   5685
            TabIndex        =   51
            Top             =   285
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   767
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
            ButtonImage     =   "FrmFinaStatem.frx":1A6AD
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   435
            Left            =   4575
            TabIndex        =   52
            Top             =   285
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   767
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
            ButtonImage     =   "FrmFinaStatem.frx":1AC47
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   435
            Left            =   6930
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   285
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   767
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
            ButtonImage     =   "FrmFinaStatem.frx":1AFE1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   435
            Left            =   3330
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   285
            Visible         =   0   'False
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   767
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
            ButtonImage     =   "FrmFinaStatem.frx":21843
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic7 
            Height          =   570
            Left            =   0
            TabIndex        =   113
            TabStop         =   0   'False
            Top             =   240
            Width           =   4485
            _cx             =   7911
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
            Begin VB.Label LabCountRec 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00C00000&
               Height          =   240
               Left            =   150
               RightToLeft     =   -1  'True
               TabIndex        =   117
               Top             =   120
               Width           =   795
            End
            Begin VB.Label LabCurrRec 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               ForeColor       =   &H00800000&
               Height          =   240
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   116
               Top             =   120
               Width           =   705
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   240
               Index           =   1
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   115
               Top             =   120
               Width           =   960
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   240
               Index           =   0
               Left            =   3135
               RightToLeft     =   -1  'True
               TabIndex        =   114
               Top             =   120
               Width           =   1200
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   780
         Left            =   240
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   8145
         Visible         =   0   'False
         Width           =   13500
         _cx             =   23813
         _cy             =   1376
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
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   270
            Left            =   8745
            TabIndex        =   56
            ToolTipText     =   "Õ–› «·’› «·Õ«·Ì"
            Top             =   225
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   476
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–› «·’› «·Õ«·Ì"
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
            ButtonImage     =   "FrmFinaStatem.frx":21BDD
            ButtonImageDisabled=   "FrmFinaStatem.frx":2843F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   270
            Left            =   7125
            TabIndex        =   57
            ToolTipText     =   "Õ–› «·ﬂ·"
            Top             =   225
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   476
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–› «·ﬂ· "
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
            ButtonImage     =   "FrmFinaStatem.frx":47629
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid Grid 
         Height          =   2055
         Left            =   120
         TabIndex        =   58
         Top             =   600
         Width           =   13500
         _cx             =   23812
         _cy             =   3625
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
         BackColorAlternate=   16777088
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmFinaStatem.frx":4DE8B
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   1815
         Left            =   6480
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   3240
         Width           =   6165
         _cx             =   10874
         _cy             =   3201
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
         Begin VB.ListBox ListActivitySelected 
            Height          =   1425
            ItemData        =   "FrmFinaStatem.frx":4DF45
            Left            =   120
            List            =   "FrmFinaStatem.frx":4DF4C
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   330
            Width           =   2715
         End
         Begin VB.ListBox ListAllActivity 
            Height          =   1425
            ItemData        =   "FrmFinaStatem.frx":4DF66
            Left            =   3300
            List            =   "FrmFinaStatem.frx":4DF6D
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   330
            Width           =   2715
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «·‰‘«ÿ"
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   2595
            TabIndex        =   75
            Top             =   0
            Width           =   1470
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   2820
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   420
            Width           =   495
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   2820
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   750
            Width           =   495
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "<<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   2820
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   1305
            Width           =   495
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   2820
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   975
            Width           =   495
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   1815
         Left            =   120
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   3240
         Width           =   6285
         _cx             =   11086
         _cy             =   3201
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
         Begin VB.ListBox ListBranchAll 
            Height          =   1425
            ItemData        =   "FrmFinaStatem.frx":4DF82
            Left            =   3360
            List            =   "FrmFinaStatem.frx":4DF89
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   330
            Width           =   2775
         End
         Begin VB.ListBox ListBranchSelected 
            Height          =   1425
            ItemData        =   "FrmFinaStatem.frx":4DF9C
            Left            =   120
            List            =   "FrmFinaStatem.frx":4DFA3
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   330
            Width           =   2775
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «·›—⁄"
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   2520
            TabIndex        =   76
            Top             =   0
            Width           =   1500
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   1095
            Width           =   495
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Caption         =   "<<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   1425
            Width           =   495
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   750
            Width           =   495
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   420
            Width           =   495
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
         Height          =   2415
         Left            =   120
         TabIndex        =   77
         Top             =   5760
         Width           =   13500
         _cx             =   23812
         _cy             =   4260
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
         BackColorAlternate=   16777088
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
         Rows            =   2
         Cols            =   10
         FixedRows       =   2
         FixedCols       =   1
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmFinaStatem.frx":4DFBB
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
            Left            =   960
            TabIndex        =   78
            Top             =   720
            Visible         =   0   'False
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   1085
            _Version        =   393216
            Appearance      =   0
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   585
         Left            =   120
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   5160
         Width           =   13500
         _cx             =   23813
         _cy             =   1032
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
         Begin VB.Frame Frame8 
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰”»… «·„ÊŸ›"
            Height          =   975
            Left            =   8160
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   1800
            Visible         =   0   'False
            Width           =   2175
            Begin VB.TextBox Text6 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   360
               TabIndex        =   99
               Top             =   600
               Width           =   810
            End
            Begin VB.TextBox Text5 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   360
               TabIndex        =   98
               Top             =   240
               Width           =   810
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " %"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   315
               Index           =   24
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   600
               Width           =   300
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " %"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   315
               Index           =   23
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   240
               Width           =   300
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " ‰”»… «·„ﬁÌ„Ì‰"
               Height          =   315
               Index           =   22
               Left            =   1275
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   600
               Width           =   1500
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " ‰”»… «·„Ê«ÿ‰Ì‰"
               Height          =   315
               Index           =   21
               Left            =   1275
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   240
               Width           =   1500
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«»«  «·„ÊŸ›"
            Height          =   975
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   1920
            Visible         =   0   'False
            Width           =   3495
            Begin MSDataListLib.DataCombo DataCombo3 
               Height          =   315
               Left            =   120
               TabIndex        =   93
               Top             =   240
               Width           =   2805
               _ExtentX        =   4948
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo4 
               Height          =   315
               Left            =   120
               TabIndex        =   94
               Top             =   600
               Width           =   2805
               _ExtentX        =   4948
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Õ”«» „œÌ‰"
               Height          =   315
               Index           =   20
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   240
               Width           =   1380
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Õ”«» œ«∆‰"
               Height          =   315
               Index           =   19
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   600
               Width           =   1380
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰”»… «·‘—ﬂ…"
            Height          =   975
            Left            =   10320
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   1920
            Visible         =   0   'False
            Width           =   1935
            Begin VB.TextBox Text4 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   360
               TabIndex        =   87
               Top             =   240
               Width           =   810
            End
            Begin VB.TextBox Text3 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   360
               TabIndex        =   86
               Top             =   600
               Width           =   810
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " ‰”»… «·„Ê«ÿ‰Ì‰"
               Height          =   315
               Index           =   18
               Left            =   1035
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   240
               Width           =   1500
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " ‰”»… «·„ﬁÌ„Ì‰"
               Height          =   315
               Index           =   17
               Left            =   1035
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   600
               Width           =   1500
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " %"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   315
               Index           =   16
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   240
               Width           =   300
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " %"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000040C0&
               Height          =   315
               Index           =   15
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   600
               Width           =   300
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«»«  «·‘—ﬂ…"
            Height          =   975
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   1920
            Visible         =   0   'False
            Width           =   3375
            Begin MSDataListLib.DataCombo DataCombo1 
               Height          =   315
               Left            =   120
               TabIndex        =   81
               Top             =   240
               Width           =   2805
               _ExtentX        =   4948
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo2 
               Height          =   315
               Left            =   120
               TabIndex        =   82
               Top             =   600
               Width           =   2805
               _ExtentX        =   4948
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Õ”«» „œÌ‰"
               Height          =   315
               Index           =   2
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   240
               Width           =   1380
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Õ”«» œ«∆‰"
               Height          =   315
               Index           =   0
               Left            =   2640
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   600
               Width           =   1380
            End
         End
         Begin MSComCtl2.DTPicker FromDate 
            Height          =   360
            Left            =   10200
            TabIndex        =   104
            Top             =   120
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   635
            _Version        =   393216
            Format          =   95879169
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DataCombo5 
            Height          =   315
            Left            =   10920
            TabIndex        =   105
            Top             =   1800
            Visible         =   0   'False
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   360
            Left            =   120
            TabIndex        =   110
            Top             =   120
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   635
            ButtonPositionImage=   1
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
            ButtonImage     =   "FrmFinaStatem.frx":4E162
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSComCtl2.DTPicker ToDate 
            Height          =   360
            Left            =   7560
            TabIndex        =   111
            Top             =   120
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   635
            _Version        =   393216
            Format          =   95879169
            CurrentDate     =   38784
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï"
            Height          =   270
            Left            =   9360
            RightToLeft     =   -1  'True
            TabIndex        =   109
            Top             =   120
            Width           =   540
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   270
            Left            =   11880
            RightToLeft     =   -1  'True
            TabIndex        =   108
            Top             =   120
            Width           =   540
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —…"
            Height          =   270
            Left            =   12720
            RightToLeft     =   -1  'True
            TabIndex        =   107
            Top             =   135
            Width           =   540
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·„›—œ"
            Height          =   255
            Left            =   12480
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Top             =   1800
            Visible         =   0   'False
            Width           =   1050
         End
      End
      Begin XtremeSuiteControls.CheckBox CheckAll 
         Height          =   1455
         Left            =   12720
         TabIndex        =   112
         Top             =   3480
         Width           =   735
         _Version        =   786432
         _ExtentX        =   1296
         _ExtentY        =   2566
         _StockProps     =   79
         Caption         =   "«·ﬂ·"
         BackColor       =   -2147483635
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   20280
      TabIndex        =   17
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
      TabIndex        =   18
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
            Picture         =   "FrmFinaStatem.frx":4E4FC
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFinaStatem.frx":4E896
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFinaStatem.frx":4EC30
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFinaStatem.frx":4EFCA
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFinaStatem.frx":4F364
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFinaStatem.frx":4F6FE
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFinaStatem.frx":4FA98
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFinaStatem.frx":50032
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
      TabIndex        =   19
      Top             =   -960
      Width           =   855
   End
End
Attribute VB_Name = "FrmFinaStatem"
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

Sub Grd()
    With Fg_Journal

        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        .MergeCol(.ColIndex("DebitValue")) = True
        .MergeCol(.ColIndex("CreditValue")) = True
        .MergeCol(.ColIndex("Account_Serial")) = True
        .MergeCol(.ColIndex("AccountName")) = True
        .MergeCol(.ColIndex("DebitValue1")) = True
        .MergeCol(.ColIndex("CreditValue1")) = True
        .MergeCol(.ColIndex("DebitValue2")) = True
        .MergeCol(.ColIndex("CreditValue2")) = True
        If SystemOptions.UserInterface = ArabicInterface Then
        .Cell(flexcpText, 0, .ColIndex("Ser"), 1, .ColIndex("Ser")) = "„"
        .Cell(flexcpText, 0, .ColIndex("Account_Serial"), 1, .ColIndex("Account_Serial")) = "ﬂÊœ «·Õ”«»"
        .Cell(flexcpText, 0, .ColIndex("AccountName"), 1, .ColIndex("AccountName")) = "«”„ «·Õ”«»"
        .Cell(flexcpText, 0, .ColIndex("DebitValue"), 0, .ColIndex("CreditValue")) = " «·—’Ìœ «·«›  «ÕÌ "
        .Cell(flexcpText, 1, .ColIndex("DebitValue"), 1, .ColIndex("DebitValue")) = "„œÌ‰"
        .Cell(flexcpText, 1, .ColIndex("CreditValue"), 1, .ColIndex("CreditValue")) = "œ«∆‰"
        .Cell(flexcpText, 0, .ColIndex("DebitValue1"), 0, .ColIndex("CreditValue1")) = " «·Õ—ﬂ… "
        .Cell(flexcpText, 1, .ColIndex("DebitValue1"), 1, .ColIndex("DebitValue1")) = "„œÌ‰"
        .Cell(flexcpText, 1, .ColIndex("CreditValue1"), 1, .ColIndex("CreditValue1")) = "œ«∆‰"
        .Cell(flexcpText, 0, .ColIndex("DebitValue2"), 0, .ColIndex("CreditValue2")) = " «·—’Ìœ "
        .Cell(flexcpText, 1, .ColIndex("DebitValue2"), 1, .ColIndex("DebitValue2")) = "„œÌ‰"
        .Cell(flexcpText, 1, .ColIndex("CreditValue2"), 1, .ColIndex("CreditValue2")) = "œ«∆‰"
        Else
        .Cell(flexcpText, 0, .ColIndex("Ser"), 1, .ColIndex("Ser")) = "Serial"
        .Cell(flexcpText, 0, .ColIndex("Account_Serial"), 1, .ColIndex("Account_Serial")) = "Account Code"
        .Cell(flexcpText, 0, .ColIndex("AccountName"), 1, .ColIndex("AccountName")) = "Account Name"
        .Cell(flexcpText, 0, .ColIndex("DebitValue"), 0, .ColIndex("CreditValue")) = " Opening Balance"
        .Cell(flexcpText, 1, .ColIndex("DebitValue"), 1, .ColIndex("DebitValue")) = "Debit"
        .Cell(flexcpText, 1, .ColIndex("CreditValue"), 1, .ColIndex("CreditValue")) = "Credit"
        .Cell(flexcpText, 0, .ColIndex("DebitValue1"), 0, .ColIndex("CreditValue1")) = " Transactions "
        .Cell(flexcpText, 1, .ColIndex("DebitValue1"), 1, .ColIndex("DebitValue1")) = "Debit"
        .Cell(flexcpText, 1, .ColIndex("CreditValue1"), 1, .ColIndex("CreditValue1")) = "Credit"
        .Cell(flexcpText, 0, .ColIndex("DebitValue2"), 0, .ColIndex("CreditValue2")) = " Balance "
        .Cell(flexcpText, 1, .ColIndex("DebitValue2"), 1, .ColIndex("DebitValue2")) = "Debit"
        .Cell(flexcpText, 1, .ColIndex("CreditValue2"), 1, .ColIndex("CreditValue2")) = "Credit"
        End If
        .Cell(flexcpAlignment, 0, .ColIndex("DebitValue1"), 0, .ColIndex("CreditValue1")) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, .ColIndex("DebitValue2"), 0, .ColIndex("CreditValue2")) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, .ColIndex("DebitValue"), 0, .ColIndex("CreditValue")) = flexAlignCenterCenter
        .ColWidth(.ColIndex("Account_Serial")) = 1500
        .ColWidth(.ColIndex("AccountName")) = 3000
        .ColWidth(.ColIndex("CreditValue")) = 1590
        .ColWidth(.ColIndex("DebitValue")) = 1590
        .ColFormat(.ColIndex("DebitValue")) = "#,###.00"
        .ColFormat(.ColIndex("CreditValue")) = "#,###.00"
        .ColFormat(.ColIndex("DebitValue1")) = "#,###.00"
        .ColFormat(.ColIndex("CreditValue1")) = "#,###.00"
        .ColFormat(.ColIndex("DebitValue2")) = "#,###.00"
        .ColFormat(.ColIndex("CreditValue2")) = "#,###.00"
       ' Set .WallPaper = GrdBck.Picture
    End With
End Sub

Private Sub CheckAll_Click()
If CheckAll.value = vbChecked Then
Label7_Click
Label9_Click
End If
End Sub

    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    conection = "select * from TblFinaStatem order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Grd
    Resize_Form Me
   'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetUsers Me.DCboUserName
FillMylist
    BtnLast_Click
    ShowTip
     If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        SwitchKeyboardLang LANG_ENGLISH
        Else
        SwitchKeyboardLang LANG_ARABIC
    End If
    FullGridData
    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If
       Me.Refresh
ErrTrap:
End Sub
' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Public Sub FiLLRec()
    On Error GoTo ErrTrap
    If TxtModFlg = "E" Then
    StrSQL = "Delete From TblFinaStatemDet Where FinStID='" & val(TxtSerial1.Text) & "'"
    Cn.Execute StrSQL, , adExecuteNoRecords
    End If
    RsSavRec.Fields("RecordDate").value = RecordDate.value
    RsSavRec.Fields("Remarks").value = TxtRemarks.Text
    RsSavRec.Fields("FromDate").value = FromDate.value
    RsSavRec.Fields("ToDate").value = ToDate.value
    If Me.CheckAll.value = vbChecked Then
    RsSavRec.Fields("SelectAll").value = 1
    Else
    RsSavRec.Fields("SelectAll").value = Null
    End If
    
    RsSavRec.Fields("UserID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    RsSavRec.update
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblFinaStatemDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    With Fg_Journal
       For i = .FixedRows To .Rows - 1
         If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                RsDevsub.AddNew
                RsDevsub("FinStID").value = Me.TxtSerial1.Text
                RsDevsub("AccountCode").value = IIf((.TextMatrix(i, .ColIndex("AccountCode"))) = "", Null, .TextMatrix(i, .ColIndex("AccountCode")))
                RsDevsub("TypeTrans").value = 0
                RsDevsub("DebitValue").value = IIf((.TextMatrix(i, .ColIndex("DebitValue"))) = "", Null, val(.TextMatrix(i, .ColIndex("DebitValue"))))
                RsDevsub("CreditValue").value = IIf((.TextMatrix(i, .ColIndex("CreditValue"))) = "", Null, val(.TextMatrix(i, .ColIndex("CreditValue"))))
                RsDevsub("DebitValue1").value = IIf((.TextMatrix(i, .ColIndex("DebitValue1"))) = "", Null, val(.TextMatrix(i, .ColIndex("DebitValue1"))))
                RsDevsub("CreditValue1").value = IIf((.TextMatrix(i, .ColIndex("CreditValue1"))) = "", Null, val(.TextMatrix(i, .ColIndex("CreditValue1"))))
                RsDevsub("DebitValue2").value = IIf((.TextMatrix(i, .ColIndex("DebitValue2"))) = "", Null, val(.TextMatrix(i, .ColIndex("DebitValue2"))))
                RsDevsub("CreditValue2").value = IIf((.TextMatrix(i, .ColIndex("CreditValue2"))) = "", Null, val(.TextMatrix(i, .ColIndex("CreditValue2"))))
                RsDevsub.update
        End If
      Next i
     End With
               Set RsDevsub = New ADODB.Recordset
   StrSQL = "SELECT     *  from dbo.TblFinaStatemDet Where (1 = -1)"
   RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        For i = 0 To Me.ListActivitySelected.ListCount - 1
             RsDevsub.AddNew
             RsDevsub("FinStID").value = val(Me.TxtSerial1.Text)
             RsDevsub("BranchID").value = val(ListActivitySelected.ItemData(i))
             RsDevsub("TypeTrans").value = 1
             RsDevsub.update
       Next i
     
     
           Set RsDevsub = New ADODB.Recordset
   StrSQL = "SELECT     *  from dbo.TblFinaStatemDet Where (1 = -1)"
   RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        For i = 0 To Me.ListBranchSelected.ListCount - 1
             RsDevsub.AddNew
             RsDevsub("FinStID").value = val(Me.TxtSerial1.Text)
             RsDevsub("BranchID").value = val(ListBranchSelected.ItemData(i))
             RsDevsub("TypeTrans").value = 2
             RsDevsub.update
       Next i
    
 
     
      Select Case Me.TxtModFlg.Text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " Saved... " & Chr(13)
                Msg = Msg + "Do you want to enter another operation?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                FullGridData
                TxtModFlg = "R"
                 If SystemOptions.UserInterface = ArabicInterface Then
             Else
               Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                FullGridData
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                FiLLTXT
                FullGridData
                TxtModFlg = "R"
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                FullGridData
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                FullGridData
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
   ' ProgressBar1.Visible = True
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value) ': ProgressBar1.value = 10
    RecordDate.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value) ': ProgressBar1.value = 20
    TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value) ': ProgressBar1.value = 30
    FromDate.value = IIf(IsNull(RsSavRec.Fields("FromDate").value), Date, RsSavRec.Fields("FromDate").value) ': ProgressBar1.value = 40
    ToDate.value = IIf(IsNull(RsSavRec.Fields("ToDate").value), Date, RsSavRec.Fields("ToDate").value) ': ProgressBar1.value = 50
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    If Not IsNull(RsSavRec.Fields("SelectAll").value) Then
    If (RsSavRec.Fields("SelectAll").value) = 1 Then
    Me.CheckAll.value = vbChecked
    Else
    Me.CheckAll.value = vbUnchecked
    End If
    Else
    Me.CheckAll.value = vbUnchecked
    End If
    LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
    LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 80
    FullGridListData
    With Grid

        For i = 1 To .Rows - 1

            If Trim(TxtSerial1.Text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial1.Text = .TextMatrix(i, .ColIndex("id"))
                .Row = i
                Exit Sub
            End If

        Next

    End With
    
   ' ProgressBar1.Visible = False
   ' ProgressBar1.value = 0
ErrTrap:
 ' ProgressBar1.Visible = False
' ProgressBar1.value = 0
End Sub
 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim Sql As String
  Sql = " SELECT     dbo.TblFinaStatem.RecordDate, dbo.TblUsers.UserName, dbo.TblFinaStatem.UserID, dbo.TblFinaStatem.Remarks, dbo.TblFinaStatem.FromDate, "
  Sql = Sql & "                     dbo.TblFinaStatem.ToDate , dbo.TblFinaStatem.ID"
  Sql = Sql & " FROM         dbo.TblFinaStatem INNER JOIN"
  Sql = Sql & "                    dbo.TblUsers ON dbo.TblFinaStatem.UserID = dbo.TblUsers.UserID"
  Rs1.Open Sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.Grid
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(Rs1("RecordDate").value), "", Rs1("RecordDate").value)
                   .TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(Rs1("UserName").value), "", Rs1("UserName").value)
                   .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(Rs1("Remarks").value), "", Rs1("Remarks").value)
                   .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(Rs1("ID").value), 0, Rs1("ID").value)
                   Rs1.MoveNext
             Next i
        End With
        Exit Sub
ErrTrap:
    End Sub
    
 Sub FullGrid()
  On Error GoTo ErrTrap
  Dim AccountTypes As Integer
  Dim total As Double
  Dim OpeningBalancebeformdateMinus1 As Double
  Dim OpeningBalancebeformStartCurrentyearTOFromDAteminus1 As Double
  Dim NewOpinning As Double
  Dim OpeningBalance As Double
  Dim ProfitBalance As Double
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim Sql As String
  Dim i As Integer
  Dim BranchID As String
  Dim HideZeroBalance As Boolean
   Dim openingBalanceDate As Date
   Dim FromdateMinus1 As Date
   Dim StartCurrentDate As Date
   
   FromdateMinus1 = DateAdd("d", -1, FromDate.value)
    getFirstPeriodDateInthisYear2 openingBalanceDate
    getFirstPeriodDateInthisYear StartCurrentDate
  
  BranchID = "0"
         If SystemOptions.UserInterface = ArabicInterface Then
                HideZeroBalance = MsgBox("Â·  —Ìœ «Œ›«¡ Õ”«»«  ’›—ÌÂ ‰⁄„ «„ ·« ", vbInformation + vbYesNoCancel)
            Else
                HideZeroBalance = MsgBox("Hide Zero Account  ", vbInformation + vbYesNoCancel)
            End If
    
            If HideZeroBalance = vbCancel Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

 ' If Me.ListBranchSelected.ListIndex > -1 Then
  For i = 0 To Me.ListBranchSelected.ListCount - 1
  BranchID = BranchID & "," & Me.ListBranchSelected.ItemData(i)
  Next i
 ' End If
  If BranchID = "-1" Then
  If SystemOptions.UserInterface = ArabicInterface Then
  MsgBox "Ì—ÃÏ  ÕœÌœ ›—⁄ Ê«Õœ ⁄·Ï «·«ﬁ·"
  Else
  MsgBox "Please select Branch"
  End If
  Exit Sub
  End If
  
  updateprofitAccount 0, 0, Me.ToDate.value, BranchID
  
  
  Sql = " SELECT    ProfitBalance, Parent_Account_Code, AccountTypes, Account_Code, Account_Serial, Account_Name, debitBalance ="
  Sql = Sql & "                         (SELECT     SUM(DEV_Value1)"
  Sql = Sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  Sql = Sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  Sql = Sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS d"
  Sql = Sql & "                                                   WHERE      (d.Credit_Or_Debit = 0 and d.Posted is null AND d.RecordDate >= " & SQLDate(Me.FromDate.value, True) & " AND d .RecordDate <= " & SQLDate(Me.ToDate.value, True) & ") AND d .Account_Code = A.Account_Code and branch_id in(" & BranchID & ")) x),"
  Sql = Sql & "                    CreditBalance ="
  Sql = Sql & "                        (SELECT     SUM(DEV_Value2)"
  Sql = Sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  Sql = Sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  Sql = Sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS d1"
  Sql = Sql & "                                                   WHERE     (d1.Credit_Or_Debit = 1 and d1.Posted is null AND d1.RecordDate >= " & SQLDate(Me.FromDate.value, True) & "  AND d1.RecordDate <= " & SQLDate(Me.ToDate.value, True) & ") AND d1.Account_Code = A.Account_Code  and branch_id in(" & BranchID & ") ) x),"
  Sql = Sql & "                     OpeningBalance ="
  Sql = Sql & "                        (SELECT     SUM(DEV_Value1) + SUM(DEV_Value2)"
  Sql = Sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  Sql = Sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  Sql = Sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS1 AS do"
  'Sql = Sql & "                                                   WHERE     ( do.RecordDate  = " & SQLDate(openingBalanceDate, True) & " and do.Account_Code = A.Account_Code  and branch_id in(" & BranchID & ") )) x),"
  Sql = Sql & "                                                   WHERE     (  do.Account_Code = A.Account_Code and do.Posted is null  and branch_id in(" & BranchID & ") )) x),"
  Sql = Sql & "    OpeningBalancebeformdateMinus1 ="
  Sql = Sql & "                        (SELECT     SUM(DEV_Value1) + SUM(DEV_Value2)"
  Sql = Sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  Sql = Sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  Sql = Sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS do"
  Sql = Sql & "                                                   WHERE     ( do.RecordDate >=" & SQLDate(openingBalanceDate, True) & " and   do.RecordDate <= " & SQLDate(FromdateMinus1, True) & ") AND  do.Posted is null and  do.Account_Code = A.Account_Code  and branch_id in(" & BranchID & ") ) x),"
  Sql = Sql & "                    OpeningBalancebeformStartCurrentyearTOFromDAteminus1 ="
  Sql = Sql & "                        (SELECT     SUM(DEV_Value1) + SUM(DEV_Value2)"
  Sql = Sql & "                           FROM         (SELECT     Account_Code, DEV_Value1 = CASE WHEN Credit_Or_Debit = 0 THEN Value * 1 ELSE 0 END,"
  Sql = Sql & "                                                                         DEV_Value2 = CASE WHEN Credit_Or_Debit = 1 THEN Value * - 1 ELSE 0 END"
  Sql = Sql & "                                                   FROM         dbo.DOUBLE_ENTREY_VOUCHERS AS do"
  Sql = Sql & "                                                   WHERE     (do.RecordDate >= " & SQLDate(StartCurrentDate, True) & " AND do.RecordDate < " & SQLDate(Me.FromDate.value, True) & ") AND do.Account_Code = A.Account_Code and do.Posted is null and branch_id in(" & BranchID & ") ) x)"
  Sql = Sql & " FROM         ACCOUNTS A"
  Sql = Sql & " WHERE     A.last_account = 1   "
  'Sql = Sql & " WHERE     A.last_account = 1 and Account_Code ='a1a2a2a2a42'  "
    Sql = Sql & "order by Account_Serial "
    
  ''--AND Account_Serial = '120101001001'"
  Rs1.Open Sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
  
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     
     With Me.Fg_Journal
     .Clear
Grd

                    For i = .FixedRows To Rs1.RecordCount + 1
                   .Rows = .FixedRows + Rs1.RecordCount
                   AccountTypes = IIf(IsNull(Rs1("AccountTypes").value), 0, Rs1("AccountTypes").value)
                   OpeningBalancebeformdateMinus1 = IIf(IsNull(Rs1("OpeningBalancebeformdateMinus1").value), 0, Rs1("OpeningBalancebeformdateMinus1").value)
                   OpeningBalance = IIf(IsNull(Rs1("OpeningBalance").value), 0, Rs1("OpeningBalance").value)
                   ProfitBalance = IIf(IsNull(Rs1("ProfitBalance").value), 0, Rs1("ProfitBalance").value)
                   OpeningBalancebeformStartCurrentyearTOFromDAteminus1 = IIf(IsNull(Rs1("OpeningBalancebeformStartCurrentyearTOFromDAteminus1").value), 0, Rs1("OpeningBalancebeformStartCurrentyearTOFromDAteminus1").value)
                   If AccountTypes = 0 Or AccountTypes = 1 Then
                   NewOpinning = OpeningBalancebeformdateMinus1 + OpeningBalance + ProfitBalance
                   Else
                   NewOpinning = OpeningBalancebeformStartCurrentyearTOFromDAteminus1 + ProfitBalance
                   End If
                   
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(Rs1("Account_Code").value), "", Rs1("Account_Code").value)
                   .TextMatrix(i, .ColIndex("Account_Serial")) = IIf(IsNull(Rs1("Account_Serial").value), "", Rs1("Account_Serial").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(Rs1("Account_Name").value), "", Rs1("Account_Name").value)
                   Else
                   .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(Rs1("Account_Name").value), "", Rs1("Account_Name").value)
                   End If
                   If NewOpinning >= 0 Then
                   .TextMatrix(i, .ColIndex("DebitValue")) = NewOpinning
                   .TextMatrix(i, .ColIndex("CreditValue")) = 0
                   Else
                   .TextMatrix(i, .ColIndex("DebitValue")) = 0
                   .TextMatrix(i, .ColIndex("CreditValue")) = NewOpinning
                   End If
                   .TextMatrix(i, .ColIndex("DebitValue1")) = IIf(IsNull(Rs1("debitBalance").value), 0, Rs1("debitBalance").value)
                   .TextMatrix(i, .ColIndex("CreditValue1")) = IIf(IsNull(Rs1("CreditBalance").value), 0, Rs1("CreditBalance").value)
                   total = NewOpinning + val(.TextMatrix(i, .ColIndex("DebitValue1"))) + val(.TextMatrix(i, .ColIndex("CreditValue1")))
                     If total >= 0 Then
                   .TextMatrix(i, .ColIndex("DebitValue2")) = total
                   .TextMatrix(i, .ColIndex("CreditValue2")) = 0
                   Else
                   .TextMatrix(i, .ColIndex("DebitValue2")) = 0
                   .TextMatrix(i, .ColIndex("CreditValue2")) = total
                   End If
                   Rs1.MoveNext
                                                .Col = Fg_Journal.ColIndex("Ser")
                             .ShowCell i, .ColIndex("Ser")
                            
                             .SetFocus
                            
             Next i
         .Rows = .Rows + 1
                     .TextMatrix(.Rows - 1, .ColIndex("DebitValue")) = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows, .ColIndex("DebitValue"))
                     .TextMatrix(.Rows - 1, .ColIndex("CreditValue")) = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows, .ColIndex("CreditValue"))
                     
                     .TextMatrix(.Rows - 1, .ColIndex("DebitValue1")) = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue1"), .Rows, .ColIndex("DebitValue1"))
                     
                     .TextMatrix(.Rows - 1, .ColIndex("CreditValue1")) = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue1"), .Rows, .ColIndex("CreditValue1"))
                     
                     
                     .TextMatrix(.Rows - 1, .ColIndex("DebitValue2")) = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue2"), .Rows, .ColIndex("DebitValue2"))
                     
                     .TextMatrix(.Rows - 1, .ColIndex("CreditValue2")) = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue2"), .Rows, .ColIndex("CreditValue2"))
                     
 

        End With
        Exit Sub
ErrTrap:
    End Sub
 Sub FullGridListData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim Sql As String
  Sql = " SELECT     dbo.TblFinaStatemDet.ID, dbo.TblFinaStatemDet.FinStID, dbo.TblFinaStatemDet.TypeTrans, dbo.TblFinaStatemDet.DebitValue, dbo.TblFinaStatemDet.CreditValue, "
  Sql = Sql & "                     dbo.TblFinaStatemDet.DebitValue1, dbo.TblFinaStatemDet.CreditValue1, dbo.TblFinaStatemDet.DebitValue2, dbo.TblFinaStatemDet.CreditValue2,"
  Sql = Sql & "                    dbo.TblFinaStatemDet.AccountCode , dbo.ACCOUNTS.account_name, dbo.ACCOUNTS.account_serial, dbo.ACCOUNTS.Account_NameEng"
  Sql = Sql & " FROM         dbo.TblFinaStatemDet INNER JOIN"
  Sql = Sql & "                    dbo.ACCOUNTS ON dbo.TblFinaStatemDet.AccountCode = dbo.ACCOUNTS.Account_Code"
  Sql = Sql & " Where (dbo.TblFinaStatemDet.TypeTrans = 0) And (dbo.TblFinaStatemDet.FinStID =" & val(TxtSerial1.Text) & ")"
  Rs1.Open Sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.Fg_Journal
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(Rs1("AccountCode").value), "", Rs1("AccountCode").value)
                   .TextMatrix(i, .ColIndex("Account_Serial")) = IIf(IsNull(Rs1("Account_Serial").value), "", Rs1("Account_Serial").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(Rs1("Account_Name").value), "", Rs1("Account_Name").value)
                   Else
                   .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(Rs1("Account_NameEng").value), "", Rs1("Account_NameEng").value)
                   End If
                   .TextMatrix(i, .ColIndex("DebitValue")) = IIf(IsNull(Rs1("DebitValue").value), 0, Rs1("DebitValue").value)
                   .TextMatrix(i, .ColIndex("CreditValue")) = IIf(IsNull(Rs1("CreditValue").value), 0, Rs1("CreditValue").value)
                   .TextMatrix(i, .ColIndex("DebitValue1")) = IIf(IsNull(Rs1("DebitValue1").value), 0, Rs1("DebitValue1").value)
                   .TextMatrix(i, .ColIndex("CreditValue1")) = IIf(IsNull(Rs1("CreditValue1").value), 0, Rs1("CreditValue1").value)
                   .TextMatrix(i, .ColIndex("DebitValue2")) = IIf(IsNull(Rs1("DebitValue2").value), 0, Rs1("DebitValue2").value)
                   .TextMatrix(i, .ColIndex("CreditValue2")) = IIf(IsNull(Rs1("CreditValue2").value), 0, Rs1("CreditValue2").value)
                   Rs1.MoveNext
             Next i
                      .Rows = .Rows + 1
                     .TextMatrix(.Rows - 1, .ColIndex("DebitValue")) = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows, .ColIndex("DebitValue"))
                     .TextMatrix(.Rows - 1, .ColIndex("CreditValue")) = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows, .ColIndex("CreditValue"))
                     
                     .TextMatrix(.Rows - 1, .ColIndex("DebitValue1")) = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue1"), .Rows, .ColIndex("DebitValue1"))
                     
                     .TextMatrix(.Rows - 1, .ColIndex("CreditValue1")) = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue1"), .Rows, .ColIndex("CreditValue1"))
                     
                     
                     .TextMatrix(.Rows - 1, .ColIndex("DebitValue2")) = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue2"), .Rows, .ColIndex("DebitValue2"))
                     
                     .TextMatrix(.Rows - 1, .ColIndex("CreditValue2")) = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue2"), .Rows, .ColIndex("CreditValue2"))
                     
                     
        End With
        ListBranchSelected.Clear
      Set Rs1 = New ADODB.Recordset
  Sql = "         SELECT     dbo.TblFinaStatemDet.ID, dbo.TblFinaStatemDet.FinStID, dbo.TblFinaStatemDet.TypeTrans, dbo.TblFinaStatemDet.AccountCode, dbo.TblFinaStatemDet.BranchID,"
  Sql = Sql & "                    dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_namee"
  Sql = Sql & "  FROM         dbo.TblFinaStatemDet INNER JOIN"
  Sql = Sql & "                    dbo.TblBranchesData ON dbo.TblFinaStatemDet.BranchID = dbo.TblBranchesData.branch_id"
  Sql = Sql & " Where (dbo.TblFinaStatemDet.TypeTrans = 2) And (dbo.TblFinaStatemDet.FinStID =" & val(TxtSerial1.Text) & ")"
  Rs1.Open Sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
   For i = 0 To Rs1.RecordCount - 1
   If SystemOptions.UserInterface = ArabicInterface Then
   Me.ListBranchSelected.AddItem IIf(IsNull(Rs1("branch_name").value), "", Rs1("branch_name").value)
   Else
   Me.ListBranchSelected.AddItem IIf(IsNull(Rs1("branch_namee").value), "", Rs1("branch_namee").value)
   End If
   Me.ListBranchSelected.ItemData(i) = IIf(IsNull(Rs1("BranchID").value), 0, Rs1("BranchID").value)
   Rs1.MoveNext
  Next i
  
  Me.ListActivitySelected.Clear
  Set Rs1 = New ADODB.Recordset
  Sql = "  SELECT     dbo.TblFinaStatemDet.FinStID, dbo.TblFinaStatemDet.ID, dbo.TblFinaStatemDet.TypeTrans, dbo.TblFinaStatemDet.BranchID, dbo.tblActivitesType.Name, "
  Sql = Sql & "                    dbo.tblActivitesType.NameE"
  Sql = Sql & "  FROM         dbo.TblFinaStatemDet INNER JOIN"
  Sql = Sql & "                    dbo.tblActivitesType ON dbo.TblFinaStatemDet.BranchID = dbo.tblActivitesType.id"
  Sql = Sql & " Where (dbo.TblFinaStatemDet.TypeTrans = 1) And (dbo.TblFinaStatemDet.FinStID =" & val(TxtSerial1.Text) & ")"
  Rs1.Open Sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
   For i = 0 To Rs1.RecordCount - 1
   If SystemOptions.UserInterface = ArabicInterface Then
   Me.ListActivitySelected.AddItem IIf(IsNull(Rs1("Name").value), "", Rs1("Name").value)
   Else
   Me.ListActivitySelected.AddItem IIf(IsNull(Rs1("namee").value), "", Rs1("namee").value)
   End If
   Me.ListActivitySelected.ItemData(i) = IIf(IsNull(Rs1("BranchID").value), 0, Rs1("BranchID").value)
   Rs1.MoveNext
  Next i
  
        Exit Sub
ErrTrap:
    End Sub
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------
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
    StrRecID = new_id("TblFinaStatem", "ID", "")
    Me.TxtSerial1.Text = StrRecID
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

Private Sub Grid_Click()
    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("id")))
    FiLLTXT
End Sub

Private Sub ISButton1_Click()
print_report
End Sub

Private Sub ISButton2_Click()
If Me.TxtModFlg.Text <> "R" Then
FullGrid
End If
End Sub

Private Sub Label10_Click()
If Me.TxtModFlg.Text <> "R" Then
Me.ListBranchSelected.Clear
End If
End Sub

Private Sub Label11_Click()
If Me.TxtModFlg.Text <> "R" Then
    If Me.ListBranchSelected.ListIndex > -1 Then
        ListBranchSelected.RemoveItem ListBranchSelected.ListIndex
    End If
End If
End Sub

Private Sub Label4_Click()
If Me.TxtModFlg.Text <> "R" Then
Dim BranchID, Sql As String
Dim Rs1  As ADODB.Recordset
Dim i, k As Integer
 If ListBranchAll.ListIndex > -1 Then
    ListBranchSelected.AddItem ListBranchAll.List(ListBranchAll.ListIndex)
    ListBranchSelected.ItemData(ListBranchSelected.NewIndex) = ListBranchAll.ItemData(ListBranchAll.ListIndex)
End If
End If
End Sub
Function FillMylist()
    Dim Sql As String
    Dim Rs2 As ADODB.Recordset
    Dim i As Integer
    Set Rs2 = New ADODB.Recordset
    Sql = " SELECT * from  tblActivitesType "
    Rs2.Open Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    Me.ListAllActivity.Clear
    Me.ListActivitySelected.Clear
    If Rs2.RecordCount > 0 Then
        For i = 1 To Rs2.RecordCount
            If SystemOptions.UserInterface = ArabicInterface Then
                ListAllActivity.AddItem IIf(IsNull(Rs2("Name").value), "", Rs2("Name").value)
            Else
                ListAllActivity.AddItem IIf(IsNull(Rs2("namee").value), "", Rs2("namee").value)
            End If
            ListAllActivity.ItemData(ListAllActivity.NewIndex) = IIf(IsNull(Rs2("ID").value), 0, Rs2("ID").value)
            Rs2.MoveNext
        Next i

    End If
    Rs2.Close
End Function
Function FillMylist2()
    Dim Sql As String
    Dim Rs2 As ADODB.Recordset
    Dim i As Integer
    Dim ActivID As String
    ActivID = "0"
    For i = 0 To Me.ListActivitySelected.ListCount - 1
    ActivID = ActivID & "," & Me.ListActivitySelected.ItemData(i)
    Next i
    Me.ListBranchAll.Clear
    Me.ListBranchSelected.Clear
    If ActivID = "0" Then Exit Function
    Set Rs2 = New ADODB.Recordset
    Sql = " SELECT * from  TblBranchesData where ActivityTypeId in(" & ActivID & ") "
    Rs2.Open Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Rs2.RecordCount > 0 Then
        For i = 1 To Rs2.RecordCount
            If SystemOptions.UserInterface = ArabicInterface Then
                ListBranchAll.AddItem IIf(IsNull(Rs2("branch_name").value), "", Rs2("branch_name").value)
            Else
                ListBranchAll.AddItem IIf(IsNull(Rs2("branch_namee").value), "", Rs2("branch_namee").value)
            End If
            ListBranchAll.ItemData(ListBranchAll.NewIndex) = IIf(IsNull(Rs2("branch_id").value), 0, Rs2("branch_id").value)
            Rs2.MoveNext
        Next i

    End If


    Rs2.Close
End Function

Private Sub Label5_Click()
If Me.TxtModFlg.Text <> "R" Then
If ListActivitySelected.ListIndex > -1 Then
ListActivitySelected.RemoveItem (ListActivitySelected.ListIndex)
End If
End If
FillMylist2
End Sub

Private Sub Label6_Click()
If Me.TxtModFlg.Text <> "R" Then
ListActivitySelected.Clear
Me.ListBranchAll.Clear
Me.ListBranchSelected.Clear
End If
End Sub

Private Sub Label7_Click()
If Me.TxtModFlg.Text <> "R" Then
    Dim i As Integer
    Me.ListBranchSelected.Clear
    For i = 0 To Me.ListAllActivity.ListCount - 1
        Me.ListActivitySelected.AddItem ListAllActivity.List(i)
        ListActivitySelected.ItemData(i) = ListAllActivity.ItemData(i)
    Next i
   End If
   FillMylist2
End Sub

Private Sub Label8_Click()
If Me.TxtModFlg.Text <> "R" Then
Dim BranchID, Sql As String
Dim Rs1  As ADODB.Recordset
Dim i, k As Integer
 If Me.ListAllActivity.ListIndex > -1 Then
    Me.ListActivitySelected.AddItem ListAllActivity.List(ListAllActivity.ListIndex)
    ListActivitySelected.ItemData(ListActivitySelected.NewIndex) = ListAllActivity.ItemData(ListAllActivity.ListIndex)
End If
End If
FillMylist2
End Sub

Private Sub Label9_Click()
If Me.TxtModFlg.Text <> "R" Then
    Dim i As Integer
    Me.ListBranchSelected.Clear
    For i = 0 To Me.ListBranchAll.ListCount - 1
        ListBranchSelected.AddItem ListBranchAll.List(i)
        ListBranchSelected.ItemData(i) = ListBranchAll.ItemData(i)
    Next i
End If
End Sub

' change id search
Private Sub TxtSerial1_Change()
    On Error GoTo ErrTrap
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
ErrTrap:
End Sub
' search for select id
Public Function FindRec(ByVal RecId As Long, Optional NoteID As Long = 0)
    On Error GoTo ErrTrap
    If NoteID = 0 Then
    RsSavRec.find "ID=" & RecId, , adSearchForward, 1
      
    Else
      RsSavRec.find "ID=" & NoteID, , adSearchForward, 1
    End If
    If Not (RsSavRec.EOF) Then
        FiLLTXT
      ' FullGridData
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
    On Error GoTo ErrTrap
    FindRec val(TxtSerial1.Text)
    Me.TxtModFlg.Text = "R"
    FiLLTXT
     BtnLast_Click
ErrTrap:
End Sub
' delet sub
Private Sub btnDelete_Click()
   ' On Error GoTo ErrTrap
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    Dim x As Integer
    If SystemOptions.UserInterface = EnglishInterface Then
        x = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        x = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If
    If x = vbNo Then Exit Sub
     If TxtSerial1.Text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                x = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                x = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
       End If
               Else
                     StrSQL = "Delete From TblFinaStatemDet Where FinStID='" & val(TxtSerial1.Text) & "'"
                 Cn.Execute StrSQL, , adExecuteNoRecords
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1

  RsSavRec.delete
  
               If SystemOptions.UserInterface = EnglishInterface Then
                x = MsgBox(" Deletion Process Success ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                x = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               End If
               LabCurrRec.Caption = 0
               LabCountRec.Caption = 0
               cleargriid
               FullGridData
     End If
                            '------------------------------ Move Next ---------------------------.
        Me.Refresh
        
        BtnNext_Click
     Exit Sub
ErrTrap:

     Select Case Err.Number
        Case -2147217873, -2147467259
        If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "You can not delete this record because of its connection with other data"
            End If
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

Private Sub TxtModFlg_Change()
    If TxtModFlg.Text = "N" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        ISButton1.Enabled = False
        Me.btnQuery.Enabled = False
        ISButton1.Enabled = False
     '   Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
              
       
    ElseIf TxtModFlg.Text = "R" Then
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False
        ISButton1.Enabled = False
        If TxtSerial1.Text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
            ISButton1.Enabled = True
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
        ISButton1.Enabled = False
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
    RsSavRec.MoveFirst
   
    FiLLTXT
  '  FullGridData
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
              Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
            Msg = Msg & "Data will be updated"
            End If
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
        
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveLast
  
    FiLLTXT
   ' FullGridData
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
        Else
            Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
            Msg = Msg & "Data will be updated"
        End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify_Click()
    Dim Msg As String
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtSerial1.Text <> "" Then
        TxtModFlg = "E"
        Me.DCboUserName.BoundText = user_id
        Frm2.Enabled = True
    End If
    Exit Sub
ErrTrap:

    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
          If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê«" & Chr(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & Chr(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
          Else
           Msg = "Sorry ..." & Chr(13)
            Msg = Msg & "You can not edit this record now" & Chr(13)
            Msg = Msg & "It is in use by another user on the network"
          End If
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
   
  '  Me.VSFlexGrid2.Rows = 1
    TxtModFlg.Text = "N"
    CheckAll.value = vbUnchecked
    Me.DCboUserName.BoundText = user_id
    Me.ListActivitySelected.Clear
    Me.ListBranchSelected.Clear
 '   Me.Grid.Clear flexClearScrollable, flexClearEverything
    Me.Fg_Journal.Rows = 2
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
   ' FullGridData
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
       If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
       Else
            Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
            Msg = Msg & "Data will be updated"
        End If
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
     
    FiLLTXT
   ' FullGridData
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
              If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
       Else
            Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
            Msg = Msg & "Data will be updated"
        End If
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
    Wrap = Chr(13) + Chr(10)
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

Function print_report(Optional NoteSerial As String)
On Error GoTo ErrTrap
    Dim Sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    Sql = " SELECT     dbo.TblFinaStatemDet.FinStID, dbo.TblFinaStatemDet.TypeTrans, dbo.tblActivitesType.Name, dbo.tblActivitesType.namee, dbo.TblFinaStatemDet.DebitValue, "
    Sql = Sql & "                   dbo.TblFinaStatemDet.CreditValue, dbo.TblFinaStatemDet.DebitValue1, dbo.TblFinaStatemDet.CreditValue1, dbo.TblFinaStatemDet.DebitValue2,"
    Sql = Sql & "                  dbo.TblFinaStatemDet.CreditValue2, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.BasicAccount, dbo.ACCOUNTS.Account_NameEng,"
    Sql = Sql & "                  dbo.TblFinaStatemDet.AccountCode, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblFinaStatemDet.BranchID,"
    Sql = Sql & "                  dbo.TblFinaStatem.ID, dbo.TblFinaStatem.RecordDate, dbo.TblFinaStatem.ToDate, dbo.TblFinaStatem.FromDate, dbo.TblFinaStatem.Remarks,"
    Sql = Sql & "                  dbo.TblFinaStatem.UserID , dbo.TblUsers.UserName , dbo.ACCOUNTS.Account_Serial"
    Sql = Sql & "    FROM         dbo.TblUsers RIGHT OUTER JOIN"
    Sql = Sql & "                  dbo.TblFinaStatem ON dbo.TblUsers.UserID = dbo.TblFinaStatem.UserID LEFT OUTER JOIN"
    Sql = Sql & "                  dbo.TblFinaStatemDet ON dbo.TblFinaStatem.ID = dbo.TblFinaStatemDet.FinStID LEFT OUTER JOIN"
    Sql = Sql & "                  dbo.ACCOUNTS ON dbo.TblFinaStatemDet.AccountCode = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
    Sql = Sql & "                  dbo.TblBranchesData ON dbo.TblFinaStatemDet.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    Sql = Sql & "                  dbo.tblActivitesType ON dbo.TblFinaStatemDet.BranchID = dbo.tblActivitesType.id"
    Sql = Sql & "  Where (dbo.TblFinaStatem.ID  =" & val(TxtSerial1.Text) & ") and dbo.TblFinaStatemDet.TypeTrans=0 "
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepFinalStatem.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepFinalStatemE.rpt"
        End If
   If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
   Set RsData = New ADODB.Recordset
    RsData.Open Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
     If RsData.BOF Or RsData.EOF Then
     If SystemOptions.UserInterface = ArabicInterface Then
       Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
     Else
     Msg = "No Data"
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
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        StrReportTitle = "" '& StrAccountName
        Else
         xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        StrReportTitle = ""
    End If
    Dim Branches As String
    Branches = ""
    Dim i As Integer
    If CheckAll.value = vbUnchecked Then
    For i = 0 To Me.ListActivitySelected.ListCount - 1
    If i = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    Branches = "··‰‘«ÿ«  "
    Else
    Branches = "Activities"
    End If
    End If
    Branches = Branches & " " & Me.ListActivitySelected.List(i)
    If i <> Me.ListActivitySelected.ListCount - 1 Then
    Branches = Branches & "-"
    End If
    Next i
    Branches = Branches & Chr(13)
    For i = 0 To Me.ListBranchSelected.ListCount - 1
    If i = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    Branches = Branches & "Ê«·›—Ê⁄"
    Else
    Branches = Branches & "Branches"
    End If
    End If
    Branches = Branches & " " & Me.ListBranchSelected.List(i)
    If i <> Me.ListBranchSelected.ListCount - 1 Then
    Branches = Branches & "-"
    End If
    Next i
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(4).AddCurrentValue Branches
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
Private Sub ChangeLang()
On Error GoTo ErrTrap
   ' form name
    Me.Caption = "Financial Statements"
    Me.Label1(2).Caption = Me.Caption
    LblCode.Caption = "No"
    Me.lbldate.Caption = "Date"
    lbl(8).Caption = "User"
Label3.Caption = "Remarks"
ISButton2.Caption = "Show"
Label13.Caption = "Select Branch"
Label12.Caption = "Select Activity"
Label16.Caption = "Period"
Label14.Caption = "From"
Label17.Caption = "To"
CheckAll.RightToLeft = False
CheckAll.Caption = "All"
    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "NO. Recordes"


    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    ISButton1.Caption = "Print"
    btnQuery.Caption = "Search"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
    With Me.Grid
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
        .TextMatrix(0, .ColIndex("UserName")) = "User"
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
    End With
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
   My_SQL = "TblFinaStatem"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub
