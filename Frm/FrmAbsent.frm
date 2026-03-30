VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmAbsent 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ”ÃÌ· «·€Ì«»"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7080
   HelpContextID   =   570
   Icon            =   "FrmAbsent.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   5475
   ScaleWidth      =   7080
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2700
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Text            =   "modflag"
      Top             =   210
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox TxtAbs_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   3660
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   210
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.TextBox TxtSerial 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4650
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   120
      Width           =   1485
   End
   Begin MSComCtl2.DTPicker DTPick 
      Height          =   330
      Left            =   810
      TabIndex        =   0
      Top             =   780
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   393216
      Format          =   96468993
      CurrentDate     =   38886
   End
   Begin VB.Frame FramBrowser 
      BackColor       =   &H00E2E9E9&
      Height          =   675
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   -90
      Width           =   2250
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   180
         TabIndex        =   31
         Top             =   210
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   14871017
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
         ButtonImage     =   "FrmAbsent.frx":038A
         ColorButton     =   14871017
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   645
         TabIndex        =   32
         Top             =   210
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   14871017
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
         ButtonImage     =   "FrmAbsent.frx":0724
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1245
         TabIndex        =   33
         Top             =   210
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   14871017
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
         ButtonImage     =   "FrmAbsent.frx":0ABE
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   1710
         TabIndex        =   34
         Top             =   210
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   14871017
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
         ButtonImage     =   "FrmAbsent.frx":0E58
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
   End
   Begin VB.Frame Frm2 
      BackColor       =   &H00E2E9E9&
      Height          =   3765
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   585
      Width           =   7005
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   105
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   210
         Width           =   690
         Begin VB.ComboBox CmbTimeType 
            BackColor       =   &H00004000&
            ForeColor       =   &H0080FFFF&
            Height          =   315
            ItemData        =   "FrmAbsent.frx":11F2
            Left            =   0
            List            =   "FrmAbsent.frx":11F4
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   0
            Visible         =   0   'False
            Width           =   1305
         End
      End
      Begin VB.ListBox LstEmployees 
         Height          =   2985
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   630
         Width           =   3105
      End
      Begin VSFlex8Ctl.VSFlexGrid Grid 
         Height          =   2610
         Left            =   105
         TabIndex        =   6
         Top             =   1035
         Width           =   3105
         _cx             =   5477
         _cy             =   4604
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
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmAbsent.frx":11F6
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Height          =   3765
         Left            =   3255
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   0
         Width           =   495
         Begin ImpulseButton.ISButton BtnAddEmp 
            Height          =   285
            Left            =   90
            TabIndex        =   3
            Top             =   1830
            Width           =   285
            _ExtentX        =   503
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
            ButtonImage     =   "FrmAbsent.frx":1279
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnDelEmp 
            Height          =   285
            Left            =   90
            TabIndex        =   4
            Top             =   2265
            Width           =   285
            _ExtentX        =   503
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
            ButtonImage     =   "FrmAbsent.frx":1613
            ColorButton     =   14871017
            ButtonToggles   =   2
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   1
            LowerToggledContent=   0   'False
         End
         Begin ImpulseButton.ISButton BtnAddAll 
            Height          =   285
            Left            =   90
            TabIndex        =   2
            Top             =   1350
            Width           =   285
            _ExtentX        =   503
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
            ButtonImage     =   "FrmAbsent.frx":19AD
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnRemoveAll 
            Height          =   285
            Left            =   90
            TabIndex        =   5
            Top             =   2745
            Width           =   285
            _ExtentX        =   503
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
            ButtonImage     =   "FrmAbsent.frx":1D47
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSComctlLib.ImageList GrdImageList 
            Left            =   -45
            Top             =   3120
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmAbsent.frx":20E1
                  Key             =   "Ser"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmAbsent.frx":247B
                  Key             =   "Emp_Name"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmAbsent.frx":2815
                  Key             =   "Vac_ID"
               EndProperty
            EndProperty
         End
      End
      Begin VB.Label LabDayName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00004000&
         Caption         =   "«·Ã„⁄…"
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Left            =   285
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   720
         Width           =   420
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00004000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " «—ÌŒ «·€Ì«»"
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Index           =   2
         Left            =   2190
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   210
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00004000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ﬁ«∆„… «·€Ì«»"
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Index           =   1
         Left            =   105
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   630
         Width           =   3105
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00004000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ﬁ«∆„… «·„ÊŸ›Ì‰ «·–Ì‰ ·„ Ì”Ã· ·Â„ Õ÷Ê—"
         ForeColor       =   &H0000FFFF&
         Height          =   390
         Index           =   0
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   210
         Width           =   3105
      End
   End
   Begin VB.Frame FramTools 
      BackColor       =   &H00E2E9E9&
      Height          =   645
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   4350
      Width           =   7005
      Begin ImpulseButton.ISButton btnNew 
         Height          =   330
         Left            =   6120
         TabIndex        =   8
         Top             =   210
         Width           =   765
         _ExtentX        =   1349
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
         ButtonImage     =   "FrmAbsent.frx":2BAF
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnSave 
         Height          =   330
         Left            =   4485
         TabIndex        =   7
         Top             =   210
         Width           =   765
         _ExtentX        =   1349
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
         ButtonImage     =   "FrmAbsent.frx":2F49
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnModify 
         Height          =   330
         Left            =   5265
         TabIndex        =   9
         Top             =   210
         Width           =   765
         _ExtentX        =   1349
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
         ButtonImage     =   "FrmAbsent.frx":32E3
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUndo 
         Height          =   330
         Left            =   3645
         TabIndex        =   10
         Top             =   210
         Width           =   765
         _ExtentX        =   1349
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
         ButtonImage     =   "FrmAbsent.frx":367D
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnDelete 
         Height          =   330
         Left            =   2790
         TabIndex        =   11
         Top             =   210
         Width           =   765
         _ExtentX        =   1349
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
         ButtonImage     =   "FrmAbsent.frx":3A17
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnQuery 
         Height          =   330
         Left            =   2475
         TabIndex        =   19
         ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
         Top             =   1635
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "»ÕÀ"
         BackColor       =   14737632
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
         ButtonImage     =   "FrmAbsent.frx":3FB1
         ColorButton     =   14737632
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUpdate 
         Height          =   330
         Left            =   3240
         TabIndex        =   20
         ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
         Top             =   1650
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   " ÕœÌÀ"
         BackColor       =   14737632
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
         ButtonImage     =   "FrmAbsent.frx":434B
         ColorButton     =   14737632
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnPrint 
         Height          =   330
         Left            =   1440
         TabIndex        =   21
         Top             =   1695
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄…"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmAbsent.frx":46E5
         ColorButton     =   14737632
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnCancel 
         Height          =   330
         Left            =   105
         TabIndex        =   12
         Top             =   210
         Width           =   765
         _ExtentX        =   1349
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
         ButtonImage     =   "FrmAbsent.frx":4A7F
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   0
      TabIndex        =   29
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   5010
      Width           =   2460
      _ExtentX        =   4339
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
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ—— »Ê«”ÿ…"
      Height          =   270
      Index           =   13
      Left            =   2490
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   5040
      Width           =   915
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄„·Ì…"
      Height          =   375
      Left            =   6180
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "FrmAbsent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim StrRecID As String

Private Sub BtnAddAll_Click()
    On Error GoTo ErrTrap
    Dim II As Integer

    For II = 0 To LstEmployees.ListCount - 1

        With Grid
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("Emp_Name")) = LstEmployees.List(II)
            .TextMatrix(.Rows - 1, .ColIndex("Emp_ID")) = LstEmployees.ItemData(II)
        End With

    Next

    LstEmployees.Clear

ErrTrap:

End Sub

Private Sub BtnAddEmp_Click()
    On Error GoTo ErrTrap

    If LstEmployees.ListIndex = -1 Then Exit Sub

    With Grid
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Name")) = LstEmployees.ItemData(LstEmployees.ListIndex)
        .TextMatrix(.Rows - 1, .ColIndex("Emp_ID")) = LstEmployees.ItemData(LstEmployees.ListIndex)
        LstEmployees.RemoveItem (LstEmployees.ListIndex)
    End With
   
ErrTrap:
End Sub

Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub BtnDelEmp_Click()
    On Error GoTo ErrTrap

    If Grid.Rows = 1 Then Exit Sub
    If Grid.Row = 0 Then Exit Sub

    With Grid
        LstEmployees.AddItem .TextMatrix(.Row, .ColIndex("Emp_Name"))
        LstEmployees.ItemData(LstEmployees.NewIndex) = .TextMatrix(.Row, .ColIndex("Emp_ID"))
        .RemoveItem .Row
    End With

ErrTrap:

End Sub

Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    Dim StrSQL As String
    On Error GoTo ErrTrap

    If DoPremis(Do_Delete, Me.name, True) = False Then
        Exit Sub
    End If

    If TxtAbs_ID.text <> "" Then
        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)
    
        If MSGType = vbYes Then
            RsSavRec.find "Abs_ID=" & val(TxtAbs_ID.text), , adSearchForward, 1
            RsSavRec.delete
            StrSQL = "Delete From tblJunkAbsent Where Abs_ID=" & val(TxtAbs_ID.text)
            Cn.Execute StrSQL
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.TxtSerial.text = ""
            '------------------------------ Move Next ---------------------------.
            DTPick_Click
            Me.TxtModFlg.text = "R"
        End If

    End If

    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147217873, -2147467259
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select

End Sub

Private Sub BtnFirst_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtAbs_ID.text)
        Me.TxtModFlg.text = "R"
    End If

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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnLast_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtAbs_ID.text)
        Me.TxtModFlg.text = "R"
    End If

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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnNew_Click()
    NewRecord
End Sub

Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtAbs_ID.text)
        Me.TxtModFlg.text = "R"
    End If

    If DoPremis(Do_New, Me.name, True) Then
        Exit Sub
    End If

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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtAbs_ID.text)
        Me.TxtModFlg.text = "R"
    End If

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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnRemoveAll_Click()
    On Error GoTo ErrTrap
    Dim II As Integer

    'LstEmployees.Clear
    With Grid

        For II = 1 To .Rows - 1
            LstEmployees.AddItem .TextMatrix(II, .ColIndex("Emp_Name"))
            LstEmployees.ItemData(LstEmployees.NewIndex) = .TextMatrix(II, .ColIndex("Emp_ID"))
        
        Next

        .Clear flexClearScrollable
        .Rows = 1
    End With

ErrTrap:

End Sub

Private Sub DTPick_Change()
    DTPick_Click
End Sub

Private Sub DTPick_Click()
    LstEmployees.Clear
    Grid.Clear flexClearScrollable
    Grid.Rows = 1

    GetPresentedEmp DTPick.value
    GetAbsentedEmp DTPick.value
    LabDayName.Caption = Format(DTPick.value, "dddd")
    GetTimeDetails
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic

    Set XPic = Me.BtnAddAll.ButtonImage
    Set Me.BtnAddAll.ButtonImage = Me.BtnRemoveAll.ButtonImage
    Set Me.BtnRemoveAll.ButtonImage = XPic
    Set XPic = Me.BtnAddEmp.ButtonImage
    Set Me.BtnAddEmp.ButtonImage = Me.BtnDelEmp.ButtonImage
    Set Me.BtnDelEmp.ButtonImage = XPic

    Me.Caption = "Absence Recording "
    'Label1(2).Caption = Me.Caption

    lbl.Caption = "OPR#"
    Label1(0).Caption = "Employee List"
    Label1(2).Caption = "Date"
    Label1(1).Caption = "Absence List"
    Label1(13).Caption = "By"

    With Me.Grid
 
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Emp Name"
        .TextMatrix(0, .ColIndex("Vac_ID")) = "Absence"

    End With

    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

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

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim My_SQL As String
    Dim BKGrndPic As ClsBackGroundPic
    Dim rs As ADODB.Recordset
    Dim StrComboList As String
    Dim Dcombos As ClsDataCombos

    Resize_Form Me
    My_SQL = "select * From  tblAbsent Order By Abs_ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    With Me.Grid
        .Rows = 1
        .RowHeight(0) = 300
        '.Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon
        .Cell(flexcpPicture, 0, .ColIndex("Emp_Name")) = Me.GrdImageList.ListImages("Emp_Name").ExtractIcon
        .Cell(flexcpPicture, 0, .ColIndex("Vac_ID")) = Me.GrdImageList.ListImages("Vac_ID").ExtractIcon
    
        For i = 0 To .Cols - 1
            .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
        Next
    
        .ExtendLastCol = True
        .WallPaper = BKGrndPic.Picture
    End With

    Set Dcombos = New ClsDataCombos
    Dcombos.GetUsers Me.DCUser
    My_SQL = "Select Vac_ID,Vac_Name From tblVacancy"
    FillFlexGrid Me.Grid, 1, My_SQL
    My_SQL = "select Emp_ID,Emp_Name From tblEmployee "
    FillListBox LstEmployees, My_SQL
    Set rs = New ADODB.Recordset
    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then

        With Me.Grid
            StrComboList = .BuildComboList(rs, "Emp_Name", "Emp_ID")

            If StrComboList <> "" Then
                .ColComboList(.ColIndex("Emp_Name")) = StrComboList
            End If

        End With

        'StrComboList
    End If

    '----------------------------------------------------------------------------
    SetDtpickerDate Me.DTPick

    '----------------------------------------------------------------------------
    With Me.CmbTimeType
        .Clear
        .AddItem "⁄„·"
        .ItemData(0) = 0
        .AddItem "⁄ÿ·…"
        .ItemData(1) = 1
    End With

    DTPick.value = Date
    DTPick_Click
    BtnLast_Click
    ShowTip
    Me.TxtModFlg.text = "R"

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If

End Sub

Public Sub NewRecord()
    Dim My_SQL As String
    Dim rs As ADODB.Recordset

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    clear_all Me
    Grid.Rows = 1
    Grid.Clear flexClearScrollable, flexClearEverything

    LstEmployees.Clear
    TxtModFlg.text = "N"

    My_SQL = "TblAbsent"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial.text = rs.RecordCount + 1
    Else
        TxtSerial.text = 1
    End If

    rs.Close
    My_SQL = "select Emp_ID,Emp_Name From tblEmployee "
    FillListBox LstEmployees, My_SQL
    GetPresentedEmp DTPick.value
    GetAbsentedEmp DTPick.value
    GetTimeDetails
    TxtSerial.SetFocus
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
       
                'SaveData
                btnSave_Click

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub Grid_EnterCell()

    If Grid.Col = 1 Then
        Grid.Editable = flexEDKbdMouse
    Else
        Grid.Editable = flexEDNone
    End If

End Sub

Private Sub TxtModFlg_Change()

    If TxtModFlg.text = "N" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
    
        '  btnNext.Enabled = False
        '  btnPrevious.Enabled = False
        '  btnFirst.Enabled = False
        '  btnLast.Enabled = False
        '    CmbTimeType.Locked = False
        '    CmbTimeType.Enabled = True
    
    ElseIf TxtModFlg.text = "R" Then
        Frm2.Enabled = False

        If val(TxtAbs_ID.text) > 0 Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
        Else
            btnModify.Enabled = False
            btnDelete.Enabled = False
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
        btnDelete.Enabled = True
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
    
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    
    End If

End Sub

Private Sub BtnUndo_Click()
    FindRec val(TxtAbs_ID.text)
    Me.TxtModFlg.text = "R"
End Sub

Private Sub btnModify_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If DoPremis(Do_Edit, Me.name, True) Then
        Exit Sub
    End If

    If TxtAbs_ID.text <> "" Then
        TxtModFlg.text = "E"
        Frm2.Enabled = True
    End If

    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147467259
            'Could not update; currently locked.
            Msg = "⁄›Ê«" & Chr(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & Chr(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
        
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If

    End Select

End Sub

Private Sub btnSave_Click()
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    Dim StrSQL As String

    On Error GoTo ErrTrap

    '---------------------- check if data Vaclete -----------------------
    If CmbTimeType.ListIndex = 0 Then

        For Each CtrlTxt In Me.Controls

            If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
                If CtrlTxt.text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                    MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.title
                    CtrlTxt.SetFocus
                    Exit Sub
                End If
            End If

        Next

        Msg = "⁄›Ê« Ì—Ã∆ «” ﬂ„«· »«ﬁÏ «·»Ì«‰« "

        If ChkData = False Then
            MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading + vbQuestion, App.title
            Exit Sub
        End If

        ' -------------------------------------- txtmodflg type -------------------
        Select Case Me.TxtModFlg.text

                '------------------------------ new record ----------------------------
            Case "N"
                '------------------------- save record -----------------------------
                AddNewRec

            Case "E"
                StrRecID = Me.TxtAbs_ID.text

                If RsSavRec("Abs_ID").value <> val(Me.TxtAbs_ID.text) Then
                    RsSavRec.find "Abs_ID=" & val(StrRecID), , adSearchForward, 1
                End If

                StrSQL = "Delete From tblJunkAbsent Where Abs_ID=" & val(TxtAbs_ID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                '----------------------------- save edit -------------------------------
                FiLLRec
        End Select

        Me.TxtModFlg.text = "R"
    Else
        Msg = "⁄›Ê« Â–« ÌÊ„ ⁄ÿ·…"
        MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If

    Exit Sub
ErrTrap:
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title

End Sub

Public Sub AddNewRec()
    On Error GoTo ErrTrap
    StrRecID = new_id("tblAbsent", "Abs_ID", "")
    Me.TxtAbs_ID.text = StrRecID
    RsSavRec.AddNew
    RsSavRec.Fields("Abs_ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

Public Sub EditRec(StrTable As String, _
                   RecID As String)
    FiLLRec

End Sub

Public Sub FiLLRec()
    Dim rs As ADODB.Recordset
    Dim II As Integer
    Dim Msg As String
    Dim My_SQL As String
    Dim BeginTrans As Boolean
    On Error GoTo ErrTrap

    Cn.BeginTrans
    BeginTrans = True
    Set rs = New ADODB.Recordset
    My_SQL = "select * From tblJunkAbsent"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdText
    RsSavRec.Fields("Abs_date").value = IIf(Me.DTPick.value <> "", Me.DTPick.value, Null)
    RsSavRec.Fields("Abs_Code").value = IIf(Trim(Me.TxtSerial.text) <> "", Trim(Me.TxtSerial.text), Null)
    RsSavRec.Fields("UserID").value = IIf(Me.DCUser.BoundText = "", user_id, Me.DCUser.BoundText)

    RsSavRec.update

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        My_SQL = "Delete * From tblJunkAbsent where Abs_ID='" & Trim(Me.TxtAbs_ID.text) & "'"
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        My_SQL = "Delete  From tblJunkAbsent Where Abs_ID='" & Trim(Me.TxtAbs_ID.text) & "'"
    End If

    Cn.Execute My_SQL

    With Grid

        For II = 1 To .Rows - 1
            rs.AddNew
            rs.Fields("Abs_ID").value = IIf(StrRecID = "", Null, StrRecID)
            rs.Fields("Emp_ID").value = IIf(.TextMatrix(II, .ColIndex("Emp_ID")) <> "", .TextMatrix(II, .ColIndex("Emp_ID")), Null)
            rs.Fields("Vac_ID").value = IIf(.TextMatrix(II, .ColIndex("Vac_ID")) <> "", .TextMatrix(II, .ColIndex("Vac_ID")), Null)
            rs.update
        Next

    End With

    Cn.CommitTrans
    BeginTrans = False
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbOKOnly + vbInformation + vbMsgBoxRtlReading + vbMsgBoxRight, App.title
    TxtModFlg.text = "R"
    DTPick_Click
    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    Msg = "⁄›Ê« ·ﬁœ ›‘·  ⁄„·Ì… «·Õ›Ÿ"
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.title

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub

Public Sub FiLLTXT()
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim i As Integer
    On Error GoTo ErrTrap

    TxtAbs_ID.text = IIf(IsNull(RsSavRec.Fields("Abs_ID").value), "", RsSavRec.Fields("Abs_ID").value)
    TxtSerial.text = IIf(IsNull(RsSavRec.Fields("Abs_Code").value), "", RsSavRec.Fields("Abs_Code").value)
    Me.DTPick.value = IIf(IsNull(RsSavRec("Abs_Date").value), Date, RsSavRec("Abs_Date").value)
    Me.DCUser.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    StrSQL = "Select tblJunkAbsent.*  From tblJunkAbsent Where Abs_ID='" & val(Me.TxtAbs_ID.text) & "'"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Grid
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows + rs.RecordCount

        If Not (rs.BOF Or rs.EOF) Then
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Emp_ID")) = rs.Fields("Emp_ID").value
                .TextMatrix(i, .ColIndex("Emp_Name")) = rs.Fields("Emp_ID").value
                .TextMatrix(i, .ColIndex("Vac_ID")) = rs.Fields("Vac_ID").value
                rs.MoveNext
            Next

        End If

        rs.Close
        Set rs = Nothing
    End With

ErrTrap:
End Sub

Public Function FindRec(ByVal RecID As Long)
    On Error GoTo ErrTrap
    RsSavRec.find "Abs_ID=" & RecID, , adSearchForward, 1

    If Not (RsSavRec.EOF) Then
        FiLLTXT
    
    End If

    Exit Function
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        BtnUndo_Click
    End If

    'RsSavRec.Filter = adFilterNone
End Function

Public Sub GetPresentedEmp(ByVal Day As Date)
    On Error GoTo ErrTrap
    'My_SQL = "SELECT Emp_ID,Emp_Name From TblEmployee WHERE Emp_ID not  in (SELECT Emp_ID " & _
     "From QryAbsentEmp Where QryAbsentEmp.AbsDate = '" & Format(Day, "dd/mm/yyyy") & "' )"
    Dim My_SQL As String

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        My_SQL = "SELECT Emp_ID,Emp_Name From TblEmployee WHERE Emp_ID not  in (SELECT Emp_ID " & "From (SELECT  Emp_ID  ,AbsDate FROM QryAbsentEmp union select  Emp_ID,PresentDate From tblPresentTime) " & "Where AbsDate ='" & Format(Day, "dd/mm/yyyy") & "')"
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        My_SQL = "SELECT Emp_ID, Emp_Name From TblEmployee "
        My_SQL = My_SQL + " WHERE Emp_ID not  in (SELECT Emp_ID " & "From (SELECT  Emp_ID  ,AbsDate FROM QryAbsentEmp" & " union " & " Select  Emp_ID,PresentDate From tblPresentTime )XTable " & " Where AbsDate ='" & SQLDate(Day) & "')"
    End If

    FillListBox LstEmployees, My_SQL

ErrTrap:
End Sub

Public Sub GetAbsentedEmp(ByVal Day As Date)
    On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    My_SQL = "SELECT * From QryAbsentEmp Where AbsDate ='" & Format(Day, "dd/mm/yyyy") & "' "
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    TxtModFlg.text = "N"

    'TxtSerial.Text = ""
    With Grid

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst
            TxtModFlg.text = "E"
            FindRec rs.Fields("Abs_ID").value
         
            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
            
                .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(rs.Fields("Emp_ID").value), "", rs.Fields("Emp_ID").value)
             
                .TextMatrix(i, .ColIndex("Vac_ID")) = IIf(IsNull(rs.Fields("Vac_ID").value), "", rs.Fields("Vac_ID").value)
            
                rs.MoveNext
            Next

        End If

    End With

ErrTrap:
End Sub

Public Function ChkData() As Boolean
    On Error GoTo ErrTrap

    Dim i As Integer

    With Grid

        For i = 1 To .Rows - 1

            If .TextMatrix(i, .ColIndex("Emp_ID")) <> "" Then
                If .TextMatrix(i, .ColIndex("Vac_ID")) = "" Then
                    ChkData = False
                    .SetFocus
                    .Row = i
                    Exit Function
                End If
            End If

        Next

        ChkData = True
    End With

ErrTrap:
End Function

Public Function ChkExistData() As Boolean
    On Error GoTo ErrTrap

    Dim i As Integer
    ChkExistData = False

    With Grid

        For i = 1 To .Rows - 1

            If .TextMatrix(i, .ColIndex("Emp_ID")) <> "" Then
                ChkExistData = True
                Exit Function
            End If

        Next

    End With

ErrTrap:
End Function

Private Sub GetTimeDetails()
    On Error GoTo ErrTrap
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From tblTimeSetting where Is_WorkDay = 0"

    'My_SQL = "select * From tblTimeSetting where day='" & Trim(LabDayName.Caption) & "'"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    'Is_WorkDay = 0
    If rs.RecordCount > 0 Then
        CmbTimeType.ListIndex = IIf(IsNull(rs.Fields("Is_WorkDay").value), -1, rs.Fields("Is_WorkDay").value)
    End If

ErrTrap:
End Sub

Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = Chr(13) + Chr(10)

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With

    '----------------------------------------------------------------------------------------------
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "‰ﬁ· «·ﬂ·" & Wrap & "·‰ﬁ· ﬂ· «·„ÊŸ›Ì‰ «·Ï ﬁ«∆„… «·€Ì«»" & Wrap & "≈÷€ÿ Â–« «·„› «Õ"
        .AddControl BtnAddAll, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "‰ﬁ· „ÊŸ›" & Wrap & "·‰ﬁ· «·„ÊŸ› «·„Ÿ·· «·Ï ﬁ«∆„… «·€Ì«»" & Wrap & "≈÷€ÿ Â–« «·„› «Õ"
        .AddControl BtnAddEmp, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› „ÊŸ›" & Wrap & "·Õ–› „ÊŸ› „‰ ﬁ«∆„… «·€Ì«»" & Wrap & "≈÷€ÿ Â–« «·„› «Õ"
        .AddControl BtnDelEmp, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·ﬂ·" & Wrap & "·Õ–› ﬂ· «·„ÊŸ›Ì‰ „‰ ﬁ«∆„… «·€Ì«»" & Wrap & "≈÷€ÿ Â–« «·„› «Õ"
        .AddControl BtnRemoveAll, Msg, True
    End With

ErrTrap:
End Sub

