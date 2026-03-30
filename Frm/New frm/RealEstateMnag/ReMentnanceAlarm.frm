VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form RSMentnanceAlarm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ‰»ÌÂ«  «·’Ì«‰Â «· Ì ·„  ﬁ›·"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17790
   Icon            =   "ReMentnanceAlarm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   17790
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "«—”«· "
      Height          =   495
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   9120
      Width           =   1455
   End
   Begin VB.Frame Frame9 
      Caption         =   "«Ã„«·Ì« "
      Height          =   915
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   9240
      Visible         =   0   'False
      Width           =   13035
      Begin VB.TextBox TxtTotalContract 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   10320
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   360
         Width           =   1065
      End
      Begin VB.TextBox TxtInsuranceValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   6240
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   360
         Width           =   1065
      End
      Begin VB.TextBox TxtWater 
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
         Left            =   4080
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   360
         Width           =   1065
      End
      Begin VB.TextBox TxtElectricity 
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
         Left            =   2160
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   360
         Width           =   945
      End
      Begin VB.TextBox TxtCommiValue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   8280
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   360
         Width           =   1065
      End
      Begin VB.TextBox TxtPhone 
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
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„… «·«ÌÃ«—"
         Height          =   195
         Index           =   6
         Left            =   11505
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   480
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " √„Ì‰"
         Height          =   195
         Index           =   19
         Left            =   7560
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„Ì«Â"
         Height          =   195
         Index           =   20
         Left            =   5385
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   480
         Width           =   750
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÂ—»«¡"
         Height          =   195
         Index           =   21
         Left            =   2985
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   480
         Width           =   750
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "”⁄Ì/—”Ê„"
         Height          =   405
         Index           =   25
         Left            =   9360
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "Â« › Ê«‰ —‰ "
         Height          =   195
         Index           =   27
         Left            =   1035
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   480
         Width           =   990
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Õœœ «·› —…"
      Height          =   1200
      Left            =   11220
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   600
      Width           =   6525
      Begin MSComCtl2.DTPicker Fromdate 
         Height          =   330
         Left            =   3135
         TabIndex        =   16
         Top             =   240
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   127336449
         CurrentDate     =   41640
      End
      Begin Dynamic_Byte.NourHijriCal Fromdateh 
         Height          =   255
         Left            =   3120
         TabIndex        =   17
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   510
         Index           =   9
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   900
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "≈÷«›…"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "ReMentnanceAlarm.frx":058A
         DrawFocusRectangle=   0   'False
      End
      Begin Dynamic_Byte.NourHijriCal todateH 
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   600
         Width           =   1755
         _ExtentX        =   3201
         _ExtentY        =   450
      End
      Begin MSComCtl2.DTPicker toDate 
         Height          =   330
         Left            =   840
         TabIndex        =   37
         Top             =   240
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   127336449
         CurrentDate     =   41640
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "≈«·Ï"
         Height          =   435
         Index           =   14
         Left            =   2460
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·› —… „‰"
         Height          =   315
         Index           =   0
         Left            =   4980
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "œ·«·«  «·«·Ê«‰"
      Height          =   495
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   -1320
      Width           =   4575
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "€Ì— „”œœ ﬂ«„·«"
         Height          =   255
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Height          =   255
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "„”œœ Ã“∆Ì"
         Height          =   255
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   -15
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   30
      Width           =   17745
      Begin VB.TextBox TxtVac_ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   240
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   510
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Text            =   "modflag"
         Top             =   120
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Frame Frmo2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   540
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   450
         Visible         =   0   'False
         Width           =   3105
         Begin MSDataListLib.DataCombo DCUser 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   -255
            TabIndex        =   2
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
            TabIndex        =   3
            Top             =   45
            Width           =   855
         End
      End
      Begin MSComctlLib.ImageList GrdImageList 
         Left            =   2760
         Top             =   0
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
               Picture         =   "ReMentnanceAlarm.frx":0924
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ReMentnanceAlarm.frx":0CBE
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ReMentnanceAlarm.frx":1058
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ReMentnanceAlarm.frx":13F2
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ReMentnanceAlarm.frx":178C
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ReMentnanceAlarm.frx":1B26
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ReMentnanceAlarm.frx":1EC0
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ReMentnanceAlarm.frx":245A
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   7560
         Picture         =   "ReMentnanceAlarm.frx":27F4
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ‰»ÌÂ«  «·’Ì«‰Â «· Ì ·„  ﬁ›·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   2
         Left            =   11640
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   120
         Width           =   5880
      End
   End
   Begin ImpulseButton.ISButton btnCancel 
      Height          =   330
      Left            =   120
      TabIndex        =   7
      Top             =   7560
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
      ButtonImage     =   "ReMentnanceAlarm.frx":645C
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton CmdPrint 
      Height          =   390
      Left            =   1080
      TabIndex        =   8
      Top             =   7560
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   688
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
      ButtonImage     =   "ReMentnanceAlarm.frx":67F6
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4035
      Left            =   1560
      TabIndex        =   9
      Top             =   9480
      Width           =   13995
      _cx             =   24686
      _cy             =   7117
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
      Rows            =   2
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"ReMentnanceAlarm.frx":6B90
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
   Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
      Height          =   5460
      Left            =   120
      TabIndex        =   22
      Top             =   1800
      Width           =   17670
      _cx             =   31168
      _cy             =   9631
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
      Rows            =   12
      Cols            =   46
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"ReMentnanceAlarm.frx":6DC9
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
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "RSMentnanceAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
     Dim My_SQL As String
Private Sub BtnCancel_Click()
    Me.Hide
End Sub

Function print_report(Optional NoteSerial As String)
     Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String



        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMentanaceAlarm.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepMentanaceAlarm.rpt"
            
       End If
           

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

   ' xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.ParameterFields(3).AddCurrentValue Fromdate.value
     xReport.ParameterFields(4).AddCurrentValue Fromdateh.value
      xReport.ParameterFields(5).AddCurrentValue toDate.value
      xReport.ParameterFields(6).AddCurrentValue todateH.value
 
       
      
    xReport.ParameterFields(7).AddCurrentValue Date
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
       ' xReport.ParameterFields(10).AddCurrentValue Date ' Me.DTPicker1.value
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
'Dim gr, order As Integer
' xReport.ParameterFields(14).AddCurrentValue Order
 'xReport.ParameterFields(15).AddCurrentValue gr
 ' xReport.ParameterFields(15).AddCurrentValue gr
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  ' xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
  Dim total As String
  Dim totl As Double
 ' totl = val(LbToTalExtra.Caption) + val(Me.lbTotalMente.Caption)
 ' total = totl
 '  xReport.ParameterFields(12).AddCurrentValue Me.lbTotalMente.Caption
 '     xReport.ParameterFields(13).AddCurrentValue LbToTalExtra.Caption
 '       xReport.ParameterFields(14).AddCurrentValue total
   ' xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
 xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function

Private Sub Cmd_Click(Index As Integer)
FillGrid
End Sub
Public Sub FillGrid(Optional str As String)

  '  On Error GoTo ErrTrap
On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset
Dim StrWhere As String
    Set rs = New ADODB.Recordset
  
'Dim notpayed As Double
'notpayed = 0
 


My_SQL = "SELECT     dbo.TblOrderMaintenanceDet.ORderID, dbo.TblOrderMaintenanceDet.TypeUnit, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblOrderMaintenanceDet.UnitNo, "
My_SQL = My_SQL & "                      dbo.TblAqarDetai.unitno AS UnitNoName, dbo.TblOrderMaintenanceDet.UnitStatus, dbo.TblRentStatus.name AS NameStatus,"
My_SQL = My_SQL & "                        dbo.TblRentStatus.namee AS NameStatusE, dbo.TblOrderMaintenanceDet.RenterID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
My_SQL = My_SQL & "                        dbo.TblOrderMaintenanceDet.Mobile, dbo.TblOrderMaintenanceDet.Ms, dbo.TblOrderMaintenance.ID, dbo.TblOrderMaintenance.RecDateH,"
My_SQL = My_SQL & "                        dbo.TblOrderMaintenance.RecDate, dbo.TblOrderMaintenance.TimOrder, dbo.TblOrderMaintenance.BranchID, dbo.TblBranchesData.branch_name,"
My_SQL = My_SQL & "                        dbo.TblBranchesData.branch_namee, dbo.TblOrderMaintenance.EmpID, TblEmployee_2.Emp_Code, TblEmployee_2.Emp_Name, TblEmployee_2.Emp_Name1,"
My_SQL = My_SQL & "                        TblEmployee_2.Emp_Name2, TblEmployee_2.Emp_Name3, TblEmployee_2.Emp_Name4, TblEmployee_2.Fullcode, TblEmployee_2.Emp_Namee4,"
My_SQL = My_SQL & "                        TblEmployee_2.Emp_Namee3, TblEmployee_2.Emp_Namee2, TblEmployee_2.Emp_Namee1, TblEmployee_2.Emp_Namee, dbo.TblOrderMaintenance.SuperVM,"
My_SQL = My_SQL & "                        TblEmployee_1.Emp_Code AS Emp_CodeSup, TblEmployee_1.Emp_Name AS Emp_NameSup, TblEmployee_1.Emp_Name1 AS Emp_Name1Sup,"
My_SQL = My_SQL & "                        TblEmployee_1.Emp_Name2 AS Emp_Name2Sup, TblEmployee_1.Emp_Name3 AS Emp_Name3Sup, TblEmployee_1.Emp_Name4 AS Emp_Name4Sup,"
My_SQL = My_SQL & "                        TblEmployee_1.Fullcode AS FullcodeSup, TblEmployee_1.Emp_Namee4 AS Emp_Namee4Sup, TblEmployee_1.Emp_Namee3 AS Emp_Namee3Sup,"
My_SQL = My_SQL & "                        TblEmployee_1.Emp_Namee2 AS Emp_Namee2Sup, TblEmployee_1.Emp_Namee1 AS Emp_Namee1Sup, TblEmployee_1.Emp_Namee AS Emp_NameeSup,"
My_SQL = My_SQL & "                        dbo.TblOrderMaintenance.AqrID, dbo.TblAqar.aqarname, dbo.TblOrderMaintenance.LocationIqar, dbo.TblOrderMaintenance.Des, dbo.TblOrderMaintenance.DMY,"
My_SQL = My_SQL & "                        dbo.TblOrderMaintenance.Cont, dbo.TblOrderMaintenance.EndFateH, dbo.TblOrderMaintenance.EndFate, dbo.TblOrderMaintenance.Lock,"
My_SQL = My_SQL & "                        dbo.TblOrderMaintenance.LockDateH , dbo.TblOrderMaintenance.LockDate"
My_SQL = My_SQL & "   FROM         dbo.TblAqarDetai RIGHT OUTER JOIN"
My_SQL = My_SQL & "                        dbo.TblEmployee TblEmployee_2 RIGHT OUTER JOIN"
My_SQL = My_SQL & "                        dbo.TblOrderMaintenance ON TblEmployee_2.Emp_ID = dbo.TblOrderMaintenance.EmpID LEFT OUTER JOIN"
My_SQL = My_SQL & "                        dbo.TblEmployee TblEmployee_1 ON dbo.TblOrderMaintenance.SuperVM = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
My_SQL = My_SQL & "                        dbo.TblAqar ON dbo.TblOrderMaintenance.AqrID = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
My_SQL = My_SQL & "                        dbo.TblBranchesData ON dbo.TblOrderMaintenance.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
My_SQL = My_SQL & "                        dbo.TblOrderMaintenanceDet ON dbo.TblOrderMaintenance.ID = dbo.TblOrderMaintenanceDet.ORderID LEFT OUTER JOIN"
My_SQL = My_SQL & "                        dbo.TblCustemers ON dbo.TblOrderMaintenanceDet.RenterID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
My_SQL = My_SQL & "                        dbo.TblRentStatus ON dbo.TblOrderMaintenanceDet.UnitStatus = dbo.TblRentStatus.id ON dbo.TblAqarDetai.Id = dbo.TblOrderMaintenanceDet.UnitNo LEFT OUTER JOIN"
My_SQL = My_SQL & "                        dbo.TblAkarUnit ON dbo.TblOrderMaintenanceDet.TypeUnit = dbo.TblAkarUnit.id"
My_SQL = My_SQL & "   Where (dbo.TblOrderMaintenance.Lock = 0)"
                      
   StrWhere = ""
   If Not IsNull(Me.Fromdate.value) Then
                   StrWhere = StrWhere & " AND dbo.TblOrderMaintenance.RecDate >=" & SQLDate(Me.Fromdate.value, True) & ""
      End If

    If Not IsNull(Me.toDate.value) Then
            StrWhere = StrWhere & " AND  dbo.TblOrderMaintenance.RecDate <=" & SQLDate(Me.toDate.value, True) & ""
     
    End If
    
  
     
   

 
  If SystemOptions.usertype <> UserAdminAll Or val(Current_branch) <> 0 Then
        StrWhere = StrWhere & " and   dbo.Notes.branch_no=" & Current_branch
    End If
 

 My_SQL = My_SQL & StrWhere
 
  My_SQL = My_SQL & " order by  dbo.TblOrderMaintenance.id "


sql = My_SQL

 
 
         
   
Dim ActualTotal As Double
rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
'    rs1.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.GridInstallments
       .Rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
           rs.MoveFirst

            For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("Ser")) = i
              .TextMatrix(i, .ColIndex("InstallNo")) = (IIf(IsNull(rs.Fields("id").value), 0, rs.Fields("id").value))
               .TextMatrix(i, .ColIndex("RecDate")) = (IIf(IsNull(rs.Fields("RecDate").value), 0, rs.Fields("RecDate").value))
               .TextMatrix(i, .ColIndex("RecDateH")) = (IIf(IsNull(rs.Fields("RecDateH").value), 0, rs.Fields("RecDateH").value))
'.TextMatrix(i, .ColIndex("NoteSerial1")) = (IIf(IsNull(rs.Fields("NoteSerial11").value), "", rs.Fields("NoteSerial11").value))
.TextMatrix(i, .ColIndex("Cus_mobile")) = (IIf(IsNull(rs.Fields("Mobile").value), "", rs.Fields("Mobile").value))
.TextMatrix(i, .ColIndex("aqarname")) = (IIf(IsNull(rs.Fields("aqarname").value), "", rs.Fields("aqarname").value))
.TextMatrix(i, .ColIndex("unitnoNam")) = (IIf(IsNull(rs.Fields("UnitNoName").value), "", rs.Fields("UnitNoName").value))
.TextMatrix(i, .ColIndex("Des")) = (IIf(IsNull(rs.Fields("Des").value), "", rs.Fields("Des").value))

                'TblCustemers.Cus_mobile
 .TextMatrix(i, .ColIndex("EndFateH")) = (IIf(IsNull(rs.Fields("EndFateH").value), ToHijriDate(Date), rs.Fields("EndFateH").value))
  .TextMatrix(i, .ColIndex("EndFate")) = IIf(IsNull(rs.Fields("EndFate").value), Date, rs.Fields("EndFate").value)
        
    .TextMatrix(i, .ColIndex("Value")) = (IIf(IsNull(rs.Fields("Earnest").value), 0, rs.Fields("Earnest").value))
     
                      '    ActualTotal = getinsttPayedTocontract(val(rs.Fields("id").value))
                      If rs("EndFate").value < Date Then
 .TextMatrix(i, .ColIndex("late")) = DateDiff("d", Date, rs("EndFate").value)
 End If
 ' .TextMatrix(i, .ColIndex("Remains")) = val(.TextMatrix(i, .ColIndex("Value"))) - ActualTotal

'If ActualTotal = 0 Then
'          .Cell(flexcpBackColor, i, 1, i, 37) = &H8080FF
'Else
'          .Cell(flexcpBackColor, i, 1, i, 37) = vbYellow
'End If
     
     
   '  .TextMatrix(i, .ColIndex("CusID")) = (IIf(IsNull(rs.Fields("CusID").value), "", rs.Fields("CusID").value))
   
   If SystemOptions.UserInterface = ArabicInterface Then
   .TextMatrix(i, .ColIndex("Emp_Name")) = (IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value))
   .TextMatrix(i, .ColIndex("Emp_NameSup")) = (IIf(IsNull(rs.Fields("Emp_NameSup").value), "", rs.Fields("Emp_NameSup").value))
   
   .TextMatrix(i, .ColIndex("Unitname")) = (IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value))
   .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusName").value), "", rs.Fields("CusName").value))
   Else
   .TextMatrix(i, .ColIndex("Emp_Name")) = (IIf(IsNull(rs.Fields("Emp_Namee").value), "", rs.Fields("Emp_Namee").value))
   .TextMatrix(i, .ColIndex("Emp_NameSup")) = (IIf(IsNull(rs.Fields("Emp_NameeSup").value), "", rs.Fields("Emp_NameeSup").value))
   
   .TextMatrix(i, .ColIndex("Unitname")) = (IIf(IsNull(rs.Fields("namee").value), "", rs.Fields("namee").value))
   .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusNamee").value), "", rs.Fields("CusNamee").value))
   End If
 '.TextMatrix(i, .ColIndex("hijri")) = (IIf(IsNull(rs.Fields("hijri").value), 0, rs.Fields("hijri").value))   '
   '.Cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked
 
 '   .TextMatrix(i, .ColIndex("RentValue")) = (IIf(IsNull(rs.Fields("RentValue").value), 0, rs.Fields("RentValue").value))
 '   .TextMatrix(i, .ColIndex("Commissions")) = (IIf(IsNull(rs.Fields("Commissions").value), 0, rs.Fields("Commissions").value))
 '   .TextMatrix(i, .ColIndex("Insurance")) = (IIf(IsNull(rs.Fields("Insurance").value), 0, rs.Fields("Insurance").value))
 '   .TextMatrix(i, .ColIndex("Water")) = (IIf(IsNull(rs.Fields("Water").value), 0, rs.Fields("Water").value))
 '   .TextMatrix(i, .ColIndex("Electric")) = (IIf(IsNull(rs.Fields("Electric").value), 0, rs.Fields("Electric").value))
 '   .TextMatrix(i, .ColIndex("TelandNet")) = (IIf(IsNull(rs.Fields("Phone").value), 0, rs.Fields("Phone").value))
 
 '
 '      .TextMatrix(i, .ColIndex("allocations")) = (IIf(IsNull(rs.Fields("allocations").value), 0, rs.Fields("allocations").value))
'.TextMatrix(i, .ColIndex("Countsofall")) = (IIf(IsNull(rs.Fields("Countsofall").value), 0, rs.Fields("Countsofall").value))
'.TextMatrix(i, .ColIndex("Doneofall")) = (IIf(IsNull(rs.Fields("Doneofall").value), 0, rs.Fields("Doneofall").value))

        rs.MoveNext
            Next i
 
            rs.Close
        End If
   .AutoSize 1, .Cols - 1, False

        .RowHeight(-1) = 300
    End With

End Sub

'
'
'Private Sub ReLineGrid()
'    Dim IntCounter As Integer
'    IntCounter = 0
'    Dim i As Integer
'
'    Dim Percenrage As Double
'
'
'    IntCounter = 0
'  Me.TxtTotalContract.text = 0
'  Me.TxtCommiValue.text = 0
'    Me.TxtInsuranceValue.text = 0
'      Me.TxtWater.text = 0
'      Me.TxtElectricity.text = 0
'        Me.TxtPhone.text = 0
'
'    With Me.GridInstallments
'
'        For i = .FixedRows To .Rows - 1
'                                   If Check17.value = vbChecked Then
'                .TextMatrix(i, .ColIndex("Send")) = -1
'                Else
'                .TextMatrix(i, .ColIndex("Send")) = 0
'
'      End If
'
'            If .TextMatrix(i, .ColIndex("Send")) <> "" Then
'                IntCounter = IntCounter + 1
'                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
'
'                '''////
'
                '''//
'
'
'                     If .Cell(flexcpChecked, i, .ColIndex("Send")) = flexChecked Then
'  Me.TxtTotalContract.text = val(Me.TxtTotalContract.text) + val(.TextMatrix(i, .ColIndex("RentValue")))
'  Me.TxtCommiValue.text = val(Me.TxtCommiValue.text) + val(.TextMatrix(i, .ColIndex("Commissions")))
'  Me.TxtInsuranceValue.text = val(Me.TxtInsuranceValue.text) + val(.TextMatrix(i, .ColIndex("Insurance")))
'  Me.TxtWater.text = val(Me.TxtWater.text) + val(.TextMatrix(i, .ColIndex("Water")))
'  Me.TxtElectricity.text = val(Me.TxtElectricity.text) + val(.TextMatrix(i, .ColIndex("Electric")))
'  Me.TxtPhone.text = val(Me.TxtPhone.text) + val(.TextMatrix(i, .ColIndex("TelandNet")))
'
'  End If
'
'
'
'            End If
'
'        Next i
'
'    End With
'End Sub


Private Sub CmdPrint_Click()
    On Error Resume Next
   ' Dim GrdBack As ClsBackGroundPic
   ' 'Grid.ExtendLastCol = True
   ' Grid.WallPaper = Nothing
   ' 'Grid.AutoSize  0, Grid.Cols - 1, False
   ' Printer.Orientation = VBRUN.PrinterObjectConstants.vbPRORLandscape
 
   ' 'Printer.RightToLeft = True
   ' 'Printer.Print ("Employee Salary Report")

   ' Me.Grid.PrintGrid " ‰»Ì…    „” Œ·’«  ·„  ”œœ »«·ﬂ«„·", True, 2, 1, 1500
   print_report sql
End Sub





'Private Sub Command1_Click()
'    Dim Numbers As String
'    Dim RowNum As Integer
'    Dim Opt As Integer
'    Dim CurrentMessage As String
'    Numbers = ""
'
'    With GridInstallments
'
'        For RowNum = .FixedRows To .Rows - 1
'
'            If .Cell(flexcpChecked, RowNum, .ColIndex("Send")) = flexChecked Then
'
'                '  MsgBox (.TextMatrix(RowNum, .ColIndex("Numbers")))
'                If (.TextMatrix(RowNum, .ColIndex("Cus_mobile"))) <> "" Then
'                    If Numbers = "" Then
'                        Numbers = (.TextMatrix(RowNum, .ColIndex("Cus_mobile")))
'                    Else
'                        Numbers = Numbers & "," & (.TextMatrix(RowNum, .ColIndex("Cus_mobile")))
'                    End If
'
'                End If
'            End If
'
'        Next RowNum
'
'        CurrentMessage = ComposMessage(Me.name)  ', 0, "", Me.TXTMessageDES.text, Opt)
'
'        If Numbers = "" Then Exit Sub
'        SMSSeTTings.SendMessage CurrentMessage, Numbers
'        SMSSeTTings.Hide
'
'    End With
'
'End Sub

Private Sub Form_Load()


 
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
     Fromdate.value = Date
      toDate.value = Date 'rentInstallmentdate
      

    If SystemOptions.UserInterface = EnglishInterface Then

        SetInterface Me
        cahngelang
    End If

End Sub

Function cahngelang()
    Label1(2).Caption = "Project Invoices Not Payed"
    Me.Caption = Label1(2).Caption
    Frame1.Caption = "Color Map"
    Label3.Caption = "Fully"
    Label5.Caption = "Partial"

    Me.Caption = Label1(2).Caption
    CmdPrint.Caption = "Print"
    btnCancel.Caption = "Cancel"

    With Me.Grid
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("id")) = " Bill ID"
        .TextMatrix(0, .ColIndex("bill_date")) = "Bill Date  "
        .TextMatrix(0, .ColIndex("Project_name")) = "Project Name"
        .TextMatrix(0, .ColIndex("End_user_name")) = "End_user_name"
        .TextMatrix(0, .ColIndex("Sub_user_name")) = "Sub_user_name"
        .TextMatrix(0, .ColIndex("total")) = "Bill Total"
        .TextMatrix(0, .ColIndex("ActualTotal")) = "Payed"
        .TextMatrix(0, .ColIndex("result")) = "Variance"
        .TextMatrix(0, .ColIndex("resultpercentage")) = "Variance%"

    End With

End Function

Private Sub FromDate_Change()
If Fromdate.value <> "" Then
 Fromdateh.value = ToHijriDate(Fromdate.value)
 End If
End Sub

Private Sub Fromdateh_LostFocus()
 VBA.Calendar = vbCalGreg
            Fromdate.value = ToGregorianDate(Fromdateh.value)

End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.name, Me.Caption, Me.Caption
End Sub

Private Sub Todate_Change()
    If toDate.value <> "" Then
    todateH.value = ToHijriDate(toDate.value)
    End If
End Sub

Private Sub todateH_LostFocus()
 VBA.Calendar = vbCalGreg
            toDate.value = ToGregorianDate(todateH.value)
 
End Sub
