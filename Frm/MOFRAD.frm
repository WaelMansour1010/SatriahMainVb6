VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form MOFRAD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " «‰Ê«⁄ „›—œ«  «·—« »  «·«”«”Ì…"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18810
   Icon            =   "MOFRAD.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9795
   ScaleWidth      =   18810
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkshowMofradAll 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÌŸÂ— ﬂ«„· ›Ì „” Õﬁ«  «·«Ã«“…"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   10320
      RightToLeft     =   -1  'True
      TabIndex        =   76
      Top             =   8400
      Width           =   1935
   End
   Begin VB.CheckBox chkshowinMosirVac 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÌŸÂ— ›Ì „”Ì— «·«Ã«“« "
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   10320
      RightToLeft     =   -1  'True
      TabIndex        =   75
      Top             =   8160
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   " ”ÃÌ· «·„›—œ«  «· Ì ‰œŒ· ›Ì Õ”«» «·„›—œ"
      Height          =   375
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   53
      ToolTipText     =   "›Ï Õ«·… «‰ «·„›—œ „ €Ì—  ÊÊÕœ‰…  «Ì«„ «Ê ”«⁄« "
      Top             =   8640
      Width           =   3495
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   15
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   0
      Width           =   18825
      Begin VB.TextBox TxtVac_ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   240
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   510
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2580
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Text            =   "modflag"
         Top             =   90
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
         TabIndex        =   10
         Top             =   450
         Visible         =   0   'False
         Width           =   3105
         Begin MSDataListLib.DataCombo DCUser 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   -255
            TabIndex        =   11
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
            TabIndex        =   12
            Top             =   45
            Width           =   855
         End
      End
      Begin MSComctlLib.ImageList GrdImageList 
         Left            =   3120
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
               Picture         =   "MOFRAD.frx":058A
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MOFRAD.frx":0924
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MOFRAD.frx":0CBE
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MOFRAD.frx":1058
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MOFRAD.frx":13F2
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MOFRAD.frx":178C
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MOFRAD.frx":1B26
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MOFRAD.frx":20C0
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   90
         TabIndex        =   15
         Top             =   30
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
         ButtonImage     =   "MOFRAD.frx":245A
         ColorButton     =   16777215
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   555
         TabIndex        =   16
         Top             =   30
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
         ButtonImage     =   "MOFRAD.frx":27F4
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1155
         TabIndex        =   17
         Top             =   30
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
         ButtonImage     =   "MOFRAD.frx":2B8E
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   1620
         TabIndex        =   18
         Top             =   30
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
         ButtonImage     =   "MOFRAD.frx":2F28
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   17760
         Picture         =   "MOFRAD.frx":32C2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«‰Ê«⁄ „›—œ«  «·—« » «·«”«”Ì…"
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
         Left            =   14055
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   90
         Width           =   3240
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   6675
      Left            =   -120
      TabIndex        =   30
      Top             =   570
      Width           =   18945
      _cx             =   33417
      _cy             =   11774
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
      Rows            =   21
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"MOFRAD.frx":5FC0
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
   Begin VB.Frame Frm2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1965
      Left            =   15
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   7395
      Width           =   18915
      Begin VB.CommandButton Command2 
         Caption         =   " ÿ»Ìﬁ ⁄·Ì «·ﬂ·"
         Height          =   255
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   77
         Top             =   840
         Width           =   1575
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÌÕ”» ⁄·Ï √”«” "
         Height          =   615
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   0
         Width           =   3015
         Begin VB.OptionButton culcopt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "√Ì«„ «·‘Â— «·„ »ﬁÌ…"
            Height          =   195
            Index           =   1
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton culcopt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "30 ÌÊ„"
            Height          =   195
            Index           =   0
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CheckBox ChAllowIntrod 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„›—œ »œ·«  „ﬁœ„…"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   12000
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CheckBox ChSalary 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„›—œ —« » «”«”Ì"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   12000
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CheckBox CheckReward 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„›—œ „ﬂ«›∆…"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   12000
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CheckBox CheckDiscount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„›—œ Œ’„"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   12000
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   960
         Width           =   1815
      End
      Begin VB.CheckBox ChekAbsent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„›—œ €Ì«»"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   12000
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox INSMofradCHK 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„›—œ  √„Ì‰"
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   13680
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CheckBox CHKinsurances 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÌŒ÷⁄ ·· √„Ì‰« "
         Enabled         =   0   'False
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   13680
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CheckBox Chkacc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "·Â  ÊÃÌÂ „Õ«”»Ì"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   13680
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   960
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox ChkADVView 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„›—œ „œ›Ê⁄ „ﬁœ„"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   13680
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   1680
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Frame Frame5 
         Caption         =   "›Ì Õ«·Â «·„‘«—Ì⁄ ›Ì Õ«·… «·Œ’„ Ì⁄·Ì «·«Ì—«œ"
         Height          =   255
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   1080
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.CheckBox chkAdvPaymentdAccount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Ì” Œœ„ Õ”«» «·„œ›Ê⁄«  «·„ﬁœ„… ··„ÊŸ›"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   15480
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   1680
         Width           =   3255
      End
      Begin VB.CheckBox ChkInCrease 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÌœŒ· ›Ì    «·“Ì«œ… «·”‰ÊÌ…"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   15600
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   1095
         Left            =   -480
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   2640
         Width           =   6255
         Begin VB.CheckBox ChkOverTime 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÌœŒ· ›Ì  Õ”«» «·«÷«›Ì"
            Height          =   255
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   240
            Width           =   3135
         End
         Begin VB.CheckBox ChkDiscount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÌœŒ· ›Ì  Õ”«» «·Œ’„"
            Height          =   255
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   840
            Width           =   3135
         End
         Begin VB.CheckBox ChkPunch 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÌœŒ· ›Ì  Õ”«» «·„ﬂ«›√…"
            Height          =   255
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   600
            Width           =   3135
         End
         Begin VB.CheckBox ChkLate 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÌœŒ· ›Ì  Õ”«» «· √ŒÌ—"
            Height          =   255
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   240
            Width           =   3135
         End
         Begin VB.CheckBox ChkAbsence 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÌœŒ· ›Ì  Õ”«» «·€Ì«»"
            Height          =   255
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   600
            Width           =   3135
         End
      End
      Begin VB.CheckBox ChkAloc2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÌœŒ· ›Ì    „Œ’’ „ﬂ«›√… ‰Â«Ì… «·Œœ„…"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   15600
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   960
         Width           =   3135
      End
      Begin VB.CheckBox ChkAloc1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÌœŒ· ›Ì    „Œ’’ «·«Ã«“« "
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   15600
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   720
         Width           =   3135
      End
      Begin VB.CheckBox ChkZmamAccount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Ì” Œœ„ Õ”«» «·–„„"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   16080
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   1440
         Width           =   2655
      End
      Begin VB.CheckBox ChkView 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÌŸÂ— ›Ì «·„”Ì—"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   13680
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   720
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»Ì⁄… «·„›—œ"
         Height          =   615
         Left            =   8520
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   0
         Width           =   2415
         Begin VB.OptionButton Option2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„ €Ì—"
            Height          =   195
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "À«» "
            Height          =   195
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÊÕœ… «·„›—œ"
         Height          =   615
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   0
         Width           =   3015
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”«⁄« "
            Height          =   195
            Index           =   2
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "√Ì«„"
            Height          =   195
            Index           =   1
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„…"
            Height          =   195
            Index           =   0
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·„›—œ"
         Height          =   615
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   0
         Width           =   2295
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«÷«›…"
            Height          =   195
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œ’„"
            Height          =   195
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
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
         Height          =   315
         Left            =   11040
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Tag             =   "Enter Component Name"
         Top             =   285
         Width           =   2400
      End
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
         Height          =   315
         Left            =   13515
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·„›—œ"
         Top             =   285
         Width           =   4080
      End
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
         Left            =   17640
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   285
         Width           =   1065
      End
      Begin VB.ComboBox CmbType 
         BackColor       =   &H80000018&
         Height          =   315
         ItemData        =   "MOFRAD.frx":6179
         Left            =   -120
         List            =   "MOFRAD.frx":6189
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1230
         Visible         =   0   'False
         Width           =   1005
      End
      Begin MSDataListLib.DataCombo DCAccounts 
         Height          =   315
         Left            =   5040
         TabIndex        =   5
         Top             =   840
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCAccounts1 
         Height          =   315
         Left            =   5040
         TabIndex        =   4
         Top             =   1320
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ”«» «·—»ÿ 2"
         Height          =   255
         Index           =   3
         Left            =   8640
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ”«» «·—»ÿ "
         Height          =   255
         Index           =   2
         Left            =   8640
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«”„ «‰Ã·Ì“Ì"
         Height          =   285
         Index           =   1
         Left            =   11505
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   0
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«”„ ⁄—»Ì"
         Height          =   285
         Index           =   0
         Left            =   15660
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   0
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·„›—œ"
         Height          =   195
         Index           =   3
         Left            =   17625
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   30
         Width           =   990
      End
   End
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   1020
      Left            =   -60
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   8775
      Width           =   18960
      _cx             =   33443
      _cy             =   1799
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
         Height          =   330
         Left            =   17415
         TabIndex        =   21
         Top             =   555
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "ÃœÌœ"
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
         ButtonImage     =   "MOFRAD.frx":61A2
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnSave 
         Height          =   330
         Left            =   12510
         TabIndex        =   22
         Top             =   555
         Width           =   1590
         _ExtentX        =   2805
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
         ButtonImage     =   "MOFRAD.frx":653C
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnModify 
         Height          =   330
         Left            =   15075
         TabIndex        =   23
         Top             =   555
         Width           =   1470
         _ExtentX        =   2593
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
         ButtonImage     =   "MOFRAD.frx":68D6
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUndo 
         Height          =   330
         Left            =   10545
         TabIndex        =   24
         Top             =   555
         Width           =   1350
         _ExtentX        =   2381
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
         ButtonImage     =   "MOFRAD.frx":6C70
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnDelete 
         Height          =   330
         Left            =   8220
         TabIndex        =   25
         Top             =   555
         Width           =   1470
         _ExtentX        =   2593
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
         ButtonImage     =   "MOFRAD.frx":700A
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnQuery 
         Height          =   330
         Left            =   5880
         TabIndex        =   26
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
         ButtonImage     =   "MOFRAD.frx":75A4
         ColorButton     =   14737632
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUpdate 
         Height          =   330
         Left            =   6045
         TabIndex        =   27
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
         ButtonImage     =   "MOFRAD.frx":793E
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnPrint 
         Height          =   285
         Left            =   4845
         TabIndex        =   28
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
         ButtonImage     =   "MOFRAD.frx":7CD8
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnCancel 
         Height          =   330
         Left            =   3945
         TabIndex        =   29
         Top             =   555
         Width           =   1470
         _ExtentX        =   2593
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
         ButtonImage     =   "MOFRAD.frx":8072
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton ISButton1 
         Height          =   285
         Left            =   6120
         TabIndex        =   58
         TabStop         =   0   'False
         ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
         Top             =   600
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
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
         ButtonImage     =   "MOFRAD.frx":840C
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
         Left            =   2265
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·”Ã·« :"
         Height          =   210
         Index           =   1
         Left            =   570
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   600
         Width           =   975
      End
      Begin VB.Label LabCurrRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   615
         Width           =   675
      End
      Begin VB.Label LabCountRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   600
         Width           =   540
      End
   End
End
Attribute VB_Name = "MOFRAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim ii As Long
Dim fillgridCursor As Integer

Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    On Error GoTo ErrTrap

    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If

    If TxtVac_ID.text <> "" Then
        If CheckDelCountry(val(Me.TxtVac_ID.text)) = False Then
            Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã·...!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If

        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.Title)

        If MSGType = vbYes Then
            RsSavRec.Find "ID=" & val(TxtVac_ID.text), , adSearchForward, 1
            RsSavRec.delete
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            '------------------------------ Move Next ---------------------------.
            FillGridWithData
            BtnNext_Click
        End If
    End If

    Exit Sub
ErrTrap:
 
    Select Case Err.Number

        Case -2147217873, -2147467259
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.Title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select

End Sub

Private Sub BtnFirst_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtVac_ID.text)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnLast_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtVac_ID.text)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
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

    If TxtVac_ID.text <> "" Then
        TxtModFlg = "E"
        Frm2.Enabled = True
        Me.TxtVacName.SetFocus
    End If

    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147467259
            'Could not update; currently locked.
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
        
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If

    End Select

End Sub

Private Sub btnNew_Click()
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    'If DoPremis(Do_New, Me.name, True) = False Then
    '    Exit Sub
    'End If
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    clear_all Me
    TxtModFlg.text = "N"

    My_SQL = "mofrad"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial.text = rs.RecordCount + 1
    Else
        TxtSerial.text = 1
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
        FindRec val(TxtVac_ID.text)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtVac_ID.text)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
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
    '---------------------- check if data Vaclete -----------------------
If TxtVacNamee.text = "" Then TxtVacNamee.text = TxtVacName.text

    For Each CtrlTxt In Me.Controls

        If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
            If CtrlTxt.text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.Title
                CtrlTxt.SetFocus
                Exit Sub
            End If
        End If

    Next

    '------------------------------ check if Empcode exist ----------------------

    StrVacName = IsRecExist("mofrad", "Name", Trim(TxtVacName.text), "Name", "ID<>'" & Trim(TxtVac_ID.text) & "'")

    If StrVacName <> "" Then
        Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
         
        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.Title
        TxtVacName.SetFocus
    
        Exit Sub

    End If

    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.text

            '------------------------------ new record ----------------------------
        Case "N"
      
            '------------------------- save record -----------------------------
            AddNewRec
            BtnLast_Click

        Case "E"

            '----------------------------- save edit -------------------------------
            FiLLRec
    End Select

    Exit Sub
ErrTrap:
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.Title

End Sub
 
Private Sub BtnUndo_Click()
    FindRec val(TxtVac_ID.text)
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

    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.Title
ErrTrap:
End Sub

Private Sub Chkacc_Click()

If Chkacc.value = vbChecked Then
Label2(3).Visible = True
DCAccounts1.Visible = True

Else

Label2(3).Visible = False
DCAccounts1.Visible = False


End If

End Sub

Private Sub chkAdvPaymentdAccount_Click()
    If chkAdvPaymentdAccount.value = vbChecked Then
        DCAccounts.BoundText = ""
ChkZmamAccount.value = vbUnchecked
    End If

End Sub

Private Sub ChkZmamAccount_Click()

    If ChkZmamAccount.value = vbChecked Then
        DCAccounts.BoundText = ""
chkAdvPaymentdAccount.value = vbUnchecked
    End If

End Sub

Private Sub Command1_Click()
FrmChangedComponentData2.mofradId = val(Me.Grid.TextMatrix(Me.Grid.row, Me.Grid.ColIndex("ID")))
    FrmChangedComponentData2.show

    Exit Sub

    If Option2.value = True And Opt(2).value = False Then

    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·« Ì„ﬂ‰  ÕœÌœ „›—œ«  «·« ··„›—œ«  «·„ €Ì—… ›ﬁÿ «· Ì ÊÕœ‰Â« ”«⁄«  «Ê «Ì«„ ›ﬁÿ ", vbInformation
        Else
            MsgBox "Can't open this Screen for Fixed Component Or value Component ", vbInformation

        End If

    End If

End Sub

 

Private Sub DataCombo1_Click(Area As Integer)

End Sub

Private Sub DataCombo1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 17815
    End If
End Sub

Private Sub Command2_Click()
updatetoallacc
End Sub

Private Sub DCAccounts_Click(Area As Integer)
    ChkZmamAccount.value = vbUnchecked
    chkAdvPaymentdAccount.value = vbUnchecked
 
End Sub

Private Sub DCAccounts_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 50150221
    End If

End Sub

Private Sub DCAccounts1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 17815
    End If

End Sub
Function updatetoallacc()
Dim str As String

str = " update MOFRAD set Account_code='" & DCAccounts.BoundText & "'"
str = str & " , Account_code1 ='" & DCAccounts1.BoundText & "'"
Cn.Execute str
'RsSavRec.Resync adAffectAllChapters

FillGridWithData
MsgBox " „ «· ÿ»Ìﬁ ⁄·Ì «·ﬂ·", vbInformation



End Function
Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetAccountingCodes Me.DCAccounts, True
    Dcombos.GetAccountingCodes Me.DCAccounts1, True
    
If SystemOptions.ProjectDiscountPolicy = 1 Then
Frame5.Visible = True
Else
Frame5.Visible = False
End If


    My_SQL = "mofrad"
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    Me.TxtModFlg.text = "R"
    Resize_Form Me
    'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL

    FillGridWithData

    With Me.Grid
        .Cell(flexcpPicture, 0, .ColIndex("Name")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
        .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon

        For i = 0 To .Cols - 1
            .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
        Next
   
        .ExtendLastCol = True
        .WallPaper = BKGrndPic.Picture
        .RowHeight(-1) = 300
    End With

    BtnFirst_Click
    ShowTip

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If

ErrTrap:
End Sub

Function ChangeLang()
    Me.Command1.Caption = "Specify Component .."

    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic
    Me.Caption = "Components Types"
    Label1(2).Caption = Me.Caption
    Label2(3).Caption = "Credit Acc"
    ChAllowIntrod.Caption = "Intro.Allowances"
    Frame1.Caption = "Type"
    Option3.Caption = "Addition"
    Option4.Caption = "Discount"
    Option1.Caption = "Fixed"
    Option2.Caption = "Changed"
    Opt(0).Caption = "Value"
    Opt(1).Caption = "Days"
    ChSalary.Caption = "Salary"
    Opt(2).Caption = "Hours"
    CheckReward.Caption = "Reward"
    CheckDiscount.Caption = "Component Discount"
    Frame2.Caption = "Component Value"
    Frame3.Caption = "Component type"
    ChkView.Caption = "Show in Report"
    ChkADVView.Caption = "Adv. Comp"
    Chkacc.Caption = "Have Acc entery "
    CHKinsurances.Caption = "Have GOSI"
    INSMofradCHK.Caption = "Have Insurances"
    ChekAbsent.Caption = "Component Absence"
    With Me.Grid
        .TextMatrix(0, .ColIndex("Aloc1")) = "Vacations Alloc"
        .TextMatrix(0, .ColIndex("Aloc2")) = "End of Ser Alloc"

        .TextMatrix(0, .ColIndex("id")) = "Code"
        .TextMatrix(0, .ColIndex("name")) = "Name Ar"
        .TextMatrix(0, .ColIndex("namee")) = "Name Eng"
        .TextMatrix(0, .ColIndex("Absence")) = "Absence"
        .TextMatrix(0, .ColIndex("Late")) = "Late"
        .TextMatrix(0, .ColIndex("discount")) = "Discount"
        .TextMatrix(0, .ColIndex("OverTime")) = "OverTime"
        .TextMatrix(0, .ColIndex("Punch")) = "Bonus"
       End With

    Label2(2).Caption = "Account"
    ChkZmamAccount.Caption = "Emp Account"
    chkAdvPaymentdAccount.Caption = "Using Adv Pay for employee"
    Label1(3).Caption = "Code"
    Label1(0).Caption = "Name AR"
    Label1(1).Caption = "Name Eng"
    Label2(0).Caption = "Curr. Rec."
    Label2(1).Caption = "Rec. Count."
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    ISButton1.Caption = "Print"
    BtnUndo.Caption = "Undo"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
    ChkOverTime.Caption = " OverTime"
    Me.ChkAbsence.Caption = "Absence"
    Me.ChkDiscount.Caption = "Discount"
    Me.ChkLate.Caption = "Late"
    Me.ChkPunch.Caption = "Bonus"
    ChkAloc1.Caption = "Vacation Allocations"
    ChkAloc2.Caption = "EOS Allocation"
    ChkInCrease.Caption = "Increase Allocation"
   '######################## khaled was here #########################
   chkshowinMosirVac.Caption = "Show in vacation Salary sheet"
   chkshowMofradAll.Caption = "Show all"
   Frame6.Caption = "Calculate based on"
   culcopt(0).Caption = "30 days"
   culcopt(1).Caption = "Remaining days"
   '##################################################################
    
End Function

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
       
                'SaveData
                btnSave_Click

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub
Private Sub Form_Terminate()
    'Set cSearchDCombo = Nothing
    'Set BKGrndPic = Nothing
    Set FrmVacancy = Nothing
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

Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("mofrad", "ID", "")
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

Public Sub FiLLRec()
    On Error GoTo ErrTrap

    RsSavRec.Fields("Name").value = IIf(TxtVacName.text <> "", Trim(TxtVacName.text), Null)
    RsSavRec.Fields("Namee").value = IIf(TxtVacNamee.text <> "", Trim(TxtVacNamee.text), Null)
    RsSavRec.Fields("Account_code").value = IIf(Me.DCAccounts.BoundText = "", Null, Me.DCAccounts.BoundText)
    RsSavRec.Fields("Account_code1").value = IIf(Me.DCAccounts1.BoundText = "", Null, Me.DCAccounts1.BoundText)
    
    '################### Khaled was here ##########################
    If chkshowinMosirVac.value = vbChecked Then
        RsSavRec.Fields("showinMosirVac").value = 1
    Else
        RsSavRec.Fields("showinMosirVac").value = 0
    End If
    
    If chkshowMofradAll.value = vbChecked Then
        RsSavRec.Fields("showMofradAll").value = 1
    Else
        RsSavRec.Fields("showMofradAll").value = 0
    End If
    
    If culcopt(0).value = True Then
        RsSavRec.Fields("culc30orRminder").value = 0
    ElseIf culcopt(1).value = True Then
        RsSavRec.Fields("culc30orRminder").value = 1
    End If
    '##############################################################
    If ChAllowIntrod.value = vbChecked Then
        RsSavRec.Fields("AllowIntrod").value = 1
    Else
        RsSavRec.Fields("AllowIntrod").value = 0
    End If
    
    If Option1.value = True Then
        RsSavRec.Fields("FixedOrChanged").value = 0
    ElseIf Option2.value = True Then
        RsSavRec.Fields("FixedOrChanged").value = 1
 
    End If

    If Option3.value = True Then
        RsSavRec.Fields("AddOrDiscount").value = 0
    ElseIf Option4.value = True Then
        RsSavRec.Fields("AddOrDiscount").value = 1
 
    End If

    If Opt(0).value = True Then
        RsSavRec.Fields("Unit").value = 0
    ElseIf Opt(1).value = True Then
        RsSavRec.Fields("Unit").value = 1
    ElseIf Opt(2).value = True Then
        RsSavRec.Fields("Unit").value = 2
    End If

    If ChkAloc1.value = vbChecked Then
        RsSavRec.Fields("Aloc1").value = 1
    Else
        RsSavRec.Fields("Aloc1").value = 0
    End If
'''///////
 If Me.ChSalary.value = vbChecked Then
        RsSavRec.Fields("Salary").value = 1
    Else
        RsSavRec.Fields("Salary").value = 0
    End If
'''//

    If ChkInCrease.value = vbChecked Then
        RsSavRec.Fields("InCrease").value = 1
    Else
        RsSavRec.Fields("InCrease").value = 0
    End If
    
    If ChkAloc2.value = vbChecked Then
        RsSavRec.Fields("Aloc2").value = 1
    Else
        RsSavRec.Fields("Aloc2").value = 0
    End If

    If ChkOverTime.value = vbChecked Then
        RsSavRec.Fields("OverTime").value = 1
    Else
        RsSavRec.Fields("OverTime").value = 0
    End If

    If ChkDiscount.value = vbChecked Then
        RsSavRec.Fields("Discount").value = 1
    Else
        RsSavRec.Fields("Discount").value = 0
    End If

    If ChkPunch.value = vbChecked Then
        RsSavRec.Fields("Punch").value = 1
    Else
        RsSavRec.Fields("Punch").value = 0
    End If

    If ChkLate.value = vbChecked Then
        RsSavRec.Fields("Late").value = 1
    Else
        RsSavRec.Fields("Late").value = 0
    End If

    If ChkAbsence.value = vbChecked Then
        RsSavRec.Fields("Absence").value = 1
    Else
        RsSavRec.Fields("Absence").value = 0
    End If

    If ChkView.value = vbChecked Then
        RsSavRec.Fields("ViewComp").value = 1
    Else
        RsSavRec.Fields("ViewComp").value = 0
    End If
    
    If Chkacc.value = vbChecked Then
        RsSavRec.Fields("acc").value = 1
    Else
        RsSavRec.Fields("acc").value = 0
    End If
    
    If ChkADVView.value = vbChecked Then
        RsSavRec.Fields("ADVView").value = 1
    Else
        RsSavRec.Fields("ADVView").value = 0
    End If
    
    If ChkZmamAccount.value = vbChecked Then
        RsSavRec.Fields("ZmamAccount").value = 1
    Else
        RsSavRec.Fields("ZmamAccount").value = 0
    End If

    If chkAdvPaymentdAccount.value = vbChecked Then
        RsSavRec.Fields("AdvPaymentdAccount").value = 1
    Else
        RsSavRec.Fields("AdvPaymentdAccount").value = 0
    End If
    ''''''''''''''''''''''''''''''
    If CHKinsurances.value = vbChecked Then
    RsSavRec.Fields("Insurances").value = 1
    Else
    RsSavRec.Fields("Insurances").value = 0
    End If
    
    If INSMofradCHK.value = vbChecked Then
    RsSavRec.Fields("INSMofrad").value = 1
    Else
    RsSavRec.Fields("INSMofrad").value = 0
    End If
   ''''''''
   If CheckReward.value = vbChecked Then
    RsSavRec.Fields("Reward").value = 1
    Else
    RsSavRec.Fields("Reward").value = 0
    End If
    ''''''''''
    If ChekAbsent.value = vbChecked Then
    RsSavRec.Fields("MofrdAbcen").value = 1
    Else
    RsSavRec.Fields("MofrdAbcen").value = 0
    End If
    If CheckDiscount.value = vbChecked Then
    RsSavRec.Fields("MofrdDiscount").value = 1
    Else
    RsSavRec.Fields("MofrdDiscount").value = 0
    End If
    ''''''''''''''''''''''''''''''''''''
     
    
    RsSavRec.update

    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Else
        MsgBox " Saved SuccessFully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    End If
fillgridCursor = 1
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
    TxtVac_ID.text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    TxtVacName.text = IIf(IsNull(RsSavRec.Fields("Name").value), "", RsSavRec.Fields("Name").value)
    TxtVacNamee.text = IIf(IsNull(RsSavRec.Fields("Namee").value), "", RsSavRec.Fields("Namee").value)
    Me.DCAccounts.BoundText = IIf(IsNull(RsSavRec("Account_Code").value), "", RsSavRec("Account_Code").value)
    Me.DCAccounts1.BoundText = IIf(IsNull(RsSavRec("Account_Code1").value), "", RsSavRec("Account_Code1").value)
    
    '############################ Khaled was here ################################
    If RsSavRec.Fields("showinMosirVac").value = True Then
        Me.chkshowinMosirVac.value = vbChecked
    Else
        chkshowinMosirVac.value = vbUnchecked
    End If
    
    If RsSavRec.Fields("showMofradAll").value = True Then
        Me.chkshowMofradAll.value = vbChecked
    Else
        chkshowMofradAll.value = vbUnchecked
    End If
    
    If IsNull(RsSavRec.Fields("culc30orRminder").value) Then
        culcopt(0).value = True
    Else
        culcopt(RsSavRec.Fields("culc30orRminder").value).value = True
    End If
    '#############################################################################
    If IsNull(RsSavRec.Fields("Unit").value) Then
        Opt(0).value = True
    Else
        Opt(RsSavRec.Fields("Unit").value).value = True
    End If

    If IsNull(RsSavRec.Fields("AddOrDiscount").value) Then
        Option3.value = True
    Else

        If RsSavRec.Fields("AddOrDiscount").value = False Then
            Option3.value = True
        Else
            Option4.value = True
        End If

    End If
    
    If RsSavRec.Fields("Salary").value = True Then
        Me.ChSalary.value = vbChecked
    Else
        ChSalary.value = vbUnchecked
    End If
        
    If IsNull(RsSavRec.Fields("FixedOrChanged").value) Then
        Option1.value = True
    Else

        If RsSavRec.Fields("FixedOrChanged").value = False Then
            Option1.value = True
        Else
            Option2.value = True
        End If

    End If

    If IsNull(RsSavRec.Fields("Aloc1").value) Then
        ChkAloc1.value = vbUnchecked
    Else

        If RsSavRec.Fields("Aloc1").value = True Then
            ChkAloc1.value = vbChecked
        Else
            ChkAloc1.value = vbUnchecked
        End If

    End If


    If IsNull(RsSavRec.Fields("InCrease").value) Then
        ChkInCrease.value = vbUnchecked
    Else

        If RsSavRec.Fields("InCrease").value = True Then
            ChkInCrease.value = vbChecked
        Else
            ChkInCrease.value = vbUnchecked
        End If

    End If
 
        If IsNull(RsSavRec.Fields("AllowIntrod").value) Then
        ChAllowIntrod.value = vbUnchecked
       Else

        If RsSavRec.Fields("AllowIntrod").value = 1 Then
            ChAllowIntrod.value = vbChecked
        Else
            ChAllowIntrod.value = vbUnchecked
        End If

    End If
    
    If IsNull(RsSavRec.Fields("Aloc2").value) Then
        ChkAloc2.value = vbUnchecked
    Else

        If RsSavRec.Fields("Aloc2").value = True Then
            ChkAloc2.value = vbChecked
        Else
            ChkAloc2.value = vbUnchecked
        End If

    End If

    If IsNull(RsSavRec.Fields("OverTime").value) Then
        ChkOverTime.value = vbUnchecked
    Else

        If RsSavRec.Fields("OverTime").value = True Then
            ChkOverTime.value = vbChecked
        Else
            ChkOverTime.value = vbUnchecked
        End If

    End If

    If IsNull(RsSavRec.Fields("Discount").value) Then
        ChkDiscount.value = vbUnchecked
    Else

        If RsSavRec.Fields("Discount").value = True Then
            ChkDiscount.value = vbChecked
        Else
            ChkDiscount.value = vbUnchecked
        End If

    End If

    If IsNull(RsSavRec.Fields("Punch").value) Then
        ChkPunch.value = vbUnchecked
    Else

        If RsSavRec.Fields("Punch").value = True Then
            ChkPunch.value = vbChecked
        Else
            ChkPunch.value = vbUnchecked
        End If

    End If

    If IsNull(RsSavRec.Fields("Late").value) Then
        ChkLate.value = vbUnchecked
    Else

        If RsSavRec.Fields("Late").value = True Then
            ChkLate.value = vbChecked
        Else
            ChkLate.value = vbUnchecked
        End If

    End If

    If IsNull(RsSavRec.Fields("ViewComp").value) Then
        ChkView.value = vbUnchecked
    Else

        If RsSavRec.Fields("ViewComp").value = True Then
            ChkView.value = vbChecked
        Else
            ChkView.value = vbUnchecked
        End If

    End If

   ''''''''''''''''''''''''
     If RsSavRec.Fields("Insurances").value = True Then
     CHKinsurances.value = vbChecked
     Else
     CHKinsurances.value = vbUnchecked
     End If
     
     If RsSavRec.Fields("INSMofrad").value = True Then
     INSMofradCHK.value = vbChecked
     Else
     INSMofradCHK.value = vbUnchecked
     End If
    
     ''''''''''''''''''''''''''''''''
 
    If IsNull(RsSavRec.Fields("acc").value) Then
        Chkacc.value = vbUnchecked
        Label2(3).Visible = False
    DCAccounts1.Visible = False


    Else

        If RsSavRec.Fields("acc").value = True Then
            Chkacc.value = vbChecked
            Label2(3).Visible = True
DCAccounts1.Visible = True
        Else
            Chkacc.value = vbUnchecked
                    Label2(3).Visible = False
DCAccounts1.Visible = False

        End If

    End If
    
    
    If IsNull(RsSavRec.Fields("ADVView").value) Then
                  ChkADVView.value = vbUnchecked
    Else

                    If RsSavRec.Fields("ADVView").value = True Then
                        ChkADVView.value = vbChecked
                    Else
                        ChkADVView.value = vbUnchecked
                    End If

    End If
    
    
    If IsNull(RsSavRec.Fields("ZmamAccount").value) Then
        ChkZmamAccount.value = vbUnchecked
    Else

        If RsSavRec.Fields("ZmamAccount").value = True Then
            ChkZmamAccount.value = vbChecked
        Else
            ChkZmamAccount.value = vbUnchecked
        End If

    End If



    If IsNull(RsSavRec.Fields("AdvPaymentdAccount").value) Then
        chkAdvPaymentdAccount.value = vbUnchecked
    Else

        If RsSavRec.Fields("AdvPaymentdAccount").value = True Then
            chkAdvPaymentdAccount.value = vbChecked
        Else
            chkAdvPaymentdAccount.value = vbUnchecked
        End If

    End If
    
    
    
    If IsNull(RsSavRec.Fields("Absence").value) Then
        ChkAbsence.value = vbUnchecked
    Else

        If RsSavRec.Fields("Absence").value = True Then
            ChkAbsence.value = vbChecked
        Else
            ChkAbsence.value = vbUnchecked
        End If

    End If
  '''///////////////
      If IsNull(RsSavRec.Fields("MofrdDiscount").value) Then
        CheckDiscount.value = vbUnchecked
    Else

        If RsSavRec.Fields("MofrdDiscount").value = True Then
            CheckDiscount.value = vbChecked
        Else
            CheckDiscount.value = vbUnchecked
        End If

    End If
    
    ''''
        If IsNull(RsSavRec.Fields("MofrdAbcen").value) Then
        ChekAbsent.value = vbUnchecked
    Else

        If RsSavRec.Fields("MofrdAbcen").value = True Then
            ChekAbsent.value = vbChecked
        Else
            ChekAbsent.value = vbUnchecked
        End If

    End If
    '''
        If IsNull(RsSavRec.Fields("Reward").value) Then
        CheckReward.value = vbUnchecked
    Else

        If RsSavRec.Fields("Reward").value = True Then
            CheckReward.value = vbChecked
        Else
            CheckReward.value = vbUnchecked
        End If

    End If

    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount

    With Grid

        For i = 1 To .rows - 1

            If Trim(TxtVac_ID.text) = .TextMatrix(i, .ColIndex("ID")) Then
                TxtSerial.text = .TextMatrix(i, .ColIndex("Ser"))
            
                .row = i
                Exit Sub
            End If

        Next

    End With

ErrTrap:

End Sub

Public Sub EditRec(StrTable As String, _
                   RecId As String)
    'My_SQL = "select * From " & StrTable & " where "
    'RsSavRec.Open My_SQL, cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
    FiLLRec

End Sub

Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap
    If fillgridCursor = 1 Then Exit Sub
    FindRec val(Me.Grid.TextMatrix(Me.Grid.row, Me.Grid.ColIndex("ID")))
ErrTrap:
End Sub
Private Sub ISButton1_Click()
On Error GoTo ErrTrap
   If val(Me.TxtSerial.text) <> 0 Then
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
    
    sql = "SELECT     id, name, nameE, Absence, Late, Punch, Discount, OverTime, AddOrDiscount, Unit, FixedOrChanged, Account_Code, ViewComp, ZmamAccount, Aloc1, Aloc2, InCrease, AdvPaymentdAccount,"
    sql = sql & "      Account_code1 , ADVView"
    sql = sql & "    From dbo.MOFRAD"
            
         
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "MOFRADRPT.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "MOFRADRPTEE.rpt"
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
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
ErrTrap:
  End Function
Private Sub TxtVac_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.text
    TxtModFlg.text = ""
    TxtModFlg = TxtMod
End Sub

Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.Find "ID=" & RecId, , adSearchForward, 1

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

'Private Sub TxtVacCode_KeyPress(KeyAscii As Integer)
'KeyAscii = DataFormat(ChrOnly, KeyAscii)
'End Sub

Private Sub TxtModFlg_Change()

    If TxtModFlg.text = "N" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        ISButton1.Enabled = False
        Me.btnQuery.Enabled = False
        Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
        ' btnNext.Enabled = False
        ' btnPrevious.Enabled = False
        ' btnFirst.Enabled = False
        ' btnLast.Enabled = False
    
    ElseIf TxtModFlg.text = "R" Then
        Frm2.Enabled = False
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False
        ISButton1.Enabled = False

        If TxtVac_ID.text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
        End If

        BtnUpdate.Enabled = True
        Me.btnQuery.Enabled = True
        '   Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
        ISButton1.Enabled = True
    
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
        ISButton1.Enabled = False
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
    My_SQL = "select * From mofrad order by ID"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
                .TextMatrix(i, .ColIndex("NAME")) = IIf(IsNull(rs.Fields("Name").value), "", rs.Fields("Name").value)
               
                .TextMatrix(i, .ColIndex("NAMEe")) = IIf(IsNull(rs.Fields("Namee").value), "", rs.Fields("Namee").value)
            
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs.Fields("ID").value), "", rs.Fields("ID").value)

                .TextMatrix(i, .ColIndex("Absence")) = IIf(IsNull(rs.Fields("Absence").value), "", rs.Fields("Absence").value)
            
                .TextMatrix(i, .ColIndex("Late")) = IIf(IsNull(rs.Fields("Late").value), "", rs.Fields("Late").value)
            
                .TextMatrix(i, .ColIndex("Punch")) = IIf(IsNull(rs.Fields("Punch").value), "", rs.Fields("Punch").value)
            
                .TextMatrix(i, .ColIndex("discount")) = IIf(IsNull(rs.Fields("discount").value), "", rs.Fields("discount").value)
            
                .TextMatrix(i, .ColIndex("OverTime")) = IIf(IsNull(rs.Fields("OverTime").value), "", rs.Fields("OverTime").value)
            
                .TextMatrix(i, .ColIndex("Aloc1")) = IIf(IsNull(rs.Fields("Aloc1").value), "", rs.Fields("Aloc1").value)
            
                .TextMatrix(i, .ColIndex("Aloc2")) = IIf(IsNull(rs.Fields("Aloc2").value), "", rs.Fields("Aloc2").value)
            
                If IsNull(rs.Fields("AddOrDiscount").value) Then
                    .Cell(flexcpBackColor, i, 1, i, 8) = &H80FF80
                Else

                    If (rs.Fields("AddOrDiscount").value) = False Then
                        .Cell(flexcpBackColor, i, 1, i, 8) = &H80FF80
                    Else
                        .Cell(flexcpBackColor, i, 1, i, 8) = &HC0C0FF
                    End If
   
                End If
        
                rs.MoveNext
            Next

            rs.Close
        End If

        .RowHeight(-1) = 300
.row = val(TxtVac_ID)
    End With
fillgridCursor = 0
ErrTrap:
End Sub

'-------------------------------------------------------------
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

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            btnNew_Click
        Else
            Sendkeys "{TAB}"
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

