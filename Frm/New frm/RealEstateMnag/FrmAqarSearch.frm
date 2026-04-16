VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmAqarSearch 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19020
   Icon            =   "FrmAqarSearch.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   6615
   ScaleWidth      =   19020
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame5 
      Caption         =   "»Ì«‰«  «·„·ﬂÌÂ"
      Height          =   1215
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   5400
      Width           =   10935
      Begin VB.TextBox txtsuckno 
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
         TabIndex        =   57
         Top             =   240
         Width           =   2145
      End
      Begin VB.TextBox txtauthorizationname 
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
         Left            =   3360
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   720
         Width           =   6465
      End
      Begin MSDataListLib.DataCombo dcsupplier 
         Height          =   315
         Left            =   3240
         TabIndex        =   58
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
         Top             =   240
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker dpsuckdate 
         Height          =   315
         Left            =   120
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   720
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy/M/d"
         Format          =   156434435
         CurrentDate     =   37140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„«·ﬂ"
         Height          =   285
         Index           =   30
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   240
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·’ﬂ"
         Height          =   285
         Index           =   31
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒÂ"
         Height          =   285
         Index           =   32
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   720
         Width           =   1050
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·ÊﬂÌ·"
         Height          =   285
         Index           =   22
         Left            =   9720
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   720
         Width           =   1050
      End
   End
   Begin VB.TextBox txtmeterRentvalue 
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
      Left            =   9840
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   5040
      Width           =   1815
   End
   Begin VB.TextBox txtnoofoffices 
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
      Left            =   15600
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   5040
      Width           =   1785
   End
   Begin VB.TextBox txtnoofapartement 
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
      TabIndex        =   49
      Top             =   5040
      Width           =   1665
   End
   Begin VB.TextBox txtnoofparking 
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
      Left            =   6720
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   5040
      Width           =   2115
   End
   Begin VB.TextBox txttotallength 
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
      Left            =   3360
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   5040
      Width           =   1815
   End
   Begin VB.ComboBox cbointerfaceid 
      Height          =   315
      ItemData        =   "FrmAqarSearch.frx":000C
      Left            =   12600
      List            =   "FrmAqarSearch.frx":0016
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   5040
      Width           =   1815
   End
   Begin VB.TextBox txtfloorcount 
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
      Left            =   6720
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   4560
      Width           =   2115
   End
   Begin VB.TextBox txtEntryCount 
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
      Left            =   9840
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   4560
      Width           =   1815
   End
   Begin VB.ComboBox dcmaintenancetypeid 
      Height          =   315
      ItemData        =   "FrmAqarSearch.frx":0026
      Left            =   12600
      List            =   "FrmAqarSearch.frx":0030
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   4560
      Width           =   1815
   End
   Begin VB.TextBox txtlastrentvalue 
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
      TabIndex        =   35
      Top             =   4560
      Width           =   1665
   End
   Begin VB.TextBox txtcurrentPrice 
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
      Left            =   3360
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   4560
      Width           =   1815
   End
   Begin VB.TextBox txtaqarage 
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
      Left            =   15600
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   4560
      Width           =   1785
   End
   Begin VB.TextBox txtstreetname 
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
      Left            =   6720
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Tag             =   "«œŒ· «”„ «·‘«—⁄"
      Top             =   4080
      Width           =   2115
   End
   Begin VB.TextBox TxtAqarid 
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
      Left            =   15600
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   3600
      Width           =   1785
   End
   Begin VB.TextBox TxtN 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   12600
      TabIndex        =   15
      Top             =   3600
      Width           =   1815
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   8400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "»ÕÀ"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmAqarSearch.frx":0048
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtremark 
      Alignment       =   1  'Right Justify
      Height          =   1020
      Left            =   21960
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4680
      Width           =   7830
   End
   Begin VB.TextBox TxtAqarName 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3600
      Width           =   2115
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   19065
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
               Picture         =   "FrmAqarSearch.frx":0064
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAqarSearch.frx":03FE
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAqarSearch.frx":0798
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAqarSearch.frx":0B32
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAqarSearch.frx":0ECC
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAqarSearch.frx":1266
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAqarSearch.frx":1600
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAqarSearch.frx":1B9A
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ «·⁄ﬁ«—« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Index           =   2
         Left            =   13335
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   90
         Width           =   5280
      End
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   315
      Left            =   10080
      TabIndex        =   5
      Top             =   7320
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "6"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2745
      Left            =   0
      TabIndex        =   9
      Top             =   720
      Width           =   19035
      _cx             =   33576
      _cy             =   4842
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
      Rows            =   50
      Cols            =   30
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmAqarSearch.frx":1F34
      ScrollTrack     =   -1  'True
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   10
      Top             =   6120
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   661
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
      BackStyle       =   0
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   1380
      TabIndex        =   11
      Top             =   6120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "„”Õ"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   480
      TabIndex        =   12
      Top             =   6120
      Width           =   855
      _ExtentX        =   1508
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
      BackStyle       =   0
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton CmdItemSearch 
      Height          =   345
      Index           =   2
      Left            =   0
      TabIndex        =   13
      Top             =   6090
      Visible         =   0   'False
      Width           =   405
      _ExtentX        =   714
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
      ButtonImage     =   "FrmAqarSearch.frx":2418
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin MSDataListLib.DataCombo dcaqartypeid 
      Height          =   315
      Left            =   9840
      TabIndex        =   20
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· ‰Ê⁄ «·⁄ﬁ«—"
      Top             =   3600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboCountryID2 
      Height          =   315
      Left            =   15600
      TabIndex        =   22
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·œÊ·…"
      Top             =   4080
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboGovernmentID 
      Height          =   315
      Left            =   12600
      TabIndex        =   24
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·„œÌ‰…"
      Top             =   4080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboCityID 
      Height          =   315
      Left            =   9840
      TabIndex        =   26
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   4080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo dcschemeid 
      Height          =   315
      Left            =   240
      TabIndex        =   30
      Tag             =   " "
      Top             =   3960
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCAkarUnit 
      Height          =   315
      Left            =   3840
      TabIndex        =   64
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· ‰Ê⁄ «·⁄ﬁ«—"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo dcBranch 
      Height          =   315
      Left            =   240
      TabIndex        =   67
      Top             =   3600
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6165
      TabIndex        =   68
      Top             =   3600
      Width           =   450
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·„Œÿÿ"
      Height          =   285
      Index           =   15
      Left            =   5565
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   4080
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄ «·ÊÕœ…"
      Height          =   285
      Index           =   24
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   6240
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "ﬁÌ„… «ÌÃ«— „2"
      Height          =   285
      Index           =   17
      Left            =   11520
      RightToLeft     =   -1  'True
      TabIndex        =   54
      Top             =   5040
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "⁄œœ «·ÊÕœ«  «· Ã«—ÌÂ"
      Height          =   285
      Index           =   10
      Left            =   17520
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   5040
      Width           =   1410
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "⁄œœ «·ÊÕœ«  «·”ﬂ‰Ì…"
      Height          =   285
      Index           =   9
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   5040
      Width           =   1410
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "⁄œœ «·„Ê«ﬁ›"
      Height          =   285
      Index           =   34
      Left            =   8760
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   5040
      Width           =   1050
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "«·„”«ÕÂ «·«Ã„«·ÌÂ"
      Height          =   255
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«·Ê«ÃÂ…"
      Height          =   285
      Index           =   19
      Left            =   14160
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   5040
      Width           =   1410
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "⁄œœ «·ÿÊ«»ﬁ"
      Height          =   285
      Index           =   21
      Left            =   8760
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   4560
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "⁄œœ «·„œ«Œ·"
      Height          =   285
      Index           =   8
      Left            =   11520
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   4560
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄ «·’Ì«‰…"
      Height          =   285
      Index           =   35
      Left            =   14520
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   4560
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«Œ— ﬁÌ„Â «ÌÃ«—ÌÂ"
      Height          =   285
      Index           =   12
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   4560
      Width           =   1410
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "”⁄— «·⁄ﬁ«— «·Õ«·Ì"
      Height          =   285
      Index           =   11
      Left            =   5325
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   4560
      Width           =   1290
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«·⁄„— «·“„‰Ì"
      Height          =   285
      Index           =   37
      Left            =   17640
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   4560
      Width           =   1290
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·‘«—⁄"
      Height          =   285
      Index           =   6
      Left            =   8760
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   4080
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·ÕÌ"
      Height          =   285
      Index           =   5
      Left            =   11520
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   4080
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·„œÌ‰Â"
      Height          =   285
      Index           =   4
      Left            =   14520
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   4080
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·œÊ·Â"
      Height          =   285
      Index           =   0
      Left            =   17040
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   4080
      Width           =   1890
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄ «·⁄ﬁ«—"
      Height          =   285
      Index           =   16
      Left            =   11520
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   3600
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "„”·”· «·⁄ﬁ«—"
      Height          =   195
      Index           =   3
      Left            =   17940
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   3600
      Width           =   990
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·⁄ﬁ«—"
      Height          =   375
      Left            =   7800
      TabIndex        =   17
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·⁄ﬁ«—"
      Height          =   375
      Left            =   14355
      TabIndex        =   16
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label lblitemid 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "„·ÕÊŸ…"
      Height          =   375
      Left            =   21480
      TabIndex        =   6
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "«·»·œ"
      Height          =   375
      Left            =   8640
      TabIndex        =   3
      Top             =   7320
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·⁄ﬁ«—"
      Height          =   375
      Left            =   8640
      TabIndex        =   2
      Top             =   3600
      Width           =   1095
   End
End
Attribute VB_Name = "FrmAqarSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As ADODB.Recordset
Dim cSearchDcbo As clsDCboSearch

Private m_DcboItems As DataCombo

Public m_RetrunType As Integer
Public WithEvents FG1 As VSFlex8UCtl.VSFlexGrid
Attribute FG1.VB_VarHelpID = -1

Public WithEvents NewGrid As VSFlex8UCtl.VSFlexGrid
Attribute NewGrid.VB_VarHelpID = -1
'Public NewGrid As New ClsGrid
 
Public LngRow As Long

Public LngCol As Long

Public mIndex  As Long


Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

           If rs.State = adStateOpen Then
                rs.Close
            End If

            rs.Open Build_Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
            If SystemOptions.UserInterface = ArabicInterface Then
                '   LblRes.Caption = "‰ ÌÃ… «·»ÕÀ = " & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                '   LblRes.Caption = "Search Result=" & rs.RecordCount
            End If
    
            If rs.RecordCount < 1 Then
                FG.Clear flexClearScrollable, flexClearEverything
                FG.Rows = 2

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    Msg = "NO Search Results Found...!!!"
                    MsgBox Msg, vbOKOnly + vbExclamation, App.title
                End If

                Exit Sub
            End If

            Retrive
            FG.SetFocus

        Case 1
            clear_all Me
            FG.Clear flexClearScrollable, flexClearEverything
            dpsuckdate.value = ""

        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

End Sub







Private Sub DcboCountryID2_Change()
LoadDataCombos True, False, True
End Sub

Private Sub DcboCountryID2_Click(Area As Integer)
  'DcboCountryID2_Change

End Sub

Private Sub fg_Click()
   ' On Error GoTo ErrTrap
     If m_RetrunType = 0 Then
     RSAkar.FindRec2 val(FG.TextMatrix(FG.Row, FG.ColIndex("Aqarid")))
ElseIf m_RetrunType = 9 Then
     RSAkar.FindRec val(FG.TextMatrix(FG.Row, FG.ColIndex("NoteSerial1")))
  ElseIf m_RetrunType = 1 Then
  RSContract.DcbIqara.BoundText = val(FG.TextMatrix(FG.Row, FG.ColIndex("Aqarid")))
  ElseIf m_RetrunType = 2 Then
  RsExpenses.DcbIqara.BoundText = val(FG.TextMatrix(FG.Row, FG.ColIndex("Aqarid")))
  ElseIf m_RetrunType = 3 Then
  RsExpenses.Fg_Journal.TextMatrix(RsExpenses.LongRow, RsExpenses.Fg_Journal.ColIndex("iqar")) = FG.TextMatrix(FG.Row, FG.ColIndex("aqarname"))
    RsExpenses.Fg_Journal.TextMatrix(RsExpenses.LongRow, RsExpenses.Fg_Journal.ColIndex("iqarid")) = val(FG.TextMatrix(FG.Row, FG.ColIndex("Aqarid")))
    RsExpenses.DcbIqara2.BoundText = val(FG.TextMatrix(FG.Row, FG.ColIndex("Aqarid")))
           If SystemOptions.NoCreatJLInRentContract Then
                   RsExpenses.FillData RsExpenses.LongRow
                End If


' RsExpenses.Fg_Journal_AfterEdit RsExpenses.LongRow, RsExpenses.Fg_Journal.ColIndex("iqar")
  
'  RsExpenses.DcbIqara.BoundText = val(Fg.TextMatrix(Fg.Row, Fg.ColIndex("Aqarid")))
  
   ElseIf m_RetrunType = 2020 Then
  FrmOrderMaintenance.DcbIqara.BoundText = val(FG.TextMatrix(FG.Row, FG.ColIndex("Aqarid")))
  End If
  
End Sub

Private Sub Retrive()
    Dim Num As Integer
    On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        FG.Rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With FG
                .TextMatrix(Num, .ColIndex("Aqarid")) = IIf(IsNull(rs("Aqarid").value), "", rs("Aqarid").value)
                .TextMatrix(Num, .ColIndex("aqarNo")) = IIf(IsNull(rs("aqarNo").value), "", rs("aqarNo").value)
                If m_RetrunType = 9 Then
                .TextMatrix(Num, .ColIndex("NoteSerial1")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
                End If
                   .TextMatrix(Num, .ColIndex("UnitName")) = IIf(IsNull(rs("UnitName").value), "", Trim(rs("UnitName").value))
                .TextMatrix(Num, .ColIndex("aqarname")) = IIf(IsNull(rs("aqarname").value), "", Trim(rs("aqarname").value))
                   .TextMatrix(Num, .ColIndex("CountryName")) = IIf(IsNull(rs("CountryName").value), "", rs("CountryName").value)
                .TextMatrix(Num, .ColIndex("GovernmentName")) = IIf(IsNull(rs("GovernmentName").value), "", rs("GovernmentName").value)
.TextMatrix(Num, .ColIndex("CityName")) = IIf(IsNull(rs("CityName").value), "", rs("CityName").value)
                .TextMatrix(Num, .ColIndex("streetname")) = IIf(IsNull(rs("streetname").value), "", rs("streetname").value)
              .TextMatrix(Num, .ColIndex("SchemeName")) = IIf(IsNull(rs("SchemeName").value), "", rs("SchemeName").value)
                .TextMatrix(Num, .ColIndex("aqarage")) = IIf(IsNull(rs("aqarage").value), "", rs("aqarage").value)
                dcmaintenancetypeid.ListIndex = val(IIf(IsNull(rs("maintenancetypeid").value), -1, rs("maintenancetypeid").value))
                 .TextMatrix(Num, .ColIndex("maintenancetypeid")) = dcmaintenancetypeid.Text
                .TextMatrix(Num, .ColIndex("EntryCount")) = IIf(IsNull(rs("EntryCount").value), "", rs("EntryCount").value)
                
              .TextMatrix(Num, .ColIndex("lastrentvalue")) = IIf(IsNull(rs("lastrentvalue").value), "", rs("lastrentvalue").value)
                .TextMatrix(Num, .ColIndex("noofapartement")) = IIf(IsNull(rs("noofapartement").value), "", rs("noofapartement").value)
                cbointerfaceid.ListIndex = val(IIf(IsNull(rs("interfaceid").value), -1, rs("interfaceid").value))
                .TextMatrix(Num, .ColIndex("interfaceid")) = cbointerfaceid.Text
                .TextMatrix(Num, .ColIndex("meterRentvalue")) = IIf(IsNull(rs("meterRentvalue").value), "", rs("meterRentvalue").value)
               .TextMatrix(Num, .ColIndex("noofparking")) = IIf(IsNull(rs("noofparking").value), "", rs("noofparking").value)
                .TextMatrix(Num, .ColIndex("totallength")) = IIf(IsNull(rs("totallength").value), "", rs("totallength").value)
                
                  .TextMatrix(Num, .ColIndex("floorcount")) = IIf(IsNull(rs("floorcount").value), "", rs("floorcount").value)
                .TextMatrix(Num, .ColIndex("currentPrice")) = IIf(IsNull(rs("currentPrice").value), "", rs("currentPrice").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(Num, .ColIndex("name")) = IIf(IsNull(rs("name").value), "", Trim(rs("name").value))
                    .TextMatrix(Num, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", Trim(rs("branch_name").value))
                Else
                 .TextMatrix(Num, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", Trim(rs("branch_namee").value))
                    .TextMatrix(Num, .ColIndex("name")) = IIf(IsNull(rs("namee").value), "", Trim(rs("namee").value))
                End If
.TextMatrix(Num, .ColIndex("noofoffices")) = IIf(IsNull(rs("noofoffices").value), "", rs("noofoffices").value)
                .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                .TextMatrix(Num, .ColIndex("suckno")) = IIf(IsNull(rs("suckno").value), "", rs("suckno").value)
                 If Not (IsNull(rs("suckdate").value)) Then
                   .TextMatrix(Num, .ColIndex("suckdate")) = Format(rs("suckdate").value, "yyyy/M/d")
                End If
             '   .TextMatrix(Num, .ColIndex("suckdate")) = IIf(IsNull(rs("suckdate").value), "", rs("suckdate").value)
                .TextMatrix(Num, .ColIndex("authorizationname")) = IIf(IsNull(rs("authorizationname").value), "", rs("authorizationname").value)
               ' .TextMatrix(Num, .ColIndex("aqarNo")) = IIf(IsNull(rs("aqarNo").value), "", rs("aqarNo").value)
                
              ' If SystemOptions.UserInterface = ArabicInterface Then
              '      .TextMatrix(Num, .ColIndex("itemname")) = IIf(IsNull(rs("itemname").value), "", Trim(rs("itemname").value))
              '  Else
              '      .TextMatrix(Num, .ColIndex("itemname")) = IIf(IsNull(rs("itemnamee").value), "", Trim(rs("itemnamee").value))
              '  End If
              '
'
'                .TextMatrix(Num, .ColIndex("Transaction_Date")) = IIf(IsNull(rs("Transaction_Date").value), "", Trim(rs("Transaction_Date").value))
                '   .TextMatrix(Num, .ColIndex("currency_code")) = IIf(IsNull(rs("Transaction_Date").value), "", Trim(rs("currency_code").value))
           
                '  .TextMatrix(Num, .ColIndex("countryid")) = IIf(IsNull(rs("countryid").value), "", (rs("countryid").value))
                '    .TextMatrix(Num, .ColIndex("CountryName")) = IIf(IsNull(rs("CountryName").value), "", Trim(rs("CountryName").value))
            
            End With

            rs.MoveNext
        Next Num

        ' Fg.AutoSize 0, Fg.Cols - 1, False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_DblClick()
    fg_Click
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim BG As New ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
'''''''''''''''''''''
    Dim i As Integer
    Dim My_SQL As String
  '  Dim Dcombos As ClsDataCombos
    Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500
dpsuckdate.value = ""


'   Dim cOptions As ClsCompanyInfo
'   Set cOptions = New ClsCompanyInfo
 '   If SystemOptions.UserInterface = ArabicInterface Then
 '   lblCompanyname.Caption = cOptions.ArabCompanyName & Chr(13) & CurrentBranchName
'Else
'lblCompanyname.Caption = cOptions.EngCompanyName & Chr(13) & CurrentBranchNameE
'End If


   ' My_SQL = "tblaqar"
   ' Set BKGrndPic = New ClsBackGroundPic
   ' Set RsSavRec = New ADODB.Recordset
   ' RsSavRec.CursorLocation = adUseClient
   ' RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    'Me.TxtModFlg.text = "R"
  '  Resize_Form Me
    'load tblUsers -----------------------------------------------
   ' My_SQL = "select UserID,UserName From tblUsers "
    'fill_combo DCUser, My_SQL
    
    Set Dcombos = New ClsDataCombos
     Dcombos.GetBranches Me.Dcbranch
     Dcombos.get«hay Me.DcboGovernmentID
     Dcombos.GetCustomersSuppliers 57, Me.dcsupplier
     Dcombos.getSchemes Me.dcschemeid
     Dcombos.getAkarType Me.dcaqartypeid
     DcboCountryID2.BoundText = 1
     Dcombos.getAkarUnit Me.DCAkarUnit

'

 

'    ModFgLib.LinkFgColWithDataCombo Grid, Grid.ColIndex("GovernmentID"), Me.DcboGovernmentID

'    FillGridWithData

'    With Me.Grid
'        .Cell(flexcpPicture, 0, .ColIndex("CityName")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
'        .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon
'
'        For i = 0 To .Cols - 1
'            .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
'        Next
   
'        .ExtendLastCol = True
'        .WallPaper = BKGrndPic.Picture
'        .RowHeight(-1) = 300
'    End With

   ' BtnFirst_Click
  '  ShowTip
'LoadDataCombos




'With UnitsGrid
'        If SystemOptions.UserInterface = ArabicInterface Then
'            .Cell(flexcpPictureAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexPicAlignRightCenter
'            .ColComboList(.ColIndex("rentType")) = "#1;«Ã„«·Ì «·ÊÕœ…|#2;»«·„ —"
'        Else
'            .ColComboList(.ColIndex("rentType")) = "#1;Totals|#2;By Meter"
'        End If
        
'End With
''''''''''''''''''''

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
 
    'Dim My_SQL As String
   ' Set Dcombos = New ClsDataCombos
   ' Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
   ' Dcombos.GetItemsNames DCboItem, , , , True
   '
   ' My_SQL = " select CountryID,CountryName from TblCountriesData"
 '
 '   fill_combo Me.DataCombo4, My_SQL
 '   RetrunType = -1
   CenterForm Me

    FormPostion Me, GetPostion
    FG.WallPaper = BG.SearchWallpaper
    Set rs = New ADODB.Recordset
    DBCboClientName.BoundText = ""
    Exit Sub
ErrTrap:

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If rs.State = adStateOpen Then
        rs.Close
        Set rs = Nothing
    End If

    Set cSearchDcbo = Nothing

    FormPostion Me, SavePostion
    Set m_DcboItems = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Function Build_Sql()
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String
    Dim BolHaveSerial As Boolean
    Dim IntHaveSerial As Integer
 
    On Error GoTo ErrTrap

 StrSQL = "SELECT     dbo.TblAqar.Aqarid, dbo.TblAqar.aqarNo, dbo.TblAqar.aqartypeid, dbo.tblAkarType.name, dbo.tblAkarType.namee, dbo.TblAqar.CountryID, "
        StrSQL = StrSQL & "                dbo.TblCountriesData.CountryName, dbo.TblAqar.cityid, dbo.TblCountriesGovernments.GovernmentName, dbo.TblAqar.heyid,"
       StrSQL = StrSQL & "                 dbo.TblCountriesGovernmentsCities.CityName, dbo.TblAqar.streetname, dbo.TblAqar.schemeid, dbo.tblSchemes.name AS SchemeName,"
      StrSQL = StrSQL & "                  dbo.tblSchemes.namee AS SchemeNameE, dbo.TblAqar.aqarage, dbo.TblAqar.currentPrice, dbo.TblAqar.lastrentvalue, dbo.TblAqar.maintenancetypeid,"
        StrSQL = StrSQL & "                dbo.TblAqar.StatusId, dbo.TblAqar.EntryCount, dbo.TblAqar.floorcount, dbo.TblAqar.noofoffices, dbo.TblAqar.noofparking, dbo.TblAqar.interfaceid,"
      StrSQL = StrSQL & "                  dbo.TblAqar.noofapartement, dbo.TblAqar.totallength, dbo.TblAqar.meterRentvalue, dbo.TblAqar.Rate, dbo.TblAqar.Price, dbo.TblAqar.ownerid,"
     StrSQL = StrSQL & "                   dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblAqar.Location, dbo.TblAqar.aqarname, dbo.TblAqar.northlength, dbo.TblAqar.eastlength,"
    StrSQL = StrSQL & "                    dbo.TblAqar.Southlength, dbo.TblAqar.Westlength, dbo.TblAqar.metersalevalue, dbo.TblAqar.GoogleMap, dbo.TblAqar.suckno, dbo.TblAqar.authorizationname,"
      StrSQL = StrSQL & "                  dbo.TblAqar.suckdateH, dbo.TblAqar.suckdate, dbo.TblAqar.statusdate, dbo.TblAqar.PriceHadW, dbo.TblAqar.PriceSomW, dbo.TblAqar.StreetNo, dbo.TblAqar.Part,"
    StrSQL = StrSQL & "                    dbo.TblAqar.UnitNo, dbo.TblAkarUnit.name AS UnitName, dbo.TblAkarUnit.namee AS UnitNamee, dbo.TblAqar.Block, dbo.TblAqar.PriceSom, dbo.TblAqar.PriceHad,"
    StrSQL = StrSQL & "                    dbo.TblAqar.BranchId , dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_nameE"
    If m_RetrunType = 9 Then
        StrSQL = StrSQL & "                    ,TblContractInstallDisco.id NoteSerial1"
    End If
  StrSQL = StrSQL & " FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
   StrSQL = StrSQL & "                     dbo.TblAqar ON dbo.TblBranchesData.branch_id = dbo.TblAqar.BranchId LEFT OUTER JOIN"
     StrSQL = StrSQL & "                   dbo.TblAkarUnit ON dbo.TblAqar.UnitNo = dbo.TblAkarUnit.id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblCustemers ON dbo.TblAqar.ownerid = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.tblSchemes ON dbo.TblAqar.schemeid = dbo.tblSchemes.id LEFT OUTER JOIN"
   StrSQL = StrSQL & "                     dbo.TblCountriesGovernments INNER JOIN"
   StrSQL = StrSQL & "                     dbo.TblCountriesGovernmentsCities ON dbo.TblCountriesGovernments.GovernmentID = dbo.TblCountriesGovernmentsCities.GovernmentID INNER JOIN"
  StrSQL = StrSQL & "                      dbo.TblCountriesData ON dbo.TblCountriesGovernments.CountryID = dbo.TblCountriesData.CountryID ON"
  StrSQL = StrSQL & "                      dbo.TblAqar.heyid = dbo.TblCountriesGovernmentsCities.CityID AND dbo.TblAqar.cityid = dbo.TblCountriesGovernments.GovernmentID AND"
  StrSQL = StrSQL & "                      dbo.TblAqar.CountryID = dbo.TblCountriesData.CountryID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                      dbo.tblAkarType ON dbo.TblAqar.aqartypeid = dbo.tblAkarType.id"
     If m_RetrunType = 9 Then
      StrSQL = StrSQL & "                      LEFT OUTER JOIN dbo.TblContractInstallDisco"
    StrSQL = StrSQL & "                      ON  dbo.TblAqar.Aqarid = dbo.TblContractInstallDisco.Iqar"
    End If
    StrSQL = StrSQL & " where 1=1 "
    If m_RetrunType = 9 Then
        StrSQL = StrSQL & " and  IsNull(TblContractInstallDisco.id,0 ) <> 0"
    End If
      If Me.txtsuckno.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.suckno ='" & Me.txtsuckno.Text & "'"
 
    End If
    
     If Me.txtauthorizationname.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.authorizationname ='" & Me.txtauthorizationname.Text & "'"
 
    End If
    
     If Me.txtnoofapartement.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.noofapartement ='" & Me.txtnoofapartement.Text & "'"
 
    End If
    
      If Me.txttotallength.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.totallength ='" & Me.txttotallength.Text & "'"
 
    End If
    
    
     If Me.txtnoofparking.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.noofparking ='" & Me.txtnoofparking.Text & "'"
 
    End If
    
    
    If Me.txtmeterRentvalue.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.meterRentvalue ='" & Me.txtmeterRentvalue.Text & "'"
 
    End If
    
   If Me.txtnoofoffices.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.noofoffices ='" & Me.txtnoofoffices.Text & "'"
 
    End If
    
     If Not IsNull(Me.dpsuckdate.value) Then
        
            StrWhere = StrWhere & " AND dbo.TblAqar.suckdate >=" & SQLDate(Me.dpsuckdate.value, True) & ""
        End If
    
    
    
     If Me.txtlastrentvalue.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.lastrentvalue ='" & Me.txtlastrentvalue.Text & "'"
 
    End If
    
        If Me.txtcurrentPrice.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.currentPrice ='" & Me.txtcurrentPrice.Text & "'"
 
    End If
    
       If Me.txtEntryCount.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.EntryCount ='" & Me.txtEntryCount.Text & "'"
 
    End If
    
       If Me.txtfloorcount.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.floorcount ='" & Me.txtfloorcount.Text & "'"
 
    End If
    
      If Me.txtstreetname.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.streetname  like '%" & Me.txtstreetname.Text & "%'"
 
    End If
    
    If Me.TxtAqarName.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.aqarname  like '%" & Me.TxtAqarName.Text & "%'"
 
    End If
      If Me.TxtAqarid.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.Aqarid ='" & Me.TxtAqarid.Text & "'"
 
    End If
       If Me.txtaqarage.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.aqarage ='" & Me.txtaqarage.Text & "'"
 
    End If
    
    If Me.TxtN.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.aqarNo ='" & Me.TxtN.Text & "'"
 
    End If
     If val(Me.Dcbranch.BoundText) <> 0 And Me.Dcbranch.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.BranchId ='" & val(Me.Dcbranch.BoundText) & "'"
 
    End If
    
 If val(Me.DCAkarUnit.BoundText) <> 0 And Me.DCAkarUnit.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.UnitNo ='" & val(Me.DCAkarUnit.BoundText) & "'"
 
    End If
    
    If val(Me.dcaqartypeid.BoundText) <> 0 And Me.dcaqartypeid.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.aqartypeid ='" & val(Me.dcaqartypeid.BoundText) & "'"
 
    End If
        
    If val(Me.DcboCountryID2.BoundText) <> 0 And Me.DcboCountryID2.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.CountryID ='" & val(Me.DcboCountryID2.BoundText) & "'"
 
    End If
 If val(Me.DcboGovernmentID.BoundText) <> 0 And Me.DcboGovernmentID.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.cityid ='" & val(Me.DcboGovernmentID.BoundText) & "'"
 
    End If
    
 If val(Me.DcboCityID.BoundText) <> 0 And Me.DcboCityID.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.heyid ='" & val(Me.DcboCityID.BoundText) & "'"
 
    End If
    If val(Me.dcschemeid.BoundText) <> 0 And Me.dcschemeid.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.schemeid ='" & val(Me.dcschemeid.BoundText) & "'"
 
    End If
       If val(Me.dcsupplier.BoundText) <> 0 And Me.dcsupplier.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.ownerid ='" & val(Me.dcsupplier.BoundText) & "'"
 
    End If
    
  If Me.dcmaintenancetypeid.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.maintenancetypeid ='" & val(Me.dcmaintenancetypeid.ListIndex) & "'"
 
    End If
    
    If Me.cbointerfaceid.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqar.interfaceid ='" & val(Me.cbointerfaceid.ListIndex) & "'"
 
    End If
    
    
    
    
    StrWhere = StrWhere + " order by dbo.TblAqar.Aqarid"
Dim s As String
s = StrSQL + StrWhere
    Build_Sql = StrSQL + StrWhere
    Exit Function
ErrTrap:
End Function
Private Sub LoadDataCombos(Optional BolExceptCountries As Boolean = False, _
                           Optional BolExceptGovern As Boolean = False, _
                           Optional BolExceptCities As Boolean = False)
    Dim Dcombo As New ClsDataCombos
    Dcombo.GetCountriesNames Me.DcboCountryID2
  
 '   If BolExceptCountries = False Then
 '       Dcombo.GetCountriesNames Me.DcboCountryID2
 '       Set cSearch(0) = New clsDCboSearch
 '       Set cSearch(0).Client = Me.DcboCountryID
 '   End If

    If BolExceptGovern = False Then
        Dcombo.getCountriesGovernments Me.DcboGovernmentID, val(Me.DcboCountryID2.BoundText)
        'Set cSearch(1) = New clsDCboSearch
        'Set cSearch(1).Client = Me.DcboGovernmentID
    End If

    If BolExceptCities = False Then
        Dcombo.GetCountriesGovernCities Me.DcboCityID, val(Me.DcboCountryID2.BoundText), val(Me.DcboGovernmentID.BoundText)
'        Set cSearch(2) = New clsDCboSearch
'        Set cSearch(2).Client = Me.DcboAqarid
    End If

     
End Sub
Private Sub DcboGovernmentID_Change()
    LoadDataCombos False, True, False
End Sub

Private Sub DcboGovernmentID_Click(Area As Integer)
    DcboGovernmentID_Change
End Sub

'Private Sub DcboAqarid_Change()
'    LoadDataCombos False, False, True
'End Sub

'Private Sub DcboAqarid_Click(Area As Integer)
'    DcboAqarid_Change

'End Sub
Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl Is FG Then
            If Not FG.TextMatrix(FG.Row, 1) = "" Then
                fg_Click
                Unload Me
            End If

        Else
            Cmd_Click (0)
        End If
    End If

    On Error GoTo ErrTrap

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Cmd_Click (2)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub






Private Sub ChangeLang()
    Me.Caption = "Search For Production Orders"
    Label1(2).Caption = Me.Caption
    Label2.Caption = "Order No"
 
    Label3.Caption = "Date"
    Label5.Caption = "Country"
    Label4.Caption = "Vendor"
    Label6.Caption = "Remark"

    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"

    'OptType(0).Caption = "Start of the name"
    'OptType(1).Caption = "any part of the name"
    With Me.FG
        .TextMatrix(0, .ColIndex("order_no")) = "order no"
        '  .TextMatrix(0, .ColIndex("remark")) = "remark  "
        .TextMatrix(0, .ColIndex("CusName")) = "Customer Name"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = " Date"
        '     .TextMatrix(0, .ColIndex("CountryName")) = "Country Name"
  
        '  .AutoSize 0, .Cols - 1, False
    End With

End Sub




