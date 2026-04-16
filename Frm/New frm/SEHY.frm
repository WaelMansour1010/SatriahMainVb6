VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form SEHY 
   Caption         =   "«·„·› «·’ÕÌ"
   ClientHeight    =   7230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11265
   Icon            =   "SEHY.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   11265
   StartUpPosition =   2  'CenterScreen
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      ToolTipText     =   "«ﬂ » —ﬁ„ «·„Ê÷Ê⁄ À„ «÷€ÿ enter"
      Top             =   840
      Width           =   5415
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7440
      TabIndex        =   19
      ToolTipText     =   "«ﬂ » —ﬁ„ «·„Ê÷Ê⁄ À„ «÷€ÿ enter"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   480
      TabIndex        =   6
      Top             =   -120
      Width           =   10455
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7440
         TabIndex        =   9
         ToolTipText     =   "«ﬂ » —ﬁ„ «·„Ê÷Ê⁄ À„ «÷€ÿ enter"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4080
         TabIndex        =   8
         ToolTipText     =   "«ﬂ » —ﬁ„ «·„Ê÷Ê⁄ À„ «÷€ÿ enter"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   7
         ToolTipText     =   "«ﬂ » —ﬁ„ «·„Ê÷Ê⁄ À„ «÷€ÿ enter"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·„ÊŸ›"
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   0
         Left            =   9360
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·„ÊŸ›"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6000
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁ”„"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CheckBox Check2 
      Height          =   195
      Left            =   8760
      TabIndex        =   5
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Left            =   8760
      TabIndex        =   4
      Top             =   960
      Width           =   255
   End
   Begin VB.TextBox Text1 
      DataField       =   "subject_no"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   7800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   960
      Top             =   8160
      Visible         =   0   'False
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   6720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "Õ–› ”ÿ—"
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
      BCOL            =   255
      BCOLO           =   255
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "SEHY.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   2835
      Left            =   480
      TabIndex        =   13
      Top             =   3720
      Width           =   10545
      _cx             =   18600
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
      Rows            =   50
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"SEHY.frx":0028
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
   Begin MSDataListLib.DataCombo DcboGovernmentID 
      Height          =   315
      Left            =   6720
      TabIndex        =   20
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   1680
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker DtDate 
      Height          =   315
      Left            =   7680
      TabIndex        =   21
      Top             =   2520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   100073473
      CurrentDate     =   38784
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   4920
      TabIndex        =   23
      Top             =   2520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   100073473
      CurrentDate     =   38784
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "‰Â«Ì… «· √„Ì‰"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   6600
      TabIndex        =   24
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "»œ«Ì… «· √„Ì‰"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   9360
      TabIndex        =   22
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "«·—ﬁ„ «· √„Ì‰Ì"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   9360
      TabIndex        =   18
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "‘—ÌÕÂ «· ‹«„Ì‰"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   9360
      TabIndex        =   17
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "·œÌ…    √„Ì‰ ÿ»Ì"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   9000
      TabIndex        =   16
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Ê’› «·Õ«·Â"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   6360
      TabIndex        =   15
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "«·„Êﬁ› «·’ÕÌ"
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   9240
      TabIndex        =   14
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   0
      X2              =   11280
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "·œÌ… „‘«ﬂ· ’ÕÌ…"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   9000
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label CASE_ID 
      Caption         =   "0"
      Height          =   495
      Left            =   -600
      TabIndex        =   1
      Top             =   7440
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "SEHY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

End Sub
