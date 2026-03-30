VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form TAKEEM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«· ﬁÌÌ„ «·‘Œ’Ì"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10560
   Icon            =   "TAKEEM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   5910
   ScaleWidth      =   10560
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
      Height          =   495
      Index           =   2
      Left            =   6480
      TabIndex        =   20
      ToolTipText     =   "«ﬂ » —ﬁ„ «·„Ê÷Ê⁄ À„ «÷€ÿ enter"
      Top             =   5160
      Width           =   1815
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
      Height          =   495
      Index           =   1
      Left            =   6480
      TabIndex        =   19
      ToolTipText     =   "«ﬂ » —ﬁ„ «·„Ê÷Ê⁄ À„ «÷€ÿ enter"
      Top             =   4560
      Width           =   1815
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
      Height          =   495
      Index           =   0
      Left            =   6480
      TabIndex        =   18
      ToolTipText     =   "«ﬂ » —ﬁ„ «·„Ê÷Ê⁄ À„ «÷€ÿ enter"
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   10455
      Begin VB.TextBox Text9 
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
         TabIndex        =   14
         ToolTipText     =   "«ﬂ » —ﬁ„ «·„Ê÷Ê⁄ À„ «÷€ÿ enter"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text8 
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
         TabIndex        =   13
         ToolTipText     =   "«ﬂ » —ﬁ„ «·„Ê÷Ê⁄ À„ «÷€ÿ enter"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text7 
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
         TabIndex        =   12
         ToolTipText     =   "«ﬂ » —ﬁ„ «·„Ê÷Ê⁄ À„ «÷€ÿ enter"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁ”„"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2520
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·„ÊŸ›"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6000
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·„ÊŸ›"
         ForeColor       =   &H00000000&
         Height          =   495
         Index           =   0
         Left            =   9360
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "«·‰ﬁ«ÿ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2055
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   3375
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Left            =   120
         TabIndex        =   8
         Text            =   "10"
         ToolTipText     =   "«ﬂ » —ﬁ„ «·„Ê÷Ê⁄ À„ «÷€ÿ enter"
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Left            =   120
         TabIndex        =   6
         Text            =   "5"
         ToolTipText     =   "«ﬂ » —ﬁ„ «·„Ê÷Ê⁄ À„ «÷€ÿ enter"
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
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
         Left            =   120
         TabIndex        =   4
         Text            =   "0"
         ToolTipText     =   "«ﬂ » —ﬁ„ «·„Ê÷Ê⁄ À„ «÷€ÿ enter"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "«·Õœ «·«ﬁ’Ï"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "«·Õœ «·«œ‰Ì"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2280
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "„ Ê”ÿ «·‰ﬁ«ÿ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.TextBox Text1 
      DataField       =   "subject_no"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   9480
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
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Õ–›  ”ÿ—"
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
      MICON           =   "TAKEEM.frx":000C
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
      Left            =   0
      TabIndex        =   10
      Top             =   960
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
      FormatString    =   $"TAKEEM.frx":0028
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
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "«· ﬁœÌ— «·⁄«„"
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   3
      Left            =   8880
      TabIndex        =   23
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "«·‰ ÌÃ…"
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   2
      Left            =   8760
      TabIndex        =   22
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "„Ã„Ê⁄ «·‰ﬁ«ÿ"
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   8760
      TabIndex        =   21
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label CASE_ID 
      Caption         =   "0"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   8040
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "TAKEEM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

End Sub
