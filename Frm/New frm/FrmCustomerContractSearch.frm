VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmCustomerContractSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·»ÕÀ ⁄‰ « ›«ﬁÌ«  «·⁄„·«¡"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13470
   Icon            =   "FrmCustomerContractSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   13470
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Dcbyear 
      Height          =   315
      ItemData        =   "FrmCustomerContractSearch.frx":6852
      Left            =   15360
      List            =   "FrmCustomerContractSearch.frx":6854
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox DcbMoth 
      Height          =   315
      ItemData        =   "FrmCustomerContractSearch.frx":6856
      Left            =   15360
      List            =   "FrmCustomerContractSearch.frx":6858
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   40
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   0
      Width           =   13665
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·»ÕÀ ⁄‰ « ›«ﬁÌ«  «·⁄„·«¡"
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
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   240
         Width           =   4680
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   12360
         Picture         =   "FrmCustomerContractSearch.frx":685A
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Height          =   3015
      Left            =   0
      TabIndex        =   36
      Top             =   720
      Width           =   13455
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   2625
         Left            =   150
         TabIndex        =   37
         Top             =   120
         Width           =   13155
         _cx             =   23204
         _cy             =   4630
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmCustomerContractSearch.frx":8813
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
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Height          =   615
      Left            =   0
      TabIndex        =   33
      Top             =   6480
      Width           =   13455
      Begin VB.Label lblL 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   0
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   240
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.Label lblL 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   10
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   120
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   0
      TabIndex        =   29
      Top             =   7080
      Width           =   13455
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   0
         Left            =   9480
         TabIndex        =   30
         Top             =   240
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   661
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
         ButtonImage     =   "FrmCustomerContractSearch.frx":8957
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         ColorToggledHoverText=   16711680
         LowerToggledContent=   0   'False
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   1
         Left            =   5160
         TabIndex        =   31
         Top             =   240
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   661
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
         ButtonImage     =   "FrmCustomerContractSearch.frx":F1B9
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
      Begin ImpulseButton.ISButton Cmd 
         Cancel          =   -1  'True
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   32
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
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
         ButtonImage     =   "FrmCustomerContractSearch.frx":15A1B
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
   End
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   3720
      Width           =   6675
      Begin VB.TextBox TxtIDTO 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·« ›«ﬁÌ…"
         Height          =   195
         Index           =   14
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   315
         Index           =   6
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   315
         Index           =   5
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame lblLW 
      BackColor       =   &H00E2E9E9&
      Height          =   2055
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   4440
      Width           =   13425
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2E9E9&
         Height          =   1095
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   6615
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   330
            Left            =   3120
            TabIndex        =   15
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   214171651
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   330
            Left            =   240
            TabIndex        =   16
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   214171651
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   330
            Left            =   3120
            TabIndex        =   17
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   214171651
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   330
            Left            =   240
            TabIndex        =   18
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   214171651
            CurrentDate     =   38887
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   0
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   360
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   2
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   600
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " ‰ ÂÌ „œ Â« „‰ "
            Height          =   195
            Index           =   8
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   600
            Width           =   1380
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " »œ√ „œ Â« „‰"
            Height          =   195
            Index           =   9
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   240
            Width           =   1785
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   1095
         Left            =   6840
         TabIndex        =   9
         Top             =   120
         Width           =   6495
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3360
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   240
            Width           =   1575
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   360
            TabIndex        =   11
            Top             =   600
            Width           =   4605
            _ExtentX        =   8123
            _ExtentY        =   556
            _Version        =   393216
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
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì·"
            Height          =   285
            Index           =   7
            Left            =   5040
            TabIndex        =   13
            Top             =   600
            Width           =   1365
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·⁄„Ì·"
            Height          =   285
            Index           =   4
            Left            =   5040
            TabIndex        =   12
            Top             =   240
            Width           =   1365
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   13215
         Begin VB.TextBox txtRemarks 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2520
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   240
            Width           =   7800
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   195
            Index           =   1
            Left            =   10800
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   240
            Width           =   840
         End
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   3720
      Width           =   6735
      Begin VB.OptionButton check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Ì„ﬂ‰ «· ⁄«„· „⁄Â"
         Height          =   315
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton check2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„„‰Ê⁄ „‰ «· ⁄«„·"
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬂ·"
         Height          =   315
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õ«·… «· ⁄«„·"
         Height          =   315
         Index           =   3
         Left            =   5640
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   840
      End
      Begin VB.Image Imgw 
         Height          =   255
         Left            =   3960
         Picture         =   "FrmCustomerContractSearch.frx":3F63D
         Stretch         =   -1  'True
         Top             =   240
         Width           =   255
      End
      Begin VB.Image Imge 
         Height          =   255
         Left            =   1800
         Picture         =   "FrmCustomerContractSearch.frx":416D8
         Stretch         =   -1  'True
         Top             =   240
         Width           =   255
      End
   End
End
Attribute VB_Name = "FrmCustomerContractSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As FrmCustomerContractSearch
Private Sub fg_Click()
FrmCustomerContract.FindRec val(Fg.TextMatrix(Fg.row, 1))
End Sub
Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
         
    Set GrdBack = New ClsBackGroundPic
    With Me.Fg
       Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With
   If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    SetDtpickerDate Me.DTPicker1
    SetDtpickerDate Me.DTPicker2
    SetDtpickerDate Me.DTPicker3
    SetDtpickerDate Me.DTPicker4
   ' clear
          clear_all Me
          Me.DTPicker1.value = ""
          Me.DTPicker2.value = ""
          Me.DTPicker3.value = ""
          Me.DTPicker4.value = ""
          Me.Option1.value = True
   End Sub
  Private Sub Cmd_Click(Index As Integer)
    Select Case Index
        Case 0
        GetData
        Case 1
        clear_all Me
          Me.DTPicker1.value = ""
          Me.DTPicker2.value = ""
          Me.DTPicker3.value = ""
          Me.DTPicker4.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
                Me.lblL(0).Caption = "Search Results"
            End If
         Me.Option1.value = True
         Case 2
            Unload Me
    End Select
End Sub
Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub
Private Sub Form_Unload(Cancel As Integer)
    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub
Public Sub GetData()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    sql = "SELECT    dbo.TblCustomerContract.TblCustomerContractD, dbo.TblCustomerContract.CustomerId, dbo.TblCustomerContract.FromDate, dbo.TblCustomerContract.Todate, dbo.TblCustomerContract.Remarks,"
    sql = sql + "     dbo.TblCustomerContract.Locked , dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Type, dbo.TblCustemers.fullcode, dbo.TblCustemers.CustomerandVendor"
    sql = sql + " FROM         dbo.TblCustomerContract LEFT OUTER JOIN"
    sql = sql + "   dbo.TblCustemers ON dbo.TblCustomerContract.CustomerId = dbo.TblCustemers.CusID"
                
       BolBegine = False
       StrWhere = ""
   ''' ID SEARCH
    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblCustomerContract.TblCustomerContractD >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCustomerContract.TblCustomerContractD >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblCustomerContract.TblCustomerContractD <=" & val(Me.TxtIDTO.text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblCustomerContract.TblCustomerContractD <=" & val(Me.TxtIDTO.text) & ""
       End If
    End If
    ''''' DATA SEARCH
     If Not IsNull(Me.DTPicker1.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCustomerContract.FromDate >=" & SQLDate(Me.DTPicker1.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCustomerContract.FromDate >=" & SQLDate(Me.DTPicker1.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DTPicker2.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblCustomerContract.FromDate <=" & SQLDate(Me.DTPicker2.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblCustomerContract.FromDate <=" & SQLDate(Me.DTPicker2.value, True) & ""
        End If
    End If
    ''''''' SECAND DATE
     If Not IsNull(Me.DTPicker3.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCustomerContract.Todate >=" & SQLDate(Me.DTPicker3.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCustomerContract.Todate >=" & SQLDate(Me.DTPicker3.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DTPicker4.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblCustomerContract.Todate <=" & SQLDate(Me.DTPicker4.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblCustomerContract.Todate <=" & SQLDate(Me.DTPicker4.value, True) & ""
        End If
    End If
    ''''' COMBOW BOX SEARCH
    
    If Me.DBCboClientName.text <> "" And (val(DBCboClientName.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  dbo.TblCustomerContract.CustomerId =" & Me.DBCboClientName.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where dbo.TblCustomerContract.CustomerId =" & Me.DBCboClientName.BoundText & ""
       End If
     End If
     '''''' SEARCH TEXT
          If Me.txtRemarks.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblCustomerContract.Remarks like '%" & Me.txtRemarks.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblCustomerContract.Remarks like '%" & Me.txtRemarks.text & "%'"
        End If
       End If
       '''''' SEARCH CHECK BOX
        If (Me.Check1.value = True) Then
        If BolBegine = True Then
        StrWhere = StrWhere & " AND  dbo.TblCustomerContract.Locked = 1 "
         Else
        BolBegine = True
        StrWhere = StrWhere & " Where  dbo.TblCustomerContract.Locked = 0 "
        End If
        End If
        
        If (Me.Check2.value = True) Then
        If BolBegine = True Then
        StrWhere = StrWhere & " AND  dbo.TblCustomerContract.Locked= 0 "
         Else
        BolBegine = True
        StrWhere = StrWhere & " Where  dbo.TblCustomerContract.Locked = 1 "
        End If
        End If
       '-----------------------------------
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblCustomerContract.TblCustomerContractD"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lblL(10).Caption = "Search Results=0"
        End If
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.Fg
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lblL(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .rows - 1
                 '''' ID CULM
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("TblCustomerContractD").value), "", rs("TblCustomerContractD").value)
                .TextMatrix(i, .ColIndex("TxtVendorCode")) = IIf(IsNull(rs("fullcode").value), "", rs("fullcode").value)
                 '''' CHANGE DATE FORMAT
                 If Not (IsNull(rs("FromDate").value)) Then
                .TextMatrix(i, .ColIndex("FromDate")) = Format(rs("FromDate").value, "yyyy/M/d")
                 End If
                 If Not (IsNull(rs("Todate").value)) Then
                .TextMatrix(i, .ColIndex("Todate")) = Format(rs("Todate").value, "yyyy/M/d")
                 End If
                
                If SystemOptions.UserInterface = ArabicInterface Then
               .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                Else
               .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
                End If
                .TextMatrix(i, .ColIndex("Locked")) = IIf(IsNull(rs("Locked").value), "", rs("Locked").value)
               .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
                                  
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
       End With
    End If
End Sub
Private Sub ChangeLang()
    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "Search"
    Cmd(2).Caption = "Exit"
     Me.Caption = "Customer Contracts Search"
    ' labell name
    Me.Label1(2).Caption = Me.Caption
    Me.lbl(14).Caption = "Contracts ID"
    Me.lbl(5).Caption = "From"
    Me.lbl(6).Caption = "To"
    Me.lbl(3).Caption = "Locked Type"
    Me.lbl(4).Caption = "Customer Code"
    Me.lbl(7).Caption = "Customer Name"
    Me.lbl(9).Caption = "Start Date"
    Me.lbl(8).Caption = "End Date"
    Me.lbl(0).Caption = "To"
    Me.lbl(2).Caption = "To"
    Me.lbl(1).Caption = "Notice"
    Me.Option1.Caption = "All"
    Me.Check1.Caption = "UnLocked"
    Me.Check2.Caption = "Locked"
   ''''''''''''''''''''''' next
     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "No Contracts"
        .TextMatrix(0, .ColIndex("TxtVendorCode")) = "Customer Code"
         .TextMatrix(0, .ColIndex("CusName")) = "Customer Name"
        .TextMatrix(0, .ColIndex("FromDate")) = "Start Date"
       .TextMatrix(0, .ColIndex("Todate")) = "End Date"
         .TextMatrix(0, .ColIndex("Locked")) = "Locked Type"
       .TextMatrix(0, .ColIndex("Remarks")) = "Notice"
      End With
  End Sub
  Private Sub DBCboClientName_Click(Area As Integer)
  On Error Resume Next
    If val(DBCboClientName.BoundText) = 0 Then Exit Sub
    Dim Fullcode  As String
    GetCustomersDetail val(DBCboClientName.BoundText), , Fullcode, 2
    TxtSearchCode.text = Fullcode
End Sub
Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
   Dim CUSTID As Integer
    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.text, 2
        DBCboClientName.BoundText = CUSTID
    End If
 End Sub
  
'''''''''''''''''''''''''''' end



