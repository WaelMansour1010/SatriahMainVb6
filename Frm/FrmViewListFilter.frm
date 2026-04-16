VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmviewListFilter 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ÕœÌœ «·»Ì«‰«  ›Ï «·⁄—÷ «·ÃœÊ·"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10485
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame FraNumber 
      BackColor       =   &H00E2E9E9&
      Caption         =   " ÕœÌœ «·ﬁÌ„"
      Height          =   1095
      Left            =   4290
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   630
      Width           =   2475
      Begin VB.TextBox TxtNumberTo 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   510
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   600
         Width           =   1395
      End
      Begin VB.TextBox TxtNumberFrom 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   510
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   345
         Index           =   3
         Left            =   2010
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   345
         Index           =   2
         Left            =   2010
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   375
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   405
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   5730
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   714
      Caption         =   "Œ—ÊÃ"
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
   Begin VB.Frame FraString 
      BackColor       =   &H00E2E9E9&
      Caption         =   "Frame1"
      Height          =   4335
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   630
      Width           =   6675
      Begin VSFlex8UCtl.VSFlexGrid FgAll 
         Height          =   3255
         Left            =   3360
         TabIndex        =   7
         Top             =   360
         Width           =   3285
         _cx             =   5794
         _cy             =   5741
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
         AllowUserResizing=   0
         SelectionMode   =   0
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
         FormatString    =   $"FrmViewListFilter.frx":0000
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
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
         Height          =   3255
         Left            =   60
         TabIndex        =   10
         Top             =   360
         Width           =   3285
         _cx             =   5794
         _cy             =   5741
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
         AllowUserResizing=   0
         SelectionMode   =   0
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
         FormatString    =   $"FrmViewListFilter.frx":006D
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
   Begin VB.Frame FraDate 
      BackColor       =   &H00E2E9E9&
      Caption         =   " ÕœÌœ ﬁÌ„… ·· «—ÌŒ"
      Height          =   1095
      Left            =   4290
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   660
      Width           =   2475
      Begin MSComCtl2.DTPicker DtpFrom 
         Height          =   345
         Left            =   150
         TabIndex        =   2
         Top             =   300
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   100073473
         CurrentDate     =   39531
      End
      Begin MSComCtl2.DTPicker DtpTo 
         Height          =   345
         Left            =   150
         TabIndex        =   5
         Top             =   660
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   100073473
         CurrentDate     =   39531
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   345
         Index           =   1
         Left            =   1860
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   660
         Width           =   375
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   345
         Index           =   0
         Left            =   1860
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   300
         Width           =   375
      End
   End
   Begin MSComctlLib.TreeView TrvCols 
      Height          =   5505
      Left            =   6810
      TabIndex        =   0
      Top             =   30
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   9710
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   90
      TabIndex        =   9
      Top             =   5160
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      Caption         =   "≈⁄ „œ"
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
   Begin VB.Label lblHeader 
      Alignment       =   1  'Right Justify
      Height          =   465
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   90
      Width           =   3975
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   10410
      X2              =   0
      Y1              =   5640
      Y2              =   5640
   End
End
Attribute VB_Name = "frmviewListFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrCurrentKey As String

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
            Unload Me

        Case 1
    End Select

End Sub

Private Sub Form_Load()
    CenterForm Me

    FormPostion Me, GetPostion
    Set Me.Icon = mdifrmmain.ImgLstTree.ListImages("Filter").Picture
    '-------------------------------------------
    Me.TrvCols.Appearance = ccFlat
    Me.TrvCols.BorderStyle = ccNone
    Me.TrvCols.Checkboxes = False
    'Me.TrvCols.FullRowSelect = True
    Me.TrvCols.LabelEdit = tvwManual
    Me.TrvCols.LineStyle = tvwRootLines
    Me.TrvCols.Style = tvwTreelinesPlusMinusPictureText
    Set TrvCols.ImageList = mdifrmmain.ImgLstTree
    Make_RightToLeft Me.TrvCols
    '-------------------------------------------
    Me.Cmd(0).ButtonStyle = impActive
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Cmd(0).ButtonPositionImage = impRightOfText
    '-------------------------------------------
    SetDtpickerDate DTPFrom
    SetDtpickerDate DTPTo
    Me.FraDate.Visible = False
    Me.FraNumber.Visible = False
    Me.FraString.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub TrvCols_NodeClick(ByVal Node As MSComctlLib.Node)
    Me.FraDate.Visible = False
    Me.FraNumber.Visible = False
    Me.FraString.Visible = False

    Select Case Node.Tag

        Case "N"
            Me.FraNumber.Visible = True

        Case "S"
            Me.FraString.Visible = True

        Case "D"
            Me.FraDate.Visible = True
    End Select

    StrCurrentKey = Node.key
    LblHeader.Caption = Node.text
End Sub
