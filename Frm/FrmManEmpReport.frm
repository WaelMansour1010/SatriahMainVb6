VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmManEmpReport 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ﬁ—Ì— „⁄«‰Ì… «·’Ì«‰…"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
   Icon            =   "FrmManEmpReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6195
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„ «»⁄«  «·”«»ﬁ… ··’‰› ( «—ÌŒ ’Ì«‰… «·’‰› )"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   2145
      Index           =   4
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   1680
      Width           =   8625
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   1875
         Left            =   60
         TabIndex        =   34
         Top             =   210
         Width           =   8505
         _cx             =   15002
         _cy             =   3307
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmManEmpReport.frx":038A
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
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3780
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   5730
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  ⁄„·Ì… «·„⁄«‰Ì…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1665
      Index           =   1
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   3840
      Width           =   8655
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Ì«‰«  «·„Ê—œ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   675
         Index           =   2
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   900
         Width           =   5925
         Begin MSDataListLib.DataCombo DcboClientName 
            Height          =   315
            Left            =   2340
            TabIndex        =   22
            Top             =   240
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbGoOutDtae 
            Height          =   345
            Left            =   60
            TabIndex        =   25
            Top             =   210
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   609
            _Version        =   393216
            Format          =   100073473
            CurrentDate     =   38784
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·—ÃÊ⁄ "
            Height          =   315
            Index           =   4
            Left            =   1350
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Ê—œ"
            Height          =   315
            Index           =   2
            Left            =   4980
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   240
            Width           =   825
         End
      End
      Begin MSDataListLib.DataCombo DcboDecs 
         Height          =   315
         Left            =   2580
         TabIndex        =   20
         Top             =   570
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label LblNextOperaType 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E2E9E9&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "«·⁄„·Ì… «· «·Ì… «·„ﬁ —Õ…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2580
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   240
         Width           =   4905
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·⁄„·Ì… :"
         Height          =   315
         Index           =   5
         Left            =   7530
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   240
         Width           =   945
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁ—«— «·„⁄«Ì‰…"
         Height          =   315
         Index           =   1
         Left            =   7530
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   570
         Width           =   945
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  ⁄„·Ì… «·„⁄«‰Ì…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1605
      Index           =   0
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   30
      Width           =   8685
      Begin MSDataListLib.DataCombo DcboCusName 
         Height          =   315
         Left            =   5250
         TabIndex        =   40
         Top             =   600
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.ComboBox CboMaintenanceType 
         Enabled         =   0   'False
         Height          =   315
         Left            =   30
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   180
         Width           =   1785
      End
      Begin VB.TextBox TxtOrgManID 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1590
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   570
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.TextBox TxtQuantity 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   360
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1215
         Width           =   555
      End
      Begin VB.TextBox TxtSerial 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   360
         Left            =   630
         MaxLength       =   20
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1215
         Width           =   1665
      End
      Begin VB.TextBox TxtTicketNO 
         Enabled         =   0   'False
         Height          =   330
         Left            =   7830
         TabIndex        =   8
         Top             =   1215
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox TxtMaintanenceID 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   7110
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   645
      End
      Begin MSComCtl2.DTPicker XPDtbGoInDtae 
         Height          =   345
         Left            =   5280
         TabIndex        =   4
         Top             =   210
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   609
         _Version        =   393216
         Format          =   100073473
         CurrentDate     =   38784
      End
      Begin MSDataListLib.DataCombo DCboStoreName 
         Height          =   315
         Left            =   2670
         TabIndex        =   6
         Top             =   180
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         ListField       =   "7"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCboItemsName 
         Height          =   315
         Left            =   2310
         TabIndex        =   11
         Top             =   1215
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCboItemsCode 
         Height          =   315
         Left            =   5760
         TabIndex        =   12
         Top             =   1215
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton CmdShowTransItems 
         Height          =   360
         Left            =   7320
         TabIndex        =   13
         Top             =   1185
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   635
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "...."
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
         ButtonImage     =   "FrmManEmpReport.frx":0513
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin MSDataListLib.DataCombo DcboEmp 
         Height          =   315
         Left            =   2670
         TabIndex        =   41
         Top             =   510
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "7"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lblReciptNumber 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„ÊŸ›"
         Height          =   255
         Index           =   25
         Left            =   4380
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   570
         Width           =   840
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·⁄„Ì·"
         Height          =   315
         Index           =   22
         Left            =   7770
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   630
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·’Ì«‰…"
         Height          =   255
         Index           =   21
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   210
         Width           =   840
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬂ„Ì…"
         Height          =   210
         Index           =   0
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   990
         Width           =   495
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·”Ì—Ì«·"
         Height          =   225
         Index           =   28
         Left            =   690
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   990
         Width           =   1575
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈”„ «·’‰›"
         Height          =   210
         Index           =   30
         Left            =   2790
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   990
         Width           =   2940
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·’‰›"
         Height          =   225
         Index           =   31
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   990
         Width           =   1530
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «· ﬂ "
         Height          =   240
         Index           =   11
         Left            =   7860
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   990
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„Œ“‰"
         Height          =   315
         Index           =   24
         Left            =   4365
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   840
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· «—ÌŒ"
         Height          =   315
         Index           =   3
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   270
         Width           =   495
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·⁄„·Ì…"
         Height          =   345
         Index           =   8
         Left            =   7740
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   885
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   2
      Left            =   1860
      TabIndex        =   27
      Top             =   5700
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   661
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   60
      TabIndex        =   28
      Top             =   5700
      Width           =   840
      _ExtentX        =   1482
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   930
      TabIndex        =   29
      Top             =   5700
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   661
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   5520
      TabIndex        =   30
      Top             =   5700
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   8655
      X2              =   30
      Y1              =   5610
      Y2              =   5625
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   300
      Index           =   6
      Left            =   7740
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   5730
      Width           =   945
   End
End
Attribute VB_Name = "FrmManEmpReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSearchDcbo(8) As clsDCboSearch
Dim Dcombos As ClsDataCombos
Dim IntOldIndex As Integer

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 1
            Unload Me

        Case 2
            SaveData
    End Select

End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim rs As ADODB.Recordset
    Dim RsNotes As ADODB.Recordset
    Dim BolBegine As Boolean
    Dim IntLastDecType As Integer
    Dim IntCurrentDecType As Integer

    On Error GoTo ErrTrap
    IntLastDecType = val(Me.FG.TextMatrix(Me.FG.Rows - 1, FG.ColIndex("SupDecID")))

    If Me.DcboEmp.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·„ÊŸ›...!!!"
        Else
            Msg = "Please Select Employee Name...!!!"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        DcboEmp.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If Me.DCboStoreName.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        
            Msg = "ÌÃ» ≈Œ Ì«— «·„Œ“‰....!!! " & Chr(13)
        Else
            Msg = "Please Select Store....!!! " & Chr(13)
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        DCboStoreName.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    '-----------------------------------------------------------------------------
    If DCboItemsCode.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
    
            Msg = "ÌÃ»  ÕœÌœ ﬂÊœ «·’‰›"
        Else
            Msg = "Please Select Item Code"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        DCboItemsCode.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If DCboItemsName.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
    
            Msg = "ÌÃ»  ÕœÌœ «”„ «·’‰›"
        Else
            Msg = "Please Select Item Name"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        DCboItemsName.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If val(TxtQuantity.text) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
    
            Msg = "ÌÃ»  ÕœÌœ ﬂ„Ì… «·’‰›"
        Else
            Msg = "Please Select Item Qty"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtQuantity.SetFocus
        Exit Sub
    End If

    If Me.TxtSerial.Enabled = True And Trim(Me.TxtSerial.text) = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "»—Ã«¡ ≈œŒ«· «·”Ì—»«· «·Œ«’ »«·’‰›...!!"
        Else
            Msg = "Please Enter Item Serial"
        End If
    
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtSerial.SetFocus
        Exit Sub
    End If

    If val(Me.DcboDecs.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ≈Œ Ì«— ﬁ—«— «·„⁄«‰Ì… ..!!!"
        Else
            Msg = "Please Enter Technical Notes"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        DcboDecs.SetFocus
        Exit Sub
    End If

    IntCurrentDecType = val(Me.DcboDecs.BoundText)

    If Me.DcboDecs.BoundText = 6 Then
        If val(Me.DcboClientName.BoundText) = 0 Then

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "»—Ã«¡ ≈Œ Ì«— ‘—ﬂ… «·’Ì«‰… «·„ÕÊ· ·Â« «·’‰› ...!!"
            Else
                Msg = "please   Select Maintenance Company"
            End If
 
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcboDecs.SetFocus
            Exit Sub
        End If

    ElseIf val(Me.DcboDecs.BoundText) = 7 Then

        If val(Me.DcboClientName.BoundText) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "»—Ã«¡ ≈Œ Ì«— «·„Ê—œ  ...!!"
            Else
                Msg = "please   Select Vendor"
            End If
 
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcboDecs.SetFocus
            Exit Sub
        End If
    End If

    If val(Me.DcboDecs.BoundText) = 8 Then
    
    End If

    If CheckDecision = False Then
        Exit Sub
    End If

    Cn.BeginTrans
    BolBegine = True

    Set rs = New ADODB.Recordset
    rs.Open "TblMainteneceNew", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    rs.AddNew
    Me.TxtMaintanenceID.text = new_id("TblMainteneceNew", "MaintananceID", "")
    rs("MaintananceID").value = val(Me.TxtMaintanenceID.text)
    rs("ReciptNumber").value = val(lblReciptNumber.Caption)
    rs("Transaction_ID").value = Null
    rs("CashCustomerName").value = Null
    rs("CashCustomerPhone").value = Null
    rs("CashCustomerMobile").value = Null
    rs("CashCustomerEmail").value = Null
    rs("CashCustomerAddress").value = Null
    rs("MType").value = Null
    rs("EmpID").value = val(Me.DcboEmp.BoundText)
    rs("StoreID").value = val(Me.DCboStoreName.BoundText)
    rs("DateGoIN").value = Me.XPDtbGoInDtae.value
    rs("DateGoOUT").value = Me.XPDtbGoOutDtae.value
    
    rs("Remarks").value = Null
    rs("UserID").value = Me.DCboUserName.BoundText
    rs("PaymentType").value = 0

    If val(Me.LblNextOperaType.Tag) = 7 Or Me.LblNextOperaType.Tag = "" Then

        '„⁄«‰Ì… «·„ÊŸ› «·„Œ ’
        If val(Me.DcboDecs.BoundText) >= 1 And val(Me.DcboDecs.BoundText) <= 5 Then
            '—›÷ Œ—ÊÃ „‰ «·÷„«‰ - —›÷ «‰ Â  › —… «·÷„«‰
            '3-⁄œ„ ≈„ﬂ«‰Ì… «· ’·ÌÕ
            '4- „ «· ’·ÌÕ œ«Œ· «·÷„«‰
            '5- „ «· ’·ÌÕ » ﬂ·›…
            rs("ManOperationTypeID").value = 7 ' ”ÃÌ· Õ—ﬂ… „ «»⁄… „ÊŸ› «·’Ì«‰…
            rs("CusID").value = Null
            rs("SupDeci").value = val(Me.DcboDecs.BoundText)
        ElseIf Me.DcboDecs.BoundText = 6 Then
            ' ÕÊÌ· ≈·Ï ‘—ﬂ… ’Ì«‰… Œ«—ÃÌ…
            rs("ManOperationTypeID").value = 5
            rs("CusID").value = val(Me.DcboClientName.BoundText)
            rs("SupDeci").value = Null
        ElseIf Me.DcboDecs.BoundText = 7 Then
            'Œ—ÊÃ ≈·Ï „Ê—œ
            rs("ManOperationTypeID").value = 2
            rs("CusID").value = val(Me.DcboClientName.BoundText)
            rs("SupDeci").value = Null
        ElseIf val(Me.DcboDecs.BoundText) = 15 Then
            ' ”·Ì„ ≈·Ï «·⁄„Ì·
            rs("ManOperationTypeID").value = 4
            rs("CusID").value = val(Me.DcboCusName.BoundText)
            rs("SupDeci").value = IntLastDecType
        End If

    ElseIf val(Me.LblNextOperaType.Tag) = 6 Then
        ' ”ÃÌ· Õ—ﬂ… —ÃÊ⁄ „‰ ‘—ﬂ… ’Ì«‰… Œ«—ÃÌ…
        rs("ManOperationTypeID").value = 6
        rs("SupDeci").value = val(Me.DcboDecs.BoundText)
        rs("CusID").value = val(Me.DcboClientName.BoundText)
        
    ElseIf val(Me.LblNextOperaType.Tag) = 3 Then
        'Õ—ﬂ… —ÃÊ⁄ ÷„«‰ „‰ „Ê—œ
        rs("ManOperationTypeID").value = 3
        rs("SupDeci").value = val(Me.DcboDecs.BoundText)
        rs("CusID").value = val(Me.DcboClientName.BoundText)
    Else
        Msg = "Â–« «·‰Ê⁄ €Ì— „œ⁄„ ·”Â ....!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

        If SystemOptions.SysRegisterState = DevelopVersion Then
            Stop
        Else
            Exit Sub
        End If
    End If

    rs("TicketNO").value = val(Me.TxtTicketNo.text)
    rs("ItemID").value = Me.DCboItemsName.BoundText

    If TxtSerial.Enabled = True Then
        rs("ItemSerial").value = Trim$(Me.TxtSerial.text)
    Else
        rs("ItemSerial").value = Null
    End If

    rs("Quantity").value = val(Me.TxtQuantity.text)
    rs("CustomerNotes").value = Null
    rs("EmpNotes").value = Null
    
    rs("RetrunOrgID").value = val(Me.TxtOrgManID.text)

    If val(Me.DcboDecs.BoundText) = 8 Or val(Me.DcboDecs.BoundText) = 15 Then
        
    Else
        rs("ReItemID").value = Null
        rs("ReItemSerial").value = Null
        rs("ReItemQuantity").value = Null
        rs("ReItemPrice").value = Null
        rs("ReItemStore").value = Null
    End If

    rs.update

    Cn.CommitTrans
    BolBegine = False

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = " „  ⁄„·Ì… «·Õ›Ÿ...!!!"
    Else
        Msg = "Saved Successfully"
    End If

    MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    FrmManStore.LoadManStore
    LoadTicketNO val(Me.TxtTicketNo.text)
    Exit Sub
ErrTrap:
    TerminateRecordset rs

    If BolBegine = True Then
        Cn.RollbackTrans
        BolBegine = False
    End If

    Msg = "ÕœÀ Œÿ« √À‰«¡  ”ÃÌ· «·»Ì«‰« "
    Msg = Msg & Chr(13) & Err.description
    Msg = Msg & Chr(13) & Err.Number
    Msg = Msg & Chr(13) & Err.Source
    Msg = Msg & Chr(13) & Err.LastDllError
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub CmdShowTransItems_Click()
    Dim Msg As String

    If Me.DCboStoreName.BoundText = "" Then
        Msg = "ÌÃ»  ÕœÌœ «·„Œ“‰...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Load FrmManChooseItems
    Set FrmManChooseItems.MyForm = Me
    FrmManChooseItems.ShowManStockItems Me.DCboStoreName.BoundText, Me.DCboStoreName.text
    FrmManChooseItems.Show vbModal

End Sub

Private Sub DcboDecs_Change()
    Dim LngCusID As Long
    Dim IntLastManType As Integer
    Dim IntLastDecType As Integer
    Dim IntCurrentDecType As Integer
    Dim BolFastReplace As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    If CheckDecision = False Then
        Exit Sub
    End If

    '«·€—÷ „‰ Â–Â «·œ«·… ÂÊ „—«Ã⁄… ﬁ—«—«  «·„ÊŸ›
    'Ê«· «ﬂœ „‰  „‘Ï »«· ”·”· «·ÿ»Ì⁄Ì ·‰Ÿ«„ «·⁄„·
    IntLastManType = val(Me.FG.TextMatrix(Me.FG.Rows - 1, FG.ColIndex("ManOperationTypeID")))
    IntLastDecType = val(Me.FG.TextMatrix(Me.FG.Rows - 1, FG.ColIndex("SupDecID")))
    IntCurrentDecType = val(Me.DcboDecs.BoundText)

    If val(Me.DcboDecs.BoundText) = 1 Then
        '—›÷ - Œ—ÊÃ „‰ «·÷„«‰
        Fra(2).Visible = False
   
    ElseIf val(Me.DcboDecs.BoundText) = 2 Then
        '—›÷ ≈‰ Â  › —… «·÷„«‰
        Fra(2).Visible = False
    
    ElseIf val(Me.DcboDecs.BoundText) = 3 Then
        '⁄œ„ ≈„ﬂ«‰Ì… «· ’·ÌÕ
        Fra(2).Visible = False
    
    ElseIf val(Me.DcboDecs.BoundText) = 4 Then
        ' „ «· ’·ÌÕ œ«Œ· «·÷„«‰
        Fra(2).Visible = False
   
    ElseIf val(Me.DcboDecs.BoundText) = 5 Then
        ' „ «· ’·ÌÕ » ﬂ·›…
        Fra(2).Visible = False

    ElseIf val(Me.DcboDecs.BoundText) = 6 Then
        ' ÕÊÌ· ≈·Ï ‘—ﬂ… ’Ì«‰… Œ«—ÃÌ…
        LoadManCompany
    ElseIf val(Me.DcboDecs.BoundText) = 7 Then
        ' ÕÊÌ· ≈·Ì ÷„«‰ „Ê—œ
        LoadManSupplier
        ' ÕœÌœ «”„ «·„Ê—œ «·Œ«’ »Â–Â «·’‰›
        LngCusID = GetItemSupplier

        If LngCusID <> 0 Then
            Me.DcboClientName.BoundText = LngCusID
            Me.DcboClientName.Enabled = False
        Else
            Me.DcboClientName.BoundText = ""
            Me.DcboClientName.Enabled = True
        End If

    ElseIf val(Me.DcboDecs.BoundText) = 8 Then

        '” »œ«· „‰ ⁄‰œ «·„Ê—œ
        If SystemOptions.UserInterface = ArabicInterface Then
            Fra(5).Caption = "»Ì«‰«  «·’‰› «·„” »œ· „‰ ⁄‰œ «·„Ê—œ"
        Else
            Fra(5).Caption = "Replace Item At Supplier"
        End If

        Fra(5).Visible = True
    ElseIf val(Me.DcboDecs.BoundText) = 12 Then
        ' Œ’Ì„ ⁄·Ï «·„Ê—œ
        lbl(17).Visible = False
        lbl(18).Visible = False
    
        PutItemCostPrice
    ElseIf val(Me.DcboDecs.BoundText) = 13 Then
    
    ElseIf val(Me.DcboDecs.BoundText) = 15 Then
        Fra(2).Visible = False
        '≈ŸÂ«— »Ì«‰«  „Õ«”»… «·⁄„Ì·
    
        PutCustomerSheet
    
        If IntLastDecType = 8 Then
       
        Else
        
        End If

    ElseIf val(Me.DcboDecs.BoundText) = 16 Then

        '≈” »œ«· »’‰› ÃœÌœ
        If SystemOptions.UserInterface = ArabicInterface Then
            Fra(5).Caption = "»Ì«‰«  «·’‰› «·„” »œ·"
        Else
            Fra(5).Caption = "Replaced Item Data"
        End If

        Fra(5).Visible = True
    End If

End Sub

Private Sub DcboDecs_Click(Area As Integer)
    'DcboDecs_Change
    IntOldIndex = val(Me.DcboDecs.BoundText)
End Sub

Private Sub ChangeLang()
    Me.Caption = "Maintenance preview"
    'Ele(0).Caption = Me.Caption
    lbl(8).Caption = "Opr#"
    Cmd(2).Caption = "save"
 
    Cmd(0).Caption = "Print"
    Cmd(1).Caption = "Exit"
    lbl(3).Caption = "Date"
    lbl(24).Caption = "Store"
    lbl(25).Caption = "Employee"
    lbl(22).Caption = "Cust. Name"
   
    lbl(21).Caption = "Main. Type"
   
    Fra(0).Caption = "Maintenance Data"
    lbl(11).Caption = "Ticket NO."
 
    lbl(1).Caption = "Select"
 
    lbl(31).Caption = "Item Code"
    lbl(30).Caption = "Item Name"
    lbl(28).Caption = "Serial"
    lbl(0).Caption = "Qty"
     
    Fra(4).Caption = "Records"
    Fra(1).Caption = "Current Status"
    lbl(2).Caption = "Opr Type"
    lbl(5).Caption = "Current Opr"
    Fra(2).Caption = "Supplier"
    lbl(2).Caption = "Supplier "
    lbl(4).Caption = "Return Date"
    lbl(6).Caption = "by"

    With FG
        .TextMatrix(0, .ColIndex("MaintananceID")) = "Opr ID"
        .TextMatrix(0, .ColIndex("DateGoIN")) = "Date"
        .TextMatrix(0, .ColIndex("ManOperationTypeName")) = "Opr Type"
        .TextMatrix(0, .ColIndex("SupDecName")) = "Des"
        .TextMatrix(0, .ColIndex("FastReplace")) = "Replace"
        .TextMatrix(0, .ColIndex("Serial")) = "Ser"
        .TextMatrix(0, .ColIndex("Comment")) = "Comment"
   
    End With
   
End Sub

Private Sub Form_Load()
    Dim GrdBack  As ClsBackGroundPic

    CenterForm Me

    FormPostion Me, GetPostion

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture

    SetDtpickerDate Me.XPDtbGoInDtae
    SetDtpickerDate Me.XPDtbGoOutDtae

    Set GrdBack = New ClsBackGroundPic

    With Me.FG
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows
        .RowHeightMin = 300
        .ExtendLastCol = True
        Set .WallPaper = GrdBack.Picture
        .ExplorerBar = flexExMove
    End With

    If SystemOptions.UserInterface = ArabicInterface Then

        With CboMaintenanceType
            .Clear
            .AddItem "œ«Œ· «·÷„«‰"
            .AddItem "Œ«—Ã «·÷„«‰"
        End With

    Else

        With CboMaintenanceType
            .Clear
            .AddItem "With Warrenty"
            .AddItem "WithOut Warrenty"
        End With

    End If

    Set Dcombos = New ClsDataCombos
    Dcombos.GetCustomersSuppliers 0, Me.DcboCusName, True
    Dcombos.GetEmployees Me.DcboEmp
    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DcboEmp

    Dcombos.GetCustomersSuppliers 0, Me.DcboClientName, True
    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DcboClientName

    Dcombos.GetStores Me.DCboStoreName
    Set cSearchDcbo(2) = New clsDCboSearch
    Set cSearchDcbo(2).Client = Me.DCboStoreName

    Dcombos.GetItemsCodes Me.DCboItemsCode
    Set cSearchDcbo(3) = New clsDCboSearch
    Set cSearchDcbo(3).Client = Me.DCboItemsCode

    Dcombos.GetItemsNames Me.DCboItemsName
    Set cSearchDcbo(4) = New clsDCboSearch
    Set cSearchDcbo(4).Client = Me.DCboItemsName

    Dcombos.GetManEmpDes Me.DcboDecs
    Dcombos.GetUsers Me.DCboUserName
    Me.DCboUserName.BoundText = user_id
    Me.TxtMaintanenceID.text = new_id("TblMainteneceNew", "MaintananceID", "")
    '------------------------------------
    Fra(2).Visible = False

    '-----------------------------------
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer

    FormPostion Me, SavePostion
    For i = LBound(cSearchDcbo) To UBound(cSearchDcbo)
        Set cSearchDcbo(i) = Nothing
    Next i

End Sub

Private Sub TxtTicketNO_Change()
    LoadTicketNO val(Me.TxtTicketNo.text)
End Sub

Private Sub LoadTicketNO(LngTicktNO As Long)
    Dim StrSQL  As String
    Dim rs As ADODB.Recordset
    Dim StrTemp  As String
    Dim i As Integer

    '------------------------------------------------------------------------------
    ' Õ„Ì· »Ì«‰«  «·’‰› «·√”«”Ì…
    StrSQL = "Select * From TblMainteneceNew Where TicketNO=" & LngTicktNO
    StrSQL = StrSQL + " AND ManOperationTypeID=1"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        Me.DCboItemsCode.BoundText = rs("ItemID").value
        Me.DCboItemsName.BoundText = rs("ItemID").value

        If IsNull(rs("ItemSerial").value) Then
            Me.TxtSerial.text = ""
            Me.TxtSerial.Enabled = False
        Else
            Me.TxtSerial.text = rs("ItemSerial").value
            Me.TxtSerial.Enabled = True
        End If

        Me.TxtQuantity.text = rs("Quantity").value
        Me.DCboStoreName.BoundText = rs("StoreID").value
        Me.CboMaintenanceType.ListIndex = rs("MType").value
        Me.DcboCusName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    End If

    rs.Close
    '------------------------------------------------------------------------------
    ' Õ„Ì· ﬂ· ⁄„·Ì«  «·„ «»⁄… «· Ï Õ’·  ··’‰›
    StrSQL = "SELECT     dbo.TblMainteneceNew.MaintananceID, dbo.TblMainteneceNew.ReciptNumber," & "dbo.TblMainteneceNew.DateGoIN, dbo.TblMainteneceNew.DateGoOUT,dbo.TblMainteneceNew.GoOut," & "dbo.TblMainteneceNew.PaymentType, dbo.TblMainteneceNew.MType, dbo.TblMainteneceNew.ManOperationTypeID," & "dbo.TblManOperations.ManOperationTypeName,dbo.TblManOperations.ManOperationTypeNamee, dbo.TblMainteneceNew.TicketNO, dbo.TblMainteneceNew.CustomerNotes," & "dbo.TblMainteneceNew.EmpNotes, dbo.TblMainteneceNew.Cost, dbo.TblMainteneceNew.SupDeci," & "dbo.TblManSupDecs.SupDecName,dbo.TblManSupDecs.SupDecNamee,dbo.TblMainteneceNew.FastReplace,dbo.TblMainteneceNew.ReItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName," & "dbo.TblMainteneceNew.ReItemSerial,dbo.TblMainteneceNew.ReItemQuantity , dbo.TblMainteneceNew.ReItemStore," & "dbo.TblStore.StoreName "
    StrSQL = StrSQL + " FROM dbo.TblMainteneceNew LEFT OUTER JOIN "
    StrSQL = StrSQL + " dbo.TblStore ON dbo.TblMainteneceNew.ReItemStore = dbo.TblStore.StoreID LEFT OUTER JOIN "
    StrSQL = StrSQL + " dbo.TblItems ON dbo.TblMainteneceNew.ReItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
    StrSQL = StrSQL + " dbo.TblManSupDecs ON dbo.TblMainteneceNew.SupDeci = dbo.TblManSupDecs.SupDecID LEFT OUTER JOIN"
    StrSQL = StrSQL + " dbo.TblManOperations ON dbo.TblMainteneceNew.ManOperationTypeID = dbo.TblManOperations.ManOperationTypeID "
    StrSQL = StrSQL + " Where dbo.TblMainteneceNew.TicketNO=" & LngTicktNO & ""
    StrSQL = StrSQL + " Order BY dbo.TblMainteneceNew.MaintananceID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.FG
        .Rows = .FixedRows

        If Not (rs.BOF Or rs.EOF) Then
            .Rows = .FixedRows + rs.RecordCount
            rs.MoveFirst

            For i = 1 To rs.RecordCount
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("MaintananceID")) = IIf(IsNull(rs("MaintananceID").value), "", rs("MaintananceID").value)

                If Not IsNull(rs("DateGoIN").value) Then
                    .TextMatrix(i, .ColIndex("DateGoIN")) = DisplayDate(rs("DateGoIN").value)
                Else
                    .TextMatrix(i, .ColIndex("DateGoIN")) = ""
                End If

                .TextMatrix(i, .ColIndex("ManOperationTypeID")) = IIf(IsNull(rs("ManOperationTypeID").value), "", rs("ManOperationTypeID").value)
                .TextMatrix(i, .ColIndex("SupDecID")) = IIf(IsNull(rs("SupDeci").value), "", rs("SupDeci").value)

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("ManOperationTypeName")) = IIf(IsNull(rs("ManOperationTypeName").value), "", rs("ManOperationTypeName").value)
                    .TextMatrix(i, .ColIndex("SupDecName")) = IIf(IsNull(rs("SupDecName").value), "", rs("SupDecName").value)
                Else
                    .TextMatrix(i, .ColIndex("ManOperationTypeName")) = IIf(IsNull(rs("ManOperationTypeNamee").value), "", rs("ManOperationTypeNamee").value)
                    .TextMatrix(i, .ColIndex("SupDecName")) = IIf(IsNull(rs("SupDecNamee").value), "", rs("SupDecNamee").value)
                End If
            
                If rs("FastReplace").value = 1 Then
                    .Cell(flexcpChecked, i, .ColIndex("FastReplace")) = flexChecked
                Else
                    .Cell(flexcpChecked, i, .ColIndex("FastReplace")) = flexUnchecked
                End If

                If rs("ManOperationTypeID").value = 1 Then

                    'œŒÊ· ’Ì«‰…
                    If Not (IsNull(rs("ReItemID").value)) Then
                        StrTemp = ""
                        StrTemp = StrTemp & "ﬂÊœ «·’‰› :" & rs("ItemCode").value
                        StrTemp = StrTemp & "«”„ «·’‰› :" & rs("ItemName").value
                        StrTemp = StrTemp & "«·”Ì—Ì«· :" & rs("ReItemSerial").value
                        .TextMatrix(i, .ColIndex("Comment")) = StrTemp
                    End If
                End If

                rs.MoveNext
            Next i

        End If

        '    .Cell(flexcpFontName, .FixedRows, .ColIndex("ManOperationTypeName"), .Rows - 1, .ColIndex("SupDecName")) = "Tahoma"
        '    .Cell(flexcpForeColor, .FixedRows, .ColIndex("ManOperationTypeName"), .Rows - 1, .ColIndex("ManOperationTypeName")) = vbRed
        '    .Cell(flexcpForeColor, .FixedRows, .ColIndex("SupDecName"), .Rows - 1, .ColIndex("SupDecName")) = vbBlue
        .AutoSize 0, .Cols - 1, False
    End With

    Set rs = Nothing
    LoadAvailableManOpera
End Sub

Private Sub LoadAvailableManOpera()
    Dim IntLastManOperType As Integer
    Dim IntLastManDesc As Integer
    Dim LngManID As Long
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        Me.LblNextOperaType.Caption = "«·⁄„·Ì… «· «·Ì…«·„ﬁ —Õ…"
    Else
        Me.LblNextOperaType.Caption = "The second operation proposed"
        
    End If
        
    ' ⁄ „œ ›ﬂ—… Â–Â «·⁄„·Ì…
    '⁄·Ï «Œ— ⁄„·Ì… „ «»⁄… ··’‰›
    'Ê ÕœÌœ „«ÂÏ «·⁄„·Ì«  «·„› —÷… «· «·Ì… ·Â«
    '-----------------------------------
    Me.LblNextOperaType.Enabled = True
    Me.DcboDecs.Enabled = True
    Me.Cmd(2).Enabled = True

    '-----------------------------------
    With Me.FG
        '‰Ê⁄ «Œ— ⁄„·Ì… ’Ì«‰…
        IntLastManOperType = val(.TextMatrix(.Rows - 1, .ColIndex("ManOperationTypeID")))
        '—ﬁ„ «Œ— ⁄„·Ì… ’Ì«‰…
        LngManID = val(.TextMatrix(.Rows - 1, .ColIndex("MaintananceID")))
        '«Œ— ﬁ—«— Œ«’ »«·’Ì«‰…
        IntLastManDesc = val(.TextMatrix(.Rows - 1, .ColIndex("SupDecID")))
    End With

    If IntLastManOperType = 1 Then

        '⁄„·Ì… œŒÊ· ’Ì«‰…
        '«·⁄„·Ì… «· «·Ì… «·„ﬁ —Õ… ÂÏ
        '√‰ ÌﬁÊ„ «·„” Œœ„ » ÕœÌÂ«
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.LblNextOperaType.Caption = "ﬁ„ » ÕœÌœ „”«— «·’‰›"
        Else
            Me.LblNextOperaType.Caption = "Specify First Decision"
        End If

        Me.LblNextOperaType.Tag = 7
        Dcombos.GetManEmpDes Me.DcboDecs
    ElseIf IntLastManOperType = 2 Then

        '«Œ— ⁄„·Ì… ÂÏ
        'Œ—ÊÃ ÷„«‰ ··„Ê—œ
        '«·⁄„·Ì… «· «·Ì… ÂÏ —ÃÊ⁄ „‰ ⁄‰œ «·„Ê—œ
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.LblNextOperaType.Caption = "—ÃÊ⁄ ÷„«‰ „‰ «·„Ê—œ"
        Else
            Me.LblNextOperaType.Caption = "Back guarantee from the supplier"
        End If

        Me.LblNextOperaType.Tag = 3
        Dcombos.GetManSupDecs Me.DcboDecs
        LoadManSupplier
        Set rs = New ADODB.Recordset
        StrSQL = "Select * From TblMainteneceNew Where MaintananceID=" & LngManID
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            If Not (IsNull(rs("CusID").value)) Then
                Me.DcboClientName.BoundText = rs("CusID").value
            End If
        End If

        Me.DcboClientName.Enabled = False
    ElseIf IntLastManOperType = 3 Then

        '«Œ— ⁄„·Ì… ÂÏ —ÃÊ⁄ ÷„«‰ „‰ «·„Ê—œ
        If IntLastManDesc = 12 Then
    
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.LblNextOperaType.Caption = " ”·Ì„ «·⁄„Ì· Ê„Õ«”»… «·⁄„Ì·"
            Else
                Me.LblNextOperaType.Caption = "Delivery and client accounting"
            End If
 
            Me.LblNextOperaType.Tag = 4
            StrSQL = " SupDecID=15"
            Dcombos.GetManDesByID Me.DcboDecs, StrSQL
        Else

            '≈–«  ﬂÊ‰ «·Õ«·… Â‰« ÂÏ ≈ «Õ…
            'Õ«·… «·„⁄«‰Ì… ··„ÊŸ› «·„Œ ’
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.LblNextOperaType.Caption = "«·⁄„·Ì… «· «·Ì…«·„ﬁ —Õ…"
            Else
                Me.LblNextOperaType.Caption = "The second operation proposed"
        
            End If

            Me.LblNextOperaType.Tag = 7
            Fra(2).Visible = False
        End If

    ElseIf IntLastManOperType = 4 Then

        ' ”·Ì„ ··⁄„Ì· √Œ— ⁄„·Ì… ÂÏ
        ' .. Ê»⁄œÂ« ·Ì” Â‰«ﬂ «Ï ‘Ï¡ «Œ—
        '    Me.LblNextOperaType.Caption = "«·⁄„·Ì… «· «·Ì…«·„ﬁ —Õ…"
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.LblNextOperaType.Caption = "«·⁄„·Ì… «· «·Ì…«·„ﬁ —Õ…"
        Else
            Me.LblNextOperaType.Caption = "The second operation proposed"
        
        End If
        
        Me.LblNextOperaType.Tag = 7
        Me.LblNextOperaType.Enabled = False
        Me.DcboDecs.Enabled = False
        Me.Cmd(2).Enabled = False
        Fra(2).Visible = False
    
    ElseIf IntLastManOperType = 5 Then
        '«Œ— ⁄„·Ì… ÂÏ ⁄„·Ì…  ÕÊÌ· ·‘—ﬂ… ’Ì«‰… Œ«—ÃÌ…
        'Ê«·⁄„·Ì… «· «·Ì… ·Â« ÂÏ —ÃÊ⁄ „‰ ‘—ﬂ… ’Ì«‰…
        'Ê·«»œ „‰  ÕœÌœ ﬁ—«— «Ê ‰ ÌÃ… ‘—ﬂ… «·’Ì«‰…
    
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.LblNextOperaType.Caption = "—ÃÊ⁄ „‰ ‘—ﬂ… ’Ì«‰… Œ«—ÃÌ…"
        Else
            Me.LblNextOperaType.Caption = "Back from external maintenance"
        
        End If
        
        Me.LblNextOperaType.Tag = 6
        Dcombos.GetManComDecs Me.DcboDecs
        '≈⁄œ«œ ⁄—÷ ‘—ﬂ«  «·’Ì«‰…
        LoadManCompany
        Set rs = New ADODB.Recordset
        StrSQL = "Select * From TblMainteneceNew Where MaintananceID=" & LngManID
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            If Not (IsNull(rs("CusID").value)) Then
                Me.DcboClientName.BoundText = rs("CusID").value
            End If
        End If

        Me.DcboClientName.Enabled = False
    ElseIf IntLastManOperType = 6 Then

        '«Œ— ⁄„·Ì… ÂÏ ⁄„·Ì… —ÃÊ⁄  „‰ ‘—ﬂ… ’Ì«‰… Œ«—ÃÌ…
        'Ê«·⁄„·Ì… «· «·Ì… ·Â«   Êﬁ› ⁄·Ï ﬁ—«— ‘—ﬂ… «·’Ì«‰…
        'Â·  „ «· ’·ÌÕ «„ ·«
        If IntLastManDesc = 13 Then ' „ «· ’·ÌÕ
            '›Ï Â–Â «·Õ«·…  ⁄—÷ ··„” Œœ„  ”·Ì„ Ê„Õ«”»… «·⁄„Ì·
            Dcombos.GetManDesByID Me.DcboDecs, "SupDecID=15"
        
        End If
    End If

End Sub

Private Sub LoadManCompany()
    Fra(2).Visible = True

    If SystemOptions.UserInterface = ArabicInterface Then
        Fra(2).Caption = "»Ì«‰«  «·‘—ﬂ…"
        lbl(2).Caption = "«”„ «·‘—ﬂ…"
    Else
        Fra(2).Caption = "Company Data"
        lbl(2).Caption = "Co. Name"

    End If

    If val(Me.LblNextOperaType.Tag) = 0 Or val(Me.LblNextOperaType.Tag) = 7 Then
        Me.DcboClientName.Enabled = True
    Else
        Me.DcboClientName.Enabled = False
    End If

    'Dcombos.GetManCompanies Me.DcboClientName
    Dcombos.GetCustomersSuppliers 0, Me.DcboClientName, True

    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DcboClientName
End Sub

Private Sub LoadManSupplier()
    Fra(2).Visible = True
 
    If SystemOptions.UserInterface = ArabicInterface Then
        Fra(2).Caption = "»Ì«‰«  «·„Ê—œ"
        lbl(2).Caption = "«”„ «·„Ê—œ"
    Else
        Fra(2).Caption = "Supplier Data"
        lbl(2).Caption = "Supplier. Name"

    End If

    Dcombos.GetCustomersSuppliers 0, Me.DcboClientName, True
    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DcboClientName
End Sub

Private Function GetItemSupplier() As Long
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    On Error GoTo ErrTrap

    If Trim$(Me.TxtSerial.text) <> "" Then
        StrSQL = "Select * From SearchSerialData()SearchSerialData "
        StrSQL = StrSQL + " Where ItemSerial='" & Trim$(Me.TxtSerial.text) & "'"
        StrSQL = StrSQL + " AND Transaction_Type=1"
        StrSQL = StrSQL + " AND ItemID=" & val(Me.DCboItemsName.BoundText)
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            If Not (IsNull(rs("CusID").value)) Then
                GetItemSupplier = rs("CusID").value
            End If
        End If
    End If

ErrTrap:
End Function

Private Function CheckDecision() As Boolean
    Dim Msg As String
    Dim IntLastManType As Integer
    Dim IntLastDecType As Integer
    Dim IntCurrentDecType As Integer
    Dim BolFastReplace As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    '«·€—÷ „‰ Â–Â «·œ«·… ÂÊ „—«Ã⁄… ﬁ—«—«  «·„ÊŸ›
    'Ê«· «ﬂœ „‰  „‘Ï »«· ”·”· «·ÿ»Ì⁄Ì ·‰Ÿ«„ «·⁄„·
    IntLastManType = val(Me.FG.TextMatrix(Me.FG.Rows - 1, FG.ColIndex("ManOperationTypeID")))
    IntLastDecType = val(Me.FG.TextMatrix(Me.FG.Rows - 1, FG.ColIndex("SupDecID")))
    IntCurrentDecType = val(Me.DcboDecs.BoundText)
    CheckDecision = True

    If IntCurrentDecType = 0 Then
        Exit Function
    End If

    If IntLastManType = 0 Then
        Exit Function
    End If

    If IntCurrentDecType = IntLastDecType Then
        If SystemOptions.UserInterface = ArabicInterface Then

            Msg = "·«Ì„ﬂ‰ ﬁ»Ê· Â–« «·ﬁ—«— ..."
            Msg = Msg & Chr(13) & "(·«Ì„ﬂ‰ «‰ ÌﬂÊ‰ Â‰«ﬂ ‰›” «·ﬁ—«—Ì‰ „  «·Ì‰)"
        Else
            Msg = " Can't Accept This Decision ..."
            Msg = Msg & Chr(13) & "Decision Aleary Exist"

        End If

        MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        CheckDecision = False
        ResetDcboDesc
        Exit Function
    End If

    If IntLastDecType = 5 Then

        '«Œ— ﬁ—«— ﬂ«‰  «· ’·ÌÕ » ﬂ·›…
        'Ê»⁄œÂ« ·«»œ „‰  ”·Ì„ «·’‰› ··⁄„Ì·
        If IntCurrentDecType <> 15 Then

            '«Œ— „⁄«Ì‰… ··’‰› ﬁ——  «· ’·ÌÕ » ﬂ·›…
            'Ê·«»œ «‰ Ì „  ”·Ì„ «·’‰› ··⁄„Ì·
            If SystemOptions.UserInterface = ArabicInterface Then
        
                Msg = "·«Ì„ﬂ‰ ﬁ»Ê· Â–« «·ﬁ—«— ..."
                Msg = Msg & Chr(13) & "·«»œ «‰ Ì „  ”·Ì„ «·’‰› ··⁄„Ì·"
            Else
                Msg = "Cant Accept this Decision ..."
                Msg = Msg & Chr(13) & " Item Must Be Delivered To Customer"
        
            End If

            MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CheckDecision = False
            Me.DcboDecs.BoundText = 15
            Exit Function
        End If
    End If

    If val(Me.CboMaintenanceType.ListIndex) = 1 Then

        '«·’‰› Œ«—Ã «·’Ì«‰…
        If IntCurrentDecType = 1 Or IntCurrentDecType = 2 Or IntCurrentDecType = 4 Or IntCurrentDecType = 7 Then

            If SystemOptions.UserInterface = ArabicInterface Then
        
                Msg = "·«Ì„ﬂ‰ ﬁ»Ê· Â–« «·ﬁ—«— ..."
        
                Msg = Msg & Chr(13) & "(Â–« «·’‰› √”«”« Œ«—Ã «·÷„«‰)"
            Else
                Msg = "Cant Accept this Decision ..."
                Msg = Msg & Chr(13) & " Item Out Warrenty"
       
            End If

            MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CheckDecision = False
            ResetDcboDesc
            Exit Function
        End If
    End If

    If IntLastManType = 1 Then '⁄„·Ì… œŒÊ· ’Ì«‰…
        If IntCurrentDecType = 15 Then ' ”·Ì„ «·⁄„Ì·

            '›Ï Õ«·… «‰ «Œ— ⁄„·Ì… ÂÏ œŒÊ· ’Ì«‰… ›«‰Â »⁄œÂ« ·«Ì„ﬂ‰  ”·Ì„ «·⁄„Ì· „»«‘—…
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·«Ì„ﬂ‰ ﬁ»Ê· Â–« «·ﬁ—«— ..."
                Msg = Msg & Chr(13) & "(€Ì— „ﬁ»Ê· «‰ Ì „  ”·Ì„ «·⁄„Ì· „»«‘—… »⁄œ œŒÊ· «·’Ì«‰…)"
            Else
                Msg = "Cant Accept this Decision ..."
                Msg = Msg & Chr(13) & " Cant Delivert To Customer Directly"
    
            End If
    
            MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CheckDecision = False
            ResetDcboDesc
            Exit Function
        ElseIf (IntCurrentDecType = 1 Or IntCurrentDecType = 2 Or IntCurrentDecType = 3 Or IntCurrentDecType = 5) Then
            Set rs = New ADODB.Recordset
            StrSQL = "Select * From TblMainteneceNew Where MaintananceID=" & val(Me.TxtOrgManID.text)
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                BolFastReplace = IIf(rs("FastReplace").value = 1, True, False)
            End If
        
            If BolFastReplace = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
        
                    Msg = "·«Ì„ﬂ‰ ﬁ»Ê· Â–« «·ﬁ—«—...!!!"
                    Msg = Msg & Chr(13) & "( „ ⁄„· ≈” »œ«· ›Ê—Ì ··’‰› )"
                Else
                    Msg = "Cant Accept this Decision ..."
                    Msg = Msg & Chr(13) & " immediate replacement For Item Was Done"
    
                End If

                MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                CheckDecision = False
                ResetDcboDesc
                Exit Function
            End If
        End If
    End If

End Function

Private Sub ResetDcboDesc()

    If IntOldIndex = 0 Then
        Me.DcboDecs.BoundText = ""
    Else
        Me.DcboDecs.BoundText = IntOldIndex
    End If

End Sub

Private Sub PutItemCostPrice()
    Dim LngItemID As Long
    Dim StrItemSerial As String
    LngItemID = val(Me.DCboItemsName.BoundText)
    StrItemSerial = Trim$(Me.TxtSerial.text)

End Sub

Private Sub PutCustomerSheet()

    Dim LngLastManID As Long
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    LngLastManID = val(Me.FG.TextMatrix(Me.FG.Rows - 1, FG.ColIndex("MaintananceID")))

    StrSQL = "Select * From TblMainteneceNew Where MaintananceID=" & LngLastManID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
    
    End If

    If Not IsNull(rs("ReItemID").value) Then
    
    End If

    rs.Close
    '«·≈” ⁄·«„ ⁄‰ ﬁÌ„… «·„ﬁœ„ «·–Ï Ìœ›⁄Â «·⁄„Ì·
    StrSQL = "Select * From NOTES Where MaintananceID=" & val(Me.TxtOrgManID.text) & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        Me.lbl(18).Caption = IIf(IsNull(rs("Note_Value").value), 0, rs("Note_Value").value)
    End If

    rs.Close
End Sub
