VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmBuySearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·»ÕÀ ⁄‰ ⁄„·Ì… ‘—«¡"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
   Icon            =   "FrmBuySearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin ImpulseButton.ISButton CmdShowMoreOptions 
      Height          =   375
      Left            =   5460
      TabIndex        =   8
      Top             =   3840
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "»ÕÀ „ ﬁœ„..."
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
      ButtonImage     =   "FrmBuySearch.frx":030A
      ColorButton     =   14871017
      ColorHoverText  =   12582912
      ButtonToggles   =   1
      DrawFocusRectangle=   0   'False
      RightToLeft     =   -1  'True
      ButtonImageToggled=   "FrmBuySearch.frx":06A4
      ColorToggledHoverText=   192
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E0E0E0&
      Caption         =   "«·›« Ê—…  Õ ÊÏ ⁄·Ï Â–« «·’‰›"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1395
      Index           =   1
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   4620
      Width           =   6765
      Begin VB.TextBox TxtItemCode 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   1275
      End
      Begin VB.CheckBox ChkSerialSearchType 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "»ÕÀ „ÿ«»ﬁ"
         Height          =   285
         Left            =   1050
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   990
         Width           =   1455
      End
      Begin VB.TextBox TxtItemSerial 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2550
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   960
         Width           =   3315
      End
      Begin VB.TextBox TxtItemPrice 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2580
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   570
         Width           =   1005
      End
      Begin VB.TextBox TxtItemQty 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   1275
      End
      Begin MSDataListLib.DataCombo DCboItem 
         Height          =   315
         Left            =   540
         TabIndex        =   10
         Top             =   240
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton CmdItemSearch 
         Height          =   345
         Left            =   90
         TabIndex        =   36
         Top             =   210
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmBuySearch.frx":0A3E
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "ﬂÊœ «·’‰›"
         Height          =   345
         Index           =   6
         Left            =   150
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   690
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "”⁄— «·’‰›"
         Height          =   315
         Index           =   5
         Left            =   3660
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   615
         Width           =   825
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "”Ì—Ì«· "
         Height          =   315
         Index           =   4
         Left            =   5940
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   1020
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "ﬂ„Ì… «·’‰›"
         Height          =   315
         Index           =   3
         Left            =   5850
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   645
         Width           =   825
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "«”„ «·’‰›"
         Height          =   315
         Index           =   2
         Left            =   5820
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   270
         Width           =   825
      End
   End
   Begin VB.ComboBox CboPaymentType 
      Height          =   315
      Left            =   60
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2370
      Width           =   2085
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ì «·› —…"
      ForeColor       =   &H00FF0000&
      Height          =   1065
      Index           =   0
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   2730
      Width           =   2085
      Begin MSComCtl2.DTPicker DTPFrom 
         Height          =   345
         Left            =   60
         TabIndex        =   6
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/m/yyyy"
         DateIsNull      =   -1  'True
         Format          =   60096513
         CurrentDate     =   38979.743287037
      End
      Begin MSComCtl2.DTPicker DTPTo 
         Height          =   375
         Left            =   60
         TabIndex        =   7
         Top             =   630
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   60096513
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   11
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   255
         Width           =   285
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   285
         Index           =   10
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   675
         Width           =   345
      End
   End
   Begin VB.TextBox TxtVal 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   3840
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox XPTxtClientsName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      Height          =   315
      Left            =   60
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1980
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.TextBox XPTxtBillNum 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4500
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2370
      Width           =   1185
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2325
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6885
      _cx             =   12144
      _cy             =   4101
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmBuySearch.frx":0FD8
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
   Begin VB.CheckBox XPChkSearchType 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·⁄„Ì· »«·ﬂ«„· ›ﬁÿ"
      Height          =   225
      Left            =   870
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   4290
      Visible         =   0   'False
      Width           =   795
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1710
      TabIndex        =   16
      Top             =   3870
      Width           =   765
      _ExtentX        =   1349
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
      Left            =   855
      TabIndex        =   17
      Top             =   3870
      Width           =   795
      _ExtentX        =   1402
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
      Left            =   60
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3870
      Width           =   735
      _ExtentX        =   1296
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
   Begin MSDataListLib.DataCombo DCboClientsName 
      Height          =   315
      Left            =   3180
      TabIndex        =   2
      Top             =   2730
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboStores 
      Height          =   315
      Left            =   3180
      TabIndex        =   3
      Top             =   3090
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboUsers 
      Height          =   315
      Left            =   3180
      TabIndex        =   4
      Top             =   3450
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„” Œœ„"
      Height          =   315
      Index           =   7
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   3450
      Width           =   1035
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰ ÌÃ… «·»ÕÀ:"
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   4500
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   4290
      Width           =   2325
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
      Height          =   285
      Index           =   5
      Left            =   2190
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   2400
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„Œ“‰"
      Height          =   315
      Index           =   0
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   3090
      Width           =   1065
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·⁄—÷"
      Height          =   315
      Index           =   4
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   2730
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·⁄—÷"
      Height          =   315
      Index           =   3
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   2370
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬁÌ„… «·⁄—÷"
      Height          =   315
      Index           =   2
      Left            =   4350
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   3870
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·›« Ê—…"
      Height          =   315
      Index           =   1
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   2370
      Width           =   1065
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·⁄„Ì·"
      Height          =   315
      Index           =   0
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   2730
      Width           =   1065
   End
End
Attribute VB_Name = "FrmBuySearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As ADODB.Recordset
Dim cSearchDcbo(3) As clsDCboSearch
Private m_DealingForm As GridTransType
Dim M_ExtraRetrunObject As Object
Public RetrunFrm As Form

Private Sub Cmd_Click(Index As Integer)
Dim Msg As String
On Error GoTo ErrTrap
Select Case Index
    Case 0
       If Rs.State = adStateOpen Then
           Rs.Close
        End If
       Rs.Open Build_Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
       If Rs.RecordCount < 1 Then
           Fg.Clear flexClearScrollable, flexClearEverything
           Fg.Rows = 2
           Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
           MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
           Exit Sub
       End If
       Me.lbl(1).Caption = "‰ ÌÃ… «·»ÕÀ: " & Rs.RecordCount
       
       Retrive
    Case 1
        clear_all Me
        Fg.Clear flexClearScrollable, flexClearEverything
        Fg.Rows = 2
    Case 2
         Unload Me
End Select
Exit Sub
ErrTrap:
    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If
End Sub



Private Sub CmdItemSearch_Click()
Load FrmItemSearch
FrmItemSearch.RetrunType = 1
Set FrmItemSearch.DcboItems = Me.DCboItem
PutFormOnTop Me.hwnd, False
FrmItemSearch.Show vbModal
PutFormOnTop Me.hwnd, True
End Sub

Private Sub CmdShowMoreOptions_Click()
If CmdShowMoreOptions.Value = True Then
    Me.Fra(1).Visible = True
    'Me.Height = Me.Fra(1).top + Fra(1).Height + 400
    Me.Height = Me.Fra(1).top + Fra(1).Height + 500 ' GetMyTitleBarHight(Me.hwnd)
    'Me.Height = Me.ScaleHeight
Else
    Me.Fra(1).Visible = False
    Me.Height = Me.Fra(1).top + 500
End If
End Sub

Private Sub FG_Click()
Dim StrSQL As String
Dim Num As Integer
Dim RowNum As Integer
Dim StrQry As String
Dim RsDetails As ADODB.Recordset
Dim DateTemp As Date
Dim Msg As String

On Error GoTo ErrTrap
 
If Not Fg.TextMatrix(Fg.Row, 1) = "" Then
    Select Case Me.DealingForm
        Case PurchaseTransaction
            If Me.ExtraRetrunObject Is Nothing Then
                RetrunFrm.Retrive Val(Fg.TextMatrix(Fg.Row, 1))
            Else
                Me.ExtraRetrunObject = Val(Fg.TextMatrix(Fg.Row, 1))
            End If
        Case InvoiceTransaction
            If Me.ExtraRetrunObject Is Nothing Then
                Me.RetrunFrm.Retrive Val(Fg.TextMatrix(Fg.Row, 1))
            Else
                Me.ExtraRetrunObject = Val(Fg.TextMatrix(Fg.Row, 1))
            '.Fg.TextMatrix(Num, .Fg.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemID")), "", Trim(RsDetails("ItemID").Value))
    FrmDiscounts.DBCboClientName.text = Fg.TextMatrix(Fg.Row, 4)
            End If
        Case ReturnTransaction     '  "xxx"
            If Me.ExtraRetrunObject Is Nothing Then
                FrmReturnpurchases.Retrive Val(Fg.TextMatrix(Fg.Row, 1))
            Else
                Me.ExtraRetrunObject = Val(Fg.TextMatrix(Fg.Row, 1))
            End If
        Case ShowPrice         '"xxxx"
            FrmShowPrice.Retrive Val(Fg.TextMatrix(Fg.Row, 1))
            '«·»ÕÀ ⁄‰ «·⁄—Ê÷ «·Ã«Â“…
        Case Template
            FrmTemplate.Retrive Val(Fg.TextMatrix(Fg.Row, 1))
            '«·»ÕÀ ⁄‰ «·«Â·«ﬂ« 
         Case Destruction
            FrmDestruction.Retrive Val(Fg.TextMatrix(Fg.Row, 1))
           '«·»ÕÀ ⁄‰ „— Ã⁄ «·„»Ì⁄« 
         Case ReturnSalling
             If Me.ExtraRetrunObject Is Nothing Then
                FrmReturnSalling.Retrive Val(Fg.TextMatrix(Fg.Row, 1))
            Else
                Me.ExtraRetrunObject = Val(Fg.TextMatrix(Fg.Row, 1))
            End If
'        Case "ZZZ"
'            FrmMoving.Retrive Val(Fg.TextMatrix(Fg.Row, 1))
        Case InsertTemplate
            If Me.Fg.TextMatrix(Fg.Row, Fg.ColIndex("Transaction_Serial")) <> "" Then
                DateTemp = CDate(Me.Fg.TextMatrix(Fg.Row, Fg.ColIndex("Transaction_Serial")))
                If DateDiff("d", Date, DateTemp) < 0 Then
                    Msg = "·ﬁœ ≈‰ ÂÌ  › —… Â–Â «·⁄—÷ ...!!!"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If
            With FrmShowPrice
                .Fg.Rows = 2
                .Fg.Clear flexClearScrollable, flexClearEverything
                .Fg.Refresh
                StrSQL = "SELECT Templates.TemplateID, Template_Details.ItemID, " & _
                "Template_Details.Quantity, Template_Details.Price, Template_Details.ItemDiscountType, " & _
                "Template_Details.ItemDiscount, Template_Details.ItemCase, TblItems.HaveSerial " & _
                "FROM TblItems INNER JOIN (Templates INNER JOIN Template_Details ON " & _
                "Templates.TemplateID = Template_Details.TemplateID) ON TblItems.ItemID = " & _
                "Template_Details.ItemID"
                StrSQL = StrSQL + " where Templates.TemplateID=" & Val(Fg.TextMatrix(Fg.Row, 1))
                Set RsDetails = New ADODB.Recordset
                RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                .XPTxtSum.text = ""
                If Not (RsDetails.EOF Or RsDetails.BOF) Then
                    .Fg.Rows = RsDetails.RecordCount + 1
                    For Num = 1 To RsDetails.RecordCount
                        .Fg.TextMatrix(Num, .Fg.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemID")), "", (RsDetails("ItemID").Value))
                        .Fg.TextMatrix(Num, .Fg.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemID")), "", Trim(RsDetails("ItemID").Value))
                        .Fg.TextMatrix(Num, .Fg.ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", (RsDetails("Quantity").Value))
                        .Fg.TextMatrix(Num, .Fg.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").Value))
                        .Fg.TextMatrix(Num, .Fg.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").Value))
                        .Fg.TextMatrix(Num, .Fg.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").Value))
                        .Fg.TextMatrix(Num, .Fg.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").Value))
                        If RsDetails("HaveSerial") = True Then
                            .Fg.TextMatrix(Num, .Fg.ColIndex("HaveSerial")) = True
                        End If
                        RsDetails.MoveNext
                    Next Num
                End If
                .Cala
            End With
        Case InsertTemplateToInvoice
            With FrmSaleBill
                .Fg.Rows = 2
                .Fg.Clear flexClearScrollable, flexClearEverything
                .Fg.Refresh
                StrSQL = "SELECT Templates.TemplateID, Template_Details.ItemID, " & _
                "Template_Details.Quantity, Template_Details.Price, Template_Details.ItemDiscountType, " & _
                "Template_Details.ItemDiscount, Template_Details.ItemCase, TblItems.HaveSerial " & _
                "FROM TblItems INNER JOIN (Templates INNER JOIN Template_Details ON " & _
                "Templates.TemplateID = Template_Details.TemplateID) ON TblItems.ItemID = " & _
                "Template_Details.ItemID"
                StrSQL = StrSQL + " where Templates.TemplateID=" & Val(Fg.TextMatrix(Fg.Row, 1))
                Set RsDetails = New ADODB.Recordset
                RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                .XPTxtSum.text = ""
                If Not (RsDetails.EOF Or RsDetails.BOF) Then
                    .Fg.Rows = RsDetails.RecordCount + 1
                    For Num = 1 To RsDetails.RecordCount
                        .Fg.TextMatrix(Num, .Fg.ColIndex("Code")) = IIf(IsNull(RsDetails("ItemID")), "", (RsDetails("ItemID").Value))
                        .Fg.TextMatrix(Num, .Fg.ColIndex("Name")) = IIf(IsNull(RsDetails("ItemID")), "", Trim(RsDetails("ItemID").Value))
                        .Fg.TextMatrix(Num, .Fg.ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", (RsDetails("Quantity").Value))
                        .Fg.TextMatrix(Num, .Fg.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").Value))
                        .Fg.TextMatrix(Num, .Fg.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").Value))
                        .Fg.TextMatrix(Num, .Fg.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").Value))
                        .Fg.TextMatrix(Num, .Fg.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").Value))
                        If RsDetails("HaveSerial") = True Then
                            .Fg.TextMatrix(Num, .Fg.ColIndex("HaveSerial")) = True
                        End If
                        RsDetails.MoveNext
                    Next Num
                End If
                .Cala
            End With
    End Select

End If
Exit Sub
ErrTrap:
End Sub
Private Sub Fg_DblClick()
FG_Click
Cmd_Click (2)
End Sub

Private Sub Fg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    If Shift = ShiftConstants.vbCtrlMask Then
            Me.Fg.ColHidden(Fg.ColIndex("Transaction_ID")) = Not _
            Me.Fg.ColHidden(Fg.ColIndex("Transaction_ID"))
    End If
End If
End Sub

Private Sub Form_Activate()
On Error GoTo ErrTrap
Dim StrSQL As String
Dim Dcombos As New ClsDataCombos
PutFormOnTop Me.hwnd
If Me.DealingForm = ReturnTransaction Then
    Me.Caption = "«·»ÕÀ ⁄‰ „— Ã⁄ «·„‘ —Ì« "
    Fg.TextMatrix(0, Fg.ColIndex("Transaction_ID")) = "—ﬁ„ «·»—‰«„Ã"
    XPLbl(1).Caption = "—ﬁ„ «·⁄„·Ì…"
    XPLbl(0).Caption = "«”„ «·„Ê—œ"
    XPChkSearchType.Caption = "«”„ «·„Ê—œ »«·ﬂ«„· ›ﬁÿ"
    Dcombos.GetCustomersSuppliers 0, DCboClientsName
    Me.XPLbl(5).Visible = True
    Me.CboPayMentType.Visible = True
    '«·⁄—Ê÷ «·Ã«Â“…
ElseIf Me.DealingForm = Template Then
    Fg.TextMatrix(0, Fg.ColIndex("Transaction_ID")) = "ﬂÊœ «·⁄—÷"
    StrSQL = "SELECT * FROM Templates"
    fill_combo DCboClientsName, StrSQL
    Me.DcboStores.Visible = False
    lbl(0).Visible = False
    CmdShowMoreOptions.Enabled = False
    CboPayMentType.Visible = False
    
ElseIf Me.DealingForm = InsertTemplate Or _
    Me.DealingForm = InsertTemplateToInvoice Then
    '«·⁄—Ê÷ «·Ã«Â“…
    Fg.TextMatrix(0, Fg.ColIndex("Transaction_ID")) = "ﬂÊœ «·⁄—÷"
    Fg.TextMatrix(0, Fg.ColIndex("BillDate")) = "«”„ «·⁄—÷"
    Fg.TextMatrix(0, Fg.ColIndex("ClientNmae")) = " «—ÌŒ «·⁄—÷"
    Fg.TextMatrix(0, Fg.ColIndex("StorName")) = "ﬁÌ„… «·⁄—÷"
    Fg.TextMatrix(0, Fg.ColIndex("Transaction_Serial")) = " «—ÌŒ ≈‰ Â«¡ «·⁄—÷"
    
    XPChkSearchType.Visible = False
    TxtVal.Visible = True
    XPLbl(2).Visible = True
    XPLbl(1).Visible = False
    XPLbl(0).Visible = False
    XPLbl(3).Visible = True
    XPLbl(4).Visible = True
    '⁄—Ê÷ «·√”⁄«—
ElseIf Me.DealingForm = ShowPrice Then
    '«·»—‰«„Ã
    Fg.TextMatrix(0, Fg.ColIndex("Transaction_ID")) = "—ﬁ„ «·»—‰«„Ã"
    Fg.TextMatrix(0, Fg.ColIndex("Transaction_Serial")) = "ﬂÊœ «·⁄—÷"
    Fg.TextMatrix(0, Fg.ColIndex("BillDate")) = " «—ÌŒ «·⁄—÷"
    XPLbl(1).Caption = "—ﬁ„ «·⁄—÷"
    'XPLbl(0).Caption = "«”„ «·⁄„Ì·"
    Dcombos.GetCustomersSuppliers 0, DCboClientsName
    '«· ·›Ì« 
ElseIf Me.DealingForm = Destruction Then
    Fg.TextMatrix(0, Fg.ColIndex("Transaction_ID")) = "—ﬁ„ «·⁄„·Ì…"
    Fg.TextMatrix(0, Fg.ColIndex("BillDate")) = " «—ÌŒ «·⁄„·Ì…"
    XPLbl(1).Caption = "—ﬁ„ «·⁄„·Ì…"
    XPLbl(0).Caption = "«”„ «·„Œ“‰"
    XPChkSearchType.Visible = False
    StrSQL = "SELECT * From TblStore"
    fill_combo DCboClientsName, StrSQL
Else
    Fg.TextMatrix(0, Fg.ColIndex("Transaction_ID")) = "—ﬁ„ «·»—‰«„Ã"
    Fg.TextMatrix(0, Fg.ColIndex("Transaction_Serial")) = "—ﬁ„ «·›« Ê—…"
    XPLbl(1).Caption = "—ﬁ„ «·›« Ê—…"
    XPLbl(0).Caption = "«”„ «·⁄„Ì·"
    XPChkSearchType.Caption = "«”„ «·⁄„Ì· »«·ﬂ«„· ›ﬁÿ"
    Dcombos.GetCustomersSuppliers 0, DCboClientsName
    Me.XPLbl(5).Visible = True
    Me.CboPayMentType.Visible = True
End If
If Me.DealingForm = InsertTemplate Or Me.DealingForm = InsertTemplateToInvoice Then
    Cmd_Click (0)
End If
'StrSql = "SELECT * From TblCustemers where type=1"
'fill_combo DCboCustemerName, StrSql
Set cSearchDcbo(0) = New clsDCboSearch
Set cSearchDcbo(0).Client = Me.DCboClientsName
Dcombos.GetStores Me.DcboStores
Dcombos.GetUsers Me.DcboUsers
Set cSearchDcbo(3) = New clsDCboSearch
Set cSearchDcbo(3).Client = Me.DcboUsers
Exit Sub
ErrTrap:
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrTrap
If KeyCode = vbKeyReturn Then
    If Fg.TextMatrix(Fg.Row, Fg.ColIndex("Transaction_ID")) <> "" And Me.ActiveControl Is Fg Then
        FG_Click
    ElseIf Shift = vbCtrlMask Then
        Cmd_Click (0)
    Else
        SendKeys "{TAB}"
    End If
End If
Exit Sub
ErrTrap:
End Sub
Private Sub Form_Load()
On Error GoTo ErrTrap
Dim BG As New ClsBackGroundPic
Dim Dcombos As New ClsDataCombos

Set Rs = New ADODB.Recordset
Set Cmd(0).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Search").Picture
Set Cmd(1).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Clear").Picture
Set Cmd(2).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Exit").Picture
CenterForm Me
FormPostion Me, GetPostion
Fg.WallPaper = BG.SearchWallpaper
Dcombos.GetItemsNames Me.DCboItem
Set cSearchDcbo(1) = New clsDCboSearch
Set cSearchDcbo(1).Client = Me.DCboItem
Fg.ColFormat(Fg.ColIndex("BillDate")) = "Medium Date"
With Me.CboPayMentType
    .Clear
    .AddItem "‰ﬁœÌ"
    .AddItem "«Ã·"
    .AddItem "«·ﬂ·"
End With
CmdShowMoreOptions.Value = False
CmdShowMoreOptions_Click
SetDtpickerDate Me.DTPFrom
SetDtpickerDate Me.DTPTo
Exit Sub
ErrTrap:
End Sub
Private Sub Retrive()
Dim Num As Integer
On Error GoTo ErrTrap
Fg.Clear flexClearScrollable, flexClearEverything
If Me.DealingForm = InvoiceTransaction Or Me.DealingForm = PurchaseTransaction Then
    Set Me.Fg.DataSource = Rs
    Fg.AutoSize 0, Fg.Cols - 1, False
ElseIf Me.DealingForm = Template Or Me.DealingForm = InsertTemplate Or Me.DealingForm = InsertTemplateToInvoice Then
    If Not (Rs.EOF Or Rs.BOF) Then
        Fg.Rows = Rs.RecordCount + 1
        For Num = 1 To Rs.RecordCount
            With Fg
                .TextMatrix(Num, .ColIndex("Count")) = Num
                .TextMatrix(Num, .ColIndex("Transaction_ID")) = IIf(IsNull(Rs("TemplateID").Value), "", Val(Rs("TemplateID").Value))
                .TextMatrix(Num, .ColIndex("BillDate")) = IIf(IsNull(Rs("TemplateName").Value), "", (Rs("TemplateName").Value))
                .TextMatrix(Num, .ColIndex("ClientNmae")) = IIf(IsNull(Rs("Date").Value), "", Format(Rs("Date").Value, "yyyy/m/d"))
                .TextMatrix(Num, .ColIndex("StorName")) = IIf(IsNull(Rs("Summition").Value), "", Trim(Rs("Summition").Value))
                If Not IsNull(Rs("TemplateTime").Value) Then
                    .TextMatrix(Num, .ColIndex("Transaction_Serial")) = DisplayDate(Rs("TemplateTime").Value)
                End If
            End With
            Rs.MoveNext
        Next Num
    End If
ElseIf Not (Rs.EOF Or Rs.BOF) Then
    Fg.Rows = Rs.RecordCount + 1
    For Num = 1 To Rs.RecordCount
        With Fg
            .TextMatrix(Num, .ColIndex("Count")) = Num
            .TextMatrix(Num, .ColIndex("Transaction_ID")) = IIf(IsNull(Rs("Transaction_ID").Value), "", Val(Rs("Transaction_ID").Value))
            .TextMatrix(Num, .ColIndex("Transaction_Serial")) = IIf(IsNull(Rs("Transaction_Serial").Value), "", Rs("Transaction_Serial").Value)
            If Not IsNull(Rs("Transaction_Date").Value) Then
                .TextMatrix(Num, .ColIndex("BillDate")) = Rs("Transaction_Date").Value
            Else
                .TextMatrix(Num, .ColIndex("BillDate")) = ""
            End If
            .TextMatrix(Num, .ColIndex("ClientNmae")) = IIf(IsNull(Rs("CusName").Value), "", Trim(Rs("CusName").Value))
            .TextMatrix(Num, .ColIndex("StorName")) = IIf(IsNull(Rs("StoreName").Value), "", Trim(Rs("StoreName").Value))
    End With
    Rs.MoveNext
    Next Num
    Fg.AutoSize 0, Fg.Cols - 1, False
End If
Fg.SetFocus
Exit Sub
ErrTrap:
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrTrap
Dim I As Integer

If Rs.State = adStateOpen Then
    Rs.Close
    Set Rs = Nothing
End If
For I = LBound(cSearchDcbo) To UBound(cSearchDcbo)
    Set cSearchDcbo(I) = Nothing
Next I
FormPostion Me, SavePostion
Set Me.ExtraRetrunObject = Nothing

Exit Sub
ErrTrap:
End Sub
Private Function Build_Sql() As String
On Error GoTo ErrTrap
Dim StrSQL As String
Dim MySQL As String
Dim m_SearchFrom As GridTransType
Dim Begin As Boolean
Dim StrWhere As String
If SystemOptions.SysDataBaseType = AccessDataBase Then
    MySQL = "SELECT DISTINCT Transactions.Transaction_ID,Transactions.Transaction_Serial," & _
    "Transactions.Transaction_Date,TblCustemers.CusName, TblStore.StoreName "
    MySQL = MySQL + " FROM (TblStore RIGHT JOIN (TblCustemers RIGHT JOIN Transactions " & _
    "ON TblCustemers.CusID=Transactions.CusID) ON TblStore.StoreID=Transactions.StoreID)" & _
    "INNER JOIN Transaction_Details ON Transactions.Transaction_ID=Transaction_Details.Transaction_ID "
ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
    MySQL = "SELECT DISTINCT Transactions.Transaction_ID,Transactions.Transaction_Serial," & _
    "convert(nvarchar(50),Transactions.Transaction_Date,111)as Transaction_Date,TblCustemers.CusName, TblStore.StoreName "
    MySQL = MySQL + " FROM (TblStore RIGHT JOIN (TblCustemers RIGHT JOIN Transactions " & _
    "ON TblCustemers.CusID=Transactions.CusID) ON TblStore.StoreID=Transactions.StoreID)" & _
    "INNER JOIN Transaction_Details ON Transactions.Transaction_ID=Transaction_Details.Transaction_ID "
End If

Dim rsOut As New ADODB.Recordset
'Dim Msg As String
Set rsOut = New ADODB.Recordset
rsOut.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
If Not (rsOut.EOF Or rsOut.BOF) Then
'  FrmOut.Show
  End If
'End If

m_SearchFrom = Me.DealingForm
Select Case m_SearchFrom
    Case PurchaseTransaction
     If rsOut!checkinpo = True Then
        StrSQL = MySQL + " WHERE Transaction_Type=22"
    Else
       StrSQL = MySQL + " WHERE Transaction_Type=1"
    End If
      
        Begin = True
    Case InvoiceTransaction
    If rsOut!checkout = True Then

        StrSQL = MySQL + " WHERE Transaction_Type=21"
    Else
        StrSQL = MySQL + " WHERE Transaction_Type=2"
    End If
        Begin = True
    Case ReturnTransaction
        StrSQL = MySQL + " WHERE Transaction_Type=5"
        Begin = True
     Case ShowPrice
        StrSQL = MySQL + " WHERE Transaction_Type=6"
        Begin = True
        '«· ·›Ì« 
    Case Destruction
        StrSQL = MySQL + " WHERE Transaction_Type=8"
        Begin = True
    Case ReturnSalling
        StrSQL = MySQL + " WHERE Transaction_Type=9"
        Begin = True
'    Case "ZZZ"  '«· ÕÊÌ· „‰ „Œ“‰ ≈·Ï „Œ“‰
'        StrSql = "select * From QRyTransSearch WHERE Transaction_Type=10"
        '«·⁄—Ê÷ «·Ã«Â“…
    Case Template, InsertTemplate, InsertTemplateToInvoice
        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "select * From TemplateSearch"
            Begin = False
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = "SELECT TemplateSearch.* FROM dbo.TemplateSearch() TemplateSearch"
            Begin = False
        End If
End Select

If m_SearchFrom = Template Or m_SearchFrom = InsertTemplate Or m_SearchFrom = InsertTemplateToInvoice Then
    If XPTxtBillNum.text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and TemplateID=" & (XPTxtBillNum.text)
        Else
            StrWhere = StrWhere + " where TemplateID=" & (XPTxtBillNum.text)
            Begin = True
        End If
    End If
    If DCboClientsName.BoundText <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and TemplateID =" & Trim(DCboClientsName.BoundText)
        Else
            StrWhere = StrWhere + " where TemplateID =" & Trim(DCboClientsName.BoundText)
            Begin = True
        End If
    End If
    If TxtVal.text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and Summition =" & TxtVal.text
        Else
            StrWhere = StrWhere + " where Summition=" & TxtVal.text
            Begin = True
        End If
    End If
    If Not IsNull(Me.DTPFrom.Value) Then
        If Begin = True Then
            StrWhere = StrWhere + " and [Date] >=" & SQLDate(Me.DTPFrom.Value, True) & ""
        Else
            StrWhere = StrWhere + " where [Date] >=" & SQLDate(Me.DTPFrom.Value, True) & ""
            Begin = True
        End If
    End If
    If Not IsNull(Me.DTPTo.Value) Then
        If Begin = True Then
            StrWhere = StrWhere + " and [Date] <=" & SQLDate(Me.DTPTo.Value, True) & ""
        Else
            StrWhere = StrWhere + " where [Date] <=" & SQLDate(Me.DTPTo.Value, True) & ""
            Begin = True
        End If
    End If
    Build_Sql = StrSQL + StrWhere + " order by TemplateID"
ElseIf m_SearchFrom = Destruction Then '«· ·›Ì« 
    If XPTxtBillNum.text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and Transaction_Serial ='" & (XPTxtBillNum.text) & "'"
        Else
            StrWhere = StrWhere + " where Transaction_Serial ='" & (XPTxtBillNum.text) & "'"
            Begin = True
        End If
    End If
    If DCboClientsName.BoundText <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and StoreID =" & Trim(DCboClientsName.BoundText)
        Else
            StrWhere = StrWhere + " where StoreID =" & Trim(DCboClientsName.BoundText)
            Begin = True
        End If
    End If
    Build_Sql = StrSQL + StrWhere + " Order by Transactions.Transaction_ID"
Else
    '---------------------------------
    If XPTxtBillNum.text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and Transaction_Serial ='" & (XPTxtBillNum.text) & "'"
        Else
            StrWhere = StrWhere + " where Transaction_Serial ='" & (XPTxtBillNum.text) & "'"
            Begin = True
        End If
    End If
    If Me.CboPayMentType.ListIndex <> -1 Then
        If Me.CboPayMentType.ListIndex = 0 Then
            StrWhere = StrWhere + " and PaymentType=0 "
        ElseIf Me.CboPayMentType.ListIndex = 1 Then
            StrWhere = StrWhere + " and PaymentType=1"
        End If
    End If
    If DCboClientsName.BoundText <> "" Then
        If XPChkSearchType.Value = Checked Then
            If Begin = True Then
                StrWhere = StrWhere + " and CusID =" & Trim(DCboClientsName.BoundText)
            Else
                StrWhere = StrWhere + " where CusID =" & Trim(DCboClientsName.BoundText)
                Begin = True
            End If
        Else
            If Begin = True Then
                StrWhere = StrWhere + " and CusName LIKE'" & Trim(DCboClientsName.text) & "%'"
            Else
                StrWhere = StrWhere + " where CusName LIKE'" & Trim(DCboClientsName.text) & "%'"
                Begin = True
            End If
        End If
    End If
    If Me.DcboUsers.BoundText <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and UserID =" & Me.DcboUsers.BoundText & ""
        Else
            StrWhere = StrWhere + " where UserID =" & Me.DcboUsers.BoundText & ""
            Begin = True
        End If
    End If
    If Not IsNull(Me.DTPFrom.Value) Then
        If Begin = True Then
            StrWhere = StrWhere + " and Transaction_date >=" & SQLDate(Me.DTPFrom.Value, True) & ""
        Else
            StrWhere = StrWhere + " where Transaction_date >=" & SQLDate(Me.DTPFrom.Value, True) & ""
            Begin = True
        End If
    End If
    If Not IsNull(Me.DTPTo.Value) Then
        If Begin = True Then
            StrWhere = StrWhere + " and Transaction_date <=" & SQLDate(Me.DTPTo.Value, True) & ""
        Else
            StrWhere = StrWhere + " where Transaction_date <=" & SQLDate(Me.DTPTo.Value, True) & ""
            Begin = True
        End If
    End If
    If Me.DcboStores.BoundText <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and Transactions.StoreID=" & Me.DcboStores.BoundText & ""
        Else
            StrWhere = StrWhere + " where Transactions.StoreID=" & Me.DcboStores.BoundText & ""
            Begin = True
        End If
    End If
    If Me.DCboItem.BoundText <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and Transaction_Details.Item_ID=" & Me.DCboItem.BoundText & ""
        Else
            StrWhere = StrWhere + " where Transaction_Details.Item_ID=" & Me.DCboItem.BoundText & ""
            Begin = True
        End If
    End If
    If Val(TxtItemQty.text) > 0 Then
        If Begin = True Then
            StrWhere = StrWhere + " and Transaction_Details.Quantity=" & Val(TxtItemQty.text) & ""
        Else
            StrWhere = StrWhere + " where Transaction_Details.Quantity=" & Val(TxtItemQty.text) & ""
            Begin = True
        End If
    End If
    If Val(TxtItemPrice.text) > 0 Then
        If Begin = True Then
            StrWhere = StrWhere + " and Transaction_Details.Price=" & Val(TxtItemPrice.text) & ""
        Else
            StrWhere = StrWhere + " where Transaction_Details.Price=" & Val(TxtItemPrice.text) & ""
            Begin = True
        End If
    End If
    If Trim(Me.TxtItemSerial.text) <> "" Then
        If ChkSerialSearchType.Value = vbChecked Then
            If Begin = True Then
                StrWhere = StrWhere + " and Transaction_Details.ItemSerial='" & Trim(TxtItemSerial.text) & "'"
            Else
                StrWhere = StrWhere + " where Transaction_Details.ItemSerial='" & Trim(TxtItemSerial.text) & "'"
                Begin = True
            End If
        ElseIf ChkSerialSearchType.Value = vbUnchecked Then
             If Begin = True Then
                StrWhere = StrWhere + " and Transaction_Details.ItemSerial like '%" & Trim(TxtItemSerial.text) & "%'"
            Else
                StrWhere = StrWhere + " where Transaction_Details.ItemSerial like '%" & Trim(TxtItemSerial.text) & "%'"
                Begin = True
            End If
        End If
    End If
    Build_Sql = StrSQL + StrWhere + " order by Transactions.Transaction_ID DESC"
End If
Exit Function
ErrTrap:
End Function
Public Property Get DealingForm() As GridTransType
DealingForm = m_DealingForm
End Property
Public Property Let DealingForm(ByVal vNewValue As GridTransType)
'If vNewValue = OpeningBalance Or vNewValue = PurchaseTransaction Or vNewValue = InvoiceTransaction Then
    m_DealingForm = vNewValue
'End If
End Property
Public Property Get ExtraRetrunObject() As Object
Set ExtraRetrunObject = M_ExtraRetrunObject
End Property

Public Property Set ExtraRetrunObject(ByVal vNewValue As Object)
'ﬁ„  »⁄„· Â–Â «·Œ«’Ì… „Œ’Ê’ Õ Ï Ì„ﬂ‰‰Ï
'«‰ «” Œœ„ ‘«‘… «·»ÕÀ ⁄‰ «·Õ—ﬂ«  «· Ã«—Ì…
'„‰ Œ·«· ‘«‘… «·„ﬁ»Ê÷«  ÕÌÀ Ì„ﬂ‰‰Ï
'«‰ «” —Ã⁄ ﬂÊœ «·Õ—ﬂ… «· Ã«—Ì…
'›Ï ‘«‘… „À· ‘«‘… «·„ﬁ»Ê÷« 
Set M_ExtraRetrunObject = vNewValue
End Property

Private Sub TxtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim StrSQL As String
Dim Rs As ADODB.Recordset

If KeyCode = vbKeyReturn Then
    If Trim(Me.TxtItemCode.text) <> "" Then
        If Trim(Me.TxtItemCode.text) = "" Then Exit Sub
        StrSQL = "Select ItemID From TblItems Where ItemCode='" & Trim(Me.TxtItemCode.text) & "'"
        Set Rs = New ADODB.Recordset
        Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        If Not (Rs.BOF Or Rs.EOF) Then
            DCboItem.BoundText = Rs("ItemID").Value
        Else
            'Msg = "·«ÌÊÃœ ’‰› „”Ã· »Â–« «·ﬂÊœ..!"
            'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        End If
    End If
End If

End Sub


Private Sub TxtItemPrice_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtItemPrice.text, 0)
End Sub


Private Sub TxtItemQty_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtItemQty.text, 0)
End Sub



