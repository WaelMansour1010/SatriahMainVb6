VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form ShipmentOrderSearch 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11880
   Icon            =   "ShipmentOrderSearch.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   11880
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   8610
      TabIndex        =   34
      Top             =   5760
      Width           =   1845
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   8640
      TabIndex        =   33
      Top             =   6120
      Width           =   1845
   End
   Begin VB.TextBox TxtStoreID 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   8610
      TabIndex        =   28
      Top             =   5415
      Width           =   1845
   End
   Begin VB.TextBox TxtAddress 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3000
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Top             =   4560
      Width           =   4785
   End
   Begin VB.TextBox TxtContactPhone 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   8640
      TabIndex        =   23
      Top             =   4560
      Width           =   1830
   End
   Begin VB.TextBox TxtEmployeeID 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   8640
      TabIndex        =   21
      Top             =   5040
      Width           =   1830
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   360
      TabIndex        =   13
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
      MICON           =   "ShipmentOrderSearch.frx":000C
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
      Left            =   14160
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   4680
      Width           =   7830
   End
   Begin VB.TextBox TxtCusID 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   4080
      Width           =   1830
   End
   Begin VB.TextBox txtorder_no 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   8640
      TabIndex        =   6
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   11745
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
               Picture         =   "ShipmentOrderSearch.frx":0028
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ShipmentOrderSearch.frx":03C2
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ShipmentOrderSearch.frx":075C
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ShipmentOrderSearch.frx":0AF6
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ShipmentOrderSearch.frx":0E90
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ShipmentOrderSearch.frx":122A
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ShipmentOrderSearch.frx":15C4
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ShipmentOrderSearch.frx":1B5E
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ ÿ·» ‘Õ‰"
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
         Left            =   6135
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   90
         Width           =   5280
      End
   End
   Begin MSDataListLib.DataCombo DBCboClientName 
      Height          =   315
      Left            =   3000
      TabIndex        =   8
      Top             =   4080
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker XPDtbBill 
      Height          =   315
      Left            =   6000
      TabIndex        =   9
      Top             =   3600
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   96731137
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   315
      Left            =   10080
      TabIndex        =   10
      Top             =   9360
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
      TabIndex        =   14
      Top             =   720
      Width           =   11835
      _cx             =   20876
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
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"ShipmentOrderSearch.frx":1EF8
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
      Left            =   2040
      TabIndex        =   15
      Top             =   6240
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
      Left            =   1020
      TabIndex        =   16
      Top             =   6240
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
      Left            =   120
      TabIndex        =   17
      Top             =   6240
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
      TabIndex        =   18
      Top             =   9090
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
      ButtonImage     =   "ShipmentOrderSearch.frx":20EE
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin MSDataListLib.DataCombo DcboEmp 
      Height          =   315
      Left            =   3000
      TabIndex        =   22
      Top             =   5040
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "7"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCboStoreName 
      Height          =   315
      Left            =   3000
      TabIndex        =   29
      Top             =   5415
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "7"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo dcBranch 
      Height          =   315
      Left            =   3000
      TabIndex        =   30
      Top             =   6120
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCboStoreName1 
      Height          =   315
      Left            =   3000
      TabIndex        =   35
      Top             =   5760
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "7"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "„‰ „Œ“‰"
      Height          =   270
      Index           =   0
      Left            =   10380
      TabIndex        =   36
      Top             =   5820
      Width           =   1110
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«·„Œ“‰ «·ÿ«·»"
      Height          =   270
      Index           =   24
      Left            =   10380
      TabIndex        =   32
      Top             =   5475
      Width           =   1110
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄"
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10440
      TabIndex        =   31
      Top             =   6240
      Width           =   960
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "Â–Â  «·‘«‘Â   ﬁÊ„ »«·»ÕÀ ⁄‰ ÿ·»«  «·‘Õ‰"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1860
      Index           =   44
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "«·⁄‰Ê«‰"
      Height          =   270
      Index           =   28
      Left            =   7440
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   4560
      Width           =   945
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·« ’«·"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   10230
      TabIndex        =   24
      Top             =   4590
      Width           =   1305
   End
   Begin VB.Label lblitemid 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«·„‰œÊ»"
      Height          =   315
      Index           =   34
      Left            =   10500
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   5010
      Width           =   1035
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "„·ÕÊŸ…"
      Height          =   375
      Left            =   14160
      TabIndex        =   11
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "«·»·œ"
      Height          =   375
      Left            =   8640
      TabIndex        =   5
      Top             =   9120
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "«·⁄„Ì·"
      Height          =   375
      Left            =   10440
      TabIndex        =   4
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "«· «—ÌŒ"
      Height          =   375
      Left            =   7800
      TabIndex        =   3
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·ÿ·»"
      Height          =   375
      Left            =   10320
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "ShipmentOrderSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As ADODB.Recordset
Dim cSearchDcbo As clsDCboSearch

Private m_DcboItems As DataCombo

Private m_RetrunType As Integer
Public WithEvents FG1 As VSFlex8UCtl.vsFlexGrid
Attribute FG1.VB_VarHelpID = -1

Public WithEvents NewGrid As VSFlex8UCtl.vsFlexGrid
Attribute NewGrid.VB_VarHelpID = -1
'Public NewGrid As New ClsGrid
 
Public LngRow As Long
Public TType As Integer
Public LngCol As Long


Private Sub BtnFirst_Click()

End Sub

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
                Fg.Clear flexClearScrollable, flexClearEverything
                Fg.Rows = 2

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
            Fg.SetFocus

        Case 1
            clear_all Me
            Fg.Clear flexClearScrollable, flexClearEverything
XPDtbBill.value = ""
        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

End Sub

 

Private Sub DBCboClientName_Change()
  TxtCusID.text = ""

    Dim DefaultSalesPersonId As Integer
    Dim fullcode As String

    GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, fullcode

    TxtCusID.text = fullcode

    'If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
 
        GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId

  '  End If
 
End Sub

Private Sub DBCboClientName_Click(Area As Integer)
    DBCboClientName_Change
End Sub

Private Sub DcboEmp_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetSalesRepData Me.DcboEmp

    End If
End Sub

Private Sub Fg_Click()
    On Error GoTo ErrTrap
       
    
   
     
    If TType = 0 Then
         FrmShipmentOrder.Retrive (Fg.TextMatrix(Fg.Row, Fg.ColIndex("Transaction_ID")))
    ElseIf TType = 1 Then
    
    FrmShipmentRegestration.Txt_order_no = (Fg.TextMatrix(Fg.Row, Fg.ColIndex("Transaction_ID1")))
          
  End If
    
ErrTrap:
End Sub

Private Sub Retrive()
    Dim Num As Integer
    On Error GoTo ErrTrap
    Fg.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        Fg.Rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With Fg
              .TextMatrix(Num, .ColIndex("Transaction_ID1")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
               .TextMatrix(Num, .ColIndex("Transaction_ID")) = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)
            
                    .TextMatrix(Num, .ColIndex("Address")) = IIf(IsNull(rs("Address").value), "", Trim(rs("Address").value))
                .TextMatrix(Num, .ColIndex("ContactPhone")) = IIf(IsNull(rs("ContactPhone").value), "", Trim(rs("ContactPhone").value))

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value))
                Else
                    .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", Trim(rs("CusNamee").value))
                End If

               If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Num, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", Trim(rs("Emp_Name").value))
                Else
                    .TextMatrix(Num, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", Trim(rs("Emp_Namee").value))
                End If
                

                .TextMatrix(Num, .ColIndex("Transaction_Date")) = IIf(IsNull(rs("Transaction_Date").value), "", Trim(rs("Transaction_Date").value))
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
    Fg_Click
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim BG As New ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    
      '  Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
XPDtbBill.value = ""
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
 
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
 
    Dim My_SQL As String
    Set Dcombos = New ClsDataCombos
    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
    Dcombos.GetSalesRepData Me.DcboEmp
       Dcombos.GetBranches Me.dcBranch
    Dcombos.GetStores Me.DCboStoreName
Dcombos.GetStores Me.DCboStoreName1

   ' Dcombos.GetItemsNames DCboItem, , , , True
    
    My_SQL = " select CountryID,CountryName from TblCountriesData"
 
    fill_combo Me.DataCombo4, My_SQL
    RetrunType = -1
 
    CenterForm Me

    FormPostion Me, GetPostion
    Fg.WallPaper = BG.SearchWallpaper
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

StrSQL = " SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.order_no, dbo.Transactions.Currency_id, dbo.Transactions.Transaction_Date, dbo.TblCustemers.CusName, "
       StrSQL = StrSQL & "               dbo.TblCustemers.CusNamee, dbo.Transactions.Transaction_Type, dbo.Transactions.Transaction_HijriDate, dbo.Transactions.Trans_Discount,"
      StrSQL = StrSQL & "                dbo.Transactions.PaymentType, dbo.Transactions.Transaction_Serial, dbo.Transactions.NoteSerial1, dbo.Transactions.RegionID,"
     StrSQL = StrSQL & "                 dbo.Transactions.CashCustomerName, dbo.Transactions.CashCustomerPhone, dbo.Transactions.CashCustomerMobile, dbo.Transactions.CashCustomerAddress,"
     StrSQL = StrSQL & "                 dbo.Transactions.CashCustomerComment, dbo.Transactions.ContactTime, dbo.Transactions.UserID, dbo.Transactions.Enterdate, dbo.Transactions.EnterTime,"
    StrSQL = StrSQL & "                  dbo.Transactions.ContactPhone, dbo.Transactions.BranchId, dbo.Transactions.oorderdate, dbo.Transactions.CBoBasedON, dbo.Transactions.PONo,"
   StrSQL = StrSQL & "                   dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
   StrSQL = StrSQL & "                   dbo.TblEmployee.Emp_Namee1, dbo.Transactions.Emp_ID, dbo.Transactions.Address, dbo.Transactions.TransactionComment,"
 StrSQL = StrSQL & "                     dbo.TblCustemers.Fullcode AS [Fullcode c]"
StrSQL = StrSQL & " FROM         dbo.TblEmployee RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.Transactions LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID ON dbo.TblEmployee.Emp_ID = dbo.Transactions.Emp_ID"
    ' StrSQL = StrSQL & " where    1=1 "
    StrSQL = StrSQL + " WHERE     (dbo.Transactions.Transaction_Type = 54)"
 
  If Me.TxtEmployeeID.text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblEmployee.Fullcode ='" & Me.TxtEmployeeID.text & "'"
 
    End If
    
      If Me.TxtContactPhone.text <> "" Then
 
        StrWhere = StrWhere + " and dbo.Transactions.ContactPhone ='" & Me.TxtContactPhone.text & "'"
 
    End If
    If TxtCusID.text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblCustemers.Fullcode ='" & Me.TxtCusID.text & "'"
 
    End If
        If TxtAddress.text <> "" Then
 
        StrWhere = StrWhere + " and dbo.Transactions.Address ='" & Me.TxtAddress.text & "'"
 
    End If
    
    If Me.DBCboClientName.BoundText <> "" And Me.DBCboClientName.text <> "" Then
 
        StrWhere = StrWhere + " and TblCustemers.CusID =" & Me.DBCboClientName.BoundText & ""
 
    End If
 If Me.DcboEmp.BoundText <> "" And Me.DcboEmp.text <> "" Then
 
        StrWhere = StrWhere + " and dbo.Transactions.Emp_ID =" & Me.DcboEmp.BoundText & ""
 
    End If
    
    
 If Me.dcBranch.BoundText <> "" And Me.dcBranch.text <> "" Then
 
        StrWhere = StrWhere + " and dbo.Transactions.BranchId =" & Me.dcBranch.BoundText & ""
 
    End If
    
    
     If Me.DCboStoreName.BoundText <> "" And Me.DCboStoreName.text <> "" Then
 
        StrWhere = StrWhere + " and dbo.Transactions.storeID =" & Me.DCboStoreName.BoundText & ""
 
    End If
    
     If Me.DCboStoreName1.BoundText <> "" And Me.DCboStoreName1.text <> "" Then
 
        StrWhere = StrWhere + " and dbo.Transactions.storeID1 =" & Me.DCboStoreName1.BoundText & ""
 
    End If
    
    
 
       ' StrWhere = StrWhere + " and dbo.Transactions.Transaction_ID ='" & Me.txtorder_no.text & "'"
 
    
    
    If Me.txtorder_no.text <> "" Then
 
        StrWhere = StrWhere + " and dbo.Transactions.NoteSerial1 ='" & Me.txtorder_no.text & "'"
 
    End If
 If Not IsNull(Me.XPDtbBill.value) Then
        
            StrWhere = StrWhere & " AND dbo.Transactions.Transaction_Date >=" & SQLDate(Me.XPDtbBill.value, True) & ""
      End If
    StrWhere = StrWhere + " order by dbo.Transactions.Transaction_ID"

    Build_Sql = StrSQL + StrWhere
    Exit Function
ErrTrap:
End Function

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl Is Fg Then
            If Not Fg.TextMatrix(Fg.Row, 1) = "" Then
                Fg_Click
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

'Public Property Get DcboItems() As DataCombo
'    Set DcboItems = m_DcboItems
'End Property

'Public Property Set DcboItems(ByVal vNewValue As DataCombo)
'    Set m_DcboItems = vNewValue
'End Property
'
Public Property Get RetrunType() As Integer
    RetrunType = m_RetrunType
End Property

Public Property Let RetrunType(ByVal vNewValue As Integer)
    m_RetrunType = vNewValue
    ' 0 = Retrun in the Items Screen
    ' 1 = Retrun in the Data Combo
End Property

Private Sub ChangeLang()
    Me.Caption = "Search For Order Shipping"
    Label1(2).Caption = Me.Caption
    Label2.Caption = "Order No"
 
    Label3.Caption = "Date"
    Label17.Caption = "ContactPhone"
lbl(28).Caption = "Address"
    Label4.Caption = "Customer"
    Label6.Caption = "Remark"
lbl(44).Caption = "This Screen Looks for the Request Shipping"
    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"
lbl(34).Caption = "Sales Person"
    'OptType(0).Caption = "Start of the name"
    'OptType(1).Caption = "any part of the name"
    With Me.Fg
        .TextMatrix(0, .ColIndex("Transaction_ID")) = "Order No"
         .TextMatrix(0, .ColIndex("Transaction_Date")) = "Date  "
        .TextMatrix(0, .ColIndex("CusName")) = "Customer Name"
        .TextMatrix(0, .ColIndex("Emp_Name")) = " Sales Person"
             .TextMatrix(0, .ColIndex("ContactPhone")) = "ContactPhone "
  .TextMatrix(0, .ColIndex("Address")) = "Address "
        '  .AutoSize 0, .Cols - 1, False
    End With

End Sub

'Private Sub lblitemid_Change()
'DCboItem.BoundText = val(Me.lblitemid.Caption)
'End Sub

'Private Sub TxtItemCode_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'          If KeyCode = vbKeyReturn Then
'                If Trim(Me.TxtItemCode(1).text) = "" Then Exit Sub
'                StrSQL = "Select ItemID From TblItems Where ItemCode='" & Trim(Me.TxtItemCode(Index).text) & "'"
'                Set rs = New ADODB.Recordset
'                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

'                If Not (rs.BOF Or rs.EOF) Then
'                    DCboItem.BoundText = rs("ItemID").value
'                Else
'                    Msg = "·«ÌÊÃœ ’‰› „”Ã· »Â–« «·ﬂÊœ..!"
'                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'                End If
'            End If
'
'End Sub

Private Sub TxtCusID_KeyPress(KeyAscii As Integer)
Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , Me.TxtCusID.text, 1
        DBCboClientName.BoundText = CUSTID
    End If
End Sub

