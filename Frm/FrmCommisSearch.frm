VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCommisSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ «·⁄„Ê·«  «·„” Õﬁ…"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12510
   Icon            =   "FrmCommisSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   12510
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtFitter 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   3240
      Width           =   6675
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄„·Ì…"
      Height          =   645
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   2580
      Width           =   3795
      Begin VB.TextBox TxtIDTOo 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox TxtIDFromo 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   9
         Left            =   1020
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   7
         Left            =   2775
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.ComboBox DcbOrderStatus 
      Height          =   315
      Left            =   12840
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   12960
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Text            =   " „ „Ê«›ﬁ… «·⁄„"
      Top             =   600
      Width           =   1035
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   645
      Index           =   0
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   3600
      Width           =   12435
      Begin VB.TextBox TxtRate 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2940
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   240
         Width           =   1635
      End
      Begin VB.TextBox TxtValue 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   240
         Width           =   1635
      End
      Begin VB.TextBox TxtOperation 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   6060
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   180
         Width           =   1995
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   9120
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   180
         Width           =   1635
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰”»… «·⁄„Ê·…"
         Height          =   195
         Index           =   12
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·⁄„Ê·…"
         Height          =   195
         Index           =   11
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   300
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄„·ÌÂ"
         Height          =   195
         Index           =   8
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬁÌ„… «·«Ã„«·Ì…"
         Height          =   195
         Index           =   0
         Left            =   11175
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   1020
      End
   End
   Begin VB.Frame lbreg 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·⁄„·ÌÂ"
      Height          =   1035
      Left            =   2580
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2520
      Width           =   2295
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   90
         TabIndex        =   6
         Top             =   270
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   408027139
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   90
         TabIndex        =   7
         Top             =   630
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   408027139
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   3
         Left            =   1695
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   660
         Width           =   480
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   4
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   330
         Width           =   540
      End
   End
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÿ·»"
      Height          =   645
      Left            =   8760
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2580
      Width           =   3795
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox TxtIDTO 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   5
         Left            =   2775
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   6
         Left            =   1020
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   525
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   60
      TabIndex        =   10
      Top             =   30
      Width           =   12435
      _cx             =   21934
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmCommisSearch.frx":038A
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
      Left            =   1650
      TabIndex        =   11
      Top             =   5280
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
      Left            =   810
      TabIndex        =   12
      Top             =   5280
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
      Left            =   30
      TabIndex        =   13
      Top             =   5280
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
   Begin VSFlex8UCtl.VSFlexGrid Grid 
      Height          =   2355
      Left            =   -450
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   13485
      _cx             =   23786
      _cy             =   4154
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
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmCommisSearch.frx":051E
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
   Begin MSDataListLib.DataCombo DBCboClientName 
      Height          =   315
      Left            =   7230
      TabIndex        =   36
      Top             =   4230
      Visible         =   0   'False
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcbEmployee 
      Height          =   285
      Left            =   7230
      TabIndex        =   37
      Top             =   4710
      Visible         =   0   'False
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„‰œÊ»"
      Height          =   285
      Index           =   12
      Left            =   11850
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   4710
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·⁄„Ì·"
      Height          =   285
      Index           =   1
      Left            =   11910
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   4290
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·≈Ã„«·Ï"
      Height          =   285
      Index           =   2
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2880
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2880
      Width           =   1665
   End
   Begin VB.Label LblClientName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·›‰Ì"
      Height          =   195
      Left            =   11760
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   3270
      Width           =   675
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   10
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   2700
      Width           =   1575
   End
End
Attribute VB_Name = "FrmCommisSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Public mType As Integer
Public ScrenFlg As Integer

Private Sub Cmd_Click(Index As Integer)

    Select Case Index
        Case 0
            If mType = 1 Then
                GetData2
            Else
                GetData
            End If
        Case 1
            clear_all Me
            Me.DtpDateFrom.value = ""
            Me.DtpDateTo.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
               ' Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
    End Select

End Sub


Private Sub Fg_Click()

   
 FrmCommisRece.Retrive (val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id"))))
 

End Sub





Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
   ' Dcombos.GetEmployees Me.DCEmp_Name
    ' Dcombos.GetClientName Me.DCEmp_Name
    
    If mType = 1 Then
        Fg.Visible = False
        Grid.Visible = True
        lbl(12).Caption = "‰”»… «·’Ì«‰…"
        DBCboClientName.Visible = True
        DcbEmployee.Visible = True
        Label1(1).Visible = True
        Label1(12).Visible = True
        TxtFitter.Visible = False
        LblClientName.Visible = False
        lbprocess.Caption = "—ﬁ„ «·⁄ﬁœ/«·›« Ê—…"
        TxtOperation.Visible = False
        lbl(8).Visible = False
    
    If ScrenFlg = 1 Then
    Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName, True
    Else
    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
    End If
    Dcombos.GetSalesRepData Me.DcbEmployee
    Else
        Fg.Visible = True
        Grid.Visible = False
    End If
      If SystemOptions.UserInterface = EnglishInterface Then
     
        Me.DcbOrderStatus.AddItem "New"
        Me.DcbOrderStatus.AddItem "Accept Customer"
        Me.DcbOrderStatus.AddItem "Final Maintenance"

             Else
  
 DcbOrderStatus.AddItem "ÃœÌœ"
DcbOrderStatus.AddItem " „ „Ê«›ﬁ… «·⁄„Ì·"
DcbOrderStatus.AddItem " „ «‰Â«¡ «·«’·«Õ"


    End If
    Set DCboSearch = New clsDCboSearch
    'Set DCboSearch.Client = Me.DCEmp_Name
    'Dcombos.GetUsers Me.DCUser
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

  '  CenterForm Me
'GetData
'    FormPostion Me, GetPostion
    Set GrdBack = New ClsBackGroundPic

    With Me.Fg
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    SetDtpickerDate Me.DtpDateFrom
    SetDtpickerDate Me.DtpDateTo

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub

Public Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
StrSQL = " SELECT     dbo.TblCommisRece.id, dbo.TblCommisReceDetails.ID_Aut, dbo.TblCommisReceDetails.DateOp, dbo.TblCommisReceDetails.Total,"
 StrSQL = StrSQL & "                     dbo.TblCommisReceDetails.Operation, dbo.TblCommisReceDetails.Fitter, dbo.TblCommisReceDetails.PerceTage, dbo.TblCommisReceDetails.PerceTageValue,"
 StrSQL = StrSQL & "                      dbo.TblCommisReceDetails.id2, dbo.TblCommisRece.DateFrom, dbo.TblCommisRece.DateTo, dbo.TblCommisRece.RecordDate, dbo.TblCommisRece.AllFit,"
  StrSQL = StrSQL & "                     dbo.TblCommisRece.LimitFit , dbo.TblCommisRece.UserID"
 StrSQL = StrSQL & " FROM         dbo.TblCommisRece INNER JOIN"
  StrSQL = StrSQL & "                     dbo.TblCommisReceDetails ON dbo.TblCommisRece.id = dbo.TblCommisReceDetails.id2"
   
    BolBegine = False
    StrWhere = ""

    If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblCommisReceDetails.ID_Aut >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCommisReceDetails.ID_Aut >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If
      If val(Me.TxtIDFromo.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblCommisRece.id >=" & val(Me.TxtIDFromo.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCommisRece.id >=" & val(Me.TxtIDFromo.Text) & ""
        End If
    End If


    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCommisReceDetails.ID_Aut <=" & val(Me.TxtIDTO.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCommisReceDetails.ID_Aut <=" & val(Me.TxtIDTO.Text) & ""
        End If
    End If
        If val(Me.TxtIDTOo.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCommisRece.id <=" & val(Me.TxtIDTOo.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCommisRece.id <=" & val(Me.TxtIDTOo.Text) & ""
        End If
    End If
    '///////////////////
   
'////////////////////////
 If Me.TxtFitter.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCommisReceDetails.Fitter like '%" & Me.TxtFitter.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCommisReceDetails.Fitter like '%" & Me.TxtFitter.Text & "%'"
        End If
    End If
   If txtTotal.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCommisReceDetails.Total=" & txtTotal.Text & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCommisReceDetails.Total=" & txtTotal.Text & ""
        End If
    End If
If TxtOperation.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCommisReceDetails.Operation like '%" & Me.TxtOperation.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCommisReceDetails.Operation like '%" & Me.TxtOperation.Text & "%'"
        End If
    End If
    If TxtRate.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCommisReceDetails.PerceTage=" & TxtRate.Text & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCommisReceDetails.PerceTage=" & TxtRate.Text & ""
        End If
    End If
        If TxtValue.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCommisReceDetails.PerceTageValue=" & TxtValue.Text & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCommisReceDetails.PerceTageValue=" & TxtValue.Text & ""
        End If
    End If
   ' If Me.DCUser.BoundText <> "" Then
   ''     If BolBegine = True Then
    ''        StrWhere = StrWhere & " AND    dbo.TblCardAuthorizationReform.UserID=" & Me.DCUser.BoundText & ""
     ''   Else
      ''      BolBegine = True
       '     StrWhere = " Where    dbo.TblCardAuthorizationReform.UserID=" & Me.DCUser.BoundText & ""
       ' End If
    'End If

    If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblCommisRece.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblCommisRece.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblCommisRece.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblCommisRece.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If

    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TblCommisRece.ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
    'Me.lbl(10).Caption = "‰ ÌÃ…«·»Õ÷="
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
    '        Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.Fg
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
              '  Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
              '  Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
              
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                 .TextMatrix(i, .ColIndex("ID_Aut")) = IIf(IsNull(rs("ID_Aut").value), "", rs("ID_Aut").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                    .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
            If Not (IsNull(rs("DateOp").value)) Then
                    .TextMatrix(i, .ColIndex("DateOp")) = Format(rs("DateOp").value, "yyyy/M/d")
                End If
                .TextMatrix(i, .ColIndex("Fitter")) = IIf(IsNull(rs("Fitter").value), "", rs("Fitter").value)
                .TextMatrix(i, .ColIndex("Total")) = val(IIf(IsNull(rs("Total").value), "", rs("Total").value))
                .TextMatrix(i, .ColIndex("Operation")) = IIf(IsNull(rs("Operation").value), "", rs("Operation").value)
              .TextMatrix(i, .ColIndex("PerceTage")) = IIf(IsNull(rs("PerceTage").value), "", rs("PerceTage").value)
               .TextMatrix(i, .ColIndex("PerceTageValue")) = IIf(IsNull(rs("PerceTageValue").value), "", rs("PerceTageValue").value)
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

    End If

End Sub


Public Sub GetData2()
 
    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String
    Dim BolBegine As Boolean
    Dim StrWhere As String
    Set rs = New ADODB.Recordset
    My_SQL = "select * From TblOLDContract"
   
    Grid.Rows = 1
  BolBegine = False
    StrWhere = ""

    If val(Me.TxtIDFrom.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblOLDContract.ContractNo >=" & val(Me.TxtIDFrom.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblOLDContract.ContractNo >=" & val(Me.TxtIDFrom.Text) & ""
        End If
    End If
      If val(Me.TxtIDFromo.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblOLDContract.ContractNo >=" & val(Me.TxtIDFromo.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblOLDContract.ContractNo >=" & val(Me.TxtIDFromo.Text) & ""
        End If
    End If


    If val(Me.TxtIDTO.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblOLDContract.ContractNo <=" & val(Me.TxtIDTO.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblOLDContract.ContractNo <=" & val(Me.TxtIDTO.Text) & ""
        End If
    End If
        If val(Me.TxtIDTOo.Text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblOLDContract.ContractNo <=" & val(Me.TxtIDTOo.Text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblOLDContract.ContractNo <=" & val(Me.TxtIDTOo.Text) & ""
        End If
    End If
    '///////////////////
   
'////////////////////////

   If txtTotal.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblOLDContract.ContractValue=" & txtTotal.Text & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblOLDContract.ContractValue=" & txtTotal.Text & ""
        End If
    End If

   
      
    If Me.DBCboClientName.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND    dbo.TblOLDContract.CusID=" & Me.DBCboClientName.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where    dbo.TblOLDContract.CusID=" & Me.DBCboClientName.BoundText & ""
        End If
    End If
    
    If Me.DcbEmployee.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND    dbo.TblOLDContract.Emp_ID=" & Me.DcbEmployee.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where    dbo.TblOLDContract.Emp_ID=" & Me.DcbEmployee.BoundText & ""
        End If
    End If

    If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblOLDContract.ContractDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblOLDContract.ContractDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblOLDContract.ContractDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblOLDContract.ContractDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
    If ScrenFlg = 1 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " And   ScrenFlg=1 "
        Else
            StrWhere = " Where ScrenFlg=1 "
        End If
    Else
        If BolBegine = True Then
            StrWhere = StrWhere & " And   ScrenFlg Is null "
        Else
            StrWhere = " Where ScrenFlg Is null  "
        End If
    End If
    My_SQL = My_SQL & StrWhere & " order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
                .TextMatrix(i, .ColIndex("ContractNo")) = IIf(IsNull(rs.Fields("ContractNo").value), "", rs.Fields("ContractNo").value)
               
               .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
           If Not (IsNull(rs.Fields("PaymentType").value)) Then
           If rs.Fields("PaymentType").value = 0 Then
           If SystemOptions.UserInterface = ArabicInterface Then
           .TextMatrix(i, .ColIndex("PaymentType")) = "ﬁÌ„…"
           Else
           .TextMatrix(i, .ColIndex("PaymentType")) = "Value"
           End If
           ElseIf rs.Fields("PaymentType").value = 1 Then
              If SystemOptions.UserInterface = ArabicInterface Then
           .TextMatrix(i, .ColIndex("PaymentType")) = "‰”»…"
           Else
           .TextMatrix(i, .ColIndex("PaymentType")) = "Percentage"
           End If
           End If
           End If
           .TextMatrix(i, .ColIndex("ContractNo")) = IIf(IsNull(rs.Fields("ContractNo").value), "", rs.Fields("ContractNo").value)
           .TextMatrix(i, .ColIndex("Vlue")) = IIf(IsNull(rs.Fields("Vlue").value), 0, rs.Fields("Vlue").value)
           .TextMatrix(i, .ColIndex("NetValue")) = IIf(IsNull(rs.Fields("NetValue").value), 0, rs.Fields("NetValue").value)
           
                .TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(rs.Fields("CusID").value), "", rs.Fields("CusID").value)
                .TextMatrix(i, .ColIndex("ContractNo")) = IIf(IsNull(rs.Fields("ContractNo").value), "", rs.Fields("ContractNo").value)
                .TextMatrix(i, .ColIndex("ContractDate")) = IIf(IsNull(rs.Fields("ContractDate").value), "", rs.Fields("ContractDate").value)
                .TextMatrix(i, .ColIndex("ContractValue")) = IIf(IsNull(rs.Fields("ContractValue").value), "", rs.Fields("ContractValue").value)
                .TextMatrix(i, .ColIndex("EndGuranteeDate")) = IIf(IsNull(rs.Fields("EndGuranteeDate").value), "", rs.Fields("EndGuranteeDate").value)
               .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(rs.Fields("Remarks").value), "", rs.Fields("Remarks").value)
               .TextMatrix(i, .ColIndex("ReturnNo")) = IIf(IsNull(rs.Fields("ReturnNo").value), "", rs.Fields("ReturnNo").value)
               .TextMatrix(i, .ColIndex("ReturnValue")) = IIf(IsNull(rs.Fields("ReturnValue").value), "", rs.Fields("ReturnValue").value)
               
                rs.MoveNext
            Next

            rs.Close
        End If

        .RowHeight(-1) = 300
    End With

ErrTrap:
 


    '-----------------------------------

   
  

End Sub


Private Sub ChangeLang()
 
    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "Search"
    Cmd(2).Caption = "Exit"
  Me.Caption = "Search CardAuthorizationReform"

Me.LblClientName.Caption = "ClientName"
lbl(4).Caption = "From"
lbl(3).Caption = "To"
lbl(7).Caption = "From"
lbl(9).Caption = "To"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
lbl(0).Caption = "Total Value"
lbl(8).Caption = "Operation"
lbl(2).Caption = "Total"
lbl(12).Caption = "Commission Rate"
lbl(11).Caption = "Commission"
Me.lbreg.Caption = "CommisSearch"
Me.lbprocess.Caption = "Order No"
Frame1.Caption = "No. Process"
     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("id")) = "No Order"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date Process"
         .TextMatrix(0, .ColIndex("ID_Aut")) = "No Process"
        .TextMatrix(0, .ColIndex("DateOp")) = "Date Order"
       .TextMatrix(0, .ColIndex("Total")) = "Total"
        .TextMatrix(0, .ColIndex("Fitter")) = "Technical"
         .TextMatrix(0, .ColIndex("Operation")) = "Operation"
        .TextMatrix(0, .ColIndex("PerceTage")) = "Commission Rate"
    .TextMatrix(0, .ColIndex("PerceTageValue")) = "Commission Payable"
    End With
  '
End Sub

Private Sub Grid_Click()

 FrmOldContract.FindRec (val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("id"))))
 
 
 


End Sub

Private Sub TxtIDFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDFrom.Text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

Private Sub TxtIDTO_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDTO.Text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

