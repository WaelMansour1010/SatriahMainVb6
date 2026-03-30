VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSearch_VehicleAllocation 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "»ÕÀ ⁄ﬁÊœ «·Ê“«—…"
   ClientHeight    =   9180
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   10380
   Icon            =   "FrmSearch_VehicleAllocation.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Height          =   852
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   8280
      Width           =   10332
      Begin ImpulseButton.ISButton Cmd 
         Height          =   432
         Index           =   0
         Left            =   4860
         TabIndex        =   30
         Top             =   240
         Width           =   996
         _ExtentX        =   1757
         _ExtentY        =   762
         ButtonPositionImage=   1
         Caption         =   "»ÕÀ"
         BackColor       =   14871017
         FontSize        =   7.8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         Height          =   432
         Index           =   1
         Left            =   3768
         TabIndex        =   31
         Top             =   240
         Width           =   1032
         _ExtentX        =   1820
         _ExtentY        =   762
         ButtonPositionImage=   1
         Caption         =   "„”Õ"
         BackColor       =   14871017
         FontSize        =   7.8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         Height          =   432
         Index           =   2
         Left            =   2760
         TabIndex        =   32
         Top             =   240
         Width           =   972
         _ExtentX        =   1715
         _ExtentY        =   762
         ButtonPositionImage=   1
         Caption         =   "Œ—ÊÃ"
         BackColor       =   14871017
         FontSize        =   7.8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
   End
   Begin VB.Frame frm_ministrycontract 
      BackColor       =   &H00E2E9E9&
      Height          =   8412
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   -120
      Visible         =   0   'False
      Width           =   10332
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2E9E9&
         Height          =   1812
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   6480
         Width           =   10332
         Begin VB.TextBox txtProcessNo 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   7668
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   240
            Width           =   1560
         End
         Begin VB.TextBox XPTxtBoxName 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   4644
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   600
            Width           =   4608
         End
         Begin VB.TextBox txtMinistryContractNo 
            Alignment       =   1  'Right Justify
            Height          =   312
            Left            =   4644
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   240
            Width           =   1608
         End
         Begin MSComCtl2.DTPicker dtpToDate 
            Height          =   312
            Left            =   1704
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   600
            Width           =   1464
            _ExtentX        =   2582
            _ExtentY        =   550
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   84606979
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker dtpFromDate 
            Height          =   312
            Left            =   1704
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   240
            Width           =   1464
            _ExtentX        =   2582
            _ExtentY        =   550
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   84606979
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal dtpFromDateH 
            Height          =   312
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1596
            _ExtentX        =   2815
            _ExtentY        =   550
         End
         Begin Dynamic_Byte.NourHijriCal dtpToDateH 
            Height          =   312
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   1596
            _ExtentX        =   2815
            _ExtentY        =   550
         End
         Begin Dynamic_Byte.NourHijriCal dtpSContractDateH 
            Height          =   252
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   1596
            _ExtentX        =   2815
            _ExtentY        =   445
         End
         Begin MSComCtl2.DTPicker dtpSContractDate 
            Height          =   252
            Left            =   1704
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   960
            Width           =   1464
            _ExtentX        =   2582
            _ExtentY        =   445
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   84606979
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker dtpEContractDate 
            Height          =   252
            Left            =   1704
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1464
            _ExtentX        =   2582
            _ExtentY        =   445
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   84606979
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal dtpEContractDateH 
            Height          =   252
            Left            =   120
            TabIndex        =   15
            Top             =   1320
            Width           =   1596
            _ExtentX        =   2815
            _ExtentY        =   445
         End
         Begin MSDataListLib.DataCombo dcVendor 
            Height          =   288
            Left            =   4644
            TabIndex        =   16
            Top             =   960
            Width           =   1608
            _ExtentX        =   2836
            _ExtentY        =   508
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcDuration 
            Height          =   288
            Left            =   7668
            TabIndex        =   17
            Top             =   1320
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   508
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcCity 
            Height          =   288
            Left            =   7668
            TabIndex        =   28
            Top             =   960
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   508
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —…"
            Height          =   252
            Index           =   1
            Left            =   9252
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   1320
            Width           =   972
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«œ«—… «· ⁄·Ì„Ì…"
            Height          =   312
            Index           =   9
            Left            =   6252
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   960
            Width           =   1212
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„‰ÿﬁ…"
            Height          =   312
            Index           =   10
            Left            =   9204
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   960
            Width           =   1020
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "  «‰ Â«¡ «· ⁄«ﬁœ „Ì·«œÏ"
            Height          =   372
            Index           =   8
            Left            =   3072
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   1320
            Width           =   1248
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "  «· ⁄«ﬁœ „Ì·«œÏ"
            Height          =   252
            Index           =   5
            Left            =   2820
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   960
            Width           =   1500
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„‰  «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   3000
            TabIndex        =   22
            Top             =   240
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ï  «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   0
            Left            =   3144
            TabIndex        =   21
            Top             =   600
            Width           =   1176
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”„Ï «· ⁄«ﬁœ"
            Height          =   252
            Index           =   3
            Left            =   8988
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   600
            Width           =   1236
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ ⁄ﬁœ «·Ê“«—…"
            Height          =   312
            Index           =   15
            Left            =   6252
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   240
            Width           =   1212
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·⁄ﬁœ"
            Height          =   288
            Index           =   0
            Left            =   9300
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   240
            Width           =   924
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   588
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   10416
         _cx             =   18373
         _cy             =   1037
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   24
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
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "       »ÕÀ ⁄ﬁÊœ «·Ê“«—…   "
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   2
         ChildSpacing    =   1
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   7
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   1
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
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   5760
         Left            =   0
         TabIndex        =   3
         Top             =   720
         Width           =   10332
         _cx             =   18224
         _cy             =   10160
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearch_VehicleAllocation.frx":038A
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
End
Attribute VB_Name = "FrmSearch_VehicleAllocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Public calltype As Integer
Public SendForm As String

Private Sub Cmd_Click(Index As Integer)
On Error Resume Next
    Select Case Index

        Case 0
                If frm_ministrycontract.Visible = True Then
                    GetData
                ElseIf frm_attribution.Visible = True Then
                    GetData_Attr
                End If
        Case 1
            clear_all Me
        Case 2
            Unload Me
    End Select

End Sub




Private Sub fg_attr_Click()

  Dim i As Integer
   i = val(fg_attr.TextMatrix(fg_attr.Row, fg_attr.ColIndex("IDMC")))
   
   If i > 0 Then
   
        If SendForm = "attributioncontract" Then
              FrmAttributionContract.Retrive (i)
        ElseIf SendForm = "AC" Then
             Set Rs_Temp = New ADODB.Recordset
            Rs_Temp.Open " select ProcessNo  from TblMinistryContract where  IDMC =  " & i, Cn, adOpenStatic, adLockOptimistic, adCmdText
             If Rs_Temp.RecordCount > 0 Then
                    FrmAttributionContract.txtMinistryContractNo.text = IIf(IsNull(Rs_Temp("ProcessNo").value), "", Rs_Temp("ProcessNo").value)
             End If
        ElseIf SendForm = "VA" Then
             Set Rs_Temp = New ADODB.Recordset
            Rs_Temp.Open " select ProcessNo  from TblMinistryContract where  IDMC =  " & i, Cn, adOpenStatic, adLockOptimistic, adCmdText
             If Rs_Temp.RecordCount > 0 Then
                    FrmVehicleAllocation.dcMinistryContract.BoundText = IIf(IsNull(Rs_Temp("IDMC").value), "", Rs_Temp("IDMC").value)
             End If
        End If
   
   
   End If


'Unload Me
ErrTrap:



End Sub

Private Sub Fg_Click()

' On Error GoTo ErrTrap
     
   Dim i As Integer
   i = val(Fg.TextMatrix(Fg.Row, Fg.ColIndex("IDMC")))
   
   If i > 0 Then
        
        
        If SendForm = "MC" Then
              FrmMinistryContract.Retrive (i)
        ElseIf SendForm = "AC" Then
              
             Set Rs_Temp = New ADODB.Recordset
             Rs_Temp.Open " select ProcessNo  from TblMinistryContract where  IDMC =  " & i, Cn, adOpenStatic, adLockOptimistic, adCmdText
             If Rs_Temp.RecordCount > 0 Then
                    FrmAttributionContract.txtMinistryContractNo.text = IIf(IsNull(Rs_Temp("ProcessNo").value), "", Rs_Temp("ProcessNo").value)
             End If
        ElseIf SendForm = "VA" Then
            ' Set Rs_Temp = New Adodb.Recordset
            ' Rs_Temp.Open " select ProcessNo  from TblMinistryContract where  IDMC =  " & i, Cn, adOpenStatic, adLockOptimistic, adCmdText
            ' If Rs_Temp.RecordCount > 0 Then
                    FrmVehicleAllocation.dcMinistryContract.BoundText = i ' IIf(IsNull(Rs_Temp("IDMC").value), "", Rs_Temp("IDMC").value)
            ' End If
        End If
   
   
   End If


'Unload Me
ErrTrap:

End Sub

Private Sub Form_Activate()
PutFormOnTop Me.hWnd, True
mdifrmmain.Enabled = False
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.getCountriesGovernments dcCity
    Dim str As String
    If SystemOptions.UserInterface = ArabicInterface Then
    str = " Select ID , Name   from TblManagerialArea "
    Else
    str = " Select ID , NameE   from TblManagerialArea "
    End If
    fill_combo dcVendor, str
    str = "select id , name  from TblDurations "
    fill_combo dcDuration, str
 

    Set DCboSearch = New clsDCboSearch
    Set GrdBack = New ClsBackGroundPic
    With Me.Fg
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With
 
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
   dtpFromDate.value = Date
   dtpToDate.value = Date
   dtpFromDateH.value = ToHijriDate(Date)
   dtpToDateH.value = ToHijriDate(Date)
   dtpSContractDate.value = Date
   dtpSContractDateH.value = ToHijriDate(Date)
   dtpEContractDate.value = Date
   dtpEContractDateH.value = ToHijriDate(Date)
   dtpFromDate1.value = Date
   dtpToDate1.value = Date
   dtpFromDateH1.value = ToHijriDate(Date)
   dtpToDateH1.value = ToHijriDate(Date)
   dtpSContractDate1.value = Date
   dtpSContractDateH1.value = ToHijriDate(Date)
   dtpEContractDate1.value = Date
   dtpEContractDateH1.value = ToHijriDate(Date)
    
   dtpFromDate.value = Null
   dtpToDate.value = Null
   dtpSContractDate.value = Null
   dtpEContractDate.value = Null
   dtpFromDate1.value = Null
   dtpToDate1.value = Null
   dtpSContractDate1.value = Null
   dtpEContractDate1.value = Null
    
    If SendForm = "MC" Or SendForm = "AC" Or SendForm = "VA" Then
            frm_ministrycontract.Visible = True
    ElseIf SendForm = "attributioncontract" Then
            frm_attribution.Visible = True
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
mdifrmmain.Enabled = True
   ' FormPostion Me, SavePostion
   ' Set DCboSearch = Nothing
End Sub

Public Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

  
    StrSQL = "  SELECT dbo.TblMinistryContract.DurationID, dbo.TblMinistryContract.CityID, dbo.TblMinistryContract.VendorID, dbo.TblMinistryContract.ProcessNo, dbo.TblMinistryContract.Name,"
    StrSQL = StrSQL & "     dbo.TblMinistryContract.FromDate, dbo.TblMinistryContract.FromDateH, dbo.TblMinistryContract.ToDate, dbo.TblMinistryContract.ToDateH,"
    StrSQL = StrSQL & "     dbo.TblMinistryContract.StudentCount, dbo.TblMinistryContract.StudentCustom, dbo.TblMinistryContract.DisCount, dbo.TblMinistryContract.StartContractDate,"
    StrSQL = StrSQL & "     dbo.TblMinistryContract.StartContractDateh, dbo.TblMinistryContract.EndContractDate, dbo.TblMinistryContract.EndContractDateh,"
    StrSQL = StrSQL & "     dbo.TblDurations.Name AS DurationName, dbo.TblCountriesGovernments.GovernmentName,dbo.TblMinistryContract.MinistryContractNo , dbo.TblManagerialArea.Name AS MAName  , dbo.TblMinistryContract.IDMC"
    StrSQL = StrSQL & "     FROM     dbo.TblMinistryContract LEFT OUTER JOIN"
    StrSQL = StrSQL & "     dbo.TblManagerialArea ON dbo.TblMinistryContract.VendorID = dbo.TblManagerialArea.ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "     dbo.TblCountriesGovernments ON dbo.TblMinistryContract.CityID = dbo.TblCountriesGovernments.GovernmentID LEFT OUTER JOIN"
    StrSQL = StrSQL & "     dbo.TblDurations ON dbo.TblMinistryContract.DurationID = dbo.TblDurations.ID"
    StrSQL = StrSQL & "     where 1 = 1  "
   
   
    If Me.txtProcessNo.text <> "" Then
            StrSQL = StrSQL & "   and  IDMC =  '" & txtProcessNo.text & "'"
    End If
    
    If Me.txtMinistryContractNo.text <> "" Then
            StrSQL = StrSQL & "   and  MinistryContractNo =  '" & txtMinistryContractNo.text & "'"
    End If

    If Me.XPTxtBoxName.text <> "" Then
            StrSQL = StrSQL & "   and  TblMinistryContract.Name like  '%" & XPTxtBoxName.text & "%'"
    End If
    
     If Me.dcDuration.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblMinistryContract.DurationID =  " & val(Me.dcDuration.BoundText)
    End If
 
    If Me.dcCity.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblMinistryContract.CityID =  " & val(Me.dcCity.BoundText)
    End If
    
     If Me.dcVendor.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblMinistryContract.VendorID =  " & val(Me.dcVendor.BoundText)
    End If
     
    If Not IsNull(Me.dtpFromDate.value) Then
            StrSQL = StrSQL & "   and  TblMinistryContract.FromDate  >=  '" & dtpFromDate.value & "'"
    End If
    
   If Not IsNull(Me.dtpToDate.value) Then
            StrSQL = StrSQL & "   and  TblMinistryContract.ToDate  >=  '" & dtpToDate.value & "'"
    End If
    
     If Not IsNull(dtpSContractDate.value) Then
            StrSQL = StrSQL & "   and  TblMinistryContract.StartContractDate  >=  '" & dtpSContractDate.value & "'"
    End If
    
    If Not IsNull(dtpEContractDate.value) Then
            StrSQL = StrSQL & "   and  TblMinistryContract.EndContractDate  >=  '" & dtpEContractDate.value & "'"
    End If
    
     
     
    StrSQL = StrSQL
    StrSQL = StrSQL & " Order By IDMC "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
          '  Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.Lbl(10).Caption = "Search Results=0"
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
                'Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                'Me.Lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                 .TextMatrix(i, .ColIndex("IDMC")) = IIf(IsNull(rs("IDMC").value), "", rs("IDMC").value)
                .TextMatrix(i, .ColIndex("ProcessNo")) = IIf(IsNull(rs("ProcessNo").value), "", rs("ProcessNo").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(i, .ColIndex("MinistryContractNo")) = IIf(IsNull(rs("MinistryContractNo").value), "", rs("MinistryContractNo").value)
                .TextMatrix(i, .ColIndex("DurationName")) = IIf(IsNull(rs("DurationName").value), "", rs("DurationName").value)
                
               .TextMatrix(i, .ColIndex("MAName")) = IIf(IsNull(rs("MAName").value), "", rs("MAName").value)
               .TextMatrix(i, .ColIndex("GovernmentName")) = IIf(IsNull(rs("GovernmentName").value), "", rs("GovernmentName").value)
               .TextMatrix(i, .ColIndex("DurationName")) = IIf(IsNull(rs("DurationName").value), "", rs("DurationName").value)
                .TextMatrix(i, .ColIndex("FromDate")) = IIf(IsNull(rs("FromDate").value), "", rs("FromDate").value)
                .TextMatrix(i, .ColIndex("FromDateH")) = IIf(IsNull(rs("FromDateH").value), "", rs("FromDateH").value)
                .TextMatrix(i, .ColIndex("ToDate")) = IIf(IsNull(rs("ToDate").value), "", rs("ToDate").value)
                .TextMatrix(i, .ColIndex("ToDateH")) = IIf(IsNull(rs("ToDateH").value), "", rs("ToDateH").value)
                .TextMatrix(i, .ColIndex("EndContractDate")) = IIf(IsNull(rs("EndContractDate").value), "", rs("EndContractDate").value)
                .TextMatrix(i, .ColIndex("EndContractDateh")) = IIf(IsNull(rs("EndContractDateh").value), "", rs("EndContractDateh").value)
                .TextMatrix(i, .ColIndex("StartContractDate")) = IIf(IsNull(rs("StartContractDate").value), "", rs("StartContractDate").value)
                .TextMatrix(i, .ColIndex("StartContractDateh")) = IIf(IsNull(rs("StartContractDateh").value), "", rs("StartContractDateh").value)
                
                 rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        
        End With

    End If

End Sub

Public Sub GetData_Attr()

    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

  
  StrSQL = "SELECT dbo.TblDurations.Name AS DurationName, dbo.TblCountriesGovernments.GovernmentName, dbo.TblManagerialArea.Name AS MAName, dbo.TblAttributionContract.IDAC,"
  StrSQL = StrSQL & "  dbo.TblAttributionContract.ProcessNo , dbo.TblAttributionContract.Name , dbo.TblAttributionContract.FromDate ,"
  StrSQL = StrSQL & "  dbo.TblAttributionContract.FromDateH , dbo.TblAttributionContract.ToDate , dbo.TblAttributionContract.ToDateH ,"
  StrSQL = StrSQL & "  dbo.TblAttributionContract.VendorID , dbo.TblAttributionContract.StudentCount , dbo.TblAttributionContract.StudentCustom ,"
  StrSQL = StrSQL & "  dbo.TblAttributionContract.EndContractDate, dbo.TblAttributionContract.EndContractDateH , dbo.TblAttributionContract.MinistryContractNo ,"
  StrSQL = StrSQL & "  dbo.TblAttributionContract.DurationID"
  StrSQL = StrSQL & " , dbo.TblAttributionContract.StartContractDate, dbo.TblAttributionContract.StartContractDateh  "
  StrSQL = StrSQL & "  FROM     dbo.TblAttributionContract INNER JOIN"
  StrSQL = StrSQL & "  dbo.TblDurations ON dbo.TblAttributionContract.DurationID = dbo.TblDurations.ID INNER JOIN"
  StrSQL = StrSQL & "  dbo.TblManagerialArea ON dbo.TblAttributionContract.VendorID = dbo.TblManagerialArea.ID INNER JOIN"
  StrSQL = StrSQL & "  dbo.TblCountriesGovernments ON dbo.TblAttributionContract.CityID = dbo.TblCountriesGovernments.GovernmentID"
    StrSQL = StrSQL & "     where 1 = 1  "
   
   
    If Me.txtProcessNo1.text <> "" Then
            StrSQL = StrSQL & "   and  IDAC =  '" & txtProcessNo1.text & "'"
    End If
    
    If Me.txtMinistryContractNo1.text <> "" Then
            StrSQL = StrSQL & "   and  MinistryContractNo =  '" & txtMinistryContractNo1.text & "'"
    End If

    If Me.XPTxtBoxName1.text <> "" Then
            StrSQL = StrSQL & "   and  TblMinistryContract.Name like  '%" & XPTxtBoxName1.text & "%'"
    End If
    
     If Me.dcDuration1.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblMinistryContract.DurationID =  " & val(Me.dcDuration1.BoundText)
    End If
 
    If Me.dcCity1.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblMinistryContract.CityID =  " & val(Me.dcCity1.BoundText)
    End If
    
     If Me.dcVendor1.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblMinistryContract.VendorID =  " & val(Me.dcVendor1.BoundText)
    End If
     
    If Not IsNull(Me.dtpFromDate1.value) Then
            StrSQL = StrSQL & "   and  TblMinistryContract.FromDate  >=  '" & dtpFromDate1.value & "'"
    End If
    
   If Not IsNull(Me.dtpToDate1.value) Then
            StrSQL = StrSQL & "   and  TblMinistryContract.ToDate  >=  '" & dtpToDate1.value & "'"
    End If
    
     If Not IsNull(dtpSContractDate1.value) Then
            StrSQL = StrSQL & "   and  TblMinistryContract.StartContractDate  >=  '" & dtpSContractDate1.value & "'"
    End If
    
    If Not IsNull(dtpEContractDate1.value) Then
            StrSQL = StrSQL & "   and  TblMinistryContract.EndContractDate  >=  '" & dtpEContractDate1.value & "'"
    End If
    
     
     
    StrSQL = StrSQL
    StrSQL = StrSQL & " Order By IDMC "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            'Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            'Me.Lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.fg_attr
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                'Me.Lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                'Me.Lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                 .TextMatrix(i, .ColIndex("IDMC")) = IIf(IsNull(rs("IDAC").value), "", rs("IDAC").value)
                .TextMatrix(i, .ColIndex("ProcessNo")) = IIf(IsNull(rs("ProcessNo").value), "", rs("ProcessNo").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(i, .ColIndex("MinistryContractNo")) = IIf(IsNull(rs("MinistryContractNo").value), "", rs("MinistryContractNo").value)
                .TextMatrix(i, .ColIndex("DurationName")) = IIf(IsNull(rs("DurationName").value), "", rs("DurationName").value)
                
               .TextMatrix(i, .ColIndex("MAName")) = IIf(IsNull(rs("MAName").value), "", rs("MAName").value)
               .TextMatrix(i, .ColIndex("GovernmentName")) = IIf(IsNull(rs("GovernmentName").value), "", rs("GovernmentName").value)
               .TextMatrix(i, .ColIndex("DurationName")) = IIf(IsNull(rs("DurationName").value), "", rs("DurationName").value)
                .TextMatrix(i, .ColIndex("FromDate")) = IIf(IsNull(rs("FromDate").value), "", rs("FromDate").value)
                .TextMatrix(i, .ColIndex("FromDateH")) = IIf(IsNull(rs("FromDateH").value), "", rs("FromDateH").value)
                .TextMatrix(i, .ColIndex("ToDate")) = IIf(IsNull(rs("ToDate").value), "", rs("ToDate").value)
                .TextMatrix(i, .ColIndex("ToDateH")) = IIf(IsNull(rs("ToDateH").value), "", rs("ToDateH").value)
                .TextMatrix(i, .ColIndex("EndContractDate")) = IIf(IsNull(rs("EndContractDate").value), "", rs("EndContractDate").value)
                .TextMatrix(i, .ColIndex("EndContractDateh")) = IIf(IsNull(rs("EndContractDateh").value), "", rs("EndContractDateh").value)
                .TextMatrix(i, .ColIndex("StartContractDate")) = IIf(IsNull(rs("StartContractDate").value), "", rs("StartContractDate").value)
                .TextMatrix(i, .ColIndex("StartContractDateh")) = IIf(IsNull(rs("StartContractDateh").value), "", rs("StartContractDateh").value)
                
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
    
  Me.Caption = "Search Models"

'Me.LblClientName.Caption = "ClientName"
lbl(4).Caption = "Code"
lbl(3).Caption = "Remarks"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
lbl(0).Caption = "Name Arb"
lbl(8).Caption = " Name ENG"
lbl(2).Caption = "Total"
'Me.lbreg.Caption = "Date Registration"
Me.lbprocess.Caption = "Process No"
     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("code")) = "Code"
       ' .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
         .TextMatrix(0, .ColIndex("name")) = " Name Arb"
        .TextMatrix(0, .ColIndex("namee")) = " Name ENG"
       .TextMatrix(0, .ColIndex("remark")) = "Remarks"
    End With
  '
End Sub

Private Sub TxtIDFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDFrom.text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

Private Sub TxtIDTO_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDTO.text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

