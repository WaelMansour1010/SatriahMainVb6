VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form PlanSearch 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12705
   Icon            =   "PlanSearch.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   6600
   ScaleWidth      =   12705
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.ComboBox CplanType 
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
      ItemData        =   "PlanSearch.frx":000C
      Left            =   5880
      List            =   "PlanSearch.frx":001C
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox txtRemarks 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   31
      Top             =   5400
      Width           =   11610
   End
   Begin VB.Frame Frame2 
      Caption         =   " »‰«¡ ⁄·Ï"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   3960
      Width           =   12495
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "√„— »Ì⁄  "
         Height          =   255
         Index           =   0
         Left            =   10800
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "›Ê« Ì— „»Ì⁄« "
         Height          =   255
         Index           =   1
         Left            =   8760
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "ŒÿÂ ”«»ﬁ…           Õœœ —ﬁ„ «·Œÿ…"
         Height          =   255
         Index           =   3
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   240
         Width           =   3015
      End
      Begin VB.Frame Frame3 
         Caption         =   "Õœœ «·› —… ·«” œ⁄«¡ «·»Ì«‰« "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6360
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   600
         Width           =   5775
         Begin MSComCtl2.DTPicker dbFromDate 
            Height          =   270
            Left            =   3840
            TabIndex        =   24
            Top             =   240
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   476
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   99155969
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker DBTo 
            Height          =   270
            Left            =   1800
            TabIndex        =   25
            Top             =   240
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   476
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   99155969
            CurrentDate     =   38784
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ï"
            Height          =   270
            Index           =   2
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   240
            Width           =   600
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„‰"
            Height          =   270
            Index           =   5
            Left            =   5190
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   240
            Width           =   465
         End
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "ÿ·»«  ‘Õ‰  "
         Height          =   255
         Index           =   2
         Left            =   6840
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox TxtOldPlanNo 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   3615
      End
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   360
      TabIndex        =   9
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
      MICON           =   "PlanSearch.frx":004E
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
      TabIndex        =   8
      Top             =   4680
      Width           =   7830
   End
   Begin VB.TextBox txtorder_no 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   9360
      TabIndex        =   4
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
      Width           =   12585
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
               Picture         =   "PlanSearch.frx":006A
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PlanSearch.frx":0404
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PlanSearch.frx":079E
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PlanSearch.frx":0B38
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PlanSearch.frx":0ED2
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PlanSearch.frx":126C
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PlanSearch.frx":1606
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PlanSearch.frx":1BA0
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ ⁄‰ ŒÿÂ"
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
   Begin MSComCtl2.DTPicker XPDtbBill 
      Height          =   315
      Left            =   3480
      TabIndex        =   5
      Top             =   3600
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   99155969
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   315
      Left            =   12840
      TabIndex        =   6
      Top             =   6120
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
      TabIndex        =   10
      Top             =   720
      Width           =   12675
      _cx             =   22357
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"PlanSearch.frx":1F3A
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
      TabIndex        =   11
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
      TabIndex        =   12
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
      TabIndex        =   13
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      Format          =   99155969
      CurrentDate     =   38784
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   1080
      TabIndex        =   17
      Top             =   3600
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   99155969
      CurrentDate     =   38784
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "„·«ÕŸ« "
      Height          =   195
      Index           =   3
      Left            =   11760
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   5520
      Width           =   705
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "‰Ê⁄ «·ŒÿÂ"
      Height          =   375
      Left            =   7800
      TabIndex        =   19
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "«·Ï"
      Height          =   375
      Left            =   2880
      TabIndex        =   18
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "«· «—ÌŒ"
      Height          =   375
      Left            =   1800
      TabIndex        =   16
      Top             =   0
      Width           =   495
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
      Left            =   14160
      TabIndex        =   7
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "„‰"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·ŒÿÂ"
      Height          =   375
      Left            =   11040
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "PlanSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As ADODB.Recordset
Dim cSearchDcbo As clsDCboSearch

Private m_DcboItems As DataCombo

Private m_RetrunType As Integer
Public WithEvents Fg1 As VSFlex8UCtl.vsFlexGrid
Attribute Fg1.VB_VarHelpID = -1

Public WithEvents NewGrid As VSFlex8UCtl.vsFlexGrid
Attribute NewGrid.VB_VarHelpID = -1
'Public NewGrid As New ClsGrid
 
Public LngRow As Long

Public LngCol As Long
Public TType As Integer




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
            dbFromDate.value = ""
            XPDtbBill.value = ""
            DTPicker2.value = ""
DBTo.value = ""
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







Private Sub Fg_Click()
  '  On Error GoTo ErrTrap
       
    
  
     
   If Me.TType = 0 Then
       FrmProductionPlan.Retrive val(FG.TextMatrix(FG.Row, FG.ColIndex("TbllProductionPlanD")))
   ElseIf Me.TType = 1 Then
          
   FrmShipmentOrder.TxtPONo = val(FG.TextMatrix(FG.Row, FG.ColIndex("TbllProductionPlanD")))
   End If
   
    
        
        
'    Exit Sub
'ErrTrap:
End Sub

Private Sub Retrive()
    Dim Num As Integer
    On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        FG.Rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With FG
               ' .TextMatrix(Num, .ColIndex("order_no")) = IIf(IsNull(rs("TbllProductionPlanD").value), "", rs("TbllProductionPlanD").value)
                .TextMatrix(Num, .ColIndex("TbllProductionPlanD")) = IIf(IsNull(rs("TbllProductionPlanD").value), "", rs("TbllProductionPlanD").value)
             .TextMatrix(Num, .ColIndex("DbFrom")) = IIf(IsNull(rs("DbFrom").value), "", (rs("DbFrom").value))
              .TextMatrix(Num, .ColIndex("Todate")) = IIf(IsNull(rs("Todate").value), "", (rs("Todate").value))
             .TextMatrix(Num, .ColIndex("FromDate")) = IIf(IsNull(rs("FromDate").value), "", (rs("FromDate").value))
              .TextMatrix(Num, .ColIndex("DBTo")) = IIf(IsNull(rs("DBTo").value), "", (rs("DBTo").value))
                   .TextMatrix(Num, .ColIndex("Remarks")) = IIf(IsNull(rs("Remarks").value), "", (rs("Remarks").value))
                .TextMatrix(Num, .ColIndex("OldPlanNo")) = IIf(IsNull(rs("OldPlanNo").value), "", (rs("OldPlanNo").value))

            Me.CplanType.ListIndex = val(IIf(IsNull(rs("planType").value), -1, (rs("planType").value)))
  .TextMatrix(Num, .ColIndex("planType")) = Me.CplanType.text
               
               If (rs("Opt").value) = 0 Or IsNull((rs("Opt").value)) = True Then
.TextMatrix(Num, .ColIndex("Opt")) = "√„— »Ì⁄"
               ElseIf (rs("Opt").value) = 1 Then
               .TextMatrix(Num, .ColIndex("Opt")) = "›Ê« Ì— „»Ì⁄« "
                ElseIf (rs("Opt").value) = 2 Then
               .TextMatrix(Num, .ColIndex("Opt")) = "ÿ·»«  ‘Õ‰"
                ElseIf (rs("Opt").value) = 3 Then
               .TextMatrix(Num, .ColIndex("Opt")) = "ŒÿÂ ”«»ﬁ…"
            End If
                '  .TextMatrix(Num, .ColIndex("Opt")) = IIf(IsNull(rs("Transaction_Date").value), "", Trim(rs("currency_code").value))
           
                '  .TextMatrix(Num, .ColIndex("countryid")) = IIf(IsNull(rs("countryid").value), "", (rs("countryid").value))
                '    .TextMatrix(Num, .ColIndex("CountryName")) = IIf(IsNull(rs("CountryName").value), "", Trim(rs("CountryName").value))
            
            End With

            rs.MoveNext
        Next Num

          FG.AutoSize 0, FG.Cols - 1, False
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
   'Me.CplanType.AddItem "Œÿ… „‘ —Ì« "
'Me.CplanType.AddItem "Œÿ… ≈‰ «Ã"
'Me.CplanType.AddItem "Œÿ… „»Ì⁄« "
'Me.CplanType.AddItem " Œÿ… ‘Õ‰"
            dbFromDate.value = ""
            XPDtbBill.value = ""
            DTPicker2.value = ""
DBTo.value = ""

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
 
    Dim My_SQL As String
    Set Dcombos = New ClsDataCombos
  '  Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
   ' Dcombos.GetItemsNames DCboItem, , , , True
    
  '  My_SQL = " select CountryID,CountryName from TblCountriesData"
 
  '  fill_combo Me.DataCombo4, My_SQL
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

StrSQL = " SELECT     TbllProductionPlanD, planType, BranchId, Todate1, FromDate1, DbFrom, DBTo, Opt, OldPlanNo, Remarks, FromDate, Todate, Locked"
StrSQL = StrSQL & " From dbo.TbllProductionPlan"
    
    StrSQL = StrSQL + " WHERE     (1 = 1)"
 
    If Me.txtorder_no.text <> "" Then
        '     FrmProductionOrder1.Retrive (Val(FG.TextMatrix(FG.Row, 3)))
     
        StrWhere = StrWhere + " and     (dbo.TbllProductionPlan.TbllProductionPlanD = " & val(txtorder_no.text) & ")"
    
    End If
    
    If Me.CplanType.text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TbllProductionPlan.planType =" & val(Me.CplanType.ListIndex) & ""
 
    End If
  If Me.Opt(0).value = True Then
 
        StrWhere = StrWhere + " and dbo.TbllProductionPlan.Opt = 0"
 
    End If
    If Me.Opt(1).value = True Then
 
        StrWhere = StrWhere + " and dbo.TbllProductionPlan.Opt = 1"
 
    End If
    If Me.Opt(2).value = True Then
 
        StrWhere = StrWhere + " and dbo.TbllProductionPlan.Opt = 2"
 
    End If
    If Me.Opt(3).value = True Then
 
        StrWhere = StrWhere + " and dbo.TbllProductionPlan.Opt = 3"
 
    End If
    If TxtRemarks.text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TbllProductionPlan.Remarks like '%" & Me.TxtRemarks.text & "%'"
 
    End If
    
    
     If Me.TxtOldPlanNo.text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TbllProductionPlan.OldPlanNo ='" & Me.TxtOldPlanNo.text & "'"
 
    End If
 If Not IsNull(Me.XPDtbBill.value) Then
        
            StrWhere = StrWhere & " AND dbo.TbllProductionPlan.FromDate >=" & SQLDate(Me.XPDtbBill.value, True) & ""
      End If
       If Not IsNull(Me.XPDtbBill.value) Then
        
            StrWhere = StrWhere & " AND dbo.TbllProductionPlan.DbFrom >=" & SQLDate(Me.dbFromDate.value, True) & ""
      End If
       If Not IsNull(Me.DTPicker2.value) Then
        
            StrWhere = StrWhere & " AND dbo.TbllProductionPlan.Todate <=" & SQLDate(Me.DTPicker2.value, True) & ""
      End If
       If Not IsNull(Me.DBTo.value) Then
        
            StrWhere = StrWhere & " AND dbo.TbllProductionPlan.DBTo <=" & SQLDate(Me.DBTo.value, True) & ""
      End If
    StrWhere = StrWhere + " order by dbo.TbllProductionPlan.TbllProductionPlanD"

    Build_Sql = StrSQL + StrWhere
    Exit Function
ErrTrap:
End Function

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl Is FG Then
            If Not FG.TextMatrix(FG.Row, 1) = "" Then
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





Public Property Get RetrunType() As Integer
    RetrunType = m_RetrunType
End Property

Public Property Let RetrunType(ByVal vNewValue As Integer)
    m_RetrunType = vNewValue
    ' 0 = Retrun in the Items Screen
    ' 1 = Retrun in the Data Combo
End Property

Private Sub ChangeLang()
    Me.Caption = "Search For Plan"
    Label1(2).Caption = Me.Caption
    Label2.Caption = "Plan No"
 Label9.Caption = "Type Plan"
    Label3.Caption = "From"
    Label8.Caption = "To"
    Frame2.Caption = "By"
 lbl(5).Caption = "From"
lbl(2).Caption = "To"
Opt(3).RightToLeft = True
Opt(3).Caption = "Previous plan"

Opt(2).RightToLeft = True
Opt(2).Caption = "Shipping Requests"
Opt(1).RightToLeft = True
Opt(1).Caption = "Sales Invoices "
Opt(0).RightToLeft = True
Frame3.Caption = ""
Opt(0).Caption = "Sell Order "
lbl(3).Caption = "Remarks"
    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"

    'OptType(0).Caption = "Start of the name"
    'OptType(1).Caption = "any part of the name"
    With Me.FG
        .TextMatrix(0, .ColIndex("TbllProductionPlanD")) = "Plan No"
         .TextMatrix(0, .ColIndex("planType")) = "PlanType  "
        .TextMatrix(0, .ColIndex("DbFrom")) = "From"
        .TextMatrix(0, .ColIndex("Todate")) = " To"
             .TextMatrix(0, .ColIndex("Opt")) = "Based on "
  .TextMatrix(0, .ColIndex("OldPlanNo")) = "Old Plan No"
  .TextMatrix(0, .ColIndex("FromDate")) = " From"
  .TextMatrix(0, .ColIndex("DBTo")) = " To"
  .TextMatrix(0, .ColIndex("Remarks")) = " Remarks"
  
  
  
        '  .AutoSize 0, .Cols - 1, False
    End With

End Sub



'Private Sub TxtItemCode_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'          If KeyCode = vbKeyReturn Then
'                If Trim(Me.TxtItemCode(1).text) = "" Then Exit Sub
'                StrSQL = "Select ItemID From TblItems Where ItemCode='" & Trim(Me.TxtItemCode(Index).text) & "'"
'                Set rs = New ADODB.Recordset
'                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
''
 '               If Not (rs.BOF Or rs.EOF) Then
 '                   DCboItem.BoundText = rs("ItemID").value
 '               Else
 '                   Msg = "·«ÌÊÃœ ’‰› „”Ã· »Â–« «·ﬂÊœ..!"
 '                   MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
 '               End If
 '           End If
 '
'End Sub
