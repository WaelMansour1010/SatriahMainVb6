VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmStagnantItems 
   Caption         =   "«·√’‰«› «·—«ﬂœ… "
   ClientHeight    =   8895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11040
   Icon            =   "FrmStagnantItems.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8895
   ScaleWidth      =   11040
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8895
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11040
      _cx             =   19473
      _cy             =   15690
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   14871017
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   2
      ChildSpacing    =   1
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   6
      GridCols        =   5
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmStagnantItems.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1125
         Index           =   3
         Left            =   2145
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   30
         Width           =   2100
         _cx             =   3704
         _cy             =   1984
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "ÿ—Ìﬁ… «·⁄—÷"
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   6
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   1
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
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
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄—÷ ÃœÊ·Ï"
            Height          =   195
            Index           =   3
            Left            =   870
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   330
            Value           =   -1  'True
            Width           =   1170
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄—÷ ‘Ã—Ì"
            Height          =   315
            Index           =   2
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   630
            Width           =   1200
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1125
         Index           =   2
         Left            =   4260
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   30
         Visible         =   0   'False
         Width           =   3225
         _cx             =   5689
         _cy             =   1984
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "ÿ—Ìﬁ… Õ”«»  ﬂ·›… «·„Œ“Ê‰"
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   6
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   1
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
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
         Begin VB.ComboBox CboCostType 
            Height          =   315
            Left            =   90
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   540
            Width           =   2355
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1125
         Index           =   1
         Left            =   7500
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   30
         Width           =   3510
         _cx             =   6191
         _cy             =   1984
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
         ForeColor       =   128
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   " ÕœÌœ  «—ÌŒ «·»ÕÀ "
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   6
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   1
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
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
         Begin VB.Frame Frame1 
            Caption         =   "Frame1"
            Height          =   465
            Left            =   150
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   180
            Visible         =   0   'False
            Width           =   1365
            Begin VB.OptionButton Option2 
               Alignment       =   1  'Right Justify
               Caption         =   "„»Ì⁄«"
               Height          =   225
               Left            =   180
               RightToLeft     =   -1  'True
               TabIndex        =   22
               Top             =   270
               Width           =   315
            End
            Begin VB.OptionButton Option1 
               Alignment       =   1  'Right Justify
               Caption         =   "—«ﬂœ"
               Height          =   255
               Left            =   840
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Top             =   210
               Value           =   -1  'True
               Width           =   375
            End
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰– »œ«Ì… «·»—‰«„Ã"
            Height          =   405
            Index           =   1
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   240
            Width           =   1545
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ «· «—ÌŒ"
            Height          =   405
            Index           =   0
            Left            =   2310
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   600
            Value           =   -1  'True
            Width           =   1035
         End
         Begin MSComCtl2.DTPicker DtpSelect 
            Height          =   345
            Left            =   780
            TabIndex        =   10
            Top             =   630
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   609
            _Version        =   393216
            Format          =   166854657
            CurrentDate     =   39209
         End
      End
      Begin MSComctlLib.ProgressBar PrgBar 
         Height          =   360
         Left            =   30
         TabIndex        =   7
         Top             =   7935
         Width           =   10980
         _ExtentX        =   19368
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
      End
      Begin ImpulseButton.ISButton CmdPrint 
         Height          =   510
         Left            =   30
         TabIndex        =   1
         Top             =   645
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   900
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
         ButtonImage     =   "FrmStagnantItems.frx":0430
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   7125
         Left            =   30
         TabIndex        =   2
         Top             =   1170
         Width           =   10980
         _cx             =   19368
         _cy             =   12568
         Appearance      =   2
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmStagnantItems.frx":07CA
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   555
         Index           =   0
         Left            =   30
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   8310
         Width           =   10980
         _cx             =   19368
         _cy             =   979
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
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
         Begin ImpulseButton.ISButton CmdExit 
            Height          =   360
            Left            =   0
            TabIndex        =   4
            Top             =   90
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   635
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
            ButtonImage     =   "FrmStagnantItems.frx":0992
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
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   2
            Left            =   5175
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   90
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ï  ﬂ·›… «·„Œ“Ê‰ «·—«ﬂœ"
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   1
            Left            =   6345
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   120
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   3
            Left            =   8685
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   90
            Width           =   1110
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·√’‰«›:-"
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   0
            Left            =   9810
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   90
            Width           =   1095
         End
         Begin VB.Image Img 
            Height          =   240
            Left            =   2775
            Picture         =   "FrmStagnantItems.frx":0D2C
            Top             =   150
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin ImpulseButton.ISButton CmdDo 
         Height          =   600
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   1058
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   " ‰›Ì–"
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
         ButtonImage     =   "FrmStagnantItems.frx":10B6
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
   End
End
Attribute VB_Name = "FrmStagnantItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CmdDo_Click()

    If Me.Opt(3).value = True Then
        LoadTableData
    ElseIf Me.Opt(2).value = True Then
        LoadTreeGrid
    End If

End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPrint_Click()
    Dim Msg As String
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer, j As Integer
    Dim cItemsReport As ClsItemsReport
    Dim XNode As VSFlex8UCtl.VSFlexNode
    Dim IntReportStyle As Integer

    If Me.Opt(2).value = False Then
        If ItemsInGrid(Me.Fg, Fg.ColIndex("ItemID")) = -1 Then
            Msg = "ÌÃ»  ÕœÌœ «·√’‰«› √Ê·« ...!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
    End If

    StrSQL = "Delete  From TempPrintStagnantItems"
    Cn.Execute StrSQL, , adExecuteNoRecords

    Set rs = New ADODB.Recordset
    rs.Open "TempPrintStagnantItems", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    With Me.Fg

        If Me.Opt(2).value = True Then

            'Save Tree Style
            For i = .FixedRows To .Rows - 1

                If .IsSubtotal(i) = False Then
                    rs.AddNew
                    rs("ItemID").value = IIf(.TextMatrix(i, .ColIndex("ItemID")) = "", Null, .TextMatrix(i, .ColIndex("ItemID")))
                    rs("ItemCode").value = IIf(.TextMatrix(i, .ColIndex("ItemCode")) = "", Null, .TextMatrix(i, .ColIndex("ItemCode")))
                    rs("ItemName").value = IIf(.TextMatrix(i, .ColIndex("ItemName")) = "", Null, .TextMatrix(i, .ColIndex("ItemName")))
                    rs("Qty").value = IIf(.TextMatrix(i, .ColIndex("Qty")) = "", Null, .TextMatrix(i, .ColIndex("Qty")))
                    rs("ItemCostPrice").value = IIf(.TextMatrix(i, .ColIndex("ItemCostPrice")) = "", Null, .TextMatrix(i, .ColIndex("ItemCostPrice")))
                    rs("StockCost").value = IIf(.TextMatrix(i, .ColIndex("StockCost")) = "", Null, .TextMatrix(i, .ColIndex("StockCost")))
                    rs("InvTransID").value = IIf(.TextMatrix(i, .ColIndex("InvTransID")) = "", Null, .TextMatrix(i, .ColIndex("InvTransID")))
                    rs("InvTransSerial").value = IIf(.TextMatrix(i, .ColIndex("InvTransSerial")) = "", Null, .TextMatrix(i, .ColIndex("InvTransSerial")))
                    rs("InvTransDate").value = IIf(.TextMatrix(i, .ColIndex("InvTransDate")) = "", Null, .TextMatrix(i, .ColIndex("InvTransDate")))
                    rs("ID").value = i

                    If Me.Opt(2).value = True Then
                        Set XNode = Fg.GetNode(i)

                        If Not XNode Is Nothing Then
                            rs("GroupID").value = val(XNode.Key)
                            rs("GroupName").value = XNode.text
                            rs("ParentGroupID").value = val(XNode.GetNode(flexNTParent).Key)
                        End If
                    End If

                    rs.update
                Else
                    Set XNode = Fg.GetNode(i)

                    If Not XNode Is Nothing Then
                        rs.AddNew
                        rs("GroupID").value = val(XNode.Key)
                        rs("GroupName").value = XNode.text

                        If Not XNode.GetNode(flexNTParent) Is Nothing Then
                            rs("ParentGroupID").value = val(XNode.GetNode(flexNTParent).Key)
                        Else
                            rs("ParentGroupID").value = Null
                        End If

                        rs("ID").value = i
                        rs.update
                    End If
                End If

            Next i

        Else

            For i = .FixedRows To .Rows - 1

                If .IsSubtotal(i) = False Then
                    rs.AddNew
                    rs("ItemID").value = IIf(.TextMatrix(i, .ColIndex("ItemID")) = "", Null, .TextMatrix(i, .ColIndex("ItemID")))
                    rs("ItemCode").value = IIf(.TextMatrix(i, .ColIndex("ItemCode")) = "", Null, .TextMatrix(i, .ColIndex("ItemCode")))
                    rs("ItemName").value = IIf(.TextMatrix(i, .ColIndex("ItemName")) = "", Null, .TextMatrix(i, .ColIndex("ItemName")))
                    rs("Qty").value = IIf(.TextMatrix(i, .ColIndex("Qty")) = "", Null, .TextMatrix(i, .ColIndex("Qty")))
                    rs("ItemCostPrice").value = IIf(.TextMatrix(i, .ColIndex("ItemCostPrice")) = "", Null, .TextMatrix(i, .ColIndex("ItemCostPrice")))
                    rs("StockCost").value = IIf(.TextMatrix(i, .ColIndex("StockCost")) = "", Null, .TextMatrix(i, .ColIndex("StockCost")))
                    rs("InvTransID").value = IIf(.TextMatrix(i, .ColIndex("InvTransID")) = "", Null, .TextMatrix(i, .ColIndex("InvTransID")))
                    rs("InvTransSerial").value = IIf(.TextMatrix(i, .ColIndex("InvTransSerial")) = "", Null, .TextMatrix(i, .ColIndex("InvTransSerial")))
                    rs("InvTransDate").value = IIf(.TextMatrix(i, .ColIndex("InvTransDate")) = "", Null, .TextMatrix(i, .ColIndex("InvTransDate")))
                    rs("ID").value = i
                    rs.update
                End If

            Next i

        End If

    End With

    rs.Close
    Set rs = Nothing
    Set cItemsReport = New ClsItemsReport

    If Me.Opt(2).value = True Then
        IntReportStyle = 1
    Else
        IntReportStyle = 0
    End If

    If Me.Opt(0).value = True Then
        cItemsReport.ShowStagnantItems DisplayDate(DtpSelect), IntReportStyle
    Else
        cItemsReport.ShowStagnantItems "", IntReportStyle
    End If

    Set cItemsReport = Nothing
End Sub

Private Sub Fg_DblClick()
'    Dim LngItemID As Long
'    Dim Frm As FrmShowItemCostPrice
'
'    With Me.Fg
'
'        If .Col = -1 Then Exit Sub
'        If .Row = -1 Then Exit Sub
'        LngItemID = val(.TextMatrix(.Row, .ColIndex("ItemID")))
'
'        If val(.TextMatrix(.Row, .ColIndex("ItemID"))) <> 0 Then
'            If .Col = .ColIndex("ItemID") Or .Col = .ColIndex("ItemCode") Or .Col = .ColIndex("ItemName") Then
'                Load FrmSelectData
'                FrmSelectData.DcboItemName.BoundText = val(.TextMatrix(.Row, .ColIndex("ItemID")))
'                FrmSelectData.TxtItemCode.text = (.TextMatrix(.Row, .ColIndex("ItemCode")))
'                FrmSelectData.show
'            ElseIf .Col = .ColIndex("Qty") Then
'                Load FrmSearchSerial
'                FrmSearchSerial.XPTxtCode.text = val(.TextMatrix(.Row, .ColIndex("ItemCode")))
'                FrmSearchSerial.DCboItemsName.BoundText = val(.TextMatrix(.Row, .ColIndex("ItemID")))
'                FrmSearchSerial.show vbModal
'            ElseIf .Col = .ColIndex("InvTransSerial") Or .Col = .ColIndex("InvTransDate") Then
'
'                If val(.TextMatrix(.Row, .ColIndex("InvTransID"))) <> 0 Then
'                    If checkApility("FrmSaleBill", True) = True Then
'                        Load frmsalebill
'                        frmsalebill.show
'                        frmsalebill.Retrive val(.TextMatrix(.Row, .ColIndex("InvTransID")))
'                        frmsalebill.ZOrder 0
'                    End If
'                End If
'
'            ElseIf .Col = .ColIndex("ItemCostPrice") Or .Col = .ColIndex("StockCost") Then
'                Set Frm = New FrmShowItemCostPrice
'                Me.MousePointer = vbArrowHourglass
'                Frm.LoadData LngItemID, val(.TextMatrix(.Row, .ColIndex("Qty")))
'                Me.MousePointer = vbDefault
'                Frm.show
'            End If
'        End If
    
'    End With

End Sub

Private Sub Fg_MouseMove(Button As Integer, _
                         Shift As Integer, _
                         X As Single, _
                         Y As Single)
    Static LngOldCol As Long
    Dim LngCol As Long

    With Me.Fg

        If .MouseRow = -1 Or .MouseCol = -1 Then
            .ToolTipText = ""
            Exit Sub
        Else
            LngCol = .MouseCol
        End If

        If val(.TextMatrix(.MouseRow, .ColIndex("ItemID"))) = 0 Then
            .ToolTipText = ""
            Exit Sub
        End If

        If LngCol = LngOldCol Then
            Exit Sub
        Else
            LngOldCol = LngCol
        End If

        Select Case LngCol

            Case .ColIndex("ItemID"), .ColIndex("ItemCode"), .ColIndex("ItemName")
                .ToolTipText = "»«·÷€ÿ „— Ì‰ „  «·Ì Ì‰ Ì„ﬂ‰ﬂ ⁄—÷  ﬁ—Ì— ﬂ«—  «·’‰› ·Â–« «·’‰› "
                Exit Sub

            Case .ColIndex("Qty")
                .ToolTipText = "»«·÷€ÿ „— Ì‰ „  «·Ì Ì‰ Ì„ﬂ‰ﬂ ⁄—÷ ≈” ⁄·«„ ﬂ„Ì«  Â–« «·’‰›"
                Exit Sub

            Case .ColIndex("InvTransSerial"), .ColIndex("InvTransDate")
                .ToolTipText = "»«·÷€ÿ „— Ì‰ „  «·Ì Ì‰ Ì„ﬂ‰ﬂ ⁄—÷ Â–Â «·›« Ê—…"
                Exit Sub

            Case Else
                .ToolTipText = ""
        End Select

    End With

End Sub

Private Sub Fg_MouseUp(Button As Integer, _
                       Shift As Integer, _
                       X As Single, _
                       Y As Single)
    Dim LngCurrentItemID As Long
    Dim LngMouseRow As Long

    If Button = vbRightButton Then

        With Me.Fg
            LngMouseRow = .MouseRow

            If LngMouseRow = -1 Then Exit Sub
            If .Col = -1 Then Exit Sub
            mdifrmmain.MnuItemTools_ItemSerial.Tag = ""
            mdifrmmain.MnuItemTools_ItemCart.Tag = ""
            mdifrmmain.MnuItemTools_ItemData.Tag = ""
            mdifrmmain.MnuItemTools_ItemQty.Tag = ""
        
            If val(.TextMatrix(LngMouseRow, .ColIndex("ItemID"))) <> 0 Then
                LngCurrentItemID = val(.TextMatrix(LngMouseRow, .ColIndex("ItemID")))
                mdifrmmain.MnuItemTools_ItemSerial.Enabled = False
                mdifrmmain.MnuItemTools_ItemSerial.Tag = ""
            
                mdifrmmain.MnuItemTools_ItemCart.Tag = LngCurrentItemID & "-" & ""
                mdifrmmain.MnuItemTools_ItemQty.Tag = LngCurrentItemID
                mdifrmmain.MnuItemTools_ItemData.Tag = LngCurrentItemID
                Me.PopupMenu mdifrmmain.MnuItemTools
            End If

        End With

    End If

End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Set GrdBack = New ClsBackGroundPic

    With Me.Fg
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    SetDtpickerDate Me.DtpSelect

    With Me.CboCostType
        .Clear
        .AddItem "«Œ— ”⁄— ‘—«¡"
        .ItemData(0) = 1
        .AddItem "«·„ Ê”ÿ «·„—ÃÕ"
        .ItemData(1) = 2
        .AddItem "«·„ Ê”ÿ «·„—ÃÕ «·ÃœÌœ"
        .ItemData(2) = 3
        '    .AddItem "«·Ê«—œ √Ê·« Ì’—› √Ê·«"
        '    .ItemData(2) = 3
        '    .AddItem "«·Ê«—œ √ŒÌ— Ì’—› √Ê·«"
        '    .ItemData(3) = 4
        '    .AddItem "«Œ— ”⁄— »Ì⁄"
        '    .ItemData(4) = 5
    End With
Option1_Click
    Me.Height = 9240
    Me.Width = 11100

    Resize_Form Me
End Sub

Private Function GetSaledItemsIDs() As String
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Dim StrTemp As String

    On Error GoTo ErrTrap

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = " SELECT DISTINCT TOP 100 PERCENT dbo.Transaction_Details.Item_ID "
        StrSQL = StrSQL + " FROM dbo.Transactions INNER JOIN"
        StrSQL = StrSQL + " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID"
        StrSQL = StrSQL + " = dbo.Transaction_Details.Transaction_ID"
        StrSQL = StrSQL + " Where (dbo.Transactions.Transaction_Type = 2 OR " & "dbo.Transactions.Transaction_Type = 1 or dbo.Transactions.Transaction_Type = 22  or dbo.Transactions.Transaction_Type = 21)"

        If Me.Opt(0).value = True Then
            StrSQL = StrSQL + " AND (dbo.Transactions.Transaction_Date > "
            StrSQL = StrSQL + SQLDate(Me.DtpSelect.value - 1, True) & ")"
        End If

        StrSQL = StrSQL + " ORDER BY dbo.Transaction_Details.Item_ID "
    Else
    
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    StrTemp = ""

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst

        For i = 0 To rs.RecordCount - 1
            StrTemp = StrTemp & rs("Item_ID").value & ","
            rs.MoveNext
        Next i

    End If

    If StrTemp <> "" Then
        StrTemp = mId(StrTemp, 1, Len(StrTemp) - 1)
    End If

    GetSaledItemsIDs = StrTemp
    Exit Function
ErrTrap:
    GetSaledItemsIDs = ""
End Function

Private Sub ChangeLang()
    Me.Caption = "Stagnant Items"
    Me.CmdExit.Caption = "Exit"
    Me.CmdPrint.Caption = "Print"
    Me.CmdDo.Caption = "Select"
    'Me.lbl(1).Caption = "Stagnant Items Since"
    'Me.lbl(2).Caption = "Day"
    Me.Opt(1).Caption = "From the Programe Start"
    Me.lbl(0).Caption = "Items Count:-"

End Sub

Private Sub LoadTableData()
    Dim i As Integer
    Dim Msg As String
    Dim StrSaledItemsID As String
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim LastItemInv As LastItemTransInfo
    Dim DblItemCost As Double
    Dim IntCostPriceType As StockCostType

    If Me.Opt(0).value = True Then
        If Me.DtpSelect.value = Date Then
            Msg = "ÌÃ»  ÕœÌœ «· «—ÌŒ ..!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
    End If

'    If Me.CboCostType.ListIndex = -1 Then
'        Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ…  ﬂ·›… «·„Œ“Ê‰..!!"
'        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'        CboCostType.SetFocus
'        Sendkeys "{F4}"
'        Exit Sub
'    ElseIf Me.CboCostType.ListIndex = 0 Then
'        IntCostPriceType = LastPurPriceType
'    ElseIf Me.CboCostType.ListIndex = 1 Then
'        IntCostPriceType = WeightAverage
'    ElseIf Me.CboCostType.ListIndex = 2 Then
'        IntCostPriceType = ModernWeightAverage
'    ElseIf Me.CboCostType.ListIndex = 3 Then
'        IntCostPriceType = FirstInFirstOut
'    ElseIf Me.CboCostType.ListIndex = 4 Then
'        IntCostPriceType = LastInFirstOut
'    ElseIf Me.CboCostType.ListIndex = 5 Then
'        IntCostPriceType = LastSalesPriceType
'    End If
IntCostPriceType = ModernWeightAverage
    Me.MousePointer = vbArrowHourglass
    StrSaledItemsID = GetSaledItemsIDs

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        
        If Option1.value = True Then
        
            StrSQL = "select ItemID,ItemCode,ItemName,Sum(Qty) as Qty1 From " & "dbo.QryGARDShort() QryGARDShort"
            StrSQL = StrSQL + " Where Qty > 0 "
    
            If StrSaledItemsID <> "" Then
                StrSQL = StrSQL + " AND ItemID  NOT IN "
                StrSQL = StrSQL + "("
                StrSQL = StrSQL + StrSaledItemsID
                StrSQL = StrSQL + ")"
            End If
    
            StrSQL = StrSQL + " Group BY ItemID,ItemCode,ItemName"
            StrSQL = StrSQL + " Order By Qty1 desc"
        Else
            StrSQL = " SELECT * FROM ("
            StrSQL = StrSQL + " SELECT"
                
            StrSQL = StrSQL + " SUM(dbo.Transaction_Details.SHOWQTY )    as     Qty1"
            StrSQL = StrSQL + "                ,dbo.TblItems.ItemID"
            StrSQL = StrSQL + "    ,dbo.TblItems.ItemCode"
               
            StrSQL = StrSQL + "                ,dbo.TblItems.ItemName"
            StrSQL = StrSQL + "                ,dbo.TblItems.GroupID"
            StrSQL = StrSQL + "                ,dbo.Groups.GroupName"
            StrSQL = StrSQL + "                ,dbo.Groups.GroupNamee"
            StrSQL = StrSQL + "                ,dbo.TblItems.ItemNameE"
            StrSQL = StrSQL + "             From dbo.transactions"
            StrSQL = StrSQL + "             INNER JOIN dbo.Transaction_Details"
            StrSQL = StrSQL + "                 ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
            StrSQL = StrSQL + "             INNER JOIN dbo.TblStore"
            StrSQL = StrSQL + "                 ON dbo.Transactions.StoreID = dbo.TblStore.StoreID"
            StrSQL = StrSQL + "             INNER JOIN dbo.TblCustemers"
            StrSQL = StrSQL + "                 ON dbo.Transactions.CusID = dbo.TblCustemers.CusID"
            StrSQL = StrSQL + "                 INNER JOIN dbo.TransactionTypes"
            StrSQL = StrSQL + "                 ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
            StrSQL = StrSQL + "             INNER JOIN dbo.TblItems"
            StrSQL = StrSQL + "                 ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
            StrSQL = StrSQL + "             INNER JOIN dbo.Groups"
            StrSQL = StrSQL + "                 ON dbo.TblItems.GroupID = dbo.Groups.GroupID"
            StrSQL = StrSQL + "             LEFT OUTER JOIN dbo.TblCustemers TblCustemers_1"
            StrSQL = StrSQL + "                 ON dbo.TblItems.DefaultSupplier = TblCustemers_1.CusID"
            StrSQL = StrSQL + "             LEFT OUTER JOIN dbo.TblBranchesData"
            StrSQL = StrSQL + "                 ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
            StrSQL = StrSQL + "             LEFT OUTER JOIN dbo.TblUnites"
            StrSQL = StrSQL + "                 ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID"
            StrSQL = StrSQL + "             LEFT OUTER JOIN dbo.TblEmployee"
            StrSQL = StrSQL + "                 ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID"
            StrSQL = StrSQL + "             LEFT OUTER JOIN dbo.TblEmployee Emp2"
            StrSQL = StrSQL + "                 ON dbo.Transactions.EmpOrderInstitute = Emp2.Emp_ID"
            StrSQL = StrSQL + "             INNER JOIN dbo.Groups AS Groups_1"
            StrSQL = StrSQL + "                 ON dbo.Groups.ParentID = Groups_1.GroupID"
            StrSQL = StrSQL + "             Where (dbo.transactions.Transaction_Type = 21)"
'            StrSQL = StrSQL + "             AND Transactions.Transaction_Date >= '01-Jan-2022'"
'            StrSQL = StrSQL + "             AND Transactions.Transaction_Date <= '05-Jan-2022'"
'            StrSQL = StrSQL + "             "
            
            If Me.Opt(0).value = True Then
                StrSQL = StrSQL + " AND (dbo.Transactions.Transaction_Date > "
                StrSQL = StrSQL + SQLDate(Me.DtpSelect.value - 1, True) & ")"
            End If

            StrSQL = StrSQL + "             GROUP BY dbo.TblItems.ItemID"
            StrSQL = StrSQL + "                ,dbo.TblItems.ItemCode"
               
            StrSQL = StrSQL + "                ,dbo.TblItems.ItemName"
            StrSQL = StrSQL + "                ,dbo.TblItems.GroupID"
            StrSQL = StrSQL + "                ,dbo.Groups.GroupName"
            StrSQL = StrSQL + "                ,dbo.Groups.GroupNamee"
            StrSQL = StrSQL + "                ,dbo.TblItems.ItemNameE   ) T ORDER BY T.Qty1 desc"
        End If
    Else
        Exit Sub
    End If

    CmdDo.Enabled = False
    CmdPrint.Enabled = False

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.Fg
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows
        Me.PrgBar.value = 0
        Me.lbl(3).Caption = "0"
        Me.lbl(2).Caption = "0"

        If Not (rs.BOF Or rs.EOF) Then
            rs.MoveFirst
            .Rows = .FixedRows + rs.RecordCount
            Me.PrgBar.Visible = True
            Me.PrgBar.Max = rs.RecordCount

            'Me.lbl(3).Caption = Rs.RecordCount
            For i = .FixedRows To rs.RecordCount

                DoEvents
                Me.PrgBar.value = i
                Me.lbl(3).Caption = i
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                .TextMatrix(i, .ColIndex("Qty")) = IIf(IsNull(rs("Qty1").value), 0, rs("Qty1").value)
                'Calculate The Cost
              '  DblItemCost = GetCostItemPrice(val(.TextMatrix(i, .ColIndex("ItemID"))), 2, , , IntCostPriceType)
                .TextMatrix(i, .ColIndex("ItemCostPrice")) = DblItemCost
                .TextMatrix(i, .ColIndex("StockCost")) = (DblItemCost * val(.TextMatrix(i, .ColIndex("Qty"))))
                LastItemInv = GetLastItemTrans(rs("ItemID").value, 21)
                .TextMatrix(i, .ColIndex("InvTransID")) = LastItemInv.Transactionid
                .TextMatrix(i, .ColIndex("InvTransDate")) = LastItemInv.TransactionDate
                .TextMatrix(i, .ColIndex("InvTransSerial")) = LastItemInv.TransactionSerial
                rs.MoveNext
            Next i
        
        End If

        Me.PrgBar.value = 0
        Me.PrgBar.Visible = False
        .AutoSize 0, .Cols - 1, False

        If .Rows > .FixedRows Then
            Me.lbl(2).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("StockCost"), .Rows - 1, .ColIndex("StockCost"))
        Else
            Me.lbl(2).Caption = 0
        End If

    End With

    CmdDo.Enabled = True
    CmdPrint.Enabled = True
    Me.MousePointer = vbDefault
End Sub

Private Sub LoadTreeGrid()
    Dim My_SQL As String
    Dim i As Integer
    Dim IntColName As Integer
    Dim BolRtl As Boolean
    Dim RsData As ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim StrSQL As String
    Dim ReCount As Long
    Dim LngRowNum As Long
    Dim LngParentRow As Long
    Dim DblNodeChildCount As Double
    Dim Msg As String
    Dim StrSaledItemsID As String
    Dim rs As ADODB.Recordset
    Dim LastItemInv As LastItemTransInfo
    Dim DblItemCost As Double
    Dim IntCostPriceType As StockCostType

    On Error GoTo ErrTrap

    If Me.Opt(0).value = True Then
        If Me.DtpSelect.value = Date Then
            Msg = "ÌÃ»  ÕœÌœ «· «—ÌŒ ..!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
    End If

    If Me.CboCostType.ListIndex = -1 Then
        Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ…  ﬂ·›… «·„Œ“Ê‰..!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CboCostType.SetFocus
        Sendkeys "{F4}"
        Exit Sub
    ElseIf Me.CboCostType.ListIndex = 0 Then
        IntCostPriceType = LastPurPriceType
    ElseIf Me.CboCostType.ListIndex = 1 Then
        IntCostPriceType = WeightAverage
    ElseIf Me.CboCostType.ListIndex = 2 Then
        IntCostPriceType = ModernWeightAverage
    ElseIf Me.CboCostType.ListIndex = 3 Then
        IntCostPriceType = FirstInFirstOut
    ElseIf Me.CboCostType.ListIndex = 4 Then
        IntCostPriceType = LastInFirstOut
    ElseIf Me.CboCostType.ListIndex = 5 Then
        IntCostPriceType = LastSalesPriceType
    End If

    Screen.MousePointer = vbArrowHourglass

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True
    Else
        BolRtl = False
    End If

    With Me.Fg
        .Redraw = flexRDNone
        .Rows = 1
        .ColPosition(.ColIndex("ItemName")) = 0

        If BolRtl = True Then
            IntColName = 1
            .AddItem "‘Ã—… «·√’‰«›"
        Else
            .AddItem "Items Tree"
            IntColName = 1
        End If

        .Rowdata(.Rows - 1) = "1G"
        .IsSubtotal(.Rows - 1) = True
        .Cell(flexcpFontBold, .Rows - 1, 1) = True
        .GridLines = flexGridFlat
        .MergeCells = flexMergeSpill
        .OutlineBar = flexOutlineBarComplete
        .AllowUserResizing = flexResizeColumns
        .ExtendLastCol = True
        '.NodeClosedPicture = MDIFrmMain.ImgLstMenuIcons.ListImages("NodeClose").Picture
        '.NodeOpenPicture = MDIFrmMain.ImgLstMenuIcons.ListImages("NodeOpen").Picture
        .RowHeightMin = 300
        .ScrollTrack = False
        .ScrollTips = True
        .SheetBorder = vbWhite
        '-----------------------------------------
        '.ColHidden(.ColIndex("GroupName")) = True
        '-----------------------------------------
        My_SQL = " SELECT Groups.GroupID, Groups.GroupName, Groups.ParentID " & "FROM Groups Where Groups.ParentID=1"
        Set RsData = New ADODB.Recordset
    
        RsData.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        Call LoadGridTree("1G", RsData, Fg, "Groups", "ParentID", "", , IntColName, vbBlue)
        .Redraw = True
        '--------------------------------------------------------------------------
        Me.MousePointer = vbArrowHourglass
        StrSaledItemsID = GetSaledItemsIDs

        If SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = "select ItemID,ItemCode,ItemName,GroupID,Sum(Qty) as Qty1 From " & "dbo.QryGARDShort() QryGARDShort"
            StrSQL = StrSQL + " Where Qty > 0 "

            If StrSaledItemsID <> "" Then
                StrSQL = StrSQL + " AND ItemID  NOT IN "
                StrSQL = StrSQL + "("
                StrSQL = StrSQL + StrSaledItemsID
                StrSQL = StrSQL + ")"
            End If

            StrSQL = StrSQL + " Group BY ItemID,ItemCode,ItemName,GroupID"
            StrSQL = StrSQL + " Order By ItemID"
        Else
            Exit Sub
        End If

        CmdDo.Enabled = False
        CmdPrint.Enabled = False

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
        Me.PrgBar.value = 0
        Me.lbl(3).Caption = "0"
        Me.lbl(2).Caption = "0"

        If Not (rs.BOF Or rs.EOF) Then
            rs.MoveFirst
            Me.PrgBar.Visible = True
            Me.PrgBar.Max = rs.RecordCount

            'Me.lbl(3).Caption = Rs.RecordCount
            For i = 1 To rs.RecordCount

                DoEvents
                Me.PrgBar.value = i
                Me.lbl(3).Caption = i
            
                LngParentRow = .FindRow(CStr(rs("GroupID").value) & "G", 0, -1, False, True)
            
                If LngParentRow <> -1 Then
            
                    .AddItem "", (LngParentRow + 1)
                    .Rowdata((LngParentRow + 1)) = rs("ItemID").value & "I"
                    .RowOutlineLevel((LngParentRow + 1)) = .RowOutlineLevel(LngParentRow) + 1
                    .Cell(flexcpPicture, LngParentRow + 1, 0) = mdifrmmain.ImgLstTree.ListImages("Item").ExtractIcon
                
                    LngRowNum = LngParentRow + 1
               
                    .TextMatrix(LngRowNum, .ColIndex("Serial")) = i
                    .TextMatrix(LngRowNum, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
                    .TextMatrix(LngRowNum, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                    .TextMatrix(LngRowNum, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                    .TextMatrix(LngRowNum, .ColIndex("Qty")) = IIf(IsNull(rs("Qty1").value), 0, rs("Qty1").value)
                    'Calculate The Cost
                    DblItemCost = GetCostItemPrice(val(.TextMatrix(LngRowNum, .ColIndex("ItemID"))), 2, , , IntCostPriceType)
                    .TextMatrix(LngRowNum, .ColIndex("ItemCostPrice")) = DblItemCost
                    .TextMatrix(LngRowNum, .ColIndex("StockCost")) = (DblItemCost * val(.TextMatrix(LngRowNum, .ColIndex("Qty"))))
                    LastItemInv = GetLastItemTrans(rs("ItemID").value)
                    .TextMatrix(LngRowNum, .ColIndex("InvTransID")) = LastItemInv.Transactionid
                    .TextMatrix(LngRowNum, .ColIndex("InvTransDate")) = LastItemInv.TransactionDate
                    .TextMatrix(LngRowNum, .ColIndex("InvTransSerial")) = LastItemInv.TransactionSerial
                
                    .Cell(flexcpPictureAlignment, LngRowNum, 0) = flexPicAlignRightCenter
                Else
                    MsgBox "Stop"
                End If

                rs.MoveNext
            Next i

            For i = Me.Fg.FixedRows To Me.Fg.Rows - 1
                Dim XNode As VSFlex8UCtl.VSFlexNode
                Dim StrTemp As String

                If .IsSubtotal(i) = True Then
                    Set XNode = Fg.GetNode(i)

                    If Not XNode Is Nothing Then
                        '⁄œœ «·√’‰«› «·„ÊÃÊœ… œ«Œ· Â–Â «·„Ã„Ê⁄…
                        DblNodeChildCount = ModFgLib.GetNodeChildTotal(Fg, XNode, flexSTCount)
                        StrTemp = XNode.text & " ( " & DblNodeChildCount & " ) "
                        XNode.text = StrTemp
                        '------------------------------------------------------
                        '≈Ã„«·Ï  ﬂ·›… «·√’‰«› «·„ÊÃÊœ… œ«Œ· Â–Â «·„Ã„Ê⁄…
                        DblNodeChildCount = ModFgLib.GetNodeChildTotal(Fg, XNode, flexSTSum, Fg.ColIndex("StockCost"))
                        StrTemp = " ( " & DblNodeChildCount & " ) "
                    
                        StrTemp = XNode.text & " ( " & DblNodeChildCount & " ) "
                        XNode.text = StrTemp
                        '                    .TextMatrix(I, .ColIndex("StockCost")) = StrTemp
                        '                    .Cell(flexcpForeColor, I, .ColIndex("StockCost")) = vbRed
                        '                    .Cell(flexcpFontBold, I, .ColIndex("StockCost")) = True
                        '                    .Cell(flexcpFontSize, I, .ColIndex("StockCost")) = 10
                    End If
                End If

            Next i

        End If
    
        Me.PrgBar.value = 0
        Me.PrgBar.Visible = False
        .AutoSize 0, .Cols - 1, False
        Me.lbl(2).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("StockCost"), .Rows - 1, .ColIndex("StockCost"))

        CmdDo.Enabled = True
        CmdPrint.Enabled = True
        Me.MousePointer = vbDefault
        '--------------------------------------------------------------------------
    
    End With

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Private Sub Option1_Click()
 If Option1.value = True Then
    Fg.TextMatrix(0, Fg.ColIndex("Qty")) = "«·ﬂ„Ì… »«·„Œ«“‰"
    FrmStagnantItems.Caption = "«·«’‰«› «·—«ﬂœ…"
 Else
     Fg.TextMatrix(0, Fg.ColIndex("Qty")) = "«·ﬂ„Ì… «·„»«⁄…"
     FrmStagnantItems.Caption = "«·«’‰«› «·«ﬂÀ— „»Ì⁄«"
 End If
End Sub

Private Sub Option2_Click()
Option1_Click
End Sub
