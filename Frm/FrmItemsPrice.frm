VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmItemsPrice 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "√”⁄«— «·√’‰«›"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5565
   Icon            =   "FrmItemsPrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5610
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtCompareValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   900
      Visible         =   0   'False
      Width           =   735
   End
   Begin ImpulseButton.ISButton XPBtnOK 
      Height          =   345
      Left            =   930
      TabIndex        =   3
      Top             =   5160
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "„Ê«›ﬁ"
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
      ButtonImage     =   "FrmItemsPrice.frx":038A
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
   Begin ImpulseButton.ISButton XPBtnCancel 
      Height          =   345
      Left            =   120
      TabIndex        =   4
      Top             =   5160
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "≈·€«¡"
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
      ButtonImage     =   "FrmItemsPrice.frx":0724
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
   Begin C1SizerLibCtl.C1Tab XPTabMain 
      Height          =   4185
      Left            =   90
      TabIndex        =   5
      Top             =   900
      Width           =   5415
      _cx             =   9551
      _cy             =   7382
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
      Appearance      =   2
      MousePointer    =   0
      Version         =   801
      BackColor       =   14871017
      ForeColor       =   0
      FrontTabColor   =   14871017
      BackTabColor    =   12648447
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   16711680
      Caption         =   "√”⁄«— «·‘—«¡|  √”⁄«— «·»Ì⁄"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   1
      Position        =   1
      AutoSwitch      =   -1  'True
      AutoScroll      =   0   'False
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   37
      Picture(0)      =   "FrmItemsPrice.frx":0ABE
      Picture(1)      =   "FrmItemsPrice.frx":0E58
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   3720
         Index           =   1
         Left            =   6060
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   45
         Width           =   5325
         _cx             =   9393
         _cy             =   6562
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
         AutoSizeChildren=   0
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
         Begin VB.TextBox TxtDealerPrice 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1710
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   390
            Width           =   825
         End
         Begin VB.TextBox TxtCustomerPrice 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3510
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   390
            Width           =   825
         End
         Begin VB.TextBox XPTxtPrice 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2400
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   60
            Width           =   1725
         End
         Begin VSFlex8UCtl.VSFlexGrid FG 
            Height          =   2235
            Left            =   60
            TabIndex        =   8
            Top             =   1020
            Width           =   5205
            _cx             =   9181
            _cy             =   3942
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
            Rows            =   2
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmItemsPrice.frx":11F2
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
            Editable        =   2
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
         Begin ImpulseButton.ISButton XPBtnRemove 
            Height          =   375
            Left            =   4410
            TabIndex        =   9
            Top             =   3270
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
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
            ButtonImage     =   "FrmItemsPrice.frx":127D
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            LowerToggledContent=   0   'False
         End
         Begin ImpulseButton.ISButton XPBtnAdd 
            Height          =   375
            Left            =   4845
            TabIndex        =   10
            Top             =   3270
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
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
            ButtonImage     =   "FrmItemsPrice.frx":1617
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            LowerToggledContent=   0   'False
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”⁄— «·œÌ·—"
            Height          =   255
            Index           =   2
            Left            =   2610
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   420
            Width           =   855
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”⁄— «·⁄„Ì·"
            Height          =   255
            Index           =   0
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   420
            Width           =   855
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‘—«∆Õ «·√”⁄«—"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   6
            Left            =   4230
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   780
            Width           =   975
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”⁄— «·„” Â·ﬂ"
            Height          =   255
            Index           =   5
            Left            =   4110
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   90
            Width           =   1095
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   3720
         Index           =   0
         Left            =   45
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   45
         Width           =   5325
         _cx             =   9393
         _cy             =   6562
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
         AutoSizeChildren=   0
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
         Begin VSFlex8UCtl.VSFlexGrid FGPurchasePrice 
            Height          =   3315
            Left            =   60
            TabIndex        =   0
            Top             =   300
            Width           =   5115
            _cx             =   9022
            _cy             =   5847
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
            Rows            =   2
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmItemsPrice.frx":19B1
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
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”⁄— «·‘—«¡"
            ForeColor       =   &H00FF0000&
            Height          =   225
            Index           =   4
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   30
            Width           =   1095
         End
      End
   End
   Begin VB.TextBox TxtQty 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   690
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label XPLblItemID 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   330
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label XPLblItemCode 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Left            =   2460
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   330
      Width           =   2025
   End
   Begin VB.Label XPLblItemName 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Left            =   450
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   30
      Width           =   4035
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·’‰›"
      Height          =   255
      Index           =   3
      Left            =   4530
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   330
      Width           =   975
   End
   Begin VB.Label XPLblHaveSerial 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Â–« «·’‰› ·Â ”Ì—Ì«·"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   2700
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   630
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·’‰›"
      Height          =   255
      Index           =   1
      Left            =   4530
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   30
      Width           =   975
   End
End
Attribute VB_Name = "FrmItemsPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim RsPrice As ADODB.Recordset
Dim TTP As clstooltip
Dim OldGrdValue As Variant

Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)
    On Error GoTo ErrTrap

    With FG

        If .TextMatrix(Row, Col) <> "" Then
            If Not IsNumeric(.TextMatrix(Row, Col)) Or Len((.TextMatrix(Row, Col))) > 50 Then
                Msg = "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ…" & Chr(13)
                Msg = Msg + "√œŒ· ﬁÌ„ —ﬁ„Ì… „ÊÃ»… "
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                .TextMatrix(Row, Col) = OldGrdValue
                Exit Sub
            ElseIf .TextMatrix(Row, Col) < 0 Then
                Msg = "√œŒ· ﬁÌ„ —ﬁ„Ì… „ÊÃ»… "
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                .TextMatrix(Row, Col) = OldGrdValue
                Exit Sub
            End If
        End If

        If Col = .ColIndex("To") Then
            If val(.TextMatrix(Row, Col)) < val(.TextMatrix(Row, .ColIndex("Form"))) Then
                Msg = "·«»œ √‰ ÌﬂÊ‰ «·Õœ «·√ﬁ’ ··‘—ÌÕ…(≈·Ï)" & Chr(13)
                Msg = Msg + "√ﬂ»— „‰ «·Õœ «·√œ‰Ï ·Â« („‰) "
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If
        End If

        If Col = .ColIndex("Form") Then
            If Row > 1 Then
                If val(.TextMatrix(Row, Col)) - val(.TextMatrix(Row - 1, .ColIndex("To"))) > 1 Or val(.TextMatrix(Row, Col)) - val(.TextMatrix(Row - 1, .ColIndex("To"))) < 1 Then
                    Msg = "ÌÃ» √‰  »œ√ «·‘—ÌÕ… «·ÃœÌœ… " & Chr(13)
                    Msg = Msg + "„‰ ÕÌÀ «‰ Â  «·‘—ÌÕ… «·”«»ﬁ… +1"
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If
        End If

        If Col = .ColIndex("Price") Then
            If Row >= 1 Then
                If val(.TextMatrix(Row, .ColIndex("Price"))) > val(XPTxtPrice.text) Then
                    Msg = "√”⁄«— «·‘—«∆Õ ·«»œ √‰  ﬂÊ‰ " & Chr(13)
                    Msg = Msg + "√ﬁ· „‰ ”⁄— «·»Ì⁄"
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If

            If Row > 1 Then
                If val(.TextMatrix(Row, .ColIndex("Price"))) > val(.TextMatrix(Row - 1, .ColIndex("Price"))) Then
                    Msg = "ÌÃ» √‰ ÌﬂÊ‰ ”⁄— «·‘—ÌÕ… √ﬁ·" & Chr(13)
                    Msg = Msg + "„‰ ”⁄— «·‘—ÌÕ… «·”«»ﬁ…" & Chr(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If
        End If

    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_KeyDown(KeyCode As Integer, _
                       Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If FG.TextMatrix(FG.Rows - 1, FG.ColIndex("Form")) <> "" And FG.TextMatrix(FG.Rows - 1, FG.ColIndex("To")) <> "" And FG.TextMatrix(FG.Rows - 1, FG.ColIndex("Price")) <> "" Then
            FG.Rows = FG.Rows + 1
            FG.TextMatrix(FG.Rows - 1, FG.ColIndex("NumIndex")) = FG.Rows - 1
            FG.Row = FG.Rows - 1
            FG.Col = FG.ColIndex("Form")
            FG.SetFocus
        Else
            FG.Row = FG.Rows - 1
            FG.Col = FG.ColIndex("Form")
        End If
    End If

    'If KeyCode = vbKeyReturn Then
    '    If FG.TextMatrix(FG.Rows - 1, FG.ColIndex("Form")) <> "" And _
    '    FG.TextMatrix(FG.Rows - 1, FG.ColIndex("To")) <> "" And _
    '    FG.TextMatrix(FG.Rows - 1, FG.ColIndex("Price")) <> "" Then
    '        FG.Rows = FG.Rows + 1
    '        FG.TextMatrix(FG.Rows - 1, FG.ColIndex("NumIndex")) = FG.Rows - 1
    '        FG.Row = FG.Rows - 1
    '        FG.Col = FG.ColIndex("Form")
    '    Else
    '        FG.Row = FG.Rows - 1
    '        FG.Col = FG.ColIndex("Form")
    '    End If
    'End If
    'Exit Sub
ErrTrap:

End Sub

Private Sub Fg_StartEdit(ByVal Row As Long, _
                         ByVal Col As Long, _
                         Cancel As Boolean)
    On Error GoTo ErrTrap
    OldGrdValue = FG.TextMatrix(Row, Col)
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsTemp As ADODB.Recordset
    Set rs = New ADODB.Recordset

    If Me.XPLblItemID = "" Then
        Exit Sub
    End If

    StrSQL = "select * From ItemsPrice where Item_ID=" & Me.XPLblItemID
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    FG.TextMatrix(1, FG.ColIndex("NumIndex")) = "1"

    If Not (rs.EOF Or rs.BOF) Then
        Retrive
    End If

    Set RsTemp = New ADODB.Recordset
    StrSQL = "select * From TblItems  where ItemID=" & Me.XPLblItemID
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        If RsTemp("HaveSerial").value = True Then
            StrSQL = "SELECT Transactions.Transaction_ID, Transactions.Transaction_Date,Transaction_Details.Item_ID, " & "Transactions.Transaction_Type, Transactions.CusID, Transactions.StoreID, " & " Transaction_Details.ItemDiscountType, Transaction_Details.ItemDiscount," & "Transaction_Details.ItemSerial, Transaction_Details.Price, Transaction_Details.Quantity, " & "Transactions.Transaction_Type, TblCustemers.CusName FROM TblCustemers RIGHT JOIN " & "(Transactions INNER JOIN Transaction_Details ON Transactions.Transaction_ID = " & "Transaction_Details.Transaction_ID) ON TblCustemers.CusID = Transactions.CusID " & "WHERE(((Transactions.Transaction_Type) = 1 Or (Transactions.Transaction_Type) = 3)) " & "and(((Transaction_Details.ItemSerial) In (select ItemSerial From QryGardComplete)))"
            StrSQL = StrSQL + " and  Item_ID=" & Me.XPLblItemID
            Set RsPrice = New ADODB.Recordset
            RsPrice.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsPrice.EOF Or RsPrice.BOF) Then
                RetrivePrice True
            End If

        Else
            StrSQL = "SELECT Transactions.Transaction_ID, Transactions.Transaction_Date, " & "Transactions.Transaction_Type, Transactions.CusID, TblCustemers.CusName, " & "Transaction_Details.ItemDiscountType, Transaction_Details.ItemDiscount," & "Transaction_Details.Item_ID, Transaction_Details.Quantity, Transaction_Details.Price FROM " & "(TblCustemers RIGHT JOIN Transactions ON TblCustemers.CusID = Transactions.CusID) " & "INNER JOIN Transaction_Details ON Transactions.Transaction_ID = " & "Transaction_Details.Transaction_ID Where (((Transactions.Transaction_Type) = 1 " & "Or (Transactions.Transaction_Type) = 3))"
            StrSQL = StrSQL + " and  Item_ID=" & Me.XPLblItemID
            StrSQL = StrSQL + " ORDER BY Transactions.Transaction_Date  DESC"
            Set RsPrice = New ADODB.Recordset
            RsPrice.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsPrice.EOF Or RsPrice.BOF) Then
                RetrivePrice False
            End If

        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyF2 Then
        XPBtnAdd_Click
    End If

    If KeyCode = vbKeyF3 Then
        XPBtnRemove_Click
    End If

    If Shift = 2 Then
        XPTabMain.SetFocus

        If KeyCode = vbKeyTab Then
            If XPTabMain.CurrTab = 0 Then
                XPTabMain.CurrTab = 1
                FG.SetFocus
            Else
                XPTabMain.CurrTab = 0
                FGPurchasePrice.SetFocus
            End If
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            XPBtnCancel_Click
        End If
    End If

ErrTrap:
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim BackGround As ClsBackGroundPic
    Set BackGround = New ClsBackGroundPic
    CenterForm Me

    FormPostion Me, GetPostion
    FG.WallPaper = BackGround.MoneyWallpaper
    FGPurchasePrice.WallPaper = BackGround.MoneyWallpaper
    XPTabMain.CurrTab = 1
    AddTip

    With FG
        .Cell(flexcpPicture, 0, .ColIndex("Form")) = mdifrmmain.ImgLstTree.ListImages("From").Picture
        .Cell(flexcpPicture, 0, .ColIndex("To")) = mdifrmmain.ImgLstTree.ListImages("To").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Price")) = mdifrmmain.ImgLstTree.ListImages("Price").Picture
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter
    End With

    With FGPurchasePrice
        .Cell(flexcpPicture, 0, .ColIndex("Serial")) = mdifrmmain.ImgLstTree.ListImages("Serial").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Price")) = mdifrmmain.ImgLstTree.ListImages("Price").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Qty")) = mdifrmmain.ImgLstTree.ListImages("qty").Picture
        .Cell(flexcpPicture, 0, .ColIndex("BillNum")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Date")) = mdifrmmain.ImgLstTree.ListImages("Date").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Supplier")) = mdifrmmain.ImgLstTree.ListImages("User").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Cost")) = mdifrmmain.ImgLstTree.ListImages("Price").Picture
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub XPBtnAdd_Click()
    On Error GoTo ErrTrap

    If FG.TextMatrix(FG.Rows - 1, FG.ColIndex("Form")) <> "" Then
        FG.Rows = FG.Rows + 1
        FG.TextMatrix(FG.Rows - 1, FG.ColIndex("NumIndex")) = FG.Rows - 1
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPBtnCancel_Click()
    Unload Me
End Sub

Private Sub Retrive()
    On Error GoTo ErrTrap
    Dim RowNum As Integer
    FG.Rows = rs.RecordCount + 1

    For RowNum = 1 To rs.RecordCount

        With FG
            .TextMatrix(RowNum, .ColIndex("NumIndex")) = RowNum
            .TextMatrix(RowNum, .ColIndex("Form")) = IIf(IsNull(rs("From").value), "", Trim(rs("From").value))
            .TextMatrix(RowNum, .ColIndex("To")) = IIf(IsNull(rs("To").value), "", Trim(rs("To").value))
            .TextMatrix(RowNum, .ColIndex("Price")) = IIf(IsNull(rs("Price").value), "", Trim(rs("Price").value))
        End With

        rs.MoveNext
    Next RowNum

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    FormPostion Me, SavePostion

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    Set rs = Nothing
    Set TTP = Nothing
    Set ItemReport = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Sub XPBtnOK_Click()
    On Error GoTo ErrTrap
    Dim RsItems As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim RsPrice As ADODB.Recordset
    Dim StrSQL As String
    Dim RowNum As Integer
    Dim ColNum As Integer
    Dim BeginTrans As Boolean
    Dim Msg As String

    If Not IsNumeric(XPTxtPrice.text) Then
        Msg = "”⁄— «·»Ì⁄ «·–Ì √œŒ· Â €Ì— ’«·Õ" & Chr(13)
        Msg = Msg + "√œŒ· ﬁÌ„… —ﬁ„Ì…"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPTxtPrice.SetFocus
        SelectText XPTxtPrice
        Exit Sub
    End If

    With FG

        For RowNum = 1 To .Rows - 1
            For ColNum = 1 To .Cols - 1

                If .TextMatrix(RowNum, ColNum) <> "" Then
                    If Not IsNumeric(.TextMatrix(RowNum, ColNum)) Or Len((.TextMatrix(RowNum, ColNum))) > 50 Then
                        Msg = "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ…" & Chr(13)
                        Msg = Msg + " √ﬂœ √‰ Ã„Ì⁄ «·ﬁÌ„ —ﬁ„Ì… "
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        XPTabMain.CurrTab = 1
                        FG.Row = RowNum
                        FG.Col = ColNum
                        FG.ShowCell RowNum, FG.ColIndex("Form")
                        FG.SetFocus
                        Exit Sub
                    End If
                End If

            Next ColNum

            If val(.TextMatrix(RowNum, .ColIndex("To"))) < val(.TextMatrix(RowNum, .ColIndex("Form"))) Then
                Msg = "·«»œ √‰ ÌﬂÊ‰ «·Õœ «·√ﬁ’ ··‘—ÌÕ…(≈·Ï)" & Chr(13)
                Msg = Msg + "√ﬂ»— „‰ «·Õœ «·√œ‰Ï ·Â« („‰) " & Chr(13)
                Msg = Msg + "—«Ã⁄ »Ì«‰«  «·‘—ÌÕ… —ﬁ„ " & RowNum
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                XPTabMain.CurrTab = 1
                FG.Row = RowNum
                FG.Col = FG.ColIndex("Form")
                FG.ShowCell RowNum, FG.ColIndex("Form")
                FG.SetFocus
                Exit Sub
            End If

            If RowNum > 1 Then
                If .TextMatrix(RowNum, .ColIndex("Form")) <> "" Then
                    If val(.TextMatrix(RowNum, .ColIndex("Form"))) - val(.TextMatrix(RowNum - 1, .ColIndex("To"))) > 1 Or val(.TextMatrix(RowNum, .ColIndex("Form"))) - val(.TextMatrix(RowNum - 1, .ColIndex("To"))) < 1 Then
                        Msg = "ÌÃ» √‰  »œ√ ﬂ· ‘—ÌÕ… " & Chr(13)
                        Msg = Msg + "„‰ ÕÌÀ «‰ Â  «·‘—ÌÕ… «·”«»ﬁ… +1" & Chr(13)
                        Msg = Msg + "—«Ã⁄ »Ì«‰«  «·‘—ÌÕ… —ﬁ„ " & RowNum
                        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                        XPTabMain.CurrTab = 1
                        FG.Row = RowNum
                        FG.Col = FG.ColIndex("To")
                        FG.ShowCell RowNum, FG.ColIndex("Form")
                        FG.SetFocus
                        Exit Sub
                    End If
                End If

                If val(.TextMatrix(RowNum, .ColIndex("Price"))) > val(.TextMatrix(RowNum - 1, .ColIndex("Price"))) Then
                    Msg = "ÌÃ» √‰ ÌﬂÊ‰ ”⁄— «·‘—ÌÕ… √ﬁ·" & Chr(13)
                    Msg = Msg + "„‰ ”⁄— «·‘—ÌÕ… «·”«»ﬁ…" & Chr(13)
                    Msg = Msg + "—«Ã⁄ »Ì«‰«  «·‘—ÌÕ… —ﬁ„ " & RowNum
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    XPTabMain.CurrTab = 1
                    FG.Row = RowNum
                    FG.Col = FG.ColIndex("Price")
                    FG.ShowCell RowNum, FG.ColIndex("Form")
                    FG.SetFocus
                    Exit Sub
                End If
            End If

            If RowNum >= 1 Then
                If val(.TextMatrix(RowNum, .ColIndex("Price"))) > val(XPTxtPrice.text) Then
                    Msg = "√”⁄«— «·‘—«∆Õ ·«»œ √‰  ﬂÊ‰ " & Chr(13)
                    Msg = Msg + "√ﬁ· „‰ ”⁄— «·»Ì⁄"
                    Msg = Msg + "—«Ã⁄ »Ì«‰«  «·‘—ÌÕ… —ﬁ„ " & RowNum
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    XPTabMain.CurrTab = 1
                    FG.Row = RowNum
                    FG.Col = FG.ColIndex("Price")
                    FG.ShowCell RowNum, FG.ColIndex("Form")
                    FG.SetFocus
                    Exit Sub
                End If
            End If

        Next RowNum

    End With

    StrSQL = "delete  From ItemsPrice where Item_ID=" & Me.XPLblItemID
    Cn.Execute StrSQL
    StrSQL = "select * From TblItems where ItemID=" & Me.XPLblItemID
    RsItems.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Cn.BeginTrans
    BeginTrans = True
    'If Val(TxtCompareValue.Text) <> Val(XPTxtPrice.Text) Then
    RsItems("SallingPrice").value = val(Me.XPTxtPrice.text)
    RsItems("CustomerPrice").value = val(Me.TxtCustomerPrice.text)
    RsItems("DealerPrice").value = val(Me.TxtDealerPrice.text)
    RsItems("UserID").value = user_id
    RsItems("LastUpdate").value = Date

    With FrmMainPriceList.FgMain
        .TextMatrix(.Row, .ColIndex("DefalutPrice")) = Me.XPTxtPrice.text
        .TextMatrix(.Row, .ColIndex("CustomerPrice")) = val(Me.TxtCustomerPrice.text)
        .TextMatrix(.Row, .ColIndex("DealerPrice")) = val(Me.TxtDealerPrice.text)
        .TextMatrix(.Row, .ColIndex("LastUpdate")) = Format(Date, "yyyy/m/d")
    End With

    'End If
    RsItems.update

    For RowNum = 1 To FG.Rows - 1

        With FG

            If .TextMatrix(RowNum, .ColIndex("Price")) <> "" Then
                rs.AddNew
                rs("PriceID").value = new_id("ItemsPrice", "PriceID", "", True)
                rs("Item_ID").value = Me.XPLblItemID
                rs("From").value = IIf(.TextMatrix(RowNum, .ColIndex("Form")) = "", "0", val(.TextMatrix(RowNum, .ColIndex("Form"))))
                rs("to").value = IIf(IsNull(.TextMatrix(RowNum, .ColIndex("To"))), "0", val(.TextMatrix(RowNum, .ColIndex("To"))))
                rs("Price").value = IIf(IsNull(.TextMatrix(RowNum, .ColIndex("Price"))), "0", val(.TextMatrix(RowNum, .ColIndex("Price"))))
                rs.update
            End If

        End With

    Next RowNum

    Cn.CommitTrans
    BeginTrans = False
    Set RsPrice = New ADODB.Recordset

    If Me.XPLblItemID <> "" Then
        StrSQL = "select * From ItemsPrice where Item_ID=" & Me.XPLblItemID
        RsPrice.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsPrice.EOF Or RsPrice.BOF) Then

            With FrmMainPriceList.FgMain
                .Cell(flexcpPicture, .Row, .ColIndex("DefalutPrice")) = mdifrmmain.ImgLstTree.ListImages("Tick").Picture
            End With

        Else

            With FrmMainPriceList.FgMain
                .Cell(flexcpPicture, .Row, .ColIndex("DefalutPrice")) = ""
            End With

        End If

        RsPrice.Close
    End If

    Unload Me
    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

End Sub

Private Sub XPBtnRemove_Click()
    On Error GoTo ErrTrap

    If FG.Rows > 1 Then
        If FG.Rows = 2 Then
            FG.Clear flexClearScrollable, flexClearEverything
        Else

            If FG.Rows > 1 Then
                If FG.Row <> FG.FixedRows - 1 Then
                    FG.RemoveItem (FG.Rows - 1)
                End If
            End If
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub RetrivePrice(HaveSerial As Boolean)
    On Error GoTo ErrTrap
    Dim RowNum As Integer

    With FGPurchasePrice

        If HaveSerial = True Then
            .ColHidden(.ColIndex("Qty")) = True
            .ColHidden(.ColIndex("Serial")) = False
            XPLblHaveSerial.Visible = True
        Else
            .ColHidden(.ColIndex("Serial")) = True
            .ColHidden(.ColIndex("Qty")) = False
            XPLblHaveSerial.Visible = False
        End If

        .Rows = RsPrice.RecordCount + 1
        .Cell(flexcpBackColor, 1, .ColIndex("Price")) = vbYellow

        For RowNum = 1 To RsPrice.RecordCount
    
            .TextMatrix(RowNum, .ColIndex("NumIndex")) = RowNum

            If HaveSerial = True Then
                .TextMatrix(RowNum, .ColIndex("Serial")) = IIf(IsNull(RsPrice("ItemSerial").value), "", Trim(RsPrice("ItemSerial").value))
            Else
                .TextMatrix(RowNum, .ColIndex("Qty")) = IIf(IsNull(RsPrice("Quantity").value), "", Trim(RsPrice("Quantity").value))
            End If

            .TextMatrix(RowNum, .ColIndex("Price")) = IIf(IsNull(RsPrice("Price").value), "", Trim(RsPrice("Price").value))
            .TextMatrix(RowNum, .ColIndex("BillNum")) = IIf(IsNull(RsPrice("Transaction_ID").value), "", Trim(RsPrice("Transaction_ID").value))
            .TextMatrix(RowNum, .ColIndex("Date")) = IIf(IsNull(RsPrice("Transaction_Date").value), "", Format((RsPrice("Transaction_Date").value), "yyyy/m/d"))

            If RsPrice("Transaction_Type").value = 3 Then
                .TextMatrix(RowNum, .ColIndex("Supplier")) = "—’Ìœ «›  «ÕÌ"
                .TextMatrix(RowNum, .ColIndex("Cost")) = val(RsPrice("Quantity").value) * val(RsPrice("Price").value)
            Else
                .TextMatrix(RowNum, .ColIndex("Supplier")) = IIf(IsNull(RsPrice("CusName").value), "", Trim(RsPrice("CusName").value))
            End If

            ' ﬂ·›… «·’‰› »⁄œ «·Œ’„
            If Not IsNull(RsPrice("ItemDiscountType").value) Then

                Select Case val(RsPrice("ItemDiscountType").value)

                    Case 0
                        .TextMatrix(RowNum, .ColIndex("Cost")) = val(RsPrice("Quantity").value) * val(RsPrice("Price").value)

                    Case 1
                        .TextMatrix(RowNum, .ColIndex("Cost")) = val(RsPrice("Quantity").value) * val(RsPrice("Price").value)

                    Case 2
                        .TextMatrix(RowNum, .ColIndex("Cost")) = (val(RsPrice("Quantity").value) * val(RsPrice("Price").value)) - val(RsPrice("ItemDiscount"))

                    Case 3
                        .TextMatrix(RowNum, .ColIndex("Cost")) = (val(RsPrice("Quantity").value) * val(RsPrice("Price").value)) * (1 - (val(RsPrice("ItemDiscount")) / 100))

                    Case 4
                        .TextMatrix(RowNum, .ColIndex("Cost")) = "„Ã«‰Ì"
                End Select

            End If

            If HaveSerial = False Then
                If val(.Aggregate(flexSTSum, 1, .ColIndex("Qty"), .Rows - 1, .ColIndex("Qty"))) >= val(txtqty.text) Then
                    Exit Sub
                End If
            End If

            RsPrice.MoveNext
        Next RowNum

    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = Chr(13) + Chr(10)

    With TTP
        .Create Me.hWnd, "√”⁄«— «·√’‰«›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnAdd, "≈÷«›… ‘—ÌÕ… ..." & Wrap & "·«÷«›… ‘—ÌÕ… √”⁄«— ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "√”⁄«— «·√’‰«›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnRemove, "Õ–› ‘—ÌÕ… ..." & Wrap & "·Õ–› ‘—ÌÕ… «·√”⁄«— «·√ŒÌ—…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "√”⁄«— «·√’‰«›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPTxtPrice, "”⁄— «·»Ì⁄ «·«› —«÷Ì ··’‰›", True
    End With

    With TTP
        .Create Me.hWnd, "√”⁄«— «·√’‰«›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl FGPurchasePrice, "√”⁄«— «·‘—«¡ ..." & Wrap & " «·„»·€ «·–Ì  „ œ›⁄Â  " & Wrap & "⁄‰œ ‘—«¡ Â–« «·’‰›", True
    End With

    With TTP
        .Create Me.hWnd, "√”⁄«— «·√’‰«›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnOK, "„Ê«›ﬁ ..." & Wrap & "·Õ›Ÿ ‘—«∆Õ «·√”⁄«— «· Ì  „ ﬂ «» Â«" & Wrap & " √Ê «· ⁄œÌ·«  «· Ì  „  ⁄·ÌÂ«", True
    End With

    With TTP
        .Create Me.hWnd, "√”⁄«— «·√’‰«›", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnCancel, "≈·€«¡  ..." & Wrap & "·≈·€«¡ «· ⁄œÌ·«  «· Ì  „  ⁄·Ï ‘—«∆Õ «·√”⁄«—" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub XPTabMain_Switch(OldTab As Integer, _
                             NewTab As Integer, _
                             Cancel As Integer)
    'On Error GoTo ErrTrap
    'If User_ID <> 1 Then
    '    If NewTab = 0 Then Cancel = True
    'End If
    'Exit Sub
    'ErrTrap:

End Sub
