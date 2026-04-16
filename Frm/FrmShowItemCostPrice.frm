VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmShowItemCostPrice 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "⁄—÷ Õ”«»  ﬂ·›… «·„Œ“Ê‰ „‰ ’›"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6975
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   6975
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
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   6135
      Index           =   0
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6975
      _cx             =   12303
      _cy             =   10821
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
      BackColor       =   -2147483633
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
      GridRows        =   4
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmShowItemCostPrice.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1440
         Index           =   2
         Left            =   30
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   30
         Width           =   6915
         _cx             =   12197
         _cy             =   2540
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
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… Õ”«» «· ﬂ·›…:"
            Height          =   270
            Index           =   17
            Left            =   5130
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   120
            Width           =   1515
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   16
            Left            =   1020
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   60
            Width           =   4095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
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
            Height          =   270
            Index           =   15
            Left            =   1785
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   1065
            Width           =   3960
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
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
            Height          =   255
            Index           =   14
            Left            =   1785
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   780
            Width           =   3960
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
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
            Height          =   270
            Index           =   13
            Left            =   1785
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   495
            Width           =   3960
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·’‰›:"
            Height          =   270
            Index           =   2
            Left            =   5760
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   1065
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·’‰›:"
            Height          =   270
            Index           =   1
            Left            =   5760
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   780
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·’‰›:"
            Height          =   270
            Index           =   0
            Left            =   5760
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   495
            Width           =   975
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   2040
         Index           =   1
         Left            =   30
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   4065
         Width           =   6915
         _cx             =   12197
         _cy             =   3598
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   885
            Index           =   3
            Left            =   30
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   30
            Width           =   2955
            _cx             =   5212
            _cy             =   1561
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            Caption         =   "≈Õ’«∆Ì… ”⁄— «·‘—«¡ ··’‰›"
            Align           =   0
            AutoSizeChildren=   7
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
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Index           =   21
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   540
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Index           =   20
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "√⁄·Ï ”⁄— ‘—«¡ "
               Height          =   255
               Index           =   19
               Left            =   1380
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   570
               Width           =   1245
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "√ﬁ· ”⁄— ‘—«¡ "
               Height          =   255
               Index           =   18
               Left            =   1380
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   240
               Width           =   1245
            End
            Begin VB.Image Img 
               Height          =   240
               Index           =   1
               Left            =   2640
               Picture         =   "FrmShowItemCostPrice.frx":0083
               Top             =   240
               Width           =   240
            End
            Begin VB.Image Img 
               Height          =   240
               Index           =   0
               Left            =   2640
               Picture         =   "FrmShowItemCostPrice.frx":040D
               Top             =   540
               Width           =   240
            End
         End
         Begin ImpulseButton.ISButton CmdExit 
            Height          =   375
            Left            =   30
            TabIndex        =   7
            Top             =   1575
            Width           =   975
            _ExtentX        =   1720
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
            ButtonImage     =   "FrmShowItemCostPrice.frx":0797
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
            Caption         =   "lbl"
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
            Height          =   405
            Index           =   22
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   1590
            Visible         =   0   'False
            Width           =   5355
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   12
            Left            =   150
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   1065
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· ﬂ·›… «·√Ã„«·Ì… ··ﬂ„Ì…:-"
            Height          =   315
            Index           =   11
            Left            =   1230
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   1065
            Width           =   2010
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   10
            Left            =   3825
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   1095
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂ„Ì… «·„ÊÃÊœ…"
            Height          =   315
            Index           =   9
            Left            =   4740
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   1095
            Width           =   2130
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            ForeColor       =   &H000000FF&
            Height          =   270
            Index           =   8
            Left            =   3225
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   780
            Width           =   1455
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   7
            Left            =   3225
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   420
            Width           =   1455
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "0"
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   6
            Left            =   3225
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„ Ê”ÿ ”⁄— «·’‰› :-"
            Height          =   315
            Index           =   5
            Left            =   4740
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   810
            Width           =   2130
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ï ﬁÌ„… «·’‰›-"
            Height          =   315
            Index           =   4
            Left            =   4740
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   450
            Width           =   2130
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ï «·ﬂ„Ì…:-"
            Height          =   315
            Index           =   3
            Left            =   4740
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   120
            Width           =   2130
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   -30
            X2              =   6930
            Y1              =   1485
            Y2              =   1500
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   2565
         Left            =   30
         TabIndex        =   1
         Top             =   1485
         Width           =   6915
         _cx             =   12197
         _cy             =   4524
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
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmShowItemCostPrice.frx":0B31
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
End
Attribute VB_Name = "FrmShowItemCostPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Ele_ResizeChildren(Index As Integer)
    Line1.X1 = 0
    Line1.X2 = Me.Ele(1).Width
End Sub

Private Sub Fg_DblClick()
    Dim IntType As Integer
    Dim IntTransID As Long

    With Me.FG

        If .Row < 1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If val(.TextMatrix(.Row, .ColIndex("Transaction_ID"))) = 0 Then Exit Sub
        If val(.TextMatrix(.Row, .ColIndex("TransType"))) = 0 Then Exit Sub
        IntType = val(.TextMatrix(.Row, .ColIndex("TransType")))
        IntTransID = val(.TextMatrix(.Row, .ColIndex("Transaction_ID")))
    
        If IntType = 1 Then
            OpenScreen PurchaseScreen, IntTransID
        ElseIf IntType = 3 Then
            OpenScreen OpenStockBalance, IntTransID
        ElseIf IntType = 9 Then
            OpenScreen RetrunSalles, IntTransID
        End If

    End With

End Sub

Private Sub Form_Activate()
    'PutFormOnTop Me.hwnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Me.Width = 7000
    Me.Height = 6225
    CenterForm Me

    FormPostion Me, GetPostion
    Set GrdBack = New ClsBackGroundPic
    Me.FG.WallPaper = GrdBack.Picture
    FG.AllowUserResizing = flexResizeColumns
    Ele(3).Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Public Sub LoadData(LngItemID As Long, _
                    DblQty As Double, _
                    Optional IntCostType As StockCostType = StockCostType.WeightAverage)

    Dim DblQtyTotal As Double
    Dim DblCostTotal As Double
    Dim DblItemCostPrice As Currency
    Dim xTemp As LastItemTransInfo
    Dim IntTempTransType As Integer
    Dim IntTempTransTypeName As String

    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim StrTemp As String
    Dim DblTemp As Double
    Dim cProgress As ClsProgress
    Dim DblTempQty As Double
    Dim DblOneUnitPrice As Double
    Dim Msg As String

    'DblQty :-«·ﬂ„Ì… «·„ »ﬁÌ… „‰ «·’‰› ( —’Ìœ «·’‰› Õ«·Ì«)...

    If IntCostType = LastPurPriceType Then
        StrTemp = "√Œ— ”⁄— ‘—«¡"
    ElseIf IntCostType = WeightAverage Then
        StrTemp = "«·„ Ê”ÿ «·„—ÃÕ"
    ElseIf IntCostType = FirstInFirstOut Then
        StrTemp = "«·Ê«—œ √Ê·« Ì’—› √Ê·«"
    Else
        Exit Sub
    End If

    Me.lbl(16).Caption = StrTemp

    If IntCostType = WeightAverage Then
        '«·„ Ê”ÿ «·„—ÃÕ
        StrSQL = "select QryItemTransactions.* " & " From dbo.QryItemTransactions()QryItemTransactions " & " Where ItemID = " & LngItemID & " And (Transaction_Type = 1 Or Transaction_Type = 3)"
        StrSQL = StrSQL + " Order By Transaction_ID"
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncExecute
        Set cProgress = New ClsProgress
        cProgress.ProgressType = Waiting
        cProgress.StartProgress

        Do While rs.State = adStateExecuting
            DoEvents
        Loop

        With Me.FG
            .Rows = .FixedRows

            If Not (rs.BOF Or rs.EOF) Then
                Me.lbl(13).Caption = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
                Me.lbl(14).Caption = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                Me.lbl(15).Caption = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
            
                .Rows = .FixedRows + rs.RecordCount

                For i = .FixedRows To .Rows - 1
                    .TextMatrix(i, .ColIndex("Serial")) = i
                    .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)
                    .TextMatrix(i, .ColIndex("TransType")) = IIf(IsNull(rs("Transaction_Type").value), "", rs("Transaction_Type").value)
                    .TextMatrix(i, .ColIndex("TransactionType")) = IIf(IsNull(rs("TransactionTypeName").value), "", rs("TransactionTypeName").value)
                    .TextMatrix(i, .ColIndex("TransactionSerial")) = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)

                    If Not IsNull(rs("Transaction_Date").value) Then
                        .TextMatrix(i, .ColIndex("TransactionDate")) = DisplayDate(rs("Transaction_Date").value)
                    End If

                    .TextMatrix(i, .ColIndex("Qty")) = IIf(IsNull(rs("TotalQty").value), 0, rs("TotalQty").value)
                    .TextMatrix(i, .ColIndex("ItemTotal")) = IIf(IsNull(rs("Total").value), 0, rs("Total").value)

                    If val(.TextMatrix(i, .ColIndex("Qty"))) <> 0 Then
                        DblItemCostPrice = val(.TextMatrix(i, .ColIndex("ItemTotal"))) / val(.TextMatrix(i, .ColIndex("Qty")))
                    Else
                        DblItemCostPrice = 0
                    End If

                    .TextMatrix(i, .ColIndex("ItemPrice")) = DblItemCostPrice
                    rs.MoveNext
                Next i

            End If

            rs.Close
            Set rs = Nothing
            .AutoSize 0, .Cols - 1, False
            DblQtyTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Qty"), .Rows - 1, .ColIndex("Qty"))
            DblCostTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("ItemTotal"), .Rows - 1, .ColIndex("ItemTotal"))
            DblItemCostPrice = DblCostTotal / DblQtyTotal
        End With

        Me.lbl(6).Caption = DblQtyTotal
        Me.lbl(7).Caption = DblCostTotal
        Me.lbl(8).Caption = Format(DblItemCostPrice, SystemOptions.SysDefCurrencyForamt)
        Me.lbl(10).Caption = DblQty
        Me.lbl(12).Caption = Format(val(Me.lbl(8).Caption) * DblQty, SystemOptions.SysDefCurrencyForamt)
        '--------------Item Purchase Price
        Me.Ele(3).Visible = True
        DblTemp = FG.Aggregate(flexSTMin, FG.FixedRows, FG.ColIndex("ItemPrice"), FG.Rows - 1, FG.ColIndex("ItemPrice"))
        Me.lbl(20).Caption = DblTemp
        DblTemp = FG.Aggregate(flexSTMax, FG.FixedRows, FG.ColIndex("ItemPrice"), FG.Rows - 1, FG.ColIndex("ItemPrice"))
        Me.lbl(21).Caption = DblTemp
        '-------------
    ElseIf IntCostType = LastPurPriceType Then
        '«Œ— ”⁄— ‘—«¡
        Set rs = New ADODB.Recordset
        StrSQL = "Select * From TblItems Where ItemID=" & LngItemID & ""
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        Me.lbl(13).Caption = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
        Me.lbl(14).Caption = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
        Me.lbl(15).Caption = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
            
        xTemp = GetLastItemTrans(LngItemID, 1)
        IntTempTransType = 1
        IntTempTransTypeName = "›« Ê—… ‘—«¡"

        If xTemp.Transactionid = 0 Then
            xTemp = GetLastItemTrans(LngItemID, 3)
            IntTempTransType = 3
            IntTempTransTypeName = "—’Ìœ ≈›  «ÕÏ"
        End If

        If xTemp.Transactionid = 0 Then
            xTemp = GetLastItemTrans(LngItemID, 9)
            IntTempTransType = 1
            IntTempTransTypeName = "„— Ã⁄ „»Ì⁄« "
        End If

        If xTemp.Transactionid <> 0 Then

            With FG
                .Rows = .FixedRows + 1
                i = 1
                .TextMatrix(i, .ColIndex("Serial")) = 1
                .TextMatrix(i, .ColIndex("Transaction_ID")) = xTemp.Transactionid
                .TextMatrix(i, .ColIndex("TransType")) = IntTempTransType
                .TextMatrix(i, .ColIndex("TransactionType")) = IntTempTransTypeName
                .TextMatrix(i, .ColIndex("TransactionSerial")) = xTemp.TransactionSerial

                If Not IsNull(xTemp.TransactionDate) Then
                    .TextMatrix(i, .ColIndex("TransactionDate")) = DisplayDate(CDate(xTemp.TransactionDate))
                End If
            
                .TextMatrix(i, .ColIndex("Qty")) = xTemp.SngItemQty
                .TextMatrix(i, .ColIndex("ItemPrice")) = xTemp.SngItemPrice
                .TextMatrix(i, .ColIndex("ItemTotal")) = xTemp.SngItemQty * xTemp.SngItemPrice
            End With

        End If

        Me.lbl(3).Enabled = False
        Me.lbl(4).Enabled = False
        Me.lbl(6).Enabled = False
        Me.lbl(7).Enabled = False
    
        Me.lbl(5).Caption = "«Œ— ”⁄— ‘—«¡ ··’‰›:"
        Me.lbl(8).Caption = xTemp.SngItemPrice
        Me.lbl(10).Caption = DblQty
        Me.lbl(12).Caption = Format(val(Me.lbl(8).Caption) * DblQty, SystemOptions.SysDefCurrencyForamt)
    ElseIf IntCostType = FirstInFirstOut Then
        '«·Ê«—œ «Ê·« Ì’—› √Ê·«
        StrSQL = "Select * From dbo.QryItemFifoTransactions(" & LngItemID & " ,'1,3,16') Order By Transaction_ID DESC  "
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncExecute
        Set cProgress = New ClsProgress
        cProgress.ProgressType = Waiting
        cProgress.StartProgress

        Do While rs.State = adStateExecuting
            DoEvents
        Loop

        With Me.FG
            .Rows = .FixedRows

            If Not (rs.BOF Or rs.EOF) Then
                Me.lbl(13).Caption = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
                Me.lbl(14).Caption = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                Me.lbl(15).Caption = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
            
                .Rows = .FixedRows + rs.RecordCount

                For i = .FixedRows To .Rows - 1
                    .TextMatrix(i, .ColIndex("Serial")) = i
                    .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)
                    .TextMatrix(i, .ColIndex("TransType")) = IIf(IsNull(rs("Transaction_Type").value), "", rs("Transaction_Type").value)
                    .TextMatrix(i, .ColIndex("TransactionType")) = IIf(IsNull(rs("TransactionTypeName").value), "", rs("TransactionTypeName").value)
                    .TextMatrix(i, .ColIndex("TransactionSerial")) = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)

                    If Not IsNull(rs("Transaction_Date").value) Then
                        .TextMatrix(i, .ColIndex("TransactionDate")) = DisplayDate(rs("Transaction_Date").value)
                    End If

                    .TextMatrix(i, .ColIndex("Qty")) = IIf(IsNull(rs("TotalQty").value), 0, rs("TotalQty").value)
                    .TextMatrix(i, .ColIndex("ItemTotal")) = IIf(IsNull(rs("Total").value), 0, rs("Total").value)
                    DblItemCostPrice = val(.TextMatrix(i, .ColIndex("ItemTotal"))) / val(.TextMatrix(i, .ColIndex("Qty")))
                    .TextMatrix(i, .ColIndex("ItemPrice")) = DblItemCostPrice
                    rs.MoveNext
                Next i

            End If

            rs.Close
            Set rs = Nothing
            .AutoSize 0, .Cols - 1, False
            DblQtyTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Qty"), .Rows - 1, .ColIndex("Qty"))
            DblCostTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("ItemTotal"), .Rows - 1, .ColIndex("ItemTotal"))
            DblItemCostPrice = DblCostTotal / DblQtyTotal
        End With
    
        Me.lbl(3).Caption = "≈Ã„«·Ï «·ﬂ„Ì… «·Ê«—œ…:-"
        Me.lbl(4).Caption = "≈Ã„«·Ï   ﬂ·›… «·ﬂ„Ì… «·Ê«—œ…:-"
        Me.lbl(6).Caption = DblQtyTotal
        Me.lbl(7).Caption = DblCostTotal
        Me.lbl(8).Caption = Format(DblItemCostPrice, SystemOptions.SysDefCurrencyForamt)
        Me.lbl(10).Caption = DblQty

        With Me.FG
            i = 1
            DblTempQty = val(.TextMatrix(i, .ColIndex("Qty")))

            If DblTempQty >= DblQty Then
                '«·ﬂ„Ì… «·„ÊÃÊœ… ›Ï Â–Â «·›« Ê—… «ﬂ»— „‰ «·—’Ìœ «·‰Â«∆Ï
                '«Ê «‰ «·ﬂ„Ì… ›Ï «·›« Ê—…  ”«ÊÏ ‰›” «·ﬂ„Ì… «·„ »ﬁÌ… ﬂ—’Ìœ ‰Â«∆Ï
                'If RsItem("Total").Value = 0 Then Stop
                'DblOneUnitPrice = DblTempQty / RsItem("Total").Value
                DblOneUnitPrice = val(.TextMatrix(i, .ColIndex("ItemTotal"))) / DblTempQty
                DblTemp = DblOneUnitPrice * DblQty
                .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = vbGreen
            ElseIf DblTempQty < DblQty Then
                DblTempQty = 0

                Do While DblTempQty < DblQty
                    DblTempQty = DblTempQty + val(.TextMatrix(i, .ColIndex("Qty")))

                    If DblTempQty <= DblQty Then
                        DblTemp = DblTemp + val(.TextMatrix(i, .ColIndex("ItemTotal")))
                    Else
                        'Stop
                        DblOneUnitPrice = val(.TextMatrix(i, .ColIndex("ItemTotal"))) / val(.TextMatrix(i, .ColIndex("Qty")))
                        DblTemp = DblTemp + (DblOneUnitPrice * (DblQty - (DblTempQty - val(.TextMatrix(i, .ColIndex("Qty"))))))
                    End If

                    .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = vbGreen
                    i = i + 1
                Loop

            End If

        End With

        Me.lbl(12).Caption = Format(DblTemp, SystemOptions.SysDefCurrencyForamt)
    
        '--------------Item Purchase Price
        Me.Ele(3).Visible = True
        DblTemp = FG.Aggregate(flexSTMin, FG.FixedRows, FG.ColIndex("ItemPrice"), FG.Rows - 1, FG.ColIndex("ItemPrice"))
        Me.lbl(20).Caption = DblTemp
        DblTemp = FG.Aggregate(flexSTMax, FG.FixedRows, FG.ColIndex("ItemPrice"), FG.Rows - 1, FG.ColIndex("ItemPrice"))
        Me.lbl(21).Caption = DblTemp
        '--------------------------------
        Msg = "„·ÕÊŸ…:-"
        Msg = Msg & "ÌÃ» „·«ÕŸ… «‰ «·›Ê« Ì— «Ê «·Õ—ﬂ«  «·„Ÿ··… »«··Ê‰ «·√Œ÷— "
        Msg = Msg & "ÂÏ «·√ Ï «⁄ „œ ⁄·ÌÂ« «·»—‰«„Ã ›Ï Õ”«» ﬁÌ„… «·„Œ“Ê‰ »‰Ÿ«„ «·Ê«—œ «Ê·« "
        Msg = Msg & "Ì’—› √Ê·«."
        Me.lbl(22).Caption = Msg
        Me.lbl(22).Visible = True
    End If

    If Not cProgress Is Nothing Then
        cProgress.StopProgess
        Set cProgress = Nothing
    End If

End Sub

