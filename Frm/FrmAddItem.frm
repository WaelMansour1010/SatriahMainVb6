VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmAddItem 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "≈÷«›… ’‰›"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4590
   Icon            =   "FrmAddItem.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   4590
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
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8190
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4590
      _cx             =   8096
      _cy             =   14446
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
      Appearance      =   0
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
      GridRows        =   2
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmAddItem.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin ImpulseButton.ISButton Cmd 
         Height          =   360
         Left            =   3570
         TabIndex        =   8
         Top             =   7800
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   635
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   " ÕœÌÀ"
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
         ButtonImage     =   "FrmAddItem.frx":03EF
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin C1SizerLibCtl.C1Tab TabMain 
         Height          =   7755
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   4530
         _cx             =   7990
         _cy             =   13679
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
         BackColor       =   12648447
         ForeColor       =   -2147483630
         FrontTabColor   =   14871017
         BackTabColor    =   12648447
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   16711680
         Caption         =   "⁄—÷ ‘Ã—Ï|⁄—÷ ÃœÊ·Ï"
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   3
         Position        =   1
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
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
         Picture(0)      =   "FrmAddItem.frx":0789
         Picture(1)      =   "FrmAddItem.frx":0B23
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   7290
            Index           =   1
            Left            =   5175
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   4440
            _cx             =   7832
            _cy             =   12859
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
            Begin VSFlex8UCtl.VSFlexGrid FgItems 
               Height          =   7200
               Left            =   30
               TabIndex        =   7
               Top             =   45
               Width           =   4380
               _cx             =   7726
               _cy             =   12700
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
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmAddItem.frx":0EBD
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
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   7290
            Index           =   0
            Left            =   45
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   45
            Width           =   4440
            _cx             =   7832
            _cy             =   12859
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
            Align           =   0
            AutoSizeChildren=   8
            BorderWidth     =   1
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
            _GridInfo       =   $"FrmAddItem.frx":0F7D
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   1065
               Index           =   3
               Left            =   15
               TabIndex        =   9
               TabStop         =   0   'False
               Top             =   6210
               Width           =   4410
               _cx             =   7779
               _cy             =   1879
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
                  Caption         =   "Label1"
                  Height          =   255
                  Index           =   5
                  Left            =   60
                  RightToLeft     =   -1  'True
                  TabIndex        =   15
                  Top             =   690
                  Width           =   1650
               End
               Begin VB.Image Img 
                  Height          =   255
                  Index           =   5
                  Left            =   1740
                  Top             =   690
                  Width           =   480
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Label1"
                  Height          =   255
                  Index           =   4
                  Left            =   60
                  RightToLeft     =   -1  'True
                  TabIndex        =   14
                  Top             =   420
                  Width           =   1650
               End
               Begin VB.Image Img 
                  Height          =   255
                  Index           =   4
                  Left            =   1740
                  Top             =   420
                  Width           =   480
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Label1"
                  Height          =   255
                  Index           =   3
                  Left            =   60
                  RightToLeft     =   -1  'True
                  TabIndex        =   13
                  Top             =   150
                  Width           =   1650
               End
               Begin VB.Image Img 
                  Height          =   255
                  Index           =   3
                  Left            =   1740
                  Top             =   150
                  Width           =   480
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Label1"
                  Height          =   255
                  Index           =   2
                  Left            =   2250
                  RightToLeft     =   -1  'True
                  TabIndex        =   12
                  Top             =   690
                  Width           =   1590
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Label1"
                  Height          =   255
                  Index           =   1
                  Left            =   2250
                  RightToLeft     =   -1  'True
                  TabIndex        =   11
                  Top             =   420
                  Width           =   1590
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Label1"
                  Height          =   255
                  Index           =   0
                  Left            =   2250
                  RightToLeft     =   -1  'True
                  TabIndex        =   10
                  Top             =   150
                  Width           =   1590
               End
               Begin VB.Image Img 
                  Height          =   255
                  Index           =   2
                  Left            =   3855
                  Top             =   690
                  Width           =   480
               End
               Begin VB.Image Img 
                  Height          =   255
                  Index           =   1
                  Left            =   3855
                  Top             =   420
                  Width           =   480
               End
               Begin VB.Image Img 
                  Height          =   255
                  Index           =   0
                  Left            =   3855
                  Top             =   150
                  Width           =   480
               End
            End
            Begin MSComctlLib.TreeView TrvItems 
               Height          =   6180
               Left            =   15
               TabIndex        =   6
               Top             =   15
               Width           =   4410
               _ExtentX        =   7779
               _ExtentY        =   10901
               _Version        =   393217
               Indentation     =   18
               LineStyle       =   1
               Style           =   7
               Appearance      =   1
            End
         End
      End
      Begin ImpulseButton.ISButton XPBtnOK 
         Default         =   -1  'True
         Height          =   360
         Left            =   1185
         TabIndex        =   4
         Top             =   7800
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   635
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
         ButtonImage     =   "FrmAddItem.frx":0FFA
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
      Begin ImpulseButton.ISButton XPBtnCancel 
         Cancel          =   -1  'True
         Height          =   360
         Left            =   30
         TabIndex        =   5
         Top             =   7800
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   635
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
         ButtonImage     =   "FrmAddItem.frx":1394
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
   End
End
Attribute VB_Name = "FrmAddItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XNode As MSComctlLib.Node

Public cGrid As ClsGrid

Private Sub Chk_Click(Index As Integer)

End Sub

Private Sub Cmd_Click()
    LoadData
End Sub

Private Sub Form_Activate()
    PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim SrSQL As String
    Dim rs As ADODB.Recordset
    Dim GrdBack As ClsBackGroundPic
    Dim i As Integer

    With Me.TrvItems
        .LineStyle = tvwRootLines
        .Sorted = False
        .LabelEdit = tvwManual
        .OLEDragMode = ccOLEDragAutomatic
        .OLEDropMode = ccOLEDropNone
        Set .ImageList = mdifrmmain.ImgLstTree
    End With

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    Else
        Make_RightToLeft Me.TrvItems
    End If

    With Me.FgItems
        Set GrdBack = New ClsBackGroundPic
        Set .WallPaper = GrdBack.Picture
        LoadData
    End With

    'Closed_Node
    Me.Img(0).Picture = mdifrmmain.ImgLstTree.ListImages("Closed_Node").Picture
    Me.lbl(0).Caption = "„Ã„Ê⁄… „€·ﬁ…"
    Me.Img(1).Picture = mdifrmmain.ImgLstTree.ListImages("Open_Node").Picture
    Me.lbl(1).Caption = "„Ã„Ê⁄… „„ œ…"
    Me.Img(2).Picture = mdifrmmain.ImgLstTree.ListImages("Item").Picture
    Me.lbl(2).Caption = "’‰› ⁄«œÏ"

    Me.Img(3).Picture = mdifrmmain.ImgLstTree.ListImages("Assblied").Picture
    Me.lbl(3).Caption = "’‰› „Ã„⁄"

    Me.Img(4).Picture = mdifrmmain.ImgLstTree.ListImages("ItemPart").Picture
    Me.lbl(4).Caption = "’‰› „‰ «·’‰› «·„Ã„⁄"

    Me.Img(5).Picture = mdifrmmain.ImgLstTree.ListImages("LinkItem").Picture
    Me.lbl(5).Caption = "’‰› „·Õﬁ"

    Me.Width = val(GetSetting(SystemOptions.SysRegsAppPath, "FormsSetting\" & Me.Name, "Width", 4170))
    Me.Height = val(GetSetting(SystemOptions.SysRegsAppPath, "FormsSetting\" & Me.Name, "Height", 5730))
    CenterForm Me

    FormPostion Me, GetPostion

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    If UnloadMode = QueryUnloadConstants.vbFormControlMenu Then
        If Not cGrid Is Nothing Then
            cGrid.GrdTBar.Buttons("Tree").value = tbrUnpressed
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    SaveSetting SystemOptions.SysRegsAppPath, "FormsSetting\" & Me.Name, "Width", Me.Width
    SaveSetting SystemOptions.SysRegsAppPath, "FormsSetting\" & Me.Name, "Height", Me.Height

End Sub

Private Sub TrvItems_DblClick()
    Dim YNode As MSComctlLib.Node
    Dim LngSelectedItemID As Long
    Dim LngGridRow As Long

    Set YNode = TrvItems.SelectedItem

    If Not YNode Is Nothing Then
        If right$(YNode.Key, 1) = "G" Then 'This is Group Node (NOT Allowd to Add To the Grid)
            Exit Sub
        End If

        If Not cGrid Is Nothing Then
            If cGrid.TxtModFlag.Text = "R" Then Exit Sub
            LngSelectedItemID = val(YNode.Key)
            LngGridRow = cGrid.FindItemInGrid(0, LngSelectedItemID)

            If LngGridRow = -1 Then
                cGrid.AddItemInGrid cGrid.NewEmptyRow, LngSelectedItemID
            End If
        End If
    End If

End Sub

Private Sub TrvItems_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    Set XNode = TrvItems.HitTest(X, Y)
End Sub

Private Sub TrvItems_OLEStartDrag(Data As MSComctlLib.DataObject, _
                                  AllowedEffects As Long)

    If Not XNode Is Nothing Then
        If right$(XNode.Key, 1) = "G" Then 'This is Group Node (NOT Allowd to Drag)
            AllowedEffects = vbDropEffectNone
            Exit Sub
        End If

        AllowedEffects = vbDropEffectCopy
        Data.SetData XNode.Key, vbCFText
        '    Load FrmItemTip
        '    ShowWindow FrmItemTip.hwnd, SW_SHOWNA
    Else
        AllowedEffects = vbDropEffectNone
    End If

End Sub

Private Sub XPBtnCancel_Click()
    Unload Me
End Sub

Private Sub ChangeLang()
    Me.Caption = "Add Item"
    TabMain.TabCaption(0) = "Tree View"
    TabMain.TabCaption(1) = "Table View"
    Me.Cmd.Caption = "Refresh"
    Me.XPBtnCancel.Caption = "Cancel"
    Me.XPBtnOK.Caption = "OK"

    With Me.FgItems
        .TextMatrix(0, .ColIndex("ItemID")) = "Item ID"
        .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
        .TextMatrix(0, .ColIndex("GoupName")) = "Group Name"
        .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
    End With

End Sub

Private Sub LoadData()
    Dim SrSQL As String
    Dim rs As ADODB.Recordset

    With Me.FgItems
        Set rs = New ADODB.Recordset
        StrSQL = "SELECT TblItems.ItemID, TblItems.ItemCode, Groups.GroupName, TblItems.ItemName"
        StrSQL = StrSQL + " FROM TblItems INNER JOIN "
        StrSQL = StrSQL + " Groups ON TblItems.GroupID = Groups.GroupID "
        StrSQL = StrSQL + ""
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        Set .DataSource = rs

        If SystemOptions.UserInterface = EnglishInterface Then

            For i = 0 To .Cols - 1
                .ColAlignment(i) = flexAlignLeftCenter
                .FixedAlignment(i) = flexAlignLeftCenter
            Next i

            .TextMatrix(0, .ColIndex("ItemID")) = "Item ID"
            .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
            .TextMatrix(0, .ColIndex("GoupName")) = "Group Name"
            .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
        End If

        .AutoSize 0, .Cols - 1, False
    End With

    Me.TrvItems.Nodes.Clear
    ModTree.LoadTreeGroups Me.TrvItems
End Sub

