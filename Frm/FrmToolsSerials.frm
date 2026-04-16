VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{D95CB779-00CB-4B49-97B9-9F0B61CAB3C1}#4.0#0"; "biokey.ocx"
Begin VB.Form FrmToolsSerials 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4470
   Icon            =   "FrmToolsSerials.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6900
   ScaleWidth      =   4470
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin ZKFPEngXControl.ZKFPEngX ZKFPEngX1 
      Left            =   1440
      Top             =   5520
      EnrollCount     =   3
      SensorIndex     =   0
      Threshold       =   10
      VerTplFileName  =   ""
      RegTplFileName  =   ""
      OneToOneThreshold=   10
      Active          =   0   'False
      IsRegister      =   0   'False
      EnrollIndex     =   0
      SensorSN        =   ""
      FPEngineVersion =   "9"
      ImageWidth      =   0
      ImageHeight     =   0
      SensorCount     =   0
      TemplateLen     =   1152
      EngineValid     =   0   'False
      ForceSecondMatch=   0   'False
      IsReturnNoLic   =   -1  'True
      LowestQuality   =   30
      FakeFunOn       =   1
   End
   Begin VB.CommandButton cmdInit 
      Caption         =   "⁄—÷ «·»’„Â"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   600
      TabIndex        =   30
      Top             =   5040
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton cmdSaveImage 
      Caption         =   "Õ›Ÿ"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   29
      Top             =   6000
      Width           =   3135
   End
   Begin VB.Frame FrameCommands 
      Height          =   7695
      Left            =   5760
      TabIndex        =   12
      Top             =   7680
      Visible         =   0   'False
      Width           =   5655
      Begin ZKFPEngXControl.ZKFPEngX ZKFPEngX2 
         Left            =   5760
         Top             =   840
         EnrollCount     =   3
         SensorIndex     =   0
         Threshold       =   10
         VerTplFileName  =   ""
         RegTplFileName  =   ""
         OneToOneThreshold=   10
         Active          =   0   'False
         IsRegister      =   0   'False
         EnrollIndex     =   0
         SensorSN        =   ""
         FPEngineVersion =   "9"
         ImageWidth      =   0
         ImageHeight     =   0
         SensorCount     =   0
         TemplateLen     =   1152
         EngineValid     =   0   'False
         ForceSecondMatch=   0   'False
         IsReturnNoLic   =   -1  'True
         LowestQuality   =   30
         FakeFunOn       =   1
      End
      Begin VB.Frame Frame2 
         Caption         =   "Image Format"
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         TabIndex        =   21
         Top             =   1560
         Width           =   2415
         Begin VB.OptionButton OptionJpg 
            Caption         =   "JPEG"
            BeginProperty Font 
               Name            =   "SimSun"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   960
            TabIndex        =   23
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptionBmp 
            Caption         =   "BMP"
            BeginProperty Font 
               Name            =   "SimSun"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox TextSensorSN 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1560
         TabIndex        =   20
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox TextSensorIndex 
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3600
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox TextSensorCount 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1560
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox TextFingerName 
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2640
         TabIndex        =   17
         Text            =   "fingerprint1"
         Top             =   2520
         Width           =   1935
      End
      Begin VB.CommandButton cmdIdentify 
         Caption         =   "identify(1:N)"
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   16
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CommandButton cmdVerify 
         Caption         =   "Verify(1:1)"
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CommandButton cmdEnroll 
         Caption         =   "Register"
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close  Sensor"
         Height          =   375
         Left            =   1920
         TabIndex        =   13
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Serial Number"
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Current"
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   27
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Sensor Cnt"
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   25
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "SimSun"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   24
         Top             =   2640
         Width           =   735
      End
   End
   Begin VB.TextBox XPTxtCode 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox TxtItemCode 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   8250
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1140
      Width           =   1575
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   585
      Left            =   6240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7125
      _cx             =   12568
      _cy             =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   18
         Charset         =   178
         Weight          =   700
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
      Picture         =   "FrmToolsSerials.frx":0A02
      Caption         =   "ﬂ‘› √Œÿ«¡ «·”Ì—Ì«· ‰„»— ··√’‰«›"
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
      PicturePos      =   1
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
   End
   Begin VSFlex8UCtl.VSFlexGrid FG 
      Height          =   3285
      Left            =   6300
      TabIndex        =   2
      Top             =   1890
      Width           =   7005
      _cx             =   12356
      _cy             =   5794
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
      Rows            =   15
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmToolsSerials.frx":16DC
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
      WallPaperAlignment=   4
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   5370
      TabIndex        =   3
      Top             =   5370
      Width           =   735
      _ExtentX        =   1296
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
      Left            =   7215
      TabIndex        =   4
      Top             =   5370
      Width           =   735
      _ExtentX        =   1296
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
      Left            =   6420
      TabIndex        =   5
      Top             =   5370
      Width           =   735
      _ExtentX        =   1296
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
   Begin MSDataListLib.DataCombo DcboItemName 
      Height          =   315
      Left            =   5340
      TabIndex        =   6
      Top             =   1530
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton CmdItemSearch 
      Height          =   345
      Left            =   7530
      TabIndex        =   7
      Top             =   1500
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
      ButtonImage     =   "FrmToolsSerials.frx":1850
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Label LBLID 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   0
      TabIndex        =   32
      Top             =   7440
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label LblPath 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   6840
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂÊœ «·’‰›"
      Height          =   315
      Index           =   7
      Left            =   9780
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1230
      Width           =   825
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·’‰›"
      Height          =   315
      Index           =   5
      Left            =   9780
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1530
      Width           =   825
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«ﬂ » ﬂÊœ «·’‰› À„ ≈÷€ÿ ≈‰ —"
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
      Height          =   225
      Index           =   36
      Left            =   5940
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1200
      Width           =   2265
   End
End
Attribute VB_Name = "FrmToolsSerials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Dim FTempLen As Integer
Dim FRegTemplate As String
Dim FRegTemp As Variant
Dim FingerCount As Long
Dim fpcHandle As Long
Dim FFingerNames() As String
Dim FMatchType As Integer

Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long


Dim cDcboSearch As clsDCboSearch

Private Sub Cmd_Click(Index As Integer)
    Dim Msg As String
    Dim StrSQL As String
    On Error GoTo ErrTrap

    Select Case Index

        Case 0
            SearchSerials

        Case 1
            XPTxtCode.Text = ""
            Me.DcboItemName.BoundText = ""
            Me.TxtItemCode.Text = ""
            LblPlace.Caption = ""
            FG.Clear flexClearScrollable, flexClearEverything

        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

End Sub

Public Sub cmdInit_Click()
   ZKFPEngX1.SensorIndex = 0
  If ZKFPEngX1.InitEngine = 0 Then
  
      Me.Caption = " „ «· ‘€Ì·"
     TextSensorCount.Text = ZKFPEngX1.SensorCount & ""
     TextSensorIndex.Text = ZKFPEngX1.SensorIndex & ""
     TextSensorSN.Text = ZKFPEngX1.SensorSN
     
     cmdInit.Enabled = False
     FMatchType = 0
  End If
  
End Sub

Private Sub ZKFPEngX1_OnImageReceived(AImageValid As Boolean)
 ZKFPEngX1.PrintImageAt hDC, 0, 0, ZKFPEngX1.ImageWidth, ZKFPEngX1.ImageHeight
 
  End Sub

Private Sub CmdItemSearch_Click()
    Load FrmItemSearch
    FrmItemSearch.RetrunType = 1
    Set FrmItemSearch.DcboItems = Me.DcboItemName
    FrmItemSearch.show vbModal
End Sub

Private Sub cmdSaveImage_Click()
Dim lastfolder As String
Dim sFileName As String

Dim ID As String
lastfolder = LblPath.Caption
ID = LBLID.Caption
 
     sFileName = App.path & "\IMAGES\FP\"   '«‰‘«¡ «·›Ê·œ— «·«”«”ÌÂ
     If Dir(sFileName) = "" Then
    MkDir sFileName
End If

    sFileName = App.path & "\IMAGES\FP\" & lastfolder & "\" '«‰‘«¡ «·›Ê·œ— «·›—⁄Ì…
    
    If Dir(sFileName) = "" Then
'    MkDir sFileName
End If
 
 
        ZKFPEngX1.SaveJPG sFileName + ID + ".jpg"
 
 RsCustomers.ISButton2_Click
DoEvents
Unload Me
End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim BG As New ClsBackGroundPic
    On Error GoTo ErrTrap

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    CenterForm Me
    Set Dcombos = New ClsDataCombos
    Dcombos.GetItemsNames DcboItemName
    Set cDcboSearch = New clsDCboSearch
    Set cDcboSearch.Client = DcboItemName
    FG.WallPaper = BG.SearchWallpaper

    FormPostion Me, GetPostion
    Exit Sub
ErrTrap:
End Sub

Private Sub TxtItemCode_KeyDown(KeyCode As Integer, _
                                Shift As Integer)
    Dim Msg As String
    Dim StrSQL As String
    Dim rs As ADODB.Recordset

    If KeyCode = vbKeyReturn Then
        If Trim(Me.TxtItemCode.Text) = "" Then Exit Sub
        StrSQL = "Select ItemID From TblItems Where ItemCode='" & Trim(Me.TxtItemCode.Text) & "'"
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            DcboItemName.BoundText = rs("ItemID").value
        
        Else
            DcboItemName.BoundText = ""
            Msg = "·«ÌÊÃœ ’‰› „”Ã· »Â–« «·ﬂÊœ..!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        End If
    End If

End Sub

Private Sub SearchSerials()
    Dim rs As ADODB.Recordset
    Dim Num As Integer

    If DcboItemName.BoundText = "" Then
        Msg = "«ﬂ » «”„ «·’‰› «·„—«œ «·»ÕÀ ⁄‰Â ...!!! "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DcboItemName.SetFocus
        Exit Sub
    End If

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "SELECT Transaction_Details.ItemSerial, Transactions.Transaction_ID, " & "Transactions.Transaction_Serial, Transactions.Transaction_Type," & "Transaction_Details.Item_ID, TransactionTypes.TransactionTypeName"
        StrSQL = StrSQL + " FROM TransactionTypes INNER JOIN (Transactions INNER JOIN " & "Transaction_Details ON Transactions.Transaction_ID = Transaction_Details.Transaction_ID)" & "ON TransactionTypes.Transaction_Type = Transactions.Transaction_Type "
        StrSQL = StrSQL + " WHERE (Transaction_Details.Item_ID=" & Me.DcboItemName.BoundText & ")  AND   Transaction_Details.ItemSerial NOT IN"
        StrSQL = StrSQL + "("
        StrSQL = StrSQL + " SELECT Transaction_Details.ItemSerial"
        StrSQL = StrSQL + " FROM TransactionTypes INNER JOIN (Transactions INNER JOIN " & "Transaction_Details ON Transactions.Transaction_ID = Transaction_Details.Transaction_ID)" & " ON TransactionTypes.Transaction_Type = Transactions.Transaction_Type"
        StrSQL = StrSQL + " Where (Transaction_Details.Item_ID = " & Me.DcboItemName.BoundText & ") And (Transactions.Transaction_Type = 1 OR Transactions.Transaction_Type = 3));"

    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT Transaction_Details.ItemSerial, Transactions.Transaction_ID, " & "Transactions.Transaction_Serial, Transactions.Transaction_Type," & "Transaction_Details.Item_ID, TransactionTypes.TransactionTypeName"
        StrSQL = StrSQL + " FROM TransactionTypes INNER JOIN (Transactions INNER JOIN " & "Transaction_Details ON Transactions.Transaction_ID = Transaction_Details.Transaction_ID)" & "ON TransactionTypes.Transaction_Type = Transactions.Transaction_Type "
        StrSQL = StrSQL + " WHERE (Transaction_Details.Item_ID=" & Me.DcboItemName.BoundText & ")  AND   Transaction_Details.ItemSerial NOT IN"
        StrSQL = StrSQL + "("
        StrSQL = StrSQL + " SELECT Transaction_Details.ItemSerial"
        StrSQL = StrSQL + " FROM TransactionTypes INNER JOIN (Transactions INNER JOIN " & "Transaction_Details ON Transactions.Transaction_ID = Transaction_Details.Transaction_ID)" & " ON TransactionTypes.Transaction_Type = Transactions.Transaction_Type"
        StrSQL = StrSQL + " Where (Transaction_Details.Item_ID = " & Me.DcboItemName.BoundText & ") And (Transactions.Transaction_Type = 1 OR Transactions.Transaction_Type = 3));"

    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        FG.Clear flexClearScrollable, flexClearEverything
        FG.Rows = 1
        Msg = "·«  ÊÃœ √Ì »Ì«‰« ..... "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
        rs.MoveFirst
        FG.Rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With FG
                '            If Opt(2).Value = True Then
                .TextMatrix(Num, .ColIndex("NumIndex")) = Num
                .TextMatrix(Num, .ColIndex("WrongSerial")) = IIf(IsNull(rs("ItemSerial").value), "", (rs("ItemSerial").value))
                .TextMatrix(Num, .ColIndex("TranseNum")) = IIf(IsNull(rs("Transaction_ID").value), "", (rs("Transaction_ID").value))
                .TextMatrix(Num, .ColIndex("TransType")) = IIf(IsNull(rs("TransactionTypeName").value), "", (rs("TransactionTypeName").value))
                .Cell(flexcpData, Num, .ColIndex("TransType")) = IIf(IsNull(rs("Transaction_Type").value), "", (rs("Transaction_Type").value))
                .TextMatrix(Num, .ColIndex("Transaction_Serial")) = IIf(IsNull(rs("Transaction_Serial").value), "", (rs("Transaction_Serial").value))
                '                .TextMatrix(Num, .ColIndex("ItemPlace")) = _
                '                IIf(IsNull(RS("StoreName").Value), "", (RS("StoreName").Value))
                '                .TextMatrix(Num, .ColIndex("TransDate")) = _
                '                IIf(IsNull(RS("Transaction_Date").Value), "", Format((RS("Transaction_Date").Value), "yyyy/m/d"))
                '                ' ÕœÌœ „ﬂ«‰ «·ﬁÿ⁄Â
                '                StrSQL = "select * From QryGardComplete where ItemSerial='" & XPTxtCode.Text & "'"
                '                If Me.DcboItemName.BoundText <> "" Then
                '                    StrSQL = StrSQL + " and ItemID=" & Me.DcboItemName.BoundText & ""
                '                End If
                '                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                '               If RsTemp("QTY").Value > 0 Then
                '                   LblPlace.Caption = "„ÊÃÊœ… ›Ì «·„Œ“‰ " & RsTemp("StoreName").Value
                '                Else
                '                    LblPlace.Caption = "€Ì— „ÊÃÊœ… ›Ì «·„Œ“‰/«·„Œ«“‰"
                '               End If
                '               RsTemp.Close
            End With

            rs.MoveNext
        Next Num

    End If

End Sub

Private Sub Retrive()
    Dim Num As Integer
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim Msg As String
    Dim StrTemp As String

    On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveFirst
        StrTemp = rs("ItemID").value

        If rs.RecordCount > 1 Then

            For Num = 1 To rs.RecordCount

                If StrTemp <> rs("ItemID").value Then
                    Msg = "Â–« «·”Ì—Ì«· ”Ã· ›Ï «·»—‰«„Ã „⁄ «ﬂÀ—"
                    Msg = Msg & CHR(13) & "„‰ ’‰› „Œ ·› Ê·–« ÌÃ»  ÕœÌœ «”„ «·’‰›...!!"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If

                rs.MoveNext
            Next Num

        End If

        rs.MoveFirst
        FG.Rows = rs.RecordCount + 1
        LblRemark.Visible = True
        'XPLblItemName.Caption = Rs("ItemName").Value
        'XPLblItemCode.Caption = IIf(IsNull(Rs("ItemCode").Value), "", Rs("ItemCode").Value)
        Me.TxtItemCode.Text = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
        Me.DcboItemName.BoundText = StrTemp

        For Num = 1 To rs.RecordCount

            With FG
                '            If Opt(2).Value = True Then
                .TextMatrix(Num, .ColIndex("NumIndex")) = Num
                .TextMatrix(Num, .ColIndex("Transaction_Serial")) = IIf(IsNull(rs("Transaction_Serial").value), "", (rs("Transaction_Serial").value))
                .TextMatrix(Num, .ColIndex("Client")) = IIf(IsNull(rs("CusName").value), "", (rs("CusName").value))
                .TextMatrix(Num, .ColIndex("TranseNum")) = IIf(IsNull(rs("Transaction_ID").value), "", (rs("Transaction_ID").value))
                .TextMatrix(Num, .ColIndex("TransType")) = IIf(IsNull(rs("TransactionTypeName").value), "", (rs("TransactionTypeName").value))
                .Cell(flexcpData, Num, .ColIndex("TransType")) = IIf(IsNull(rs("Transaction_Type").value), "", (rs("Transaction_Type").value))
                .TextMatrix(Num, .ColIndex("ItemPlace")) = IIf(IsNull(rs("StoreName").value), "", (rs("StoreName").value))
                .TextMatrix(Num, .ColIndex("TransDate")) = IIf(IsNull(rs("Transaction_Date").value), "", Format((rs("Transaction_Date").value), "yyyy/m/d"))
                ' ÕœÌœ „ﬂ«‰ «·ﬁÿ⁄Â
                StrSQL = "select * From QryGardComplete where ItemSerial='" & XPTxtCode.Text & "'"

                If Me.DcboItemName.BoundText <> "" Then
                    StrSQL = StrSQL + " and ItemID=" & Me.DcboItemName.BoundText & ""
                End If

                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp("QTY").value > 0 Then
                    LblPlace.Caption = "„ÊÃÊœ… ›Ì «·„Œ“‰ " & RsTemp("StoreName").value
                Else
                    LblPlace.Caption = "€Ì— „ÊÃÊœ… ›Ì «·„Œ“‰/«·„Œ«“‰"
                End If

                RsTemp.Close
            End With

            rs.MoveNext
        Next Num

    Else
        FG.Clear flexClearScrollable, flexClearEverything
        FG.Rows = 1
        LblPlace.Caption = ""
        LblRemark.Visible = False
        Msg = "·«  ÊÃœ √Ì »Ì«‰«  ⁄‰ Â–Â «·ﬁÿ⁄… "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Exit Sub
ErrTrap:
End Sub

