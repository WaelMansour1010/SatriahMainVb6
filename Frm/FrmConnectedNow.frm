VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{88F7F54F-F24B-4B64-B0E0-2454E1E6DA40}#1.0#0"; "ciaXPButton30.ocx"
Object = "{81FEA250-2DA5-40F7-A3F1-6F8532B748DB}#1.0#0"; "ciaXPPanel30.ocx"
Object = "{A7E76481-D275-422D-A506-9F4C890EE7D7}#1.0#0"; "ciaXPFrame30.ocx"
Object = "{8CD7576A-DE38-4ABA-A9D1-206DB09FED98}#1.0#0"; "ciaXPLabel30.ocx"
Begin VB.Form FrmConnectedNow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·„ ’·Ì‰ «·√‰"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   Icon            =   "FrmConnectedNow.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5565
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin ciaXPPanel30.XPPanel30 XPPanel301 
      Height          =   4995
      Left            =   0
      TabIndex        =   0
      Top             =   570
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   8811
      LicValid        =   -1  'True
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   3495
         Left            =   0
         TabIndex        =   2
         Top             =   30
         Width           =   6585
         _cx             =   11615
         _cy             =   6165
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmConnectedNow.frx":038A
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
         ExplorerBar     =   0
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
      Begin ciaXPButton30.XPButton30 Cmd 
         Cancel          =   -1  'True
         Height          =   405
         Index           =   6
         Left            =   60
         TabIndex        =   1
         Top             =   4500
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   714
         AutoSelectTheme =   -1  'True
         Caption         =   "062E06310648062C"
         PicturePosition =   2
         Picture         =   "FrmConnectedNow.frx":046E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseImageShadow  =   0   'False
         LicValid        =   -1  'True
      End
      Begin ciaXPFrame30.XPFrame30 XPFra 
         Height          =   1395
         Left            =   3690
         Top             =   3570
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   2461
         Alignment       =   2
         Caption         =   "0628064A062706460627062A00200627064406230634062A063106270643"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   1
         Radius          =   7
         CaptionPicture  =   "FrmConnectedNow.frx":0808
         CaptionPictureOnRight=   -1  'True
         LicValid        =   -1  'True
         Begin ciaXPLabel30.XPLabel30 XPLbl 
            Height          =   285
            Index           =   2
            Left            =   990
            Top             =   270
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   503
            BackStyle       =   1
            BackColor       =   16777215
            Border          =   0   'False
            CaptionPosition =   1
            Caption         =   "0639062F062F0020062706440645062A06350644064A064600200627064406230646"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            LicValid        =   -1  'True
         End
         Begin ciaXPLabel30.XPLabel30 XPLbl 
            Height          =   285
            Index           =   8
            Left            =   990
            Top             =   900
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   503
            BackStyle       =   1
            BackColor       =   16777215
            Border          =   0   'False
            CaptionPosition =   1
            Caption         =   "0639062F062F0020062706440623063906360627062100200627064406390627062F064A0646"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            LicValid        =   -1  'True
         End
         Begin ciaXPLabel30.XPLabel30 XPLbl 
            Height          =   285
            Index           =   9
            Left            =   990
            Top             =   600
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   503
            BackStyle       =   1
            BackColor       =   16777215
            Border          =   0   'False
            CaptionPosition =   1
            Caption         =   "0639062F062F002006270644062E06280631062706210020062706440645062A06350644064A0646"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            LicValid        =   -1  'True
         End
      End
   End
   Begin ciaXPLabel30.XPLabel30 XPLblHeader 
      Height          =   615
      Left            =   -30
      Top             =   -30
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   1085
      BackColor       =   16777215
      CaptionPosition =   1
      Caption         =   "062706440645062A06350644064A064600200627064406230646"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      LicValid        =   -1  'True
   End
End
Attribute VB_Name = "FrmConnectedNow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
CenterForm Me
With FG
    .AutoSize 1, .Cols - 1, False
End With
End Sub

