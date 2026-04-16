VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{A3550A07-56EC-11D3-8DC5-00409503C9B8}#1.0#0"; "axbarcode.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmParent 
   BackColor       =   &H00C0C0FF&
   Caption         =   "ãÚÇíäÉ"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9030
   Icon            =   "FrmParent.frx":0000
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5970
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   5640
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   330
      Width           =   9030
      _cx             =   15928
      _cy             =   9948
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
      BackColor       =   16777215
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
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
      GridRows        =   2
      GridCols        =   2
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmParent.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.VScrollBar VScrol 
         Height          =   5355
         Left            =   8775
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   15
         Width           =   240
      End
      Begin VB.HScrollBar HScrol 
         Height          =   240
         Left            =   15
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   5385
         Width           =   8745
      End
      Begin VB.PictureBox PicParent 
         BackColor       =   &H00FFFFFF&
         Height          =   5355
         Left            =   15
         ScaleHeight     =   5295
         ScaleWidth      =   8685
         TabIndex        =   6
         Top             =   15
         Width           =   8745
         Begin VB.PictureBox PicStikers 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4050
            Left            =   345
            RightToLeft     =   -1  'True
            ScaleHeight     =   71.438
            ScaleMode       =   6  'Millimeter
            ScaleWidth      =   139.965
            TabIndex        =   9
            Top             =   -15
            Width           =   7935
            Begin AXBARCODELib.Axbarcode AXBar 
               Height          =   855
               Left            =   3735
               TabIndex        =   10
               Top             =   300
               Visible         =   0   'False
               Width           =   1695
               _Version        =   65536
               _ExtentX        =   2990
               _ExtentY        =   1508
               _StockProps     =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ShowLightMargins=   0   'False
               ShowText        =   0   'False
            End
            Begin VB.Image IMG 
               Height          =   225
               Index           =   0
               Left            =   405
               Stretch         =   -1  'True
               Top             =   225
               Width           =   345
            End
            Begin VB.Label UpLbl 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Height          =   240
               Index           =   0
               Left            =   510
               TabIndex        =   12
               Top             =   825
               Visible         =   0   'False
               Width           =   390
            End
            Begin VB.Label DNLbl 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Height          =   240
               Index           =   0
               Left            =   495
               TabIndex        =   11
               Top             =   825
               Visible         =   0   'False
               Width           =   390
            End
         End
      End
   End
   Begin MSComctlLib.Toolbar Tbar 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9030
      _ExtentX        =   15928
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sp"
            Style           =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbrFirst"
            ImageKey        =   "First"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbrPrev"
            ImageKey        =   "Prev"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbrNext"
            ImageKey        =   "Next"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbrLast"
            ImageKey        =   "Last"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   4
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Setup"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbrExt"
            ImageKey        =   "Ext"
         EndProperty
      EndProperty
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1260
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   30
         Width           =   495
         Begin VB.Label LblPageNum 
            Caption         =   "Lab"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   105
            TabIndex        =   4
            Top             =   30
            Width           =   315
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   0
         Width           =   435
         Begin VB.Label LblPageCount 
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   75
            TabIndex        =   2
            Top             =   60
            Width           =   285
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6555
      Top             =   2505
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParent.frx":03D7
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParent.frx":0771
            Key             =   "First"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParent.frx":0B0B
            Key             =   "Prev"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParent.frx":0EA5
            Key             =   "Next"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParent.frx":123F
            Key             =   "Last"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParent.frx":15D9
            Key             =   "Ext"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParent.frx":1973
            Key             =   "Setup"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    FrmParent.LblPageNum.Caption = cBarcode.PageNumber
    FrmParent.LblPageCount.Caption = cBarcode.PageCount
End Sub

Private Sub Form_Load()
    PicStikers.Width = 21 * 567
    PicStikers.Height = 29 * 567
End Sub

Private Sub Form_Resize()
    On Error GoTo ErrTrap

    PicStikers.top = 0
    PicStikers.left = 1800

    If PicStikers.Height > Me.Height Then
        Me.VScrol.Max = (PicStikers.Height - (Me.Height)) + 2600
        Me.VScrol.LargeChange = PicStikers.Height / 10
        Me.VScrol.SmallChange = PicStikers.Height / 20
        Me.VScrol.value = 0
        Me.VScrol.Visible = True
    Else
        Me.VScrol.Visible = False
    End If

    If PicStikers.Width + 1800 > Me.Width Then
        Me.HScrol.Max = (PicStikers.Width - (Me.Width)) + 4000
        Me.HScrol.LargeChange = PicStikers.Width / 10
        Me.HScrol.SmallChange = PicStikers.Width / 20
        Me.HScrol.value = 0
        PicStikers.left = (3000 - HScrol.value) * -1
        Me.HScrol.Visible = True
    Else
        Me.HScrol.Visible = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub HScrol_Change()
    PicStikers.left = (3000 - HScrol.value) * -1
End Sub

Private Sub Tbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo ErrTrap
    Dim clprnt As ClsBarcode

    Select Case Button.key

        Case "Print"
            cBarcode.PrintPage 0

        Case "tbrExt"
            Unload Me
            Exit Sub

        Case "tbrNext"
            cBarcode.MoveNext

        Case "tbrFirst"
            cBarcode.MoveFirst

        Case "tbrPrev"
            cBarcode.MovePrev

        Case "tbrLast"
            cBarcode.MoveLast

        Case "Setup"
            'FrmSetting.show vbModal
            'cBarcode.Preview

    End Select

    'PicStikers.Top = 0
    'PicStikers.Left = 1800
    LblPageNum.Caption = cBarcode.PageNumber
    LblPageCount.Caption = cBarcode.PageCount
ErrTrap:
End Sub

Private Sub VScrol_Change()
    PicStikers.top = -VScrol.value
End Sub
