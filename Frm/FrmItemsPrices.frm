VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmItemsPrices 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ê÷⁄ Œÿ… √”⁄«—  „Ã„Ê⁄«  «·√’‰«›"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8130
   Icon            =   "FrmItemsPrices.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6225
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   4785
      Left            =   30
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   930
      Width           =   4155
      _cx             =   7329
      _cy             =   8440
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
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   2
         Left            =   1590
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   2430
         Width           =   975
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "”⁄— À«» "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   405
         Index           =   2
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   2010
         Width           =   4035
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   1
         Left            =   1590
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1590
         Width           =   975
      End
      Begin VB.TextBox Txt 
         Alignment       =   1  'Right Justify
         Height          =   375
         Index           =   0
         Left            =   1590
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "”⁄— ‘—«¡ «·’‰› + ‰”»… „∆ÊÌ… „‰ ”⁄— «·‘—«¡"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   405
         Index           =   1
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1140
         Width           =   4035
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "”⁄— ‘—«¡ «·’‰› + „»·€ À«» "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   0
         Left            =   540
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   390
         Value           =   -1  'True
         Width           =   3525
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   420
         Left            =   60
         TabIndex        =   11
         Top             =   4320
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   741
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
         BackStyle       =   0
         ButtonImage     =   "FrmItemsPrices.frx":058A
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
         Caption         =   "«·”⁄— «·À«» "
         Height          =   285
         Index           =   3
         Left            =   2580
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   2460
         Width           =   1425
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "%"
         Height          =   255
         Index           =   2
         Left            =   1290
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1650
         Width           =   285
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·‰”»… «·„∆ÊÌ…"
         Height          =   285
         Index           =   1
         Left            =   2580
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1650
         Width           =   1425
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·„»·€ «·À«» "
         Height          =   285
         Index           =   0
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   750
         Width           =   1245
      End
   End
   Begin MSComctlLib.TreeView TreeGroups 
      Height          =   4785
      Left            =   4200
      TabIndex        =   0
      Top             =   930
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   8440
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin ImpulseButton.ISButton CmdExit 
      Height          =   390
      Left            =   60
      TabIndex        =   12
      Top             =   5760
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   688
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
      ButtonImage     =   "FrmItemsPrices.frx":0924
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "÷⁄ ⁄·«„… ’Õ «„«„ «·„Ã„Ê⁄«  «· Ï  —Ìœ  ÕœÌœ √”⁄«—Â«"
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
      Left            =   4170
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   660
      Width           =   3885
   End
   Begin VB.Label LblHeader 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ê÷⁄ Œÿ… √”⁄«— „Ã„Ê⁄«  «·√’‰«›"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   585
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   30
      Width           =   8055
   End
End
Attribute VB_Name = "FrmItemsPrices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd_Click()
    Dim StrSQL As String
    Dim i As Integer
    Dim Msg As String

    If Me.Opt(0).value = True Then
        If val(Me.Txt(0).text) = 0 Then
            Exit Sub
        End If

        For i = 1 To Me.TreeGroups.Nodes.count

            If Me.TreeGroups.Nodes(i).Checked = True Then
                StrSQL = "Update TblItems Set SallingPrice=(PurchasePrice +" & val(Me.Txt(0).text)
                StrSQL = StrSQL + ") Where TblItems.GroupID=" & val(Me.TreeGroups.Nodes(i).key)
                Cn.Execute StrSQL, , adExecuteNoRecords
            End If

        Next i

    ElseIf Me.Opt(1).value = True Then

        For i = 1 To Me.TreeGroups.Nodes.count

            If Me.TreeGroups.Nodes(i).Checked = True Then
                StrSQL = "Update TblItems Set SallingPrice=(PurchasePrice + (PurchasePrice * (" & val(Me.Txt(1).text) / 100
                StrSQL = StrSQL + "))) Where TblItems.GroupID=" & val(Me.TreeGroups.Nodes(i).key)
                Cn.Execute StrSQL, , adExecuteNoRecords
            End If

        Next i

    ElseIf Me.Opt(2).value = True Then

        For i = 1 To Me.TreeGroups.Nodes.count

            If Me.TreeGroups.Nodes(i).Checked = True Then
                StrSQL = "Update TblItems Set SallingPrice=" & val(Me.Txt(2).text)
                StrSQL = StrSQL + " Where TblItems.GroupID=" & val(Me.TreeGroups.Nodes(i).key)
                Cn.Execute StrSQL, , adExecuteNoRecords
            End If

        Next i

    End If

    Msg = " „  ⁄„·Ì… «· ÕœÌÀ"
    MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Set TreeGroups.ImageList = mdifrmmain.ImgLstTree
    CenterForm Me
    LoadTreeGroups Me.TreeGroups
End Sub

Private Sub LoadTreeGroups(ItemsTree As MSComctlLib.TreeView)
    Dim Rs_items As ADODB.Recordset
    Dim My_SQL As String
    Dim nodX As Node
    Dim nodz As Node
    Dim RsOptions As ADODB.Recordset
    Dim my_ch_rs As ADODB.Recordset
    Dim BolDisplayArabic As Boolean
    Dim LngLoop As Long
    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = ArabicInterface Then
        BolDisplayArabic = True
        ItemsTree.Tag = "A"
        Make_RightToLeft ItemsTree
        '''''''''''''''''''''''''''add root
        Set nodX = ItemsTree.Nodes.Add(, , "1G", "„Ã„Ê⁄… «·√’‰«›", "Root")
        ItemsTree.Nodes("1G").Expanded = True
    Else
        BolDisplayArabic = False
        '''''''''''''''''''''''''''add root
        ItemsTree.Tag = "E"
        Set nodX = ItemsTree.Nodes.Add(, , "1G", "Groups of items", "Root")
        ItemsTree.Nodes("1G").Expanded = True
    End If

    Me.TreeGroups.Sorted = False
    '''''''''''''''''''''''''''' add group
    My_SQL = " SELECT Groups.* "
    My_SQL = My_SQL + "  From Groups "
    My_SQL = My_SQL + " where (ParentID =1); "
    Set my_ch_rs = New ADODB.Recordset
    my_ch_rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    BolDisplayArabic = True

    If BolDisplayArabic = True Then
        Call fill_my_Children("1G", my_ch_rs, ItemsTree, "GROUPS", "ParentID")
    Else
        Call fill_my_Children("1G", my_ch_rs, ItemsTree, "GROUPS", "ParentID", , 2)
    End If

    ItemsTree.Refresh
    Exit Sub
ErrTrap:
End Sub

Private Sub Txt_KeyPress(Index As Integer, _
                         KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.Txt(Index).text, 0)
End Sub
