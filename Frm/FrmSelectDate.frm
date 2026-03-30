VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSelectDate 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " ÕœÌœ ŒÌ«—«   ﬁ«—Ì— «·⁄„·«¡"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5085
   Icon            =   "FrmSelectDate.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5355
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "ŒÌ«—«  «Œ—Ï ·· ﬁ—Ì—"
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
      Height          =   2265
      Index           =   3
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1830
      Width           =   4995
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈Œ Ì«— «·’‰›"
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
         Height          =   915
         Index           =   6
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1290
         Width           =   4875
         Begin VB.TextBox TxtItemCode 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1950
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   210
            Width           =   2025
         End
         Begin MSDataListLib.DataCombo DcboItemName 
            Height          =   315
            Left            =   90
            TabIndex        =   29
            Top             =   570
            Width           =   3885
            _ExtentX        =   6853
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«ﬂ » ﬂÊœ «·’‰› À„ ≈‰ —"
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
            Height          =   285
            Index           =   9
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   240
            Width           =   1845
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·’‰›"
            Height          =   315
            Index           =   8
            Left            =   3990
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   570
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·’‰›"
            Height          =   285
            Index           =   7
            Left            =   3990
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   270
            Width           =   825
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈Œ Ì«— «·„Œ“‰"
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
         Height          =   615
         Index           =   5
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   660
         Width           =   4875
         Begin VB.TextBox TxtStoreID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3390
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   210
            Width           =   525
         End
         Begin MSDataListLib.DataCombo DcboStores 
            Height          =   315
            Left            =   60
            TabIndex        =   25
            Top             =   210
            Width           =   3315
            _ExtentX        =   5847
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   315
            Index           =   6
            Left            =   3930
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   240
            Width           =   885
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
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
         Height          =   465
         Index           =   4
         Left            =   1650
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   180
         Width           =   3285
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰ﬁœÏ"
            Height          =   195
            Index           =   0
            Left            =   2250
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   210
            Width           =   795
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "√Ã·"
            Height          =   195
            Index           =   1
            Left            =   1530
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   210
            Width           =   675
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂ·"
            Height          =   195
            Index           =   2
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   210
            Value           =   -1  'True
            Width           =   795
         End
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "≈Œ — «”„ «·⁄„Ì· «Ê «·„Ê—œ"
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
      Height          =   765
      Index           =   2
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   30
      Width           =   5025
      Begin MSDataListLib.DataCombo DcboCusName 
         Height          =   315
         Left            =   570
         TabIndex        =   16
         Top             =   390
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   556
         _Version        =   393216
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin VB.ComboBox CboDealerType 
         Height          =   315
         Left            =   3720
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   390
         Width           =   1215
      End
      Begin ImpulseButton.ISButton CmdCusSearch 
         Height          =   345
         Left            =   60
         TabIndex        =   17
         Top             =   330
         Width           =   465
         _ExtentX        =   820
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
         ButtonImage     =   "FrmSelectDate.frx":038A
         DrawFocusRectangle=   0   'False
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «· ﬁ—Ì—"
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
      Height          =   945
      Index           =   1
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   840
      Width           =   5025
      Begin VB.ComboBox CboReportStyle 
         Height          =   315
         Left            =   60
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   570
         Width           =   3795
      End
      Begin VB.ComboBox CboReportType 
         Height          =   315
         Left            =   60
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   3795
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœœ «· ﬁ—Ì—"
         Height          =   315
         Index           =   5
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰„ÿ «· ﬁ—Ì—"
         Height          =   315
         Index           =   4
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   570
         Width           =   735
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ï «·› —…"
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
      Height          =   645
      Index           =   0
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   4140
      Width           =   4965
      Begin MSComCtl2.DTPicker DtpFrom 
         Height          =   345
         Left            =   2190
         TabIndex        =   3
         Top             =   240
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy/M/d"
         Format          =   100073475
         CurrentDate     =   39095.4614930556
      End
      Begin MSComCtl2.DTPicker DtpTo 
         Height          =   345
         Left            =   90
         TabIndex        =   4
         Top             =   210
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy/M/d"
         Format          =   100073475
         CurrentDate     =   39095.4616435185
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   2
         Left            =   3930
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   270
         Width           =   285
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   285
         Index           =   3
         Left            =   1770
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   270
         Width           =   345
      End
   End
   Begin ImpulseButton.ISButton XPBtnOK 
      Height          =   375
      Left            =   930
      TabIndex        =   6
      Top             =   4920
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
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
      ButtonImage     =   "FrmSelectDate.frx":0724
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
      CausesValidation=   0   'False
      Height          =   375
      Left            =   90
      TabIndex        =   7
      Top             =   4920
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
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
      ButtonImage     =   "FrmSelectDate.frx":0ABE
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
      Caption         =   "ﬁ„ » ÕœÌœ «·› —… «·“„‰Ì… ·⁄—÷ «· ﬁ—Ì—"
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
      Height          =   255
      Index           =   0
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   3930
      Width           =   195
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   375
      Index           =   1
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   4920
      Width           =   3105
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   5070
      Y1              =   4860
      Y2              =   4860
   End
End
Attribute VB_Name = "FrmSelectDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_UserCanceled As Boolean
Dim cSearch(2) As clsDCboSearch
Dim Dcombos As ClsDataCombos
Dim TTP As clstooltipdemand

Private Sub CboDealerType_Change()

    If CboDealerType.ListIndex = -1 Then
        Exit Sub
    End If

    If CboDealerType.ListIndex = 0 Then
        Dcombos.GetCustomersSuppliers 1, Me.DcboCusName, True
        Me.CboReportType.Enabled = True
    ElseIf CboDealerType.ListIndex = 1 Then
        Dcombos.GetCustomersSuppliers 2, Me.DcboCusName, True
        Me.CboReportType.Enabled = True
    ElseIf CboDealerType.ListIndex = 2 Then
        Dcombos.GetCustomersSuppliers 0, Me.DcboCusName, True
        Me.CboReportType.Enabled = True
    ElseIf CboDealerType.ListIndex = 3 Then
        Dcombos.GetPersons Me.DcboCusName
        Me.CboReportType.ListIndex = 0
        Me.CboReportType.Enabled = False
    End If

    cSearch(0).Refresh
End Sub

Private Sub CboDealerType_Click()
    CboDealerType_Change
End Sub

Private Sub CboReportType_Change()

    '.AddItem 0"ﬂ‘› Õ”«»"
    '.AddItem 1"------------------------------------"
    '.AddItem 2"›Ê« Ì— «·„»Ì⁄«  «·Œ«’… »«·⁄„Ì·(„Ê—œ)"
    '.AddItem 3"›Ê« Ì— „— Ã⁄ «·„»Ì⁄«  «·Œ«’… »«·⁄„Ì·(„Ê—œ)"
    '.AddItem 4"------------------------------------"
    '.AddItem 5"›Ê« Ì— «·„‘ —Ì«  «·Œ«’… »«·⁄„Ì·(„Ê—œ)"
    '.AddItem 6"›Ê« Ì— „Ê Ã⁄ «·„‘ —Ì«  «·Œ«’… »«·⁄„Ì·(„Ê—œ)"
    '.AddItem 7"------------------------------------"
    '.AddItem 8"«·ﬁÌ„ «·„«·Ì… «·√Ã·… ··⁄„Ì· «Ê «·„Ê—œ"
    '.AddItem 9"«·ﬁÌ„ «·„«·Ì… «·√Ã·… ⁄·Ï «·⁄„Ì· «Ê «·„Ê—œ"
    '.AddItem 10"------------------------------------"
    '.AddItem 11"«·„ﬁ»Ê÷«  «· Ï Õ’·  „‰ «·⁄„Ì· «Ê «·„Ê—œ"
    '.AddItem 12"«·„œ›Ê⁄«  «· Ï œ›⁄  ··⁄„Ì· «Ê «·„Ê—œ"
    '.AddItem 13"------------------------------------"
    '.AddItem 14"√ﬁ”«ÿ „” Õﬁ… ⁄·Ï «·⁄„Ì· «Ê «·„Ê—œ"
    '.AddItem 15"√ﬁ”«ÿ „” Õﬁ… ··⁄„Ì· «Ê «·„Ê—œ"
    '.AddItem 16"------------------------------------"
    '.AddItem 17"„»Ì⁄«  «·√’‰«› ≈·Ï «·⁄„Ì·(«·„Ê—œ)"
    '.AddItem 18"„‘ —Ì«  «·√’‰«› „‰ «·⁄„Ì·(«·„Ê—œ)"
    '.AddItem 19"„— Ã⁄ „»Ì⁄«  «·√’‰«› „‰ «·⁄„Ì·(«·„Ê—œ)"
    '.AddItem 20"„— Ã⁄ „‘ —Ì«  «·√’‰«› ≈·Ï «·⁄„Ì·(«·„Ê—œ)"
    
    If CboReportType.ListIndex = 0 Then
        'ﬂ‘› Õ”«» ⁄„Ì·
        Me.CboReportStyle.Enabled = True
        Me.lbl(4).Enabled = True
    
        Fra(3).Enabled = False
        Fra(4).Enabled = False
        Me.Opt(0).Enabled = False
        Me.Opt(1).Enabled = False
        Me.Opt(2).Enabled = False
    
        '«·Ã“¡ «·Œ«’ »»Ì«‰«  «·„Œ“‰
        Fra(5).Enabled = False
        Me.lbl(6).Enabled = False
        Me.TxtStoreID.Enabled = False
        Me.DcboStores.Enabled = False
        '«·Ã“¡ «·Œ«’ »Ì«‰«  «·’‰›
        Fra(6).Enabled = False
        Me.lbl(7).Enabled = False
        Me.lbl(8).Enabled = False
        Me.lbl(9).Enabled = False
        Me.TxtItemCode.Enabled = False
        Me.DcboItemName.Enabled = False
    
    ElseIf Me.CboReportType.ListIndex = 2 Or Me.CboReportType.ListIndex = 3 Or Me.CboReportType.ListIndex = 5 Or Me.CboReportType.ListIndex = 6 Then
    
        Me.CboReportStyle.Enabled = False
        Me.lbl(4).Enabled = False
    
        Fra(3).Enabled = True
        '«·Ã“¡ «·Œ«’ »ÿ—Ìﬁ… «·œ›⁄
        Fra(4).Enabled = True
        Me.Opt(0).Enabled = True
        Me.Opt(1).Enabled = True
        Me.Opt(2).Enabled = True
        
        '«·Ã“¡ «·Œ«’ »»Ì«‰«  «·„Œ“‰
        Fra(5).Enabled = True
        Me.lbl(6).Enabled = True
        Me.TxtStoreID.Enabled = True
        Me.DcboStores.Enabled = True
        '«·Ã“¡ «·Œ«’ »Ì«‰«  «·’‰›
        Fra(6).Enabled = False
        Me.lbl(7).Enabled = False
        Me.lbl(8).Enabled = False
        Me.lbl(9).Enabled = False
        Me.TxtItemCode.Enabled = False
        Me.DcboItemName.Enabled = False
        
    ElseIf Me.CboReportType.ListIndex = 17 Or Me.CboReportType.ListIndex = 18 Or Me.CboReportType.ListIndex = 19 Or Me.CboReportType.ListIndex = 20 Then
    
        Me.CboReportStyle.Enabled = False
        Me.lbl(4).Enabled = False
    
        Fra(3).Enabled = True
        '«·Ã“¡ «·Œ«’ »ÿ—Ìﬁ… «·œ›⁄
        Fra(4).Enabled = False
        Me.Opt(0).Enabled = False
        Me.Opt(1).Enabled = False
        Me.Opt(2).Enabled = False
        
        '«·Ã“¡ «·Œ«’ »»Ì«‰«  «·„Œ“‰
        Fra(5).Enabled = True
        Me.lbl(6).Enabled = True
        Me.TxtStoreID.Enabled = True
        Me.DcboStores.Enabled = True
        '«·Ã“¡ «·Œ«’ »Ì«‰«  «·’‰›
        Fra(6).Enabled = True
        Me.lbl(7).Enabled = True
        Me.lbl(8).Enabled = True
        Me.lbl(9).Enabled = True
        Me.TxtItemCode.Enabled = True
        Me.DcboItemName.Enabled = True
    
    Else
        'Disable ALL
        Me.CboReportStyle.Enabled = False
        Me.lbl(4).Enabled = False
        Fra(3).Enabled = True
        '«·Ã“¡ «·Œ«’ »ÿ—Ìﬁ… «·œ›⁄
        Fra(4).Enabled = False
        Me.Opt(0).Enabled = False
        Me.Opt(1).Enabled = False
        Me.Opt(2).Enabled = False
        
        '«·Ã“¡ «·Œ«’ »»Ì«‰«  «·„Œ“‰
        Fra(5).Enabled = False
        Me.lbl(6).Enabled = False
        Me.TxtStoreID.Enabled = False
        Me.DcboStores.Enabled = False
        '«·Ã“¡ «·Œ«’ »Ì«‰«  «·’‰›
        Fra(6).Enabled = False
        Me.lbl(7).Enabled = False
        Me.lbl(8).Enabled = False
        Me.lbl(9).Enabled = False
        Me.TxtItemCode.Enabled = False
        Me.DcboItemName.Enabled = False
    
    End If

End Sub

Private Sub CboReportType_Click()
    CboReportType_Change
End Sub

Private Sub CboReportType_Validate(Cancel As Boolean)
    Dim StrMSG As String

    If CboReportType.text Like "----*" Then
        Set TTP = New clstooltipdemand
        TTP.Style = TTBalloon
        TTP.Icon = TTIconError
        TTP.Centered = True
        TTP.RightToLeft = True
        TTP.CreateToolTip CboReportType.hWnd
        TTP.DelayTime = 250
        TTP.VisibleTime = 5000
        StrMSG = "Œÿ« ›Ï ≈Œ Ì«— «· ﬁ—Ì—...!!!"
        TTP.Title = StrMSG
        StrMSG = "ÌÃ» «‰  ﬁÊ„ »≈Œ ÌÌ«— «· ﬁ—Ì— «·„—«œ ⁄—÷Â"
        TTP.TipText = StrMSG
        TTP.PopupOnDemand = True
        TTP.Show (CboReportType.Width / Screen.TwipsPerPixelY), (CboReportType.Height / Screen.TwipsPerPixelX - 1)     '//In Pixel only
        Cancel = True
    End If

End Sub

Private Sub CmdCusSearch_Click()
    Load FrmCustemerSearch
    FrmCustemerSearch.SearchType = 1
    FrmCustemerSearch.RetrunType = 1
    Set FrmCustemerSearch.DcboCustomers = Me.DcboCusName
    FrmCustemerSearch.Show vbModal
End Sub

Private Sub DcboCusName_Change()
    Dim IntDealerType As Integer
    On Local Error GoTo ErrTrap

    If val(Me.DcboCusName.BoundText) = 0 Then
        Exit Sub
    End If

    IntDealerType = GetDealerType(Me.DcboCusName.BoundText)

    If IntDealerType = 1 Then
        Me.CboDealerType.ListIndex = 0
    ElseIf IntDealerType = 2 Then
        Me.CboDealerType.ListIndex = 1
    ElseIf IntDealerType = 3 Then
        Me.CboDealerType.ListIndex = 3
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub DcboCusName_Click(Area As Integer)
    DcboCusName_Change
End Sub

Private Sub Form_Activate()
    PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()

    Set Dcombos = New ClsDataCombos
    Dcombos.GetCustomersSuppliers 0, Me.DcboCusName, True
    Set cSearch(0) = New clsDCboSearch
    cSearch(0).AllowWriting = False
    Set cSearch(0).Client = Me.DcboCusName

    Dcombos.GetStores Me.DcboStores
    Set cSearch(1) = New clsDCboSearch
    cSearch(1).AllowWriting = False
    Set cSearch(1).Client = Me.DcboStores

    Dcombos.GetItemsNames Me.DcboItemName
    Set cSearch(2) = New clsDCboSearch
    cSearch(2).AllowWriting = False
    Set cSearch(2).Client = Me.DcboItemName

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    CenterForm Me

    FormPostion Me, GetPostion

    SetDtpickerDate Me.DTPFrom
    SetDtpickerDate Me.DTPTo

    With Me.CboDealerType
        .Clear
        .AddItem "⁄„Ì·"
        .AddItem "„Ê—œ"
        .AddItem "⁄„·«¡ Ê„Ê—œÌ‰"
        .AddItem "„ﬁ«Ê· »«ÿ‰"
    End With

    With Me.CboReportType
        .Clear
        .AddItem "ﬂ‘› Õ”«»"
        .AddItem "------------------------------------"
        .AddItem "›Ê« Ì— «·„»Ì⁄«  «·Œ«’… »«·⁄„Ì·(„Ê—œ)"
        .AddItem "›Ê« Ì— „— Ã⁄ «·„»Ì⁄«  «·Œ«’… »«·⁄„Ì·(„Ê—œ)"
        .AddItem "------------------------------------"
        .AddItem "›Ê« Ì— «·„‘ —Ì«  «·Œ«’… »«·⁄„Ì·(„Ê—œ)"
        .AddItem "›Ê« Ì— „Ê Ã⁄ «·„‘ —Ì«  «·Œ«’… »«·⁄„Ì·(„Ê—œ)"
        .AddItem "------------------------------------"
        .AddItem "«·ﬁÌ„ «·„«·Ì… «·√Ã·… ··⁄„Ì· «Ê «·„Ê—œ"
        .AddItem "«·ﬁÌ„ «·„«·Ì… «·√Ã·… ⁄·Ï «·⁄„Ì· «Ê «·„Ê—œ"
        .AddItem "------------------------------------"
        .AddItem "«·„ﬁ»Ê÷«  «· Ï Õ’·  „‰ «·⁄„Ì· «Ê «·„Ê—œ"
        .AddItem "«·„œ›Ê⁄«  «· Ï œ›⁄  ··⁄„Ì· «Ê «·„Ê—œ"
        .AddItem "------------------------------------"
        .AddItem "√ﬁ”«ÿ „” Õﬁ… ⁄·Ï «·⁄„Ì· «Ê «·„Ê—œ"
        .AddItem "√ﬁ”«ÿ „” Õﬁ… ··⁄„Ì· «Ê «·„Ê—œ"
        .AddItem "------------------------------------"
        .AddItem "„»Ì⁄«  «·√’‰«› ≈·Ï «·⁄„Ì·(«·„Ê—œ)"
        .AddItem "„‘ —Ì«  «·√’‰«› „‰ «·⁄„Ì·(«·„Ê—œ)"
        .AddItem "„— Ã⁄ „»Ì⁄«  «·√’‰«› „‰ «·⁄„Ì·(«·„Ê—œ)"
        .AddItem "„— Ã⁄ „‘ —Ì«  «·√’‰«› ≈·Ï «·⁄„Ì·(«·„Ê—œ)"
    End With

    With Me.CboReportStyle
        .Clear
        .AddItem "‰Ÿ«„ «·œ«∆‰ Ê«·„œÌ‰"
        .AddItem "‰Ÿ«„   «·Ï «·⁄„·Ì«  »«· ”·”·"
        .AddItem "‰Ÿ«„ «·œ«∆‰ Ê«·„œÌ‰(»«·≈÷«›… ≈·Ï ⁄—÷ «·√’‰«›)"
        .ListIndex = 0
    End With

    CboReportType_Change
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    If UnloadMode <> VBRUN.QueryUnloadConstants.vbFormCode Then
        Me.UserCanceled = True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer

    FormPostion Me, SavePostion
    For i = LBound(cSearch) To UBound(cSearch)
        Set cSearch(i) = Nothing
    Next i

End Sub

Public Property Get UserCanceled() As Boolean
    UserCanceled = m_UserCanceled
End Property

Public Property Let UserCanceled(ByVal vNewValue As Boolean)
    m_UserCanceled = vNewValue
End Property

Private Sub TxtItemCode_KeyDown(KeyCode As Integer, _
                                Shift As Integer)
    Dim LngTempID As Long

    If KeyCode = vbKeyReturn Then
        If Trim(Me.TxtItemCode.text) = "" Then Exit Sub
        LngTempID = GetItemID(Trim(Me.TxtItemCode.text))

        If LngTempID = 0 Then
            Me.DcboItemName.BoundText = ""
            Msg = "·«ÌÊÃœ ’‰› „”Ã· »Â–« «·ﬂÊœ..!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        ElseIf val(Me.DcboItemName.BoundText) <> LngTempID Then
            DcboItemName.BoundText = LngTempID
        End If
    End If

End Sub

Private Sub XPBtnCancel_Click()
    Me.UserCanceled = True
    Me.Hide
End Sub

Private Sub XPBtnOK_Click()
    Dim cReport As ClsCustemerReport
    Dim cItemsReport As ClsItemsReport
    Dim LngCusID As Long
    Dim LngItemID As Long
    Dim LngStoreID As Long
    Dim BolReturn As Boolean
    Dim Msg As String
    Dim StrSQL As String
    Dim StrDesReport As String
    Dim Reports As ClsRepoerts

    LngCusID = val(Me.DcboCusName.BoundText)
    LngItemID = val(Me.DcboItemName.BoundText)
    LngStoreID = val(Me.DcboStores.BoundText)

    PutFormOnTop Me.hWnd, False

    If Me.DcboCusName.BoundText = "" Then
        Msg = "ÌÃ» ≈Œ Ì«— «”„ «·⁄„Ì· ﬁ»· ⁄—÷ «· ﬁ—Ì—...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        PutFormOnTop Me.hWnd, True
        DcboCusName.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If Me.CboReportType.ListIndex = -1 Then
        Msg = "ÌÃ» ≈Œ Ì«— «· ﬁ—Ì— «·„—«œ ⁄—÷Â...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        PutFormOnTop Me.hWnd, True
        CboReportType.SetFocus
        SendKeys "{F4}"
        Exit Sub
    ElseIf Me.CboReportType.ListIndex = 0 Then

        If Me.CboReportStyle.ListIndex = -1 Then
            Msg = "ÌÃ» ≈Œ Ì«— ‰„ÿ «· ﬁ—Ì— «·„—«œ ⁄—÷Â...!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            PutFormOnTop Me.hWnd, True
            CboReportStyle.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
    End If

    PutFormOnTop Me.hWnd, False

    If Me.CboReportType.ListIndex = 0 Then
        'ﬂ‘› Õ”«» «·⁄„Ì·
        Set cReport = New ClsCustemerReport
        cReport.ClientAccount LngCusID, Me.DTPFrom.value, Me.DTPTo.value, WindowTarget, Me.CboReportStyle.ListIndex
    ElseIf Me.CboReportType.ListIndex = 2 Then
        ' ﬁ—Ì— „»Ì⁄«  «·⁄„Ì·
        Set cReport = New ClsCustemerReport
        cReport.ShowCusTransactions LngCusID, Trim(Me.DcboCusName.text), 2, Me.DTPFrom.value, Me.DTPTo.value
        Set cReport = Nothing
    ElseIf Me.CboReportType.ListIndex = 3 Then
        ' ﬁ—Ì— „— Ã⁄ „»Ì⁄«  «·⁄„Ì·
        Set cReport = New ClsCustemerReport
        cReport.ShowCusTransactions LngCusID, Trim(Me.DcboCusName.text), 9, Me.DTPFrom.value, Me.DTPTo.value
        Set cReport = Nothing
    ElseIf Me.CboReportType.ListIndex = 5 Then
        ' ﬁ—Ì— „‘ —Ì«  «·⁄„Ì·
        Set cReport = New ClsCustemerReport
        cReport.ShowCusTransactions LngCusID, Trim(Me.DcboCusName.text), 1, Me.DTPFrom.value, Me.DTPTo.value
        Set cReport = Nothing
    ElseIf Me.CboReportType.ListIndex = 6 Then
        ' ﬁ—Ì— „— Ã⁄  „‘ —Ì«  «·⁄„Ì·
        Set cReport = New ClsCustemerReport
        cReport.ShowCusTransactions LngCusID, Trim(Me.DcboCusName.text), 5, Me.DTPFrom.value, Me.DTPTo.value
        Set cReport = Nothing
    ElseIf Me.CboReportType.ListIndex = 11 Then
        ' ﬁ—Ì— »«·„ﬁ»Ê÷«  «· Ï Õ’·  „‰ «·⁄„Ì·
        StrSQL = "Select * From CahingReport "
        StrSQL = StrSQL + " Where  CahingReport.NOTEID <> 0"
        StrSQL = StrSQL + " AND CusID=" & LngCusID & ""

        If Not IsNull(Me.DTPFrom.value) Then
            StrSQL = StrSQL + " AND NoteDate >=" & SQLDate(DTPFrom.value, True) & ""
        End If

        If Not IsNull(Me.DTPTo.value) Then
            StrSQL = StrSQL + " AND NoteDate <=" & SQLDate(DTPTo.value, True) & ""
        End If

        StrSQL = StrSQL + " Order by NoteID"
        Set Reports = New ClsRepoerts
        StrDesReport = " ﬁ—Ì— »«·„ﬁ»Ê÷«  „‰ :" & Me.DcboCusName.text
        Reports.CashingReports StrSQL, WindowTarget, StrDesReport, True
    ElseIf Me.CboReportType.ListIndex = 12 Then
        ' ﬁ—Ì— »«·„œ›Ê⁄«  «· Ï Õ’·  „‰ «·⁄„Ì·
        StrSQL = "Select * From PaymentsReport "
        StrSQL = StrSQL + " Where  PaymentsReport.NOTEID <> 0"
        StrSQL = StrSQL + " AND CusID=" & LngCusID & ""

        If Not IsNull(Me.DTPFrom.value) Then
            StrSQL = StrSQL + " AND NoteDate >=" & SQLDate(DTPFrom.value, True) & ""
        End If

        If Not IsNull(Me.DTPTo.value) Then
            StrSQL = StrSQL + " AND NoteDate <=" & SQLDate(DTPTo.value, True) & ""
        End If

        StrSQL = StrSQL + " Order by NoteID"
        Set Reports = New ClsRepoerts
        StrDesReport = " ﬁ—Ì— »«·œ›Ê⁄«  ≈·Ì :" & Me.DcboCusName.text
        Reports.PaymentsReports StrSQL, WindowTarget, StrDesReport, True
    ElseIf Me.CboReportType.ListIndex = 14 Then
        '√ﬁ”«ÿ „” Õﬁ… ⁄·Ï «·⁄„Ì· «Ê «·„Ê—œ
    
    ElseIf Me.CboReportType.ListIndex = 15 Then
        '√ﬁ”«ÿ „” Õﬁ… ··⁄„Ì· «Ê «·„Ê—œ
    ElseIf Me.CboReportType.ListIndex = 17 Then
        Set cItemsReport = New ClsItemsReport
        BolReturn = cItemsReport.ShowItemTransCustomer(LngItemID, LngCusID, 2, LngStoreID, Me.DTPFrom.value, Me.DTPTo.value, WindowTarget)
    ElseIf Me.CboReportType.ListIndex = 18 Then
        Set cItemsReport = New ClsItemsReport
        BolReturn = cItemsReport.ShowItemTransCustomer(LngItemID, LngCusID, 1, LngStoreID, Me.DTPFrom.value, Me.DTPTo.value, WindowTarget)
    ElseIf Me.CboReportType.ListIndex = 19 Then
        Set cItemsReport = New ClsItemsReport
        BolReturn = cItemsReport.ShowItemTransCustomer(LngItemID, LngCusID, 9, LngStoreID, Me.DTPFrom.value, Me.DTPTo.value, WindowTarget)
    ElseIf Me.CboReportType.ListIndex = 20 Then
        Set cItemsReport = New ClsItemsReport
        BolReturn = cItemsReport.ShowItemTransCustomer(LngItemID, LngCusID, 5, LngStoreID, Me.DTPFrom.value, Me.DTPTo.value, WindowTarget)
    End If

    PutFormOnTop Me.hWnd, True
    Me.UserCanceled = False
    'Unload Me
End Sub

Private Sub ChangeLang()
    Me.Caption = "Select the date interval"
    'Fra.Caption = "Interval"
    lbl(2).Caption = "From"
    lbl(3).Caption = "To"
    Me.XPBtnOK.Caption = "Ok"
    Me.XPBtnCancel.Caption = "Cancel"
End Sub

Public Sub Retrive(LngCusID As Long)
    Dim IntType As Integer
    IntType = GetDealerType(LngCusID)

    If IntType = 1 Or IntType = 2 Then
        Me.CboDealerType.ListIndex = 2
    Else
        Me.CboDealerType.ListIndex = 3
    End If

    Me.DcboCusName.BoundText = LngCusID
End Sub
