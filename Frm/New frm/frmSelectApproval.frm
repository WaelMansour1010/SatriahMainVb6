VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FrmSelectApproval 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ÕœÌœ ÃÂ… «·«⁄ „«œ"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9465
   ControlBox      =   0   'False
   Icon            =   "frmSelectApproval.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   9465
   Begin VB.CommandButton CMDCancel 
      Caption         =   "«·€«¡"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox TxtItem 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   -960
      Width           =   3165
   End
   Begin VB.TextBox TxtCodeItem 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3600
      TabIndex        =   4
      Top             =   -1440
      Width           =   2205
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "„Ê«›ﬁ"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   5280
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid1 
      Height          =   4275
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   9225
      _cx             =   16272
      _cy             =   7541
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
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSelectApproval.frx":000C
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
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   765
      Index           =   5
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   9435
      _cx             =   16642
      _cy             =   1349
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   24
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
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Picture         =   "frmSelectApproval.frx":0128
      Caption         =   " ÕœÌœ ÃÂ… «·«⁄ „«œ  "
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   0
      ChildSpacing    =   0
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   6
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   0
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
      Begin VB.CheckBox Check17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   " ÕœÌœ «·ﬂ·"
         Height          =   195
         Left            =   7920
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·’‰›"
      Height          =   345
      Index           =   0
      Left            =   1440
      TabIndex        =   7
      Top             =   -1200
      Width           =   930
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "ﬂÊœ «·’‰›"
      Height          =   345
      Index           =   32
      Left            =   4200
      TabIndex        =   6
      Top             =   -1680
      Width           =   930
   End
End
Attribute VB_Name = "FrmSelectApproval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private m_UserCanceled As Boolean
Public myfrmname As String
Public Transaction_ID As Double
Public NoteSerial1 As String



Public Property Get UserCanceled() As Boolean
    UserCanceled = m_UserCanceled
End Property

Public Property Let UserCanceled(ByVal vNewValue As Boolean)
    m_UserCanceled = vNewValue
End Property
Private Sub Check17_Click()
    Dim i As Integer

    If Check17.value = vbChecked Then

        With Me.Grid1
 
            For i = 1 To .Rows - 1
        
           .Cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked
            Next i

        End With

    Else

        With Me.Grid1

            For i = 1 To .Rows - 1
        
              .Cell(flexcpChecked, i, .ColIndex("Select")) = flexUnchecked
            Next i

        End With

    End If


End Sub

Function checkselected() As Boolean
 checkselected = False
  With Me.Grid1

        For i = 1 To .Rows - 1

            If val(.TextMatrix(i, .ColIndex("id"))) <> 0 And .Cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
          checkselected = True
          Exit Function
                         
                            
                     
                    
            End If
            
            '
        Next i
   
        
  End With
  
End Function
Sub save()
 
  With Me.Grid1

        For i = 1 To .Rows - 1

            If val(.TextMatrix(i, .ColIndex("id"))) <> 0 And .Cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
         
                       '
                           FillApprovedTableNew Me.myfrmname, Me.Transaction_ID, Me.NoteSerial1, val(.TextMatrix(i, .ColIndex("id")))
                             Me.UserCanceled = False
                             Unload Me
                         
                            
                     
                    
            End If
            
            '
        Next i
   
        
  End With
End Sub

Private Sub CmdCancel_Click()
    On Error GoTo ErrTrap
    Me.UserCanceled = True
    Me.Hide
    Exit Sub
ErrTrap:
End Sub

Private Sub CMDOK_Click()
If checkselected = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Õœœ «·«⁄ „«œ ", vbCritical
        Else
        MsgBox "Select Approve ", vbCritical
        End If


Exit Sub
End If

save
  Unload Me
End Sub
Sub Retrive(Optional frmname As String)
  Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Dim StrItemID As String
   StrItemID = FrmComparePrices.StrItemID
StrSQL = "SELECT    id,ApprovName, ApprovNamee"
StrSQL = StrSQL & " From dbo.TblApprovalDef"
StrSQL = StrSQL & " WHERE     (ScreenName = N'" & frmname & "')"
 
 
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or RsDev.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid1
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
  '
              .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(RsDev("id").value), 0, val(RsDev("id").value))
                 
                
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("ApprovName")) = IIf(IsNull(RsDev("ApprovName").value), "", RsDev("ApprovName").value)
                Else
                    .TextMatrix(i, .ColIndex("ApprovName")) = IIf(IsNull(RsDev("ApprovNamee").value), "", RsDev("ApprovNamee").value)
                End If
              
             
                
                .Cell(flexcpChecked, i, .ColIndex("Select")) = flexUnchecked
            
                RsDev.MoveNext
            Next i
    .AutoSize 0, .Cols - 1, False
        End With

    End If
End Sub
 


 
Private Sub Form_Load()
    Resize_Form Me
 
Retrive myfrmname
 
'If FrmComparePrices.TxtModFlg.text <> "N" Then
 
'End If

End Sub

