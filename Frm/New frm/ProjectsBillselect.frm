VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ProjectsBillselect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«Œ Ì«— «·„” Œ·’«  «·„—«œ  ”œÌœÂ«  "
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14520
   Icon            =   "ProjectsBillselect.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5115
   ScaleWidth      =   14520
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   -15
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   30
      Width           =   14505
      Begin VB.TextBox TxtVac_ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   240
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   510
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Text            =   "modflag"
         Top             =   120
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Frame Frmo2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   540
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   450
         Visible         =   0   'False
         Width           =   3105
         Begin MSDataListLib.DataCombo DCUser 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   -255
            TabIndex        =   3
            Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
            Top             =   15
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483624
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "«·„” Œœ„"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   13
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   45
            Width           =   855
         End
      End
      Begin MSComctlLib.ImageList GrdImageList 
         Left            =   480
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ProjectsBillselect.frx":058A
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ProjectsBillselect.frx":0924
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ProjectsBillselect.frx":0CBE
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ProjectsBillselect.frx":1058
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ProjectsBillselect.frx":13F2
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ProjectsBillselect.frx":178C
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ProjectsBillselect.frx":1B26
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ProjectsBillselect.frx":20C0
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«Œ Ì«— «·„” Œ·’«  «·„—«œ  ”œÌœÂ«  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   2
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   120
         Width           =   4560
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4035
      Left            =   -1080
      TabIndex        =   0
      Top             =   570
      Width           =   15555
      _cx             =   27437
      _cy             =   7117
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"ProjectsBillselect.frx":245A
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
   Begin ImpulseButton.ISButton btnCancel 
      Height          =   330
      Left            =   0
      TabIndex        =   8
      Top             =   4680
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   582
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
      ButtonImage     =   "ProjectsBillselect.frx":2691
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
End
Attribute VB_Name = "ProjectsBillselect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub FillGridWithData(project_no As Integer)

'On Error GoTo ErrTrap
 
Dim I As Integer
Dim x As Integer
Dim rs As ADODB.Recordset

 
Dim ActualTotal As Double
Dim result As Double
Dim resultpercentage As Double
Dim sql As String
sql = "SELECT  * FROM     project_billl     where project_no = " & project_no
Set rs = New ADODB.Recordset
rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
 If rs.RecordCount = 0 Then
 
 Exit Sub
 End If
I = 0
With Me.Grid
    .Rows = 1
    .Clear flexClearScrollable
  
        rs.MoveFirst
        For x = 1 To rs.RecordCount
       
 ActualTotal = getBillPayedToproject(Val(rs.Fields("id").value))
 result = Val(rs.Fields("total").value) - ActualTotal
 resultpercentage = Round((ActualTotal / Val(rs.Fields("total").value)) * 100, 2)
 
If Val(rs.Fields("total").value) > ActualTotal Then
 I = I + 1
  .Rows = .Rows + 1
             .TextMatrix(I, .ColIndex("Ser")) = I
            .TextMatrix(I, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), _
            "", rs.Fields("id").value)
            
            .TextMatrix(I, .ColIndex("bill_date")) = IIf(IsNull(rs.Fields("bill_date").value), _
            "", rs.Fields("bill_date").value)
                                           .TextMatrix(I, .ColIndex("project_no")) = IIf(IsNull(rs.Fields("project_no").value), _
            "", rs.Fields("project_no").value)
                         .TextMatrix(I, .ColIndex("Project_name")) = IIf(IsNull(rs.Fields("project_name").value), _
            "", rs.Fields("project_name").value)
            
            .TextMatrix(I, .ColIndex("End_user_name")) = IIf(IsNull(rs.Fields("End_user_name").value), _
            "", rs.Fields("End_user_name").value)
            
            .TextMatrix(I, .ColIndex("Sub_user_name")) = IIf(IsNull(rs.Fields("Sub_user_name").value), _
            "", rs.Fields("Sub_user_name").value)
            
            .TextMatrix(I, .ColIndex("bill_to")) = IIf(IsNull(rs.Fields("bill_to").value), _
            "", rs.Fields("bill_to").value)
 
                                       .TextMatrix(I, .ColIndex("total")) = IIf(IsNull(rs.Fields("total").value), _
            "", rs.Fields("total").value)
            
             .TextMatrix(I, .ColIndex("ActualTotal")) = ActualTotal
                .TextMatrix(I, .ColIndex("ResultPercentage")) = resultpercentage
                .TextMatrix(I, .ColIndex("Result")) = result

     End If
             rs.MoveNext
        Next
       rs.Close
 
    .RowHeight(-1) = 300
End With
ErrTrap:
End Sub


Private Sub btnCancel_Click()
Me.Hide
End Sub

 

Private Sub Form_Load()
 
 Me.left = (MDIFrmMain.Width - Me.Width) / 2
    Me.top = (MDIFrmMain.Height - Me.Height) / 2 - 500
 
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)


Dim IntResult As String
Dim StrMSG As String
On Error GoTo ErrTrap
If Me.TxtModFlg.text <> "R" Then
Select Case Me.TxtModFlg.text
    Case "N"
    
        If SystemOptions.UserInterface = EnglishInterface Then
                 StrMSG = "You will close this screen before save " & Chr(13)
                StrMSG = StrMSG & " the new data  " & Chr(13)
                StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)

 
    
        Else
                StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
                StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
                StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
        
        End If
        
        
    Case "E"
            If SystemOptions.UserInterface = EnglishInterface Then
              StrMSG = "You will close this screen before save  " & Chr(13)
                StrMSG = StrMSG & " the Modifications  " & Chr(13)
                StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
 
    
        Else
                StrMSG = "·Ê  „ «€·«ﬁ Â–… «·‘«‘… ﬁ»· Õ›Ÿ «·ﬁ»Ê÷« " & Chr(13)
                StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
                StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
                StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
                
        End If
End Select
IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)
Select Case IntResult
    Case vbYes
        Cancel = True
       
         
    Case vbCancel
        Cancel = True
End Select
End If
Exit Sub
ErrTrap:

End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Grid
 If .Col <> 1 Then
 Cancel = True
 End If
End With
End Sub
