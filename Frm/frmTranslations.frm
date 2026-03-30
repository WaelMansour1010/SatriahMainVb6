VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form frmTranslations 
   Caption         =   "Form1"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   ScaleHeight     =   9165
   ScaleWidth      =   10305
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   9135
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   9495
      _Version        =   786432
      _ExtentX        =   16748
      _ExtentY        =   16113
      _StockProps     =   68
      PaintManager.Position=   2
      ItemCount       =   2
      Item(0).Caption =   "Item"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "G"
      Item(1).Caption =   "Item"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "VSFlexGrid1"
      Begin VSFlex8Ctl.VSFlexGrid G 
         Height          =   7005
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   9405
         _cx             =   16589
         _cy             =   12356
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Traditional Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   16711680
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   15329769
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   255
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   2
         GridLinesFixed  =   13
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   250
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmTranslations.frx":0000
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   -1  'True
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   0
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   7
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   -1  'True
         PictureType     =   0
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   0
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   3
         BackColorFrozen =   12648447
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
         Height          =   8325
         Left            =   -70000
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   9405
         _cx             =   16589
         _cy             =   14684
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Traditional Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   16711680
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   15329769
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   255
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   2
         GridLinesFixed  =   13
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   250
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmTranslations.frx":015F
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   -1  'True
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   0
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   7
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   -1  'True
         PictureType     =   0
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   -1  'True
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   0
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   3
         BackColorFrozen =   12648447
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox PicToolbar 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   9165
      Left            =   9750
      ScaleHeight     =   9165
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   0
      Width           =   555
      Begin XtremeSuiteControls.PushButton Buttons 
         Cancel          =   -1  'True
         Height          =   600
         Index           =   8
         Left            =   0
         TabIndex        =   1
         ToolTipText     =   "Exit"
         Top             =   6105
         Width           =   555
         _Version        =   786432
         _ExtentX        =   979
         _ExtentY        =   1058
         _StockProps     =   79
         Enabled         =   -1  'True
         UseVisualStyle  =   -1  'True
         Picture         =   "frmTranslations.frx":02BE
      End
      Begin XtremeSuiteControls.PushButton Buttons 
         Height          =   600
         Index           =   1
         Left            =   0
         TabIndex        =   2
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   555
         _Version        =   786432
         _ExtentX        =   979
         _ExtentY        =   1058
         _StockProps     =   79
         Enabled         =   -1  'True
         UseVisualStyle  =   -1  'True
         Picture         =   "frmTranslations.frx":0FD0
      End
   End
End
Attribute VB_Name = "frmTranslations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' **************************
'Security Generals
Public IsBusy   As Boolean
Public NewYes As Boolean
Public DisplayYes As Boolean
Public ModifyYes As Boolean
Public PrintYes As Boolean
Public PostYes As Boolean
Public DeleteYes As Boolean
' **************************
Dim mOldArabicValue As String
    Public Frm As Form

Private Sub SaveCurrentRecord()
    On Error GoTo eh
    
    Screen.MousePointer = vbHourglass
    
   
  
    
    For i = 1 To G.rows - 1
        mArabic = Trim(G.TextMatrix(i, G.ColIndex("Arabic")))
        mEnglish = Trim(G.TextMatrix(i, G.ColIndex("English")))
        mControlName = Trim(G.TextMatrix(i, G.ColIndex("ControlName")))
        mIDD = Trim(G.TextMatrix(i, G.ColIndex("ID")))
        mControlIndex = Trim(G.TextMatrix(i, G.ColIndex("ControlIndex")))
        mFormName = Trim(G.TextMatrix(i, G.ColIndex("FormName")))
        mOldArabic = Trim(G.TextMatrix(i, G.ColIndex("OldArabic")))
        mIsVisible = CBool(G.ValueMatrix(i, G.ColIndex("IsVisible")))
        If Trim(G.TextMatrix(i, G.ColIndex("ID"))) = 572 Then
           G.TextMatrix(i, G.ColIndex("ID")) = 572
        End If
        '------------
        If mIDD <> "" Then
            s = "Update Translations Set English=N'" & mEnglish & "' ,"
            s = s & "ControlName=N'" & mControlName & "', "
            s = s & "ControlIndex=N'" & mControlIndex & "', "
            s = s & "FormName=N'" & mFormName & "', "
            s = s & "OldArabic=N'" & mOldArabic & "', "
            s = s & "IsVisible =" & IIf(mIsVisible, 1, 0)
            s = s & " Where Id=" & mIDD
    
        '-------------
        Cn.Execute s
        End If
    Next
    Screen.MousePointer = vbDefault
    '---------------
    MsgBox "Data Saved..."
    Exit Sub
eh:
    Screen.MousePointer = vbDefault
    MsgBox MyErrorHandler(Err)
End Sub

Private Sub Buttons_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
        If IsObject(rs) And Not IsEmpty(newrecord) Then
        ButtonsMouseUp Me, index, Button, Shift, X, Y, Me, rs, newrecord
    Else
        ButtonsMouseUp Me, index, Button, Shift, X, Y
    End If
End Sub

Private Sub Buttons_Click(index As Integer)
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    'check Control Access
'    If Not UserFormRights(Me, Index, NewRecord) Then
'        Exit Sub
'    End If
    '^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
    Select Case index
    Case 1
        SaveCurrentRecord
    Case 8
        Unload Me
    End Select
End Sub

Private Sub G_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)

    If Col = G.ColIndex("Arabic") Then
        G.EditSelLength = 0
    Else
        G.EditMaxLength = 50
    End If
    
    
End Sub


Private Sub G_StartEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
    '''''''''''''''''''''''''''''
    'KeyboardLang MyCurrentLanguage
    '''''''''''''''''''''''''''''
End Sub

Private Sub G_KeyDown(KeyCode As Integer, Shift As Integer)
    GridKeyDown G, KeyCode, Shift, True, True
End Sub

Private Sub G_KeyDownEdit(ByVal row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    G_KeyDown KeyCode, Shift
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    rs.Requery
End Sub

Private Sub Form_Load()
    CenterForm Me
    '------------
    G.rows = 1
    If IsObject(Frm) Then
        Screen.MousePointer = vbHourglass
     
            RefreshControls Frm
    
        Screen.MousePointer = vbDefault
    End If
End Sub
Function GetEnglishTranslation(ByVal arabicName As String, ByVal PEnglishName As String) As String
    Dim rs As Object
    Dim sql As String
    Dim EnglishName As String
    If PEnglishName <> "" Then GetEnglishTranslation = PEnglishName: Exit Function
    '  ⁄—Ì› ﬂ«∆‰ Recordset
    Set rs = CreateObject("ADODB.Recordset")
    
    ' ﬂ «»… Ã„·… SQL ··»ÕÀ ⁄‰ «·«”„ «·≈‰Ã·Ì“Ì »‰«¡ ⁄·Ï «·«”„ «·⁄—»Ì
    sql = "SELECT English FROM Translations WHERE Arabic Like '%" & Trim(arabicName) & "%' and IsNull(English,'') <> ''"
    
    ' › Õ «·« ’«· Ê ‰›Ì– «·«” ⁄·«„
    rs.Open sql, Cn, 1, 1 ' 1 = adOpenKeyset, 1 = adLockReadOnly
    
    ' «· Õﬁﬁ „‰ √‰ «·«” ⁄·«„ √⁄«œ ‰ ÌÃ…
    If Not rs.EOF Then
        EnglishName = Trim(rs.Fields("English").value & "")
    Else
        EnglishName = ""
    End If
    
    ' ≈€·«ﬁ Recordset
    rs.Close
    Set rs = Nothing
    
    ' ≈—Ã«⁄ «·«”„ «·≈‰Ã·Ì“Ì √Ê —”«·…  Ê÷Õ ⁄œ„ «·⁄ÀÊ— ⁄·ÌÂ
    GetEnglishTranslation = EnglishName
End Function

Private Sub RefreshControls(Frm As Object)
    On Error Resume Next
    '------------------------
    Dim Translations As ADODB.Recordset
    Dim Ctr As Control
    Dim mText As String
    Dim mR As Long
    '------------------------
    mText = Trim(Frm.Caption)
     If SystemOptions.UserInterface = ArabicInterface Then
        mText = left(IIf(left(Trim(mText), 40) = "‘—ﬂ… œ«Ì‰«„ﬂ: ", mId(Trim(mText), 41), Trim(mText)), 50)
    Else
        mText = left(IIf(left(Trim(mText), 22) = "Dynamic Byte: ", mId(Trim(mText), 23), Trim(mText)), 50)
    End If
    
    If SystemOptions.UserInterface = ArabicInterface Then ' «·«÷«›… ›ﬁÿ ›Ï Õ«·… «·⁄—»Ï
        Cond = "Arabic=N'" & Trim(mText) & "'"
    Else
        Cond = "English=N'" & Trim(mText) & "'"
    End If
    s = "Select * from Translations where " & Cond
    Set Translations = New ADODB.Recordset
    Translations.Open s, Cn, adOpenKeyset, adLockOptimistic
    Dim rsD As New ADODB.Recordset
    rsD.Open s, Cn, adOpenKeyset, adLockOptimistic
    '***************************
    If Trim(mText) <> "" Then
        G.rows = G.rows + 1
        r = G.rows - 1
        '------------------------
        If rsD.EOF Then
        
            If SystemOptions.UserInterface = ArabicInterface Then ' «·«÷«›… ›ﬁÿ ›Ï Õ«·… «·⁄—»Ï
                rsD.AddNew
                rsD!Arabic = Trim(mText)
            '    rsD!ID = r
'UpdateRecordSet Translations
                rsD.update
            End If
        Else
            If SystemOptions.UserInterface = ArabicInterface Then
                G.TextMatrix(r, G.ColIndex("Arabic")) = Trim(mText)
                G.TextMatrix(r, G.ColIndex("English")) = GetEnglishTranslation(Trim(mText), Trim(Translations!English & ""))
            Else
                G.TextMatrix(r, G.ColIndex("Arabic")) = Trim(Translations!Arabic & "")
                G.TextMatrix(r, G.ColIndex("English")) = GetEnglishTranslation(Trim(Translations!English & ""), Trim(mText)) '       Trim(mText)
            End If
        End If
    End If
    '------------------------
    For Each Ctr In Frm.Controls
        If Ctr.Name = "Fg_Journal" Then
            mText = mText
        End If
        If (TypeOf Ctr Is Label) _
            Or (TypeOf Ctr Is CheckBox) _
            Or (TypeOf Ctr Is OptionButton) _
            Or (TypeOf Ctr Is XtremeSuiteControls.CheckBox) _
            Or (TypeOf Ctr Is RadioButton) _
            Or (TypeOf Ctr Is frame) _
            Or (TypeOf Ctr Is ISButton) _
                Or (TypeOf Ctr Is C1Elastic) _
            Or (TypeOf Ctr Is GroupBox) _
            Or (TypeOf Ctr Is XtremeSuiteControls.PushButton) _
            Or (TypeOf Ctr Is CommandButton) Then
            '********************************
            Translations.Close
            mText = left(Trim(Ctr.Caption), 300)
'            If SystemOptions.UserInterface = ArabicInterface Then
'                Cond = "Arabic=N'" & Trim(mText) & "'"
'            Else
'                Cond = "English=N'" & Trim(mText) & "'"
'            End If
        If Trim(mText) <> "" Then
Cond = ""
            Cond = Cond & "  FormName = N'" & Trim(Frm.Name) & "'"
             Cond = Cond & " and ControlName = N'" & Trim(Ctr.Name) & "'"
              Cond = Cond & " and ControlIndex = N'" & Trim(Ctr.index) & "'"
            
            s = "Select * from Translations where " & Cond
             rsD.Open s, Cn, adOpenKeyset, adLockOptimistic
             Set Translations = New ADODB.Recordset
            Translations.Open s, Cn, adOpenKeyset, adLockOptimistic
            '***************************
            
              '  x = G.FindRow(mText, 0, IIf(ArabicInterface, G.ColIndex("Arabic"), G.ColIndex("English")))
              X = -1
                If X = -1 Then ' €Ì— „ÊÃÊœ
                    G.rows = G.rows + 1
                    r = G.rows - 1
                    '------------------------
                    If Translations.EOF Then
                        If SystemOptions.UserInterface = ArabicInterface Then ' «·«÷«›… ›ﬁÿ ›Ï Õ«·… «·⁄—»Ï
                            Translations.AddNew
                            Translations!Arabic = Trim(mText)
                            Translations!controlname = Trim(Ctr.Name)
                            Translations!OldArabic = Trim(mText)
                            Translations!formname = Trim(Frm.Name)
                            Translations!Arabic = Trim(mText)
                           ' Translations!ID = r
                            Translations.update
                            G.TextMatrix(r, G.ColIndex("Arabic")) = Trim(mText)
                            
                            G.TextMatrix(r, G.ColIndex("FormName")) = Frm.Name
                            
                            G.TextMatrix(r, G.ColIndex("ControlName")) = Ctr.Name
                            G.TextMatrix(r, G.ColIndex("OldArabic")) = Trim(mText)
                            G.TextMatrix(r, G.ColIndex("ID")) = Trim(Translations!ID & "")
                            
                            
                            If FindIndex(Frm, Ctr) >= 0 Then
                                G.TextMatrix(r, G.ColIndex("ControlIndex")) = Ctr.index
                            End If

'UpdateRecordSet Translations
                        End If
                    Else

                         G.TextMatrix(r, G.ColIndex("Arabic")) = Trim(Translations!Arabic & "")
                         G.TextMatrix(r, G.ColIndex("English")) = GetEnglishTranslation(Trim(Translations!Arabic & ""), Trim(Translations!English & ""))
                    
                        G.TextMatrix(r, G.ColIndex("FormName")) = Trim(Translations!formname & "")
                        
                        G.TextMatrix(r, G.ColIndex("ControlName")) = Trim(Translations!controlname & "")
                        G.TextMatrix(r, G.ColIndex("OldArabic")) = Trim(mText)
                        G.TextMatrix(r, G.ColIndex("ID")) = Trim(Translations!ID & "")
                        
                        G.TextMatrix(r, G.ColIndex("ControlIndex")) = Trim(Translations!controlindex & "")
                                                  If Trim(G.TextMatrix(r, G.ColIndex("ID"))) = 572 Then
           G.TextMatrix(r, G.ColIndex("ID")) = 572
        End If
                        G.TextMatrix(r, G.ColIndex("IsVisible")) = (Translations!IsVisible & "")
                        
                    End If
                End If
            End If
            '***************
        ElseIf TypeOf Ctr Is C1Tab Then
        ' ****************************
            For j = 0 To Ctr.Tabs - 1
                mText = left(Trim(Ctr.TabCaption(j)), 300)
                If SystemOptions.UserInterface = ArabicInterface Then
                    Cond = "Arabic=N'" & Trim(mText) & "'"
                Else
                    Cond = "English=N'" & Trim(mText) & "'"
                End If
                s = "Select * from Translations where " & Cond
                Set Translations = New ADODB.Recordset
                Translations.Open s, Cn, adOpenKeyset, adLockOptimistic
                
                '***************************
                If Trim(mText) <> "" Then
                   ' x = G.FindRow(mText, 0, IIf(ArabicInterface, G.ColIndex("Arabic"), G.ColIndex("English")))
                   X = -1
                    If X = -1 Then ' €Ì— „ÊÃÊœ
                        G.rows = G.rows + 1
                        r = G.rows - 1
                        '------------------------
                        If Translations.EOF Then
                            If SystemOptions.UserInterface = ArabicInterface Then ' «·«÷«›… ›ﬁÿ ›Ï Õ«·… «·⁄—»Ï
                                Translations.AddNew
                                Translations!Arabic = Trim(mText)
                                Translations!controlname = Trim(Ctr.Name)
                                Translations!OldArabic = Trim(mText)
                                Translations!formname = Trim(Frm.Name)
                                Translations!Arabic = Trim(mText)
                                Translations.update
'UpdateRecordSet Translations
                            End If
                        Else
                            If SystemOptions.UserInterface = ArabicInterface Then
                                G.TextMatrix(r, G.ColIndex("Arabic")) = Trim(mText)
                                G.TextMatrix(r, G.ColIndex("English")) = GetEnglishTranslation(Trim(mText), Trim(Translations!English & ""))
                            Else
                                G.TextMatrix(r, G.ColIndex("Arabic")) = Trim(Translations!Arabic & "")
                                G.TextMatrix(r, G.ColIndex("English")) = Trim(mText)
                            End If
                        End If
                    End If
                End If
            Next
        ElseIf TypeOf Ctr Is VSFlexGrid Or typename(Ctr) = VSFlexGrid1 Or StartsWithKeywords(Ctr.Name) Then
        ' ****************************
'            For r = 0 To Ctr.FixedRows - 1
'                For j = 0 To Ctr.Cols - 1
'                    mText = Left(Trim(Ctr.TextMatrix(r, j)), 300)
'                    If SystemOptions.UserInterface = ArabicInterface Then
'                        cond = "Arabic=N'" & Trim(mText) & "'"
'                    Else
'                        cond = "English=N'" & Trim(mText) & "'"
'                    End If
'                    s = "Select * from Translations where " & cond
'                    Translations.Open s, Cn, adOpenKeyset, adLockOptimistic
'                    '***************************
'                    If Trim(mText) <> "" Then
'                        x = G.FindRow(mText, 0, IIf(ArabicInterface, G.ColIndex("Arabic"), G.ColIndex("English")))
'                        If x = -1 Then ' €Ì— „ÊÃÊœ
'                            G.Rows = G.Rows + 1
'                            mRow = G.Rows - 1
'                            '------------------------
'                            If Translations.EOF Then
'                                If SystemOptions.UserInterface = ArabicInterface Then ' «·«÷«›… ›ﬁÿ ›Ï Õ«·… «·⁄—»Ï
'                                    Translations.AddNew
'                                    Translations!arabic = Trim(mText)
'UpdateRecordSet Translations
'                                End If
'                            Else
'                                If SystemOptions.UserInterface = ArabicInterface Then
'                                    G.TextMatrix(mRow, G.ColIndex("Arabic")) = Trim(mText)
'                                    G.TextMatrix(mRow, G.ColIndex("English")) = Trim(Translations!English & "")
'                                Else
'                                    G.TextMatrix(mRow, G.ColIndex("Arabic")) = Trim(Translations!arabic & "")
'                                    G.TextMatrix(mRow, G.ColIndex("English")) = Trim(mText)
'                                End If
'                            End If
'                        End If
'                    End If
'                Next
            'Next
           '******************************
           For mR = 0 To Ctr.rows - 1
                For j = 0 To Ctr.Cols - 1
                    If mR > Ctr.FixedRows And j > 0 Then
                        GoTo NextCol
                    End If
                    
                    mText = left(Trim(Ctr.TextMatrix(mR, j)), 300)
                    If Trim(mText) = "" Then GoTo NextCol
                    If IsNumeric(mText) Then GoTo NextCol '  „‘ „Õ «ÃÌ‰ ‰ —Ã„ «·«—ﬁ«„
                    
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Cond = "Arabic=N'" & Trim(mText) & "'"
                    Else
                        Cond = "English=N'" & Trim(mText) & "'"
                    End If
                    
                    s = "Select * from Translations where " & Cond
                    Set Translations = New ADODB.Recordset
                    Translations.Open s, Cn, adOpenKeyset, adLockOptimistic
                    '***************************
'                    If Trim(mText) <> "" Then
'                        X = G.FindRow(mText, 0, IIf(ArabicInterface, G.ColIndex("Arabic"), G.ColIndex("English")))
'                        If X = -1 Then ' €Ì— „ÊÃÊœ
'                            G.rows = G.rows + 1
'                            mRow = G.rows - 1
'                            '------------------------
'                            If Translations.EOF Then
'                                If SystemOptions.UserInterface = ArabicInterface Then ' «·«÷«›… ›ﬁÿ ›Ï Õ«·… «·⁄—»Ï
'                                    Translations.AddNew
'                                    Translations!Arabic = Trim(mText)
'                                    Translations.update
''UpdateRecordSet Translations
'                                End If
'                            Else
'                                If SystemOptions.UserInterface = ArabicInterface Then
'                                    G.TextMatrix(mRow, G.ColIndex("Arabic")) = Trim(mText)
'                                    G.TextMatrix(mRow, G.ColIndex("English")) = GetEnglishTranslation(Trim(mText), Trim(Translations!English & ""))
'                                Else
'                                    G.TextMatrix(mRow, G.ColIndex("Arabic")) = Trim(Translations!Arabic & "")
'                                    G.TextMatrix(mRow, G.ColIndex("English")) = Trim(mText)
'                                End If
'                            End If
'                        End If
'                    End If
'NextCol:
'                Next
'            Next

        If Trim(mText) <> "" Then
Cond = ""
            Cond = Cond & "  FormName = N'" & Trim(Frm.Name) & "'"
             Cond = Cond & " and ControlName = N'" & Trim(Ctr.Name) & "'"
              Cond = Cond & " and ControlIndex = N'" & Trim(Ctr.index) & "'"
               Cond = Cond & " and Arabic = N'" & Trim(mText) & "'"
            
            s = "Select * from Translations where " & Cond
             rsD.Open s, Cn, adOpenKeyset, adLockOptimistic
             Set Translations = New ADODB.Recordset
            Translations.Open s, Cn, adOpenKeyset, adLockOptimistic
            '***************************
            
              '  x = G.FindRow(mText, 0, IIf(ArabicInterface, G.ColIndex("Arabic"), G.ColIndex("English")))
              X = -1
                If X = -1 Then ' €Ì— „ÊÃÊœ
                    G.rows = G.rows + 1
                    r = G.rows - 1
                    '------------------------
                    If Translations.EOF Then
                        If SystemOptions.UserInterface = ArabicInterface Then ' «·«÷«›… ›ﬁÿ ›Ï Õ«·… «·⁄—»Ï
                            Translations.AddNew
                            Translations!Arabic = Trim(mText)
                            
                            Translations!controlname = Trim(Ctr.Name)
                            Translations!OldArabic = Trim(mText)
                            Translations!formname = Trim(Frm.Name)
                            Translations!Arabic = Trim(mText)
                            
                           ' Translations!ID = r
                            Translations.update
                            G.TextMatrix(r, G.ColIndex("Arabic")) = Trim(mText)
                            
                            G.TextMatrix(r, G.ColIndex("FormName")) = Frm.Name
                            
                            G.TextMatrix(r, G.ColIndex("ControlName")) = Ctr.Name
                            G.TextMatrix(r, G.ColIndex("OldArabic")) = Trim(mText)
                            G.TextMatrix(r, G.ColIndex("ID")) = Trim(Translations!ID & "")
                            
                            
                            If FindIndex(Frm, Ctr) >= 0 Then
                                G.TextMatrix(r, G.ColIndex("ControlIndex")) = Ctr.index
                            End If

'UpdateRecordSet Translations
                        End If
                    Else

                         G.TextMatrix(r, G.ColIndex("Arabic")) = Trim(Translations!Arabic & "")
                         G.TextMatrix(r, G.ColIndex("English")) = GetEnglishTranslation(Trim(Translations!Arabic & ""), Trim(Translations!English & ""))
                    
                        G.TextMatrix(r, G.ColIndex("FormName")) = Trim(Translations!formname & "")
                        
                        G.TextMatrix(r, G.ColIndex("ControlName")) = Trim(Translations!controlname & "")
                        G.TextMatrix(r, G.ColIndex("OldArabic")) = Trim(mText)
                        G.TextMatrix(r, G.ColIndex("ID")) = Trim(Translations!ID & "")
                        
                        G.TextMatrix(r, G.ColIndex("ControlIndex")) = Trim(Translations!controlindex & "")
                                                  If Trim(G.TextMatrix(r, G.ColIndex("ID"))) = 572 Then
           G.TextMatrix(r, G.ColIndex("ID")) = 572
        End If
                        G.TextMatrix(r, G.ColIndex("IsVisible")) = (Translations!IsVisible & "")
                        
                    End If
                End If
            End If
            
NextCol:
                Next
            Next
        '******************************
        
        '*************************************
        ElseIf TypeOf Ctr Is MSChart Then  ' SaMi 1/11/2009
            For i = 1 To 4
                Ctr.Column = i
                mText1 = left(Trim(Ctr.ColumnLabel), 300)
                If SystemOptions.UserInterface = ArabicInterface Then
                    Cond = "Arabic=N'" & Trim(mText1) & "'"
                Else
                    Cond = "English=N'" & Trim(mText1) & "'"
                End If
                s = "Select * from Translations where " & Cond
                Set Translations = New ADODB.Recordset
                Translations.Open s, Cn, adOpenKeyset, adLockOptimistic
                '***************************
                If Trim(mText1) <> "" Then
                    X = G.FindRow(mText1, 0, IIf(ArabicInterface, G.ColIndex("Arabic"), G.ColIndex("English")))
                    If X = -1 Then ' €Ì— „ÊÃÊœ
                        G.rows = G.rows + 1
                        r = G.rows - 1
                        '------------------------
                        If Translations.EOF Then
                            If SystemOptions.UserInterface = ArabicInterface Then ' «·«÷«›… ›ﬁÿ ›Ï Õ«·… «·⁄—»Ï
                                Translations.AddNew
                                Translations!Arabic = Trim(mText1)
UpdateRecordSet Translations
                            End If
                        Else
                            If SystemOptions.UserInterface = ArabicInterface Then
                                G.TextMatrix(r, G.ColIndex("Arabic")) = Trim(mText1)
                                G.TextMatrix(r, G.ColIndex("English")) = GetEnglishTranslation(Trim(mText1), Trim(Translations!English & ""))
                            Else
                                G.TextMatrix(r, G.ColIndex("Arabic")) = Trim(Translations!Arabic & "")
                                G.TextMatrix(r, G.ColIndex("English")) = Trim(mText1)
                            End If
                        End If
                    End If
                End If
            Next
            '****************************************************
            mText2 = left(Trim(Ctr.RowLabel), 300)
            
            If SystemOptions.UserInterface = ArabicInterface Then
                Cond = "Arabic=N'" & Trim(mText2) & "'"
            Else
                Cond = "English=N'" & Trim(mText2) & "'"
            End If
            
            s = "Select * from Translations where " & Cond
            Set Translations = New ADODB.Recordset
            Translations.Open s, Cn, adOpenKeyset, adLockOptimistic
            '***************************
            If Trim(mText2) <> "" Then
                X = G.FindRow(mText2, 0, IIf(ArabicInterface, G.ColIndex("Arabic"), G.ColIndex("English")))
                If X = -1 Then ' €Ì— „ÊÃÊœ
                    G.rows = G.rows + 1
                    r = G.rows - 1
                    '------------------------
                    If Translations.EOF Then
                        If SystemOptions.UserInterface = ArabicInterface Then ' «·«÷«›… ›ﬁÿ ›Ï Õ«·… «·⁄—»Ï
                            Translations.AddNew
                            Translations!Arabic = Trim(mText2)
UpdateRecordSet Translations
                        End If
                    Else
                        If SystemOptions.UserInterface = ArabicInterface Then
                            G.TextMatrix(r, G.ColIndex("Arabic")) = Trim(mText2)
                            G.TextMatrix(r, G.ColIndex("English")) = GetEnglishTranslation(Trim(mText2), Trim(Translations!English & ""))
                        Else
                            G.TextMatrix(r, G.ColIndex("Arabic")) = Trim(Translations!Arabic & "")
                            G.TextMatrix(r, G.ColIndex("English")) = Trim(mText2)
                        End If
                    End If
                End If
            End If
        ElseIf typename(Ctr) = "CtlCostCenters" Or typename(Ctr) = "CtlBalancesPeriods" Then
        '--------------------------------
            RefreshControls Ctr
        End If
        ' ************************************
NextControl:
    Next
    '***********************
    GridSerial G
    Exit Sub
eh:
    If Err.Number <> 438 And Err.Number <> 343 Then ' Object doesn't support this property or method ---- Object not an array
        'MsgBox err.Number & " : " & err.Description
        Debug.Print Err.Number, Err.Description
    End If
    MsgBox Err.Number & " : " & Err.Description
    'Debug.Print err.Number, err.Description
    Resume Next
End Sub

Private Function FindIndex(ByRef F As Form, ByRef ctl As Control) As Integer
    Dim ctlTest As Control
    For Each ctlTest In F.Controls
        If (ctlTest.Name = ctl.Name) And (Not (ctlTest Is ctl)) Then
            'if the object is the same name but is not the same object we can assume it is a control array
            FindIndex = ctl.index
            Exit Function
        End If
    Next
    'if we get here then no controls on the form have the same name so can't be a control array
    FindIndex = -99
End Function


Private Sub RefreshControlsRPT()
    On Error GoTo eh
    '«⁄«œÂ »—„ÃÂ
  
    '------------------------
    '    Dim oSubreportObject As CRAXDRT.SubreportObject
    '    Dim myObject As Object
    '    Dim oSection As CRAXDRT.Section

    '    For Each oSection In Rpt.Sections
    '        For Each myObject In oSection.ReportObjects
    '            Select Case myObject.Kind
    '            Case crSubreportObject
    '                '---------------------
    '                Set oSubreportObject = myObject
    '                '---------------------
    '                For Each oSection1 In oSubreportObject.OpenSubreport.Sections
    '                    For Each myObject1 In oSection1.ReportObjects
    '                        Select Case myObject1.Kind
    '                        Case crTextObject
    '                            mText = Trim(myObject1.Text)
    '                            GetCaptionTranslate mText
    '                        End Select
    '                    Next myObject1
    '                Next oSection1
    '
    '            Case crTextObject
    '                '---------------------
    '                mText = Trim(myObject.Text)
    '                '--------------------------
    '                GetCaptionTranslate mText
    '                '------------------------
    '            Case crFieldObject
    '            '---------------------
    '            Case crLineObject
    '            '---------------------
    '            Case crBoxObject, crOLEObject
    '            '---------------------
    '            Case crGraphObject
    '            Case crBlobFieldObject
    '            Case crCrossTabObject
    '            Case crMapObject
    '            Case crOlapGridObject
    '            Case crOLEObject
    '            End Select
    '        Next myObject
    '    Next oSection
    ' *****************
    ' Dim mText As String
'    Dim myRpt As TopReport
'    Set myRpt = frm.CurrentRpt
'    For Each mText In myRpt.GetAllObjectText
'        GetCaptionTranslate mText
'    Next mText
'    GridSerial G
    Exit Sub
eh:
    MsgBox MyErrorHandler(Err)
End Sub
Private Sub GetCaptionTranslate(ByVal Txt As String)
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Cond = "Arabic=N'" & Trim(Txt) & "'"
    Else
        Cond = "English=N'" & Trim(Txt) & "'"
    End If
    s = "Select * from Translations2 where " & Cond
    Set Translations2 = OpenRecordSet(s, adOpenKeyset, adLockOptimistic)
    '***************************
    If Trim(Txt) <> "" Then
        G.rows = G.rows + 1
        r = G.rows - 1
        '------------------------
        If Translations2.EOF Then
            If SystemOptions.UserInterface = ArabicInterface Then ' «·«÷«›… ›ﬁÿ ›Ï Õ«·… «·⁄—»Ï
                Translations2.AddNew
                Translations2!Arabic = Trim(Txt)
UpdateRecordSet Translations2
            End If
        Else
            If SystemOptions.UserInterface = ArabicInterface Then
                G.TextMatrix(r, G.ColIndex("Arabic")) = Trim(Txt)
                G.TextMatrix(r, G.ColIndex("English")) = Trim(Translations2!English & "")
            Else
                G.TextMatrix(r, G.ColIndex("Arabic")) = Trim(Translations2!Arabic & "")
                G.TextMatrix(r, G.ColIndex("English")) = Trim(Txt)
            End If
        End If
    End If
End Sub




Public Function UpdateRecordSet(ByRef CurrentRS As Variant, _
                                Optional IsAudit As Boolean = False) As Boolean
    CurrentRS.update
    UpdateRecordSet = True
    On Error GoTo ff
    CurrentRS.AbsolutePosition = CurrentRS.AbsolutePosition
    UpdateRecordSet = True
    '*************************************Aduit*****************************************************
    '    If isDebugModeNow Then
    '        If Not IsAudit Then
    '            Dim MySource As String
    '            MySource = Trim(Replace(CurrentRS.source, "  ", " "))
    '            If AuditIsTable(MySource) Then
    '                If AduitBeginStart Then
    '                    AduitBeginStart = False
    '                    'AduitMasterTable = MySource
    '                    'AduitMasterTable = AduitLastSQL
    '                    If AduitMasterTable = "" Then
    '                        If AuditIsTableRowID(MySource) Then
    '                            AduitMasterTable = MySource
    '                        End If
    '                    End If
    '                End If
    '                If InStr(1, AduitOtherTable, MySource) = 0 Then
    '                    AduitOtherTable = AduitOtherTable & vbNewLine & MySource
    '                End If
    '            End If
    '        End If
    '    End If
    '*************************************Aduit*****************************************************
ff:
    UpdateRecordSet = False
End Function





Sub ButtonsMouseUp(Frm As Form, _
                   index As Integer, _
                   Button As Integer, _
                   Shift As Integer, _
                 X As Single, _
                 Y As Single, _
                   Optional tFrm As Form, _
                   Optional rs As Variant = Nothing, _
                   Optional newrecord As Variant = False)
    On Error GoTo eh
    If Shift = vbShiftMask And Button = 1 And index = CmdPrint Then
        IsPrintPreview = True
    ElseIf Shift = 0 And Button = 1 And index = CmdPrint Then
        IsPrintPreview = False
    End If

    If Shift = 0 And Button = vbRightButton And index = CmdPrint And Not newrecord Then
        If Not tFrm Is Nothing Then
            If Not rs Is Nothing Then
                If Not rs.EOF Then
                    'ShowPrintButtonMenu tFrm, RS
                End If
            End If
        End If
    End If

    Exit Sub
eh:
    MsgBox MyErrorHandler(Err)
End Sub





Public Function MyErrorHandler(ErrNo As Long) As String
    Mmsg = ""
    Select Case ErrNo
    
    Case 0
        MyErrorHandler = ""
        Exit Function
 
    Case -2147217864

        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = " „ ≈Ã—«¡  ⁄œÌ·«  ⁄·Ï Â–Â «·‘«‘Â „‰ ÃÂ«“ ¬Œ—- „‰ ›÷·ﬂ «⁄œ  Õ„Ì· «·Õ—ﬂÂ À„ Õ«Ê· „—Â «Œ—Ï" & " - Optimistic concurrency erorr "
        Else
            Mmsg = "This Form is editing from another computer- realod and Try again " & " - Optimistic concurrency erorr "
        End If

    Case -2147467259
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = "«·ÃÂ«“ «·Œ«œ„ «·—∆Ì”Ì „€·ﬁ √Ê €Ì— „ÊÃÊœ ⁄·Ï Â–Â «·‘»ﬂ…" & " - " & ErrNo
        Else
            Mmsg = "Couldn't close file , Because it's already in use by another person or process" & " - " & ErrNo
        End If
    Case -2147352567
        'If SystemOptions.UserInterface = ArabicInterface Then
        '    mMsg = "ÌÃ»  Œ’Ì’ «·ÿ«»⁄«  „‰ ≈œ«—… «·‰Ÿ«„" & " - " & ErrNo
        'Else
        '    mMsg = "Select Correct Report Printer Device" & " - " & ErrNo
        'End If
    Case 3155, 3022, -2147217873, -2147217900    ' insert fail
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = " ·«Ì„ﬂ‰ «÷«›… Â–« «·”Ã· ° Â–Â «·»Ì«‰«   „  ”ÃÌ·Â« „‰ ﬁ»·" & " - " & ErrNo
        Else
            Mmsg = "You Can not Add this Record , May be there is Dublicated values" & " - " & ErrNo
        End If
    Case 3200    ' Change Or Delete Failed
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = " ·«Ì„ﬂ‰ «·€«¡ √Ê  ⁄œÌ· Â–« «·”Ã·  »”»» ÊÃÊœ »Ì«‰«  √Œ—Ï „— »ÿ… »Â ÊÌÃ» «·€«¡Â« √Ê·«" & " - " & ErrNo
        Else
            Mmsg = "You Can not Delete Or Modify this Record , Because There Is Some Data Depends On It " & " - " & ErrNo
        End If
    Case 3157, 3046, 3202, 3218    ' Update Fail
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = " Â‰«ﬂ ›‘· ›Ï  Œ“Ì‰ «· ⁄œÌ·«  ° ﬁœ ÌﬂÊ‰ «·”Ã· „ﬁ›· »Ê«”ÿ… „” Œœ„ ¬Œ—° Õ«Ê· „—… √Œ—Ï" & " - " & ErrNo
        Else
            Mmsg = "Update Failed , May be The record is locked by another User , Try Again " & " - " & ErrNo
        End If
    Case 3186, 3187, 3188
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = "”Ã· „€·ﬁ »Ê«”ÿ… „” Œœ„ ¬Œ—" & " - " & ErrNo
        Else
            Mmsg = "Current Record locked by Another user" & " - " & ErrNo
        End If
    Case 3167
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = " „ «·€«¡ Â–« «·”Ã· »«·›⁄· " & " - " & ErrNo
        Else
            Mmsg = "Record Already Deleted" & " - " & ErrNo
        End If
    Case 3314
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = "„‰ ›÷·ﬂ √ﬂ„· «·»Ì«‰«  ﬁ»· «· Œ“Ì‰" & " - " & ErrNo
        Else
            Mmsg = "Please Complete the data before saving" & " - " & ErrNo
        End If
    Case 3262, 3211, 3212    ' Locked by another user and wait
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = "·« Ì„ﬂ‰ ≈€·«ﬁ «·„·› »”»» ÊÃÊœ „” Œœ„ ¬Œ— ÌﬁÊ„ »≈” Œœ«„Â √Ê ﬁ«„ »≈€·«ﬁÂ" & " - " & ErrNo
        Else
            Mmsg = "Couldn't close file , Because it's already in use by another person or process" & " - " & ErrNo
        End If
    Case 3197    ' Couldn't repaire this files
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = "√ﬂÀ— „‰ „” Œœ„ Õ«Ê·Ê«  €ÌÌ— ‰›” «·»Ì«‰«  ›Ï ‰›” «·Êﬁ " & " - " & ErrNo
        Else
            Mmsg = "Another Users are attempting to change the same data at the same time" & " - " & ErrNo
        End If
    Case 3056    ' Couldn't repaire this files
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = "·« Ì„ﬂ‰  ’·ÌÕ «·„·›«  «·„” Œœ„…" & " - " & ErrNo
        Else
            Mmsg = "Couldn't repaire this files" & " - " & ErrNo
        End If
    Case 3014, 3037    ' Can't open any more files
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = "·« Ì„ﬂ‰ › Õ „·›«  √Œ—Ï" & " - " & ErrNo
        Else
            Mmsg = "Can't open any more files" & " - " & ErrNo
        End If
    Case 3356, 3260, 3261, 3189, 3008, 3164, 3006    ' Table or Database Locked
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = "«·„·› „€·ﬁ »Ê«”ÿ… „” Œœ„ ¬Œ—" & " - " & ErrNo
        Else
            Mmsg = "The File is Locked by Another User" & " - " & ErrNo
        End If
    Case 3201    ' Add Or Edit Fail
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = " ·«Ì„ﬂ‰ «÷«›… Â–« «·”Ã· √Ê «· ⁄œÌ· ›ÌÂ ° ·√‰Â „— »ÿ »„·› ·„ Ì „ «·≈÷«›… √Ê «· ⁄œÌ· ›ÌÂ Õ Ï «·¬‰" & " - " & ErrNo
        Else
            Mmsg = "You Can not Add this Record or Change it , Because it's Linked to a File that has not been Added or Changed till Now" & " - " & ErrNo
        End If
    Case -2147217887
        If SystemOptions.UserInterface = ArabicInterface Then
            Mmsg = "Œÿ√ €Ì— „⁄—Ê› ° Õ«Ê·  ‰›Ì– ‰›” «·⁄„·Ì… „—… √Œ—Ï" & " - " & ErrNo
        Else
            Mmsg = "Undefined Error , Try again : " & ErrNo
        End If
    Case 3704
        '********By Khalid
        On Error Resume Next
        db.Close
        Exit Function
    Case -1000000001
        '«—Ê—— » «⁄ «·«Ê Ê ﬂÊ„»Ì·Ì  ›ﬂﬂ „‰Â
        MyErrorHandler = ""
        Exit Function
    End Select
    '*************************
    If Err.Number = vbObjectError + 1000 Then
        If Not ArabicInterface Then
            mText = Trim(Mmsg)
            If Trim(mText) <> "" Then
                Cond = "Arabic = N'" & Trim(mText) & "'"

                s = "Select * from Translations where " & Cond
                Set Translations = OpenRecordSet(s, adOpenStatic, adLockReadOnly)
                '------------------------
                If Not Translations.EOF Then
                    Mmsg = IIf(Trim(Translations!English & "") <> "", Trim(Translations!English & ""), Mmsg)
                End If
            End If
        End If

        Mmsg = Mmsg & vbNewLine & Err.Description
    Else
        Mmsg = Mmsg & vbNewLine & Err.Description & " : " & Err.Number
    End If
    '*************************
    If ErrNo <> -2147217864 Then  '  Ã«Â· «—Ê—— «·ﬂÊ‰ﬂ—‰”Ï  ‘Ìﬂ
        If Cn.Errors.count > 0 Then
            ss = ""
            Dim adoErr As ADODB.Error
            j = 1
            On Error GoTo EEE
            For Each adoErr In db.Errors
                If adoErr.Number <> 0 Then
                    If j = 1 Then ss = vbNewLine & "-------SQL Errors-------"
                    ss = ss & vbNewLine & "Error (" & j & ")=" & adoErr.Description & " : " & adoErr.Number & ":" & adoErr.SQLState & ":" & adoErr.NativeError
                    j = j + 1
                End If
            Next adoErr
EEE:
            ' for this rand error Not enough storage is available to process this command.
            If Err.Number = 48 Then
                Set adoErr = db.Errors(0)
                ss = ss & vbNewLine & "Error (48)=" & adoErr.Description & " : " & adoErr.Number & ":" & adoErr.SQLState & ":" & adoErr.NativeError
            End If
            On Error GoTo 0
            Mmsg = Mmsg & vbNewLine & ss
        End If
    End If
    '*************************
    'If Trim(mMsg) <> "()(0)" Then MyErrorHandler = mMsg Else MyErrorHandler = ""
    MyErrorHandler = Mmsg & ":" & Erl
    IsAboutError = True

End Function



Function ContainsKeywords(ByVal Name As String) As Boolean
    '  ÕÊÌ· «·‰’ ≈·Ï Õ«·… ’€Ì—… ·⁄„· „ﬁ«—‰… €Ì— Õ”«”… ·Õ«·… «·Õ—Ê›
    Dim lowerName As String
    lowerName = LCase(Name)

    ' «· Õﬁﬁ „‰ ÊÃÊœ √Ì „‰ «·ﬂ·„«  «·„Õœœ…
    If InStr(lowerName, "fg") > 0 Or InStr(lowerName, "grid") > 0 Or InStr(lowerName, "grd") > 0 Then
        ContainsKeywords = True
    Else
        ContainsKeywords = False
    End If
End Function

Function StartsWithKeywords(ByVal Name As String) As Boolean
    '  ÕÊÌ· «·‰’ ≈·Ï Õ«·… ’€Ì—… ·⁄„· „ﬁ«—‰… €Ì— Õ”«”… ·Õ«·… «·Õ—Ê›
    Dim lowerName As String
    lowerName = LCase(Name)
    
    ' «· Õﬁﬁ „‰ √‰ «·‰’ Ì»œ√ »√Ì „‰ «·ﬂ·„«  «·„Õœœ…
    If left(lowerName, 2) = "fg" Or left(lowerName, 4) = "grid" Or left(lowerName, 3) = "grd" Then
        StartsWithKeywords = True
    Else
        StartsWithKeywords = False
    End If
End Function


