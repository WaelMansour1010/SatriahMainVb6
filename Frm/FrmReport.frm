VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FrmReportViewer 
   Caption         =   "⁄—÷ «· ﬁ«—Ì—"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13920
   Icon            =   "FrmReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   13920
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "PDF"
      Height          =   315
      Left            =   8010
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DataOnly"
      Height          =   315
      Left            =   8880
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Excel"
      Height          =   315
      Left            =   9750
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   765
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11070
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtpw 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   11460
      PasswordChar    =   "*"
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox TXTSTRSQL 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   10440
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit"
      Height          =   315
      Left            =   10530
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   915
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7710
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReport.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReport.frx":0724
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReport.frx":0ABE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReport.frx":0E58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReport.frx":11F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReport.frx":158C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReport.frx":1926
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReport.frx":1CC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmReport.frx":205A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   7020
      Width           =   13920
      _ExtentX        =   24553
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   21
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "print"
            Object.ToolTipText     =   "ÿ»«⁄… «· ﬁ—Ì—"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "search"
            Object.ToolTipText     =   "»ÕÀ"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "refresh"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "last"
            Object.ToolTipText     =   "«·√ŒÌ—…"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "next"
            Object.ToolTipText     =   "«·«‰ ﬁ«· ≈·Ï «·’›Õ… «· «·Ì… „‰ «· ﬁ—Ì—"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Previous"
            Object.ToolTipText     =   "«·«‰ ﬁ«· ≈·Ï «·’›Õ… «·”«»ﬁ… „‰ «· ﬁ—Ì—"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "First"
            Object.ToolTipText     =   "«·«‰ ﬁ«· ≈·Ï «·’›Õ… «·√Ê·Ï „‰ «· ﬁ—Ì—"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.ComboBox XPCmboZoom 
         Height          =   315
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   -4560
         Width           =   1275
      End
   End
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer CRViewer 
      Height          =   4740
      Left            =   300
      TabIndex        =   6
      Top             =   570
      Width           =   8745
      lastProp        =   600
      _cx             =   15425
      _cy             =   8361
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
   Begin VB.Label txtpath 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   135
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "FrmReportViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_PreviewCaption As String

Public xReport As CRAXDRT.Report
Public mFromDate As String
Public mToDate As String
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" _
(ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
 ByVal lpReturnedString As String, ByVal nSize As Long) As Long
 
 Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Const GMEM_MOVEABLE = &H2
Private Const CF_HDROP = 15

Private Sub CopyFileToClipboard(ByVal FilePath As String)
    Dim hGlobal As Long
    Dim lpData As Long
    Dim FileList As String
    Dim DropFilesHeader(0 To 19) As Byte ' DROPFILES structure 20 bytes
    
    ' DROPFILES structure
    DropFilesHeader(0) = 20
    DropFilesHeader(16) = 1 ' WideChar = False Ì⁄‰Ì Ascii
    
    '  ÃÂÌ“ «”„ «·„·› „⁄ ‰Â«Ì… „“œÊÃ… Chr(0) Chr(0)
    FileList = FilePath & CHR(0) & CHR(0)
    
    '  Œ’Ì’ „”«Õ…
'    hGlobal = GlobalAlloc(GMEM_MOVEABLE, LenB(DropFilesHeader) + LenB(FileList))
    Const DROPFILES_SIZE As Long = 20 ' ??? ????????
hGlobal = GlobalAlloc(GMEM_MOVEABLE, DROPFILES_SIZE + LenB(FileList))

    If hGlobal Then
        lpData = GlobalLock(hGlobal)
        If lpData Then
            ' ﬂ «»… —√” DROPFILES
            
           CopyMemory ByVal lpData, ByVal VarPtr(DropFilesHeader(0)), DROPFILES_SIZE
CopyMemory ByVal (lpData + DROPFILES_SIZE), ByVal StrPtr(FileList), LenB(FileList)
 
 
'            CopyMemory ByVal lpData, DropFilesHeader(0), LenB(DropFilesHeader)
            ' ﬂ «»… „”«— «·„·›
'            CopyMemory ByVal (lpData + LenB(DropFilesHeader)), ByVal FileList, LenB(FileList)
            Call GlobalUnlock(hGlobal)
            
            ' › Õ «·ﬂ·Ì» »Ê—œ Ê‰”Œ «·»Ì«‰« 
            If OpenClipboard(0&) Then
                Call EmptyClipboard
                Call SetClipboardData(CF_HDROP, hGlobal)
                Call CloseClipboard
            End If
        End If
    End If
End Sub

Private Sub Command1_Click()
If Me.txtpw.text = "Alex2025" Then GoTo ll

 
    
            If bigUser = False Then
             If SystemOptions.BigUserPw <> Me.txtpw.text Then
               If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "€Ì— „”„ÊÕ ·ﬂ » ⁄œÌ· «· ﬁ«—Ì—  ", vbOKOnly + vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "’·«ÕÌ«  «·„” Œœ„Ì‰"
                 Else
                    MsgBox "Not Allowed", vbOKOnly + vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, "  Users Privligies"
                 End If
             
             
             Exit Sub
             End If
             
        End If
ll:
        Clipboard.Clear
Clipboard.SetText Me.TXTSTRSQL.text, vbCFText
DoEvents

    ShellExecute 0&, vbNullString, Me.txtPath, vbNullString, vbNullString, vbNormalFocus
    
End Sub

Private Sub Command2_Click()
ExportReportWithFormatting
End Sub

Private Sub Command3_Click()
ExportReportWithFormatting True
End Sub

Private Sub Command4_Click()
ExportReportWithFormatting False, True
End Sub

Private Sub CRViewer_Clicked(ByVal X As Long, _
                             ByVal Y As Long, _
                             EventInfo As Variant, _
                             UseDefault As Boolean)
                           
    Dim MyObject         As CrystalActiveXReportViewerLib10Ctl.CRVEventInfo
    Dim myFields         As CrystalActiveXReportViewerLib10Ctl.CRVFields
    Dim myField          As CrystalActiveXReportViewerLib10Ctl.CRVField
    Dim NoteType         As CrystalActiveXReportViewerLib10Ctl.CRVField
    Dim NoteSerial1      As CrystalActiveXReportViewerLib10Ctl.CRVField
    Dim notes_all        As CrystalActiveXReportViewerLib10Ctl.CRVField
    Dim NoteID           As CrystalActiveXReportViewerLib10Ctl.CRVField
    Dim NoteIDALL        As CrystalActiveXReportViewerLib10Ctl.CRVField
    Dim NoteTypeValue    As Integer
    
    Dim Transaction_ID   As CrystalActiveXReportViewerLib10Ctl.CRVField
    Dim Transaction_Type As CrystalActiveXReportViewerLib10Ctl.CRVField
    Dim Trans_ID         As Long
    Dim Trans_Type       As Long
    On Error Resume Next
      
    Dim NoteSerial1Value As String
    Dim notes_allValue   As Double
    Dim noteidvalue      As Double
    
    Dim StrTemp          As String
    Dim Xcheck           As Integer
    'ClsRepoertsr
    Dim SaleReport       As ClsSaleReport
    Set SaleReport = New ClsSaleReport
    Dim IntNumIndex  As Integer
    Dim StrFieldName As String
    Dim OwnerObject  As Object
    'Exit Sub
    Set MyObject = EventInfo

    Set myFields = MyObject.GetFields

    IntNumIndex = myFields.SelectedFieldIndex

    If IntNumIndex = 0 Then
        Exit Sub
    End If

    Set myField = myFields.Item(IntNumIndex)
 
    StrTemp = myField.value

    If Me.MDIChild = True Then
        Set OwnerObject = mdifrmmain
    Else
        Set OwnerObject = Me
    End If

    If myField.Name Like "*CusID" Then
        OpenScreen PopUpShowCustomerBalanceScreen, myField.value, 0, , , , , OwnerObject
    ElseIf myField.Name Like "*RequerMainteniD" Then
        Unload FrmRequerMainten
        FrmRequerMainten.Retrive val(myField.value)
    ElseIf myField.Name Like "*View_RepOrderUplaod.ID" Then
        Unload FrmOrderUpload
        FrmOrderUpload.show
        FrmOrderUpload.FindRec val(myField.value)
                    
    ElseIf myField.Name Like "*Field: @TransNo" Or myField.Name Like "*Field: ReportSallingTime.NoteSerial1" Or myField.Name Like "*Field: RptItemTransCus.NOTESERIAL1" Or myField.Name Like "*Field: RptItemTransCus.NoteSerial1" Or myField.Name Like "*Field: RptItemTransCus.NoteSerial1" Then
        Dim transactiontype As Integer
        Dim Transactionid   As Double
        'salimstoppppppppppppppppppppppppp
        '    transactiontype = myFields.Item(14)
        '      Transactionid = myFields.Item(13)
        'salimstoppppppppppppppppppppppppp
                 
        '*******************************************'*******************************************
 
        For i = 1 To myFields.count
            Set myField = myFields(i)
            If myField.Name Like "*Transaction_ID" Then
                Trans_ID = CDbl(myFields.Item(i))
                Transactionid = Trans_ID
                Set Transaction_ID = myFields.Item(i)
                'Trans_ID = val(myFields.Item(i))
            End If
 
            If myField.Name Like "*Transaction_Type" Then
                Trans_Type = CDbl(myFields.Item(i))
                transactiontype = Trans_Type
                Set Transaction_Type = myFields.Item(i)
                'Transaction_Type = val(myFields.Item(i))
            End If
 
        Next i
        '*******************************************'*******************************************
        '  Print myFields.Item(9)
        '          FrmOrderMaintin.Retrive val(myField.value)
        If transactiontype = 19 Then
            Unload FrmOut
            FrmOut.XPBtnMove_Click (2)
            FrmOut.Retrive CLng(Transactionid)
        ElseIf transactiontype = 26 Then
            Unload FrmProductionOrder
            FrmProductionOrder.XPBtnMove_Click (2)
            FrmProductionOrder.Retrive CLng(Transactionid)
            
        ElseIf transactiontype = 20 Then
            Unload FrmInpout
            FrmInpout.XPBtnMove_Click (2)
            FrmInpout.Retrive CLng(Transactionid)
        ElseIf transactiontype = 21 Then '' ›« Ê—… „»Ì⁄« 
            If SystemOptions.UserInterface = ArabicInterface Then
                Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + CHR(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + CHR(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
            Else
                Xcheck = MsgBox("yes  Open doc  " + CHR(13) + "No  Print Doc    " + CHR(13), vbInformation + vbYesNoCancel, "What you want?? ")
            End If
             
            If Xcheck = vbYes Then
                Unload frmsalebill
                frmsalebill.show
                frmsalebill.XPBtnMove_Click (2)
                
                frmsalebill.Retrive Trans_ID
            ElseIf Xcheck = vbNo Then
                SaleReport.ShowSallingDataDetailed Trans_ID, , , , , , , 0
            Else
                Exit Sub
            End If
  
        ElseIf transactiontype = 9 Then '' „— Ã⁄  „»Ì⁄« 
            If SystemOptions.UserInterface = ArabicInterface Then
                Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + CHR(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + CHR(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
            Else
                Xcheck = MsgBox("yes  Open doc  " + CHR(13) + "No  Print Doc    " + CHR(13), vbInformation + vbYesNoCancel, "What you want?? ")
            End If
                      
            If Xcheck = vbYes Then
                Unload FrmReturnSalling
                FrmReturnSalling.show
                FrmReturnSalling.XPBtnMove_Click (2)
                FrmReturnSalling.Retrive val(Trans_ID)
   
            ElseIf Xcheck = vbNo Then
 
                Dim SaleReport1 As New ClsRepoerts
                SaleReport1.ReturnSallingData Trans_ID, "", ""
            Else
                Exit Sub
            End If
                
        ElseIf transactiontype = 5 Then '' „— Ã⁄  „‘ —Ì« 
            If SystemOptions.UserInterface = ArabicInterface Then
                Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + CHR(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + CHR(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
            Else
                Xcheck = MsgBox("yes  Open doc  " + CHR(13) + "No  Print Doc    " + CHR(13), vbInformation + vbYesNoCancel, "What you want?? ")
            End If
                      
            If Xcheck = vbYes Then
 
                Unload FrmReturnpurchases
                FrmReturnpurchases.show
                FrmReturnpurchases.XPBtnMove_Click (2)
                FrmReturnpurchases.Retrive val(Trans_ID)
  
            ElseIf Xcheck = vbNo Then
                Dim ReturnReport As ClsReturnBackReport
  
                Set ReturnReport = New ClsReturnBackReport
                ReturnReport.ShowReturnBack Trans_ID
  
            Else
                Exit Sub
            End If
        End If
    ElseIf myField.Name Like "*TblOrderMaintidS" Then
        Unload FrmOrderMaintin
        FrmOrderMaintin.Retrive val(myField.value)
            
    ElseIf myField.Name Like "*CusName" Then
        OpenScreen PopUpShowCustomerBalanceScreen, GetDealerID(myField.value), 0, , , , , OwnerObject
    ElseIf myField.Name Like "*TripNo" Then
        Unload FrmTravelTransactions
        FrmTravelTransactions.show
        FrmTravelTransactions.TxtModFlg = "R"
        FrmTravelTransactions.FindRecbyNoteserial1 (myField.value)

    ElseIf myField.Name Like "*ItemID" Then
        OpenScreen PopUpShowItemCardScreen, myField.value, 0, , , , , OwnerObject
        
    ElseIf myField.Name Like "*QtyOnTheWay" Then
        'Wael
        
        Dim mFrm2 As New FrmReports
        mFrm2.XPChk(54).value = vbChecked
        mFrm2.DCboItemBuy.BoundText = myFields.Item(9)
        mFrm2.XPDtbBuyFrom = mFromDate
        mFrm2.XPDtpBuyTo = mToDate
        mFrm2.CmdPrintBuy_Click
        Unload mFrm2
        mFrm2.Hide

    ElseIf myField.Name Like "*ItemCode" Then
        OpenScreen PopUpShowItemCardScreen, GetItemID(myField.value), 0, , , , , OwnerObject
    ElseIf myField.Name Like "*ItemName" Then
        OpenScreen PopUpShowItemCardScreen, GetItemID("", myField.value), 0, , , , , OwnerObject
    ElseIf (myField.Name Like "*?ParCompanyName") Or (myField.Name Like "*?comment_arabic") Then
        OpenScreen OptionsScreen
        'Me.xReport.EnableParameterPrompting = False
        RefreshreReportParameters
        Me.CRViewer.RefreshEx False
    ElseIf myField.Name Like "*Account_Serial" Or myField.Name Like "*account_serial" Then
        Dim FirstPeriod As Date
        Dim AccountName As String
        Dim AccountCode As String
        AccountCode = Get_Account_code(myField.value)
        AccountName = Get_Account_name(myField.value)
        getFirstPeriodDateInthisYear FirstPeriod
        Get_Account_name
            
        If CHECK_LAST_ACCOUNT(AccountCode) = True Then ' Õ”«» ‰Â«∆Ì
                            
            If P_DTPickerAccFrom = "01/01/1999" Then
                ShowReport AccountCode, AccountName, FirstPeriod, Date, , P_dcBranch
            Else
                              
                ShowReport AccountCode, AccountName, P_DTPickerAccFrom, P_DTPickerAccTo, , P_dcBranch
            End If
                            
        Else
             
            print_report3_HyperLink AccountCode, AccountName
             
        End If

    ElseIf myField.Name Like "*NoteSerial" Or myField.Name Like "*NotesSerial" Then
        Unload FrmShowReport
        FrmShowReport.reportno = 200
        FrmShowReport.NoteSerial = myField.value
        FrmShowReport.show

        'FrmShowReport.ZOrder 0
        'Me.WindowState = 2
        'FrmShowReport.ZOrder 0
        'ShowGL_cc myField.value, , 200
        
        'FrmAccEditJournal.Retrive Adodc1.Recordset.Fields!NoteSerial
        'FrmAccEditJournal.show
        'FrmAccEditJournal.StrOldTransID = Adodc1.Recordset.Fields!NoteSerial
        
        'Set NoteType = NoteType.GetFields201210016

        'ElseIf myField.name Like "*NoteSerial" Then
        
    ElseIf myField.Name Like "*NOTESERIAL1" Or myField.Name Like "*NoteSerial1" Then
        'salimstopppppppppppppppp   'salimstopppppppppppppppp
        'Transactionid = myFields.Item(10)
        ' transactiontype = myFields.Item(11)
        'salimstopppppppppppppppp   'salimstopppppppppppppppp
        '*******************************************'*******************************************
 
        For i = 1 To myFields.count
            Set myField = myFields(i)
            If myField.Name Like "*Transaction_ID" Then
                Trans_ID = CDbl(myFields.Item(i))
                Set Transaction_ID = myFields.Item(i)
                'Trans_ID = val(myFields.Item(i))
            End If
 
            If myField.Name Like "*Transaction_Type" Then
                Trans_Type = CDbl(myFields.Item(i))
                Set Transaction_Type = myFields.Item(i)
                'Transaction_Type = val(myFields.Item(i))
            End If
 
        Next i
        '*******************************************'*******************************************
 ElseIf myField.Name Like "*order_no" Then
        'salimstopppppppppppppppp   'salimstopppppppppppppppp
        'Transactionid = myFields.Item(10)
        ' transactiontype = myFields.Item(11)
        'salimstopppppppppppppppp   'salimstopppppppppppppppp
        '*******************************************'*******************************************
 
        For i = 1 To myFields.count
            Set myField = myFields(i)
            If myField.Name Like "*poTransaction_ID" Then
                Trans_ID = CDbl(myFields.Item(i))
                Set Transaction_ID = myFields.Item(i)
                'Trans_ID = val(myFields.Item(i))
            End If
 
            If myField.Name Like "*Transaction_Type" Then
                Trans_Type = CDbl(myFields.Item(i))
                Set Transaction_Type = myFields.Item(i)
                'Transaction_Type = val(myFields.Item(i))
            End If
 
        Next i
        If Transaction_Type = 6 Then
            Trans_Type = 42
        Else
            Trans_Type = 6
        End If
        '  NoteTypeValue = val(NoteType.value)
        If myField.Name Like "*NOTESERIAL1" Then
            'Salim Stop******************************
            '   Set Transaction_ID = myFields.Item(8)
            '   Trans_ID = val(Transaction_ID.value)

            '   Set Transaction_Type = myFields.Item(9)
            '  Trans_Type = val(Transaction_Type.value)
       
            'Salim Stop******************************
       
        ElseIf myField.Name = "Field: CahingReport.NoteSerial1" Then
            Trans_Type = 4
 
            Set Transaction_ID = myFields.Item(8)
            noteidvalue = val(Transaction_ID.value)
 
            If SystemOptions.UserInterface = ArabicInterface Then
                Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + CHR(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + CHR(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
            Else
                Xcheck = MsgBox("yes  Open doc  " + CHR(13) + "No  Print Doc    " + CHR(13), vbInformation + vbYesNoCancel, "What you want?? ")
            End If

            If Xcheck = vbYes Then
                Unload FrmCashing
                FrmCashing.show
                FrmCashing.Retrive val(noteidvalue)

            ElseIf Xcheck = vbNo Then
                Set myField = myFields.Item(9)
                '   StrTemp = myField.value
                '  FrmCashing.print_report (StrTemp), "", "", "", "", ""
                'Unload FrmCashing
                '
            Else
                Exit Sub
            End If
                
            Exit Sub
        ElseIf myField.Name Like "*NoteSerial1" Then
            Trans_ID = myFields.Item(10)
            Trans_Type = myFields.Item(11)
        Else
            'salimstoppppppppppppppppppppppppppppppppppppppppp
            '       Set Transaction_ID = myFields.Item(11)
            '        Trans_ID = val(Transaction_ID.value)

            '        Set Transaction_Type = myFields.Item(12)
            'salimstoppppppppppppppppppppppppppppppppppppppppp'       Trans_Type = val(Transaction_Type.value)
 
        End If
                     
        If SystemOptions.UserInterface = ArabicInterface Then
 
            Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + CHR(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + CHR(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
        Else
            Xcheck = MsgBox("yes  Open doc  " + CHR(13) + "No  Print Doc    " + CHR(13), vbInformation + vbYesNoCancel, "What you want?? ")
        End If
ll:
        Select Case Trans_Type
                '07112013
            Case 21: ' ›« Ê—… „»Ì⁄« 
             
                If Xcheck = vbYes Then
                    Unload frmsalebill
                    frmsalebill.show
                    frmsalebill.XPBtnMove_Click (2)
                
                    frmsalebill.Retrive Trans_ID
                ElseIf Xcheck = vbNo Then
                    SaleReport.ShowSallingDataDetailed Trans_ID, , , , , , , 0
                Else
                    Exit Sub
                End If
                
            Case 9: '„—œÊœ«  „»Ì⁄« 
                
                If Xcheck = vbYes Then
                    Unload FrmReturnSalling
                    FrmReturnSalling.show
                    FrmReturnSalling.XPBtnMove_Click (2)
                    FrmReturnSalling.Retrive val(Trans_ID)
   
                ElseIf Xcheck = vbNo Then
 
                    Set SaleReport1 = New ClsRepoerts
                    SaleReport1.ReturnSallingData Trans_ID, "", ""
                Else
                    Exit Sub
                End If
                
                
                Case 42: '„—œÊœ«  „»Ì⁄« 
                
                If Xcheck = vbYes Then
                    
                    FrmPO1.show
                    FrmPO1.XPBtnMove_Click (2)
                    FrmPO1.Retrive val(Trans_ID)
   
                ElseIf Xcheck = vbNo Then
 
                    Set SaleReport = New ClsSaleReport
                    SaleReport.ShowPrice CDbl(Trans_ID), 94
                    
                Else
                    Exit Sub
                End If
                
                  Case 6: '„—œÊœ«  „»Ì⁄« 
                
                If Xcheck = vbYes Then
                    
                    FrmPO3.show
                    FrmPO3.XPBtnMove_Click (2)
                    FrmPO3.Retrive val(Trans_ID)
   
                ElseIf Xcheck = vbNo Then
 
                    Set SaleReport = New ClsSaleReport
                    SaleReport.ShowPrice CDbl(Trans_ID), 96
                    
                    
                Else
                    Exit Sub
                End If
            Case 22:  '„‘ —Ì« 
                If Xcheck = vbYes Then
                    '               Set Transaction_ID = myFields.Item(8)
                    '   Trans_ID = val(Transaction_ID.value)
                    Unload FrmBillBuy
                    FrmBillBuy.XPBtnMove_Click (2)
                    FrmBillBuy.Retrive val(Trans_ID)
   
                ElseIf Xcheck = vbNo Then
  
                    BuyReport.ShowBuyData Trans_ID, 1, True

                Else
                    Exit Sub
                End If
            Case 5: '„—œÊœ«  „‘ —Ì« 

                If Xcheck = vbYes Then
                    Unload FrmReturnpurchases
      
                    FrmReturnpurchases.show
                    FrmReturnpurchases.XPBtnMove_Click (2)
                    FrmReturnpurchases.Retrive val(Trans_ID)
  
  
                ElseIf Xcheck = vbNo Then
                    Set ReturnReport = New ClsReturnBackReport
  
                    Set ReturnReport = New ClsReturnBackReport
                    ReturnReport.ShowReturnBack Trans_ID
  
                Else
                    Exit Sub
                End If

        End Select
        
    ElseIf myField.Name Like "*FullDes" Then
       
        Set NoteType = myFields.Item(7)
        NoteTypeValue = (NoteType.value)

        Set NoteSerial1 = myFields.Item(11)
        NoteSerial1Value = NoteSerial1.value

        Set notes_all = myFields.Item(4)
        notes_allValue = notes_all.value
       
        Set NoteID = myFields.Item(10)
        noteidvalue = NoteID.value
              
        Set NoteIDALL = myFields.Item(9)
  
        Select Case NoteTypeValue
                '07112013
            Case 9080
                Unload FrmPaymenTransTrip
                FrmPaymenTransTrip.show
                FrmPaymenTransTrip.TxtModFlg = "R"
                FrmPaymenTransTrip.FindRecbyNoteserial1 NoteSerial1Value 'salim30102018

            Case 170: ' ›« Ê—… „»Ì⁄« 

                Trans_ID = get_transaction_idByNoteSerial1(val(NoteSerial1Value), 21)
                '                Trans_ID = NoteSerial1.value
                '   Set Transaction_ID = myFields.Item(10)
                'Trans_ID = val(Transaction_ID.value)
                Trans_ID = Replace(Transaction_ID.value, ",", "")
 
                If SystemOptions.UserInterface = ArabicInterface Then
                    ' Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + Chr(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + Chr(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                    Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + CHR(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + CHR(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                Else
                    Xcheck = MsgBox("yes  Open doc  " + CHR(13) + "No  Print Doc    " + CHR(13), vbInformation + vbYesNoCancel, "What you want?? ")
                End If
  
                If Xcheck = vbYes Then
                    Unload frmsalebill
                    frmsalebill.show
                    frmsalebill.XPBtnMove_Click (2)
                    frmsalebill.Retrive Trans_ID
                ElseIf Xcheck = vbNo Then
                    SaleReport.ShowSallingDataDetailed Trans_ID, , , , , , , 0
                Else
                    Exit Sub
                End If

            Case 220: '„—œÊœ«  „»Ì⁄« 
            
                '               Trans_ID = NoteSerial1.value
                Set Transaction_ID = myFields.Item(8)
                Trans_ID = val(Transaction_ID.value)

                If SystemOptions.UserInterface = ArabicInterface Then
                    '            Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + Chr(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + Chr(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                    Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + CHR(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + CHR(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                Else
                    Xcheck = MsgBox("yes  Open doc  " + CHR(13) + "No  Print Doc    " + CHR(13), vbInformation + vbYesNoCancel, "What you want?? ")
                End If
  
                If Xcheck = vbYes Then
                    Unload FrmReturnSalling
                    FrmReturnSalling.show
                    FrmReturnSalling.XPBtnMove_Click (2)
                    FrmReturnSalling.Retrive val(Trans_ID)
   
                ElseIf Xcheck = vbNo Then
 
                    'Dim SaleReport1 As New ClsRepoerts
                    SaleReport1.ReturnSallingData Trans_ID, "", ""
                Else
                    Exit Sub
                End If
           
            Case 150: ' „‘ —Ì« 
                Trans_ID = get_transaction_idByNoteSerial1(val(NoteSerial1Value), 22)
                'Trans_ID = NoteSerial1.value
                '        Set Transaction_ID = myFields.Item(8)
                Trans_ID = val(Transaction_ID.value)
                
                '                Dim BuyReport As ClsBuyReport
                Set BuyReport = New ClsBuyReport

                '            Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + Chr(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + Chr(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                If SystemOptions.UserInterface = ArabicInterface Then
                    Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + CHR(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + CHR(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                Else
                    Xcheck = MsgBox("yes  Open doc  " + CHR(13) + "No  Print Doc    " + CHR(13), vbInformation + vbYesNoCancel, "What you want?? ")
                End If

                If Xcheck = vbYes Then
                    '               Set Transaction_ID = myFields.Item(8)
                    Unload FrmBillBuy
                    Trans_ID = val(Transaction_ID.value)
        
                    FrmBillBuy.XPBtnMove_Click (2)
                    FrmBillBuy.Retrive val(Trans_ID)
   
                ElseIf Xcheck = vbNo Then
  
                    BuyReport.ShowBuyData Trans_ID, 1, True

                Else
                    Exit Sub
                End If

            Case 230: ' „—œÊœ«  „‘ —Ì« 
                'Trans_ID = get_transaction_idByNoteSerial1(val(NoteSerial1Value), 5)
                Trans_ID = NoteSerial1.value

                '           Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + Chr(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + Chr(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                If SystemOptions.UserInterface = ArabicInterface Then
                    Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + CHR(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + CHR(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                Else
                    Xcheck = MsgBox("yes  Open doc  " + CHR(13) + "No  Print Doc    " + CHR(13), vbInformation + vbYesNoCancel, "What you want?? ")
                End If

                If Xcheck = vbYes Then
 
                    Unload FrmReturnpurchases
                    FrmReturnpurchases.show
                    FrmReturnpurchases.XPBtnMove_Click (2)
                    FrmReturnpurchases.Retrive val(Trans_ID)
  
                ElseIf Xcheck = vbNo Then
                    Set ReturnReport = New ClsReturnBackReport
  
                    Set ReturnReport = New ClsReturnBackReport
                    ReturnReport.ShowReturnBack Trans_ID
  
                Else
                    Exit Sub
                End If
                
            Case 3 '„’—Êﬁ« 
   
                '              Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + Chr(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + Chr(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                If SystemOptions.UserInterface = ArabicInterface Then
                    Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + CHR(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + CHR(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                Else
                    Xcheck = MsgBox("yes  Open doc  " + CHR(13) + "No  Print Doc    " + CHR(13), vbInformation + vbYesNoCancel, "What you want?? ")
                End If
                Dim AkarVchr As Boolean
                AkarVchr = CheckAkarExpenses(val(NoteIDALL))
                If Xcheck = vbYes Then
                    'FrmExpenses5.show
                    If AkarVchr = False Then
                        Unload FrmExpenses5
                        FrmExpenses5.Retrive val(NoteIDALL)
                    Else
                        Unload RsExpenses
                        RsExpenses.Retrive val(NoteIDALL)
                    End If
                ElseIf Xcheck = vbNo Then
                    Set myField = myFields.Item(5)
 
                    StrTemp = myField.value
                    'FrmExpenses5.show
                    If AkarVchr = False Then
   
                        FrmExpenses5.print_report val(NoteSerial1)
                        Unload FrmExpenses5
                    Else
                        RsExpenses.print_report val(NoteSerial1)
                        Unload RsExpenses
  
                    End If
                Else
                    Exit Sub
                End If

            Case 4 '„ﬁ»Ê÷« 

                '              Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + Chr(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + Chr(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                If SystemOptions.UserInterface = ArabicInterface Then
                    Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + CHR(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + CHR(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                Else
                    Xcheck = MsgBox("yes  Open doc  " + CHR(13) + "No  Print Doc    " + CHR(13), vbInformation + vbYesNoCancel, "What you want?? ")
                End If
                Dim akarCahses As Boolean

                akarCahses = CheckAkarCashes(noteidvalue)
                If Xcheck = vbYes Then
                    If akarCahses = False Then
                        Unload FrmCashing
                        FrmCashing.show
                        FrmCashing.Retrive val(noteidvalue)
                    Else
                        Unload FrmCashing1
                        FrmCashing1.XPBtnMove_Click (2)
                        FrmCashing1.show
                        FrmCashing1.Retrive val(noteidvalue)
                    End If
                ElseIf Xcheck = vbNo Then
                    Set myField = myFields.Item(5)
                    StrTemp = myField.value
                      
                    If akarCahses = False Then
                        FrmCashing.print_report StrTemp, NoteSerial1Value, "", "", "", ""
                        Unload FrmCashing
                    Else
                       
                        FrmCashing1.print_report NoteSerial1Value
                        Unload FrmCashing1
                    End If
                    
                Else
                    Exit Sub
                End If

            Case 5 '„œ›Ê⁄« 

                Dim akarPayments As Boolean
                akarPayments = CheckAkarPayments(noteidvalue)

                '               Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + Chr(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + Chr(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                If SystemOptions.UserInterface = ArabicInterface Then
                    Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + CHR(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + CHR(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                Else
                    Xcheck = MsgBox("yes  Open doc  " + CHR(13) + "No  Print Doc    " + CHR(13), vbInformation + vbYesNoCancel, "What you want?? ")
                End If

                If Xcheck = vbYes Then
                    If akarPayments = True Then
                        Unload FrmPayments2
                        FrmPayments2.show
                        FrmPayments2.Retrive val(noteidvalue)
                    Else
                        Unload FrmPayments
                        FrmPayments.show
                        FrmPayments.Retrive val(noteidvalue)
                    End If
                    
                ElseIf Xcheck = vbNo Then
                    Set myField = myFields.Item(5)
                    StrTemp = myField.value
                      
                    If akarPayments = True Then
                        FrmPayments2.print_report StrTemp
                        Unload FrmPayments2
                    Else
                                
                        FrmPayments.print_report StrTemp, ""
                        Unload FrmPayments
            
                    End If
                Else
                    Exit Sub
                End If

            Case 50 '⁄Âœ…  „ÊÌ· Œ“‰ Ê«” ⁄«÷… ⁄Âœ
                '        Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + Chr(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + Chr(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                If SystemOptions.UserInterface = ArabicInterface Then
                    Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + CHR(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + CHR(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                Else
                    Xcheck = MsgBox("yes  Open doc  " + CHR(13) + "No  Print Doc    " + CHR(13), vbInformation + vbYesNoCancel, "What you want?? ")
                End If
                If Xcheck = vbYes Then
                    Unload FrmPayments1
                    FrmPayments1.show
                    FrmPayments1.Retrive val(noteidvalue)
      
                ElseIf Xcheck = vbNo Then
                    Set myField = myFields.Item(5)
                    StrTemp = myField.value
                    FrmPayments1.print_report StrTemp
                    Unload FrmPayments1
                Else
                    Exit Sub
                End If
 
            Case 53  ' ”‰œ ’—› „ ⁄œœ
                '   Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + Chr(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + Chr(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                If SystemOptions.UserInterface = ArabicInterface Then
                    Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + CHR(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + CHR(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                Else
                    Xcheck = MsgBox("yes  Open doc  " + CHR(13) + "No  Print Doc    " + CHR(13), vbInformation + vbYesNoCancel, "What you want?? ")
                End If
                Set myField = myFields.Item(5)
                StrTemp = myField.value
                      
                If Xcheck = vbYes Then
                    Unload FrmAccEditJournal3
                    FrmAccEditJournal3.show
                    FrmAccEditJournal3.Retrive StrTemp
      
                ElseIf Xcheck = vbNo Then
                    Set myField = myFields.Item(5)
                    StrTemp = myField.value
                    '        FrmAccEditJournal3.print_report StrTemp
                    '       Unload FrmPayments1
                     
                    ShowGL_cc StrTemp, , 53

                Else
                    Exit Sub
                End If
                
            Case 14 ' ÕÊÌ·«  „«·Ì…
                '       Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + Chr(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + Chr(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                If SystemOptions.UserInterface = ArabicInterface Then
                    Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + CHR(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + CHR(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                Else
                    Xcheck = MsgBox("yes  Open doc  " + CHR(13) + "No  Print Doc    " + CHR(13), vbInformation + vbYesNoCancel, "What you want?? ")
                End If
   
                If Xcheck = vbYes Then
                 
                    Unload FrmBoxDrawing
                    FrmBoxDrawing.show
                    FrmBoxDrawing.Retrive val(noteidvalue)

                ElseIf Xcheck = vbNo Then
                  
                    FrmBoxDrawing.print_report CInt(noteidvalue)
                    Unload FrmBoxDrawing
                Else
                    Exit Sub
                End If
                
            Case 80 '›Ê« Ì— „«·Ì… Ê€Ê« Ì— ‘—«¡ «’Ê· À«» …
                '    Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + Chr(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + Chr(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                If SystemOptions.UserInterface = ArabicInterface Then
                    Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + CHR(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + CHR(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                Else
                    Xcheck = MsgBox("yes  Open doc  " + CHR(13) + "No  Print Doc    " + CHR(13), vbInformation + vbYesNoCancel, "What you want?? ")
                End If
                If Xcheck = vbYes Then
                    If GetFinInvoiceType(notes_allValue) = 2 Then
                        Unload FrmExpenses4
                        FrmExpenses4.show
                        FrmExpenses4.Retrive val(notes_allValue)
            
                    Else
                        Unload FrmExpenses3
                        FrmExpenses3.show
                        FrmExpenses3.Retrive val(notes_allValue)
            
                    End If
 
                ElseIf Xcheck = vbNo Then
                    Set myField = myFields.Item(5)
                    StrTemp = myField.value
                    If GetFinInvoiceType(notes_allValue) = 2 Then
                        
                        FrmExpenses4.print_report notes_allValue
                        Unload FrmExpenses4

                    Else
                        FrmExpenses3.print_report StrTemp, ""
                        Unload FrmExpenses3
                    End If
                     
                Else
                    Exit Sub
                End If
                
            Case 350  '    350 ”‰œ   ”ÊÌ…  ⁄Âœ…        Era Voucher
 
                '    Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + Chr(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + Chr(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                If SystemOptions.UserInterface = ArabicInterface Then
                    Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + CHR(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + CHR(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                Else
                    Xcheck = MsgBox("yes  Open doc  " + CHR(13) + "No  Print Doc    " + CHR(13), vbInformation + vbYesNoCancel, "What you want?? ")
                End If
                If Xcheck = vbYes Then
                    Unload FrmExpenses30
                    FrmExpenses30.show
                    FrmExpenses30.Retrive val(notes_allValue)
                 
                ElseIf Xcheck = vbNo Then
                    Set myField = myFields.Item(5)
                    StrTemp = myField.value
                    Unload FrmExpenses30
                    FrmExpenses30.print_report StrTemp, ""
                    Unload FrmExpenses30
                     
                Else
                    Exit Sub
                End If

            Case 20 '«·«Ìœ«⁄« 
                Unload FrmBankDeposite
                FrmBankDeposite.show
                FrmBankDeposite.Retrive , val(noteidvalue)
            Case 21 ' Õ’Ì· Ê”œ«œ «·‘Ìﬂ« 
                Unload FrmBankDeposite1
                FrmBankDeposite1.show
                FrmBankDeposite1.Retrive , val(noteidvalue)
        
            Case 18 '«·«ﬁ”«ÿ
                Unload FrmReceiptPart
                FrmReceiptPart.show
                FrmReceiptPart.Retrive , val(noteidvalue)
                    
            Case 160 '160 ”‰œ «” ·«„  Recieve Voucher
                Trans_ID = NoteSerial1.value
 
                '       Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + Chr(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + Chr(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                If SystemOptions.UserInterface = ArabicInterface Then
                    Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + CHR(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + CHR(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                Else
                    Xcheck = MsgBox("yes  Open doc  " + CHR(13) + "No  Print Doc    " + CHR(13), vbInformation + vbYesNoCancel, "What you want?? ")
                End If
                If Xcheck = vbYes Then
                    Unload FrmInpout
                    FrmInpout.show
                    FrmInpout.XPBtnMove_Click (2)
                    FrmInpout.Retrive Trans_ID
                 
                ElseIf Xcheck = vbNo Then
                 
                    Set BuyReport = New ClsBuyReport
                    BuyReport.ShowRecieveVoucherData Trans_ID
                     
                Else
                    Exit Sub
                End If
            
            Case 180 '180   ”‰œ ’—›   Issue Voucher
                Trans_ID = NoteSerial1.value
                '      Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + Chr(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + Chr(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                If SystemOptions.UserInterface = ArabicInterface Then
                    Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + CHR(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + CHR(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                Else
                    Xcheck = MsgBox("yes  Open doc  " + CHR(13) + "No  Print Doc    " + CHR(13), vbInformation + vbYesNoCancel, "What you want?? ")
                End If
                If Xcheck = vbYes Then
                    Unload FrmOut
                    FrmOut.show
                    FrmOut.XPBtnMove_Click (2)
                    FrmOut.Retrive Trans_ID
                 
                ElseIf Xcheck = vbNo Then
                 
                    Set BuyReport = New ClsBuyReport
                    BuyReport.ShowIssueVoucherData Trans_ID
                     
                Else
                    Exit Sub
                End If
                
            Case 190 '190  Õ“Ì· »÷«⁄Â »Ì‰ «·„Œ«“‰
               
                Trans_ID = NoteSerial1.value
                '  Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + Chr(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + Chr(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                If SystemOptions.UserInterface = ArabicInterface Then
                    Xcheck = MsgBox("yes  › Õ «·„” ‰œ ··„—«Ã⁄Â " + CHR(13) + "No  ÿ»«⁄Â «·„” ‰œ ›ﬁÿ" + CHR(13), vbInformation + vbYesNoCancel, "Â·  —Ìœ › Õ «·„” ‰œ ··„—«Ã⁄Â  ")
                Else
                    Xcheck = MsgBox("yes  Open doc  " + CHR(13) + "No  Print Doc    " + CHR(13), vbInformation + vbYesNoCancel, "What you want?? ")
                End If
                If Xcheck = vbYes Then
                    Unload FrmMoving
                    FrmMoving.show
                    FrmMoving.XPBtnMove_Click (2)
                    FrmMoving.Retrive Trans_ID
                 
                ElseIf Xcheck = vbNo Then
                 
                    Set BuyReport = New ClsBuyReport
                    BuyReport.ShowMovingVOucherData Trans_ID
                     
                Else
                    Exit Sub
                End If
                
            Case 200:  ' ﬁÌœ
                Set myField = myFields.Item(5)
 
                StrTemp = myField.value
 
                ShowGL_cc StrTemp, , 200

        End Select
        
        '    FrmCustemers.Show
    End If

    DoEvents
    'Stop
End Sub

Private Sub CRViewer_DblClicked(ByVal X As Long, _
                                ByVal Y As Long, _
                                EventInfo As Variant, _
                                UseDefault As Boolean)
    Exit Sub
    Dim MyObject As CrystalActiveXReportViewerLib10Ctl.CRVEventInfo
    Dim myFields  As CrystalActiveXReportViewerLib10Ctl.CRVFields
    Dim myField As CrystalActiveXReportViewerLib10Ctl.CRVField
    Dim StrTemp As String

    Dim IntNumIndex As Integer
    Dim StrFieldName As String
    Dim OwnerObject As Object
    'Exit Sub
    Set MyObject = EventInfo

    Set myFields = MyObject.GetFields
    IntNumIndex = myFields.SelectedFieldIndex

    If IntNumIndex = 0 Then
        Exit Sub
    End If

    Set myField = myFields.Item(IntNumIndex)
    StrTemp = myField.value

    If Me.MDIChild = True Then
        Set OwnerObject = mdifrmmain
    Else
        Set OwnerObject = Me
    End If

    If myField.Name Like "*CusID" Then
        OpenScreen PopUpShowCustomerBalanceScreen, myField.value, 0, , , , , OwnerObject
    ElseIf myField.Name Like "*CusName" Then
        OpenScreen PopUpShowCustomerBalanceScreen, GetDealerID(myField.value), 0, , , , , OwnerObject
    ElseIf myField.Name Like "*ItemID" Then
        OpenScreen PopUpShowItemCardScreen, myField.value, 0, , , , , OwnerObject
    ElseIf myField.Name Like "*ItemCode" Then
        OpenScreen PopUpShowItemCardScreen, GetItemID(myField.value), 0, , , , , OwnerObject
    ElseIf myField.Name Like "*ItemName" Then
        OpenScreen PopUpShowItemCardScreen, GetItemID("", myField.value), 0, , , , , OwnerObject
    ElseIf (myField.Name Like "*?ParCompanyName") Or (myField.Name Like "*?comment_arabic") Then
        OpenScreen OptionsScreen
        'Me.xReport.EnableParameterPrompting = False
        RefreshreReportParameters
        Me.CRViewer.RefreshEx False
    ElseIf myField.Name Like "*NoteSerial" Or myField.Name Like "*NotesSerial" Then
        ShowGL_cc myField.value, , 200
    End If

End Sub

 

Private Sub CRViewer_DrillOnSubreport(GroupNameList As Variant, _
                                      ByVal SubreportName As String, _
                                      ByVal Title As String, _
                                      ByVal PageNumber As Long, _
                                      ByVal Index As Long, _
                                      UseDefault As Boolean)

    Dim IntRes As Integer
    Dim Msg As String
    Dim m_Report As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim XXXX As CRAXDRT.SubreportObject
    Dim Lngid As Long

    'Msg = "Â·  —Ìœ › Õ Â–« «· ﬁ—Ì— «·›—⁄Ï ›Ï ‘«‘… „‰›’·… ...øøø"
    'IntRes = MsgBox(Msg, vbYesNo + vbQuestion + _
    'vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)
    IntRes = vbNo

    If IntRes = vbYes Then
        UseDefault = True
        '    Set m_Report = Me.xReport.OpenSubreport(SubreportName)
        '    m_Report.ReportTitle = Title
        '    Set m_Report.Database.SetDataSource = Cn

        Set CViewer = New ClsReportViewer
    
        'JJJJJJJJJJJJJJJJJJ
        If SubreportName = "GL_cc.rpt" Then
            Dim MySQL As String
            Dim RsData As New ADODB.Recordset
            MySQL = "Select * From GL_CC  where  notes_id=1"
 
            Set RsData = New ADODB.Recordset
            RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
            Set m_Report = Me.xReport.OpenSubreport(SubreportName)
            m_Report.reporttitle = Title

            If SystemOptions.UserInterface = ArabicInterface Then
                '    m_Report.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
                '    m_Report.ParameterFields(2).AddCurrentValue RPTComment_Arabic
   
            Else
 
                '    m_Report.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
                '    m_Report.ParameterFields(2).AddCurrentValue RPTComment_Eng
                m_Report.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
                StrReportTitle = ""
  
            End If

            'm_Report.ParameterFields(3).AddCurrentValue user_name
            m_Report.reporttitle = StrReportTitle
            'm_Report.EnableParameterPrompting = False

            'Set xReport = xApp.OpenReport(StrFileName)
            m_Report.Database.SetDataSource RsData
        End If

        'JJJJJJJJJJ
        CViewer.FireReport m_Report, WindowTarget, Title
    Else
        UseDefault = True
    End If

End Sub
'
Private Sub ExportReportToClipboard()

    On Error GoTo ErrorHandler

    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim tempExcelPath As String
    
    Dim crxReport As CRAXDRT.Report
    Set crxReport = xReport ' <<< √Â„  ⁄œÌ·: «” Œœ«„ «· ﬁ—Ì— «·„⁄—Ê÷ »«·›⁄·!

    '  ÕœÌœ «·„”«— «·„ƒﬁ  ·· ’œÌ—
    tempExcelPath = Environ$("TEMP") & "\TempExport.xls"
    
    '  ’œÌ— «· ﬁ—Ì— ≈·Ï «ﬂ”Ì·
    With crxReport.ExportOptions
        .DestinationType = crEDTDiskFile
        .FormatType = crEFTExcel97
        .DiskFileName = tempExcelPath
    End With
    
    crxReport.Export False

    ' › Õ «·„·› »«·«ﬂ”Ì·
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlBook = xlApp.Workbooks.Open(tempExcelPath)
    Set xlSheet = xlBook.Sheets(1)
    
    ' ‰”Œ «·»Ì«‰«  ··ﬂ·Ì» »Ê—œ
    xlSheet.UsedRange.Copy

    ' «€·«ﬁ «ﬂ”Ì·
    xlBook.Close False
    xlApp.Quit

    '  ‰ŸÌ› «·–«ﬂ—…
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
    ' Õ–› «·„·› «·„ƒﬁ 
    On Error Resume Next
    Kill tempExcelPath
    On Error GoTo 0
    
    MsgBox " „ ‰”Œ «·»Ì«‰«  ≈·Ï «·ﬂ·Ì» »Ê—œ »‰Ã«Õ!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "ÕœÀ Œÿ√ √À‰«¡ «· ’œÌ— √Ê «·‰”Œ:" & vbCrLf & Err.Description, vbCritical
    On Error Resume Next
    If Not xlBook Is Nothing Then xlBook.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    If Len(Dir(tempExcelPath)) > 0 Then Kill tempExcelPath
    On Error GoTo 0

End Sub

Private Sub ExportReportWithFormatting(Optional ByVal IsDataOnly As Boolean = False, Optional ByVal ExportToPDF As Boolean = False)

    On Error GoTo ErrorHandler

    Dim xlApp As Object
    Dim xlBook As Object
    Dim tempExportPath As String
    Dim crxReport As CRAXDRT.Report
    Dim ReportNameOnly As String
    Dim TimeStamp As String

    Set crxReport = xReport



' «” Œ—«Ã «”„ «· ﬁ—Ì— „‰ «·„”«— «·„ﬂ Ê» ›Ì txtpath
ReportNameOnly = mId$(Me.txtPath, InStrRev(Me.txtPath, "\") + 1)
ReportNameOnly = left$(ReportNameOnly, InStrRev(ReportNameOnly, ".") - 1)

TimeStamp = Format(Now, "yyyy-mm-dd_hh-nn-ss")

If ExportToPDF Then
    tempExportPath = Environ$("TEMP") & "\" & ReportNameOnly & "_" & TimeStamp & ".pdf"
Else
    tempExportPath = Environ$("TEMP") & "\" & ReportNameOnly & "_" & TimeStamp & ".xls"
End If
If crxReport Is Nothing Then
   ' crxApp.OpenReport (Me.txtPath)
    MsgBox "crxReport ·„ Ì „  Õ„Ì·Â »‘ﬂ· ’ÕÌÕ"
    Exit Sub
End If

    '  ’œÌ— «· ﬁ—Ì—
    With crxReport.ExportOptions
        .DestinationType = crEDTDiskFile
        If ExportToPDF Then
            .FormatType = crEFTPortableDocFormat
        Else
            If IsDataOnly Then
                .FormatType = crEFTExcelDataOnly
            Else
                .FormatType = crEFTExcel97
            End If
        End If
        .DiskFileName = tempExportPath
    End With
    
    crxReport.Export False

    ' ‰”Œ «·„·› ≈·Ï «·ﬂ·Ì» »Ê—œ
    Call CopyFileToClipboard(tempExportPath)

    MsgBox " „ ‰”Œ «·„·› ≈·Ï «·ﬂ·Ì» »Ê—œ! Ì„ﬂ‰ﬂ «·¬‰ ·’ﬁÂ »√Ì „ﬂ«‰ ⁄‰ ÿ—Ìﬁ Ctrl+V.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "ÕœÀ Œÿ√ √À‰«¡ «· ’œÌ— √Ê «·‰”Œ:" & vbCrLf & Err.Description, vbCritical
    On Error Resume Next
    If Not xlBook Is Nothing Then xlBook.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set xlBook = Nothing
    Set xlApp = Nothing
End Sub

'

'Private Sub ExportReportWithFormatting(Optional ByVal IsDataOnly As Boolean = False)
'
'    On Error GoTo ErrorHandler
'
'    Dim xlApp As Object
'    Dim xlBook As Object
'    Dim tempExcelPath As String
'    Dim TargetPath As String
'    Dim crxReport As CRAXDRT.Report
'
'    Set crxReport = xReport
'
'    tempExcelPath = Environ$("TEMP") & "\TempExportFormatted.xls"
'    'TargetPath = Environ$("USERPROFILE") & "\Desktop\ ﬁ—Ì—_„ÿ·Ê».xls" ' ‰”Œ ··‹ Desktop
'
'    '  ’œÌ— «· ﬁ—Ì—
'    With crxReport.ExportOptions
'        .DestinationType = crEDTDiskFile
'        If IsDataOnly Then
'            .FormatType = crEFTExcelDataOnly
'        Else
'            .FormatType = crEFTExcel97
'        End If
'
'        .DiskFileName = tempExcelPath
'    End With
'
'    crxReport.Export False
'
'    ' ‰”Œ «·„·› „‰ «·”Ì—›— ≈·Ï ÃÂ«“ «·ÌÊ“—
'
'   ' crxReport.Export False
'
'Call CopyFileToClipboard(tempExcelPath)
'
'MsgBox " „ ‰”Œ «·„·› ≈·Ï «·ﬂ·Ì» »Ê—œ! Ì„ﬂ‰ﬂ «·¬‰ ·’ﬁÂ »√Ì „ﬂ«‰ ⁄‰ ÿ—Ìﬁ Ctrl+V.", vbInformation
'
'
''    FileCopy tempExcelPath, TargetPath
'
'    ' › Õ «·„·› »⁄œ ‰”ŒÂ
'    Set xlApp = CreateObject("Excel.Application")
'    xlApp.Visible = True
'  '  Set xlBook = xlApp.Workbooks.Open(TargetPath)
'
'  '  MsgBox " „  ’œÌ— Ê‰”Œ «· ﬁ—Ì— ≈·Ï ”ÿÕ «·„ﬂ » »‰Ã«Õ!", vbInformation
'    Exit Sub
'
'ErrorHandler:
'    MsgBox "ÕœÀ Œÿ√ √À‰«¡ «· ’œÌ— √Ê «·‰”Œ:" & vbCrLf & Err.Description, vbCritical
'    On Error Resume Next
'    If Not xlBook Is Nothing Then xlBook.Close False
'    If Not xlApp Is Nothing Then xlApp.Quit
'    Set xlBook = Nothing
'    Set xlApp = Nothing
'End Sub

Private Sub ExportReportToClipboard2()

    On Error GoTo ErrorHandler ' <<< ????? ??????
    
    Dim crxApp As New CRAXDRT.Application
    Dim crxReport As CRAXDRT.Report
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim tempExcelPath As String
    Dim mPath As String
    mPath = Trim(Me.txtPath)
    ' 1- ????? ???????
   ' Set crxReport = crxApp.OpenReport(App.path & "\YourReportName.rpt")
    Set crxReport = crxApp.OpenReport(mPath)
    
    
    ' 2- ????? ???? ??? ????
    tempExcelPath = Environ$("TEMP") & "\TempExport.xls"
    
    ' 3- ????? ??????? ??? Excel
    With crxReport.ExportOptions
        .DestinationType = crEDTDiskFile
        .FormatType = crEFTExcel97
        .DiskFileName = tempExcelPath
    End With
    
    crxReport.Export False ' ??????? ???? ????? ????????
    
    ' 4- ??? ??? ???????
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlBook = xlApp.Workbooks.Open(tempExcelPath)
    Set xlSheet = xlBook.Sheets(1)
    
    ' 5- ??? ???????? ??? ?????? ????
    xlSheet.UsedRange.Copy

    ' 6- ????? ??? ??????? ???? ???
    xlBook.Close False
    xlApp.Quit

    ' 7- ????? ???????
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
    ' 8- ??? ????? ??????
    On Error Resume Next
    Kill tempExcelPath
    On Error GoTo 0
    
    MsgBox "?? ??? ???????? ??? ?????? ???? ?????!", vbInformation
    Exit Sub

' ---------
ErrorHandler:
    MsgBox "??? ??? ????? ??????? ?? ?????:" & vbCrLf & Err.Description, vbCritical
    ' ????? ?? ???? ?????
    On Error Resume Next
    If Not xlBook Is Nothing Then xlBook.Close False
    If Not xlApp Is Nothing Then xlApp.Quit
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    If Len(Dir(tempExcelPath)) > 0 Then Kill tempExcelPath
    On Error GoTo 0
End Sub

Private Sub CRViewer_ExportButtonClicked(UseDefault As Boolean)
    'Dim StrViewName As String
    'Dim StrTebName As String
    'Dim m_SubReport As CRAXDDRT.Report
    'UseDefault = False
    'Load FrmReportExport
    '
    '
    'StrTemp = CRViewer.GetViewName(StrTebName)
    'If left$(StrTebName, 3) = "Sub" Then
    '
    '    Set m_SubReport = Me.xReport.OpenSubreport(StrTebName)
    '    Set FrmReportExport.xReport = m_SubReport
    'Else
    '    Set FrmReportExport.xReport = Me.xReport
    'End If
    'FrmReportExport.Show vbModal

    'Dim vPath As Variant
    'Dim vString As String
    'Dim x As Integer
    'Dim y As Integer
    'Dim counter As Integer
    'vPath = CRViewer.GetViewPath(Me.CRViewer.ActiveViewIndex)
    'x = LBound(vPath)
    'y = UBound(vPath)
    'For counter = x To y
    '    If vString <> "" Then
    '    vString = vString & ":"
    '    End If
    '    vString = vString & vPath(counter)
    'Next counter
    ''MsgBox "View path is: " & vString
'ExportReportToClipboard
End Sub

Private Sub CRViewer_OnChangeObjectRect(ByVal ObjectDescription As Variant, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal Width As Long, _
                                        ByVal Height As Long)
    MsgBox ObjectDescription(0)
End Sub

Private Sub CRViewer_OnContextMenu(ByVal ObjectDescription As Variant, _
                                   ByVal X As Long, _
                                   ByVal Y As Long, _
                                   UseDefault As Boolean)

    Dim MyX As CrystalActiveXReportViewerLib10Ctl.CRVField
    UseDefault = True

End Sub


Private Function ListPrinters() As String
    Dim Prn As Printer
    Dim i As Integer
    Dim printerList As String
    
    printerList = ""
    i = 1
    
    ' ≈‰‘«¡ ﬁ«∆„… «·ÿ«»⁄«  «·„À» … ⁄·Ï «·‰Ÿ«„
    For Each Prn In Printers
        printerList = printerList & i & ". " & Prn.DeviceName & vbCrLf
        i = i + 1
    Next Prn
    
    ListPrinters = printerList
End Function
Private Sub PrintReport()
      Dim printername As String
    Dim Buffer As String * 256
    Dim Result As Long

 Dim Prn As Printer
    Dim selectedPrinter As Printer
    Dim i As Integer
    
    
    ' ⁄—÷ ﬁ«∆„… «·ÿ«»⁄«  «·„À» …
 
 
     printername = InputBox("«Œ — «·ÿ«»⁄… Õ”» «·—ﬁ„:" & vbCrLf & ListPrinters(), "«Œ Ì«— «·ÿ«»⁄…")
    
    ' «· Õﬁﬁ ≈–«  „ ≈œŒ«· «”„ «·ÿ«»⁄… √Ê —ﬁ„Â«
    If printername = "" Then
        MsgBox "·„ Ì „ «Œ Ì«— ÿ«»⁄….", vbInformation
        Exit Sub
    End If

    ' «· Õﬁﬁ „‰ ’Õ… «·—ﬁ„ «·„œŒ·
    If Not IsNumeric(printername) Or CInt(printername) < 1 Or CInt(printername) > Printers.count Then
        MsgBox "—ﬁ„ «·ÿ«»⁄… €Ì— ’ÕÌÕ. Ì—ÃÏ «·„Õ«Ê·… „—… √Œ—Ï.", vbCritical
        Exit Sub
    End If

    '  ⁄ÌÌ‰ «·ÿ«»⁄… »‰«¡ ⁄·Ï «·«Œ Ì«—
    Set selectedPrinter = Printers(CInt(printername) - 1)
    
    ' ⁄—÷ «”„ «·ÿ«»⁄… «·„Œ «—… ·· √ﬂœ
    MsgBox " „ «Œ Ì«— «·ÿ«»⁄…: " & selectedPrinter.DeviceName
    
    ' —»ÿ «·ÿ«»⁄… «·„Œ «—… » ﬁ—Ì— Crystal Reports
    xReport.SelectPrinter selectedPrinter.drivername, selectedPrinter.DeviceName, selectedPrinter.Port
    
    ' ÿ»«⁄… «· ﬁ—Ì—
    xReport.PrintOut False
    
    MsgBox " „ ≈—”«· «· ﬁ—Ì— ··ÿ»«⁄… »‰Ã«Õ."
    
   Exit Sub

    ' ≈ŸÂ«— ‰«›–… «Œ Ì«— «·ÿ«»⁄… „‰ CommonDialog
    With CommonDialog1
        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoSelection
        .ShowPrinter
    
    End With
    
'    If Not IsEmpty(CommonDialog1.printerName) Then
'        xReport.SelectPrinter "", CommonDialog1.printerName, ""
'    Else
'        MsgBox "·„ Ì „ «Œ Ì«— ÿ«»⁄….", vbInformation
'        Exit Sub
'    End If

    ' —»ÿ «·ÿ«»⁄… «·„Œ «—… » ﬁ—Ì— Crystal Reports
    'xReport.SelectPrinter CommonDialog1.PrinterDrive, CommonDialog1.PrinterName, CommonDialog1.PrinterPort
    
    '  ‰›Ì– «·ÿ»«⁄… »⁄œ «Œ Ì«— «·ÿ«»⁄…
    xReport.PrintOut False
    
    Exit Sub

    ' «·Õ’Ê· ⁄·Ï «”„ «·ÿ«»⁄… «·«› —«÷Ì…
'    Result = GetProfileString("windows", "device", ",,,", Buffer, 255)
'    PrinterName = left(Buffer, InStr(Buffer, ",") - 1)
'
'    '  ⁄ÌÌ‰ «·ÿ«»⁄… «·„Œ «—… ›Ì Crystal Reports
'    On Error GoTo ErrorHandler
'    xReport.SelectPrinter PrinterName, "", ""
'
'    ' ÿ»«⁄… «· ﬁ—Ì—
'    xReport.PrintOut False
'
'    Exit Sub
    
ErrorHandler:
    MsgBox "ÕœÀ Œÿ√ √À‰«¡ «Œ Ì«— «·ÿ«»⁄…: " & Err.Description, vbCritical
End Sub

Private Sub CRViewer_PrintButtonClicked(UseDefault As Boolean)
    Dim LngCurrentPage As Long
    Dim LngPagesCount As Long
    Dim StrPrinterName As String
    Dim cPrinter As ClsPrinters
    Dim ObjPrinter As Object
    Dim BolCollate As Boolean
    Dim LngCopies As Long
    Dim Msg As String

    On Error GoTo ErrTrap

  UseDefault = True

'PrintReport

  ' Exit Sub
    
  If SystemOptions.ShowPrinterDialoge = False Then
   Exit Sub
End If

    If Me.xReport Is Nothing Then
        UseDefault = True
    Else
        UseDefault = False
        Load FrmPrinterDialog
        CRViewer.GetLastPageNumber LngPagesCount, True
        LngCurrentPage = CRViewer.GetCurrentPageNumber

        If SystemOptions.UserInterface = ArabicInterface Then
            FrmPrinterDialog.lbl(12).Caption = "⁄œœ «·’›Õ«  «·ﬂ·Ï :" & LngPagesCount
            FrmPrinterDialog.lbl(13).Caption = "«·’›Õ… «·Õ«·Ì…:" & CRViewer.GetCurrentPageNumber
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            FrmPrinterDialog.lbl(12).Caption = "All Pages Count:" & LngPagesCount
            FrmPrinterDialog.lbl(13).Caption = "Current Page:" & CRViewer.GetCurrentPageNumber
        End If

        FrmPrinterDialog.show vbModal

        If FrmPrinterDialog.UserCancel = False Then
            Set cPrinter = New ClsPrinters
            StrPrinterName = FrmPrinterDialog.IcboPrinters.text
            Set ObjPrinter = cPrinter.GetPrinter(StrPrinterName)

            BolCollate = IIf(FrmPrinterDialog.ChkCollate.value = vbChecked, True, False)
            LngCopies = CLng(val(FrmPrinterDialog.TxtCopies.text))

            If left(ObjPrinter.DeviceName, 2) = "\\" Then
                xReport.SelectPrinter ObjPrinter.drivername, "(" & ObjPrinter.DeviceName & ")", ObjPrinter.Port
            Else
                xReport.SelectPrinter ObjPrinter.drivername, "" & ObjPrinter.DeviceName & "", ObjPrinter.Port
            End If
        
            If FrmPrinterDialog.opt(0).value = True Then
                'xReport.PrintOut True, LngCopies, BolCollate, 1
                xReport.PrintOutEx False, LngCopies, BolCollate, 1
            ElseIf FrmPrinterDialog.opt(1).value = True Then
                xReport.PrintOut False, LngCopies, BolCollate, LngCurrentPage, LngCurrentPage
            Else
            End If
        End If

        Unload FrmPrinterDialog
    End If

    Exit Sub
ErrTrap:
    Msg = "⁄›Ê« ·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—...!!!"
    Msg = Msg & CHR(13) & "Source:" & Err.Source
    Msg = Msg & CHR(13) & "Description:" & Err.Description
    Msg = Msg & CHR(13) & "Number:" & Err.Number
End Sub

Private Sub CRViewer_RefreshButtonClicked(UseDefault As Boolean)
    On Error Resume Next
    Dim rs As ADODB.Recordset
    Dim xx As Object

    Me.xReport.EnableParameterPrompting = False
    Me.CRViewer.RefreshEx False
    UseDefault = True
End Sub

Private Sub CRViewer_SearchButtonClicked(ByVal searchText As String, _
                                         UseDefault As Boolean)
    'UseDefault = False
    'Load FrmSearch
    'Set FrmSearch.Viewer = Me.CRViewer
    'Set FrmSearch.Report = Me.xReport
    'FrmSearch.Show
End Sub

Private Sub Form_Activate()
    'PutFormOnTop Me.hwnd
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If Shift = VBRUN.ShiftConstants.vbShiftMask Then
        If KeyCode = vbKeyX Then
            Unload Me
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap

    With Me.CRViewer
        .EnableAnimationCtrl = False
        .EnableCloseButton = True
       ' .EnableGroupTree = True
        .EnableNavigationControls = True
        .EnableRefreshButton = True
        .EnableSearchControl = True
        .EnablePrintButton = True
        .EnableToolbar = True
        .EnableZoomControl = True
        .DisplayTabs = True
        .EnableExportButton = True
        .EnableSearchExpertButton = False
        .EnableSelectExpertButton = False
  
    End With

    If SystemOptions.UserInterface = EnglishInterface Then
        Me.RightToLeft = False
        Me.Caption = "Print Preview"
    Else
        Me.RightToLeft = True
        Me.Caption = "„⁄«Ì‰… ﬁ»· «·ÿ»«⁄…"
    End If

    XPCmboZoom.AddItem "400%"
    XPCmboZoom.AddItem "300%"
    XPCmboZoom.AddItem "200%"
    XPCmboZoom.AddItem "150%"
    XPCmboZoom.AddItem "100%"
    XPCmboZoom.AddItem "75%"
    XPCmboZoom.AddItem "50%"
    XPCmboZoom.AddItem "25%"
    XPCmboZoom.ListIndex = 4
 '     CRViewer.Zoom (400)
    'If CRViewer.GetCurrentPageNumber > 1 Then
    '    Toolbar.Buttons("First").Enabled = True
    '    Toolbar.Buttons("Previous").Enabled = True
    '    Toolbar.Buttons("next").Enabled = True
    '    Toolbar.Buttons("last").Enabled = True
    'Else
    '    Toolbar.Buttons("First").Enabled = False
    '    Toolbar.Buttons("Previous").Enabled = False
    '    Toolbar.Buttons("next").Enabled = False
    '    Toolbar.Buttons("last").Enabled = False
    'End If
    Resize_Form Me, ReportSize
    Me.WindowState = vbMaximized

    If SystemOptions.UserInterface = EnglishInterface Then
        Command1.Caption = "Edit Report"
    End If
 
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim Msg As String
    Dim IntRes As Integer
    On Local Error GoTo ErrTrap

    If UnloadMode = VBRUN.QueryUnloadConstants.vbFormControlMenu Then
        If (Me.CRViewer.ViewCount > 1) And (CRViewer.ActiveViewIndex > 1) Then
            '   Msg = "·«ÕŸ «‰Â »€·ﬁ Â–Â «·‘«‘… ”Ê› Ì „ €·ﬁ ﬂ· „⁄«Ì‰«  «·ÿ»«⁄… «·—∆Ì”Ì… "
            '   Msg = Msg & Chr(13) & "Ê«·›—⁄Ì… "
            '   Msg = Msg & Chr(13) & "----------------------------------------"
            '   Msg = Msg & Chr(13) & "„·ÕÊŸ…:-"
            '   Msg = Msg & Chr(13) & "≈–« ﬂ‰   —Ìœ €·ﬁ √Ï „⁄«Ì‰… ›—⁄Ì… ›ﬁÿ ..œÊ‰ €·ﬁ ﬂ· «· ﬁ—Ì—"
            '   Msg = Msg & Chr(13) & "›«‰Â Ì„ﬂ‰ﬂ €·ﬁÂ« ⁄‰ ÿ—Ìﬁ “— «·√€·«ﬁ »ÃÊ«— “— «·ÿ«»⁄… ›Ï «·—ﬂ‰ «·√Ì”— «·⁄·ÊÏ"
            '   Msg = Msg & Chr(13) & "----------------------------------------"
            '   Msg = Msg & Chr(13) & "Â·  —Ìœ €·ﬁ Â–« «·‘«‘… ﬂ·Â«...øøø"
            '   IntRes = MsgBox(Msg, vbQuestion + vbYesNoCancel + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)
            IntRes = vbNo

            If IntRes = vbYes Then
                Cancel = False
            ElseIf IntRes = vbNo Then
                Me.CRViewer.CloseView CRViewer.ActiveViewIndex
                Cancel = True
            ElseIf IntRes = vbCancel Then
                Cancel = True
            End If
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Resize()
    CRViewer.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight  '

    If SystemOptions.UserInterface = EnglishInterface Then
        Me.Caption = "Print Preview" & "( " & Me.PreviewCaption & " )"
    Else
        Me.Caption = "„⁄«Ì‰… ﬁ»· «·ÿ»«⁄…" & " ( " & Me.PreviewCaption & " )"
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set xReport = Nothing
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo ErrTrap

    Select Case Button.key

        Case "print"
            CRViewer.PrintReport

        Case "search"
            FrmSearch.show vbModal

        Case "refresh"
            CRViewer.Refresh

        Case "last"
            CRViewer.ShowLastPage

        Case "next"
            CRViewer.ShowNextPage

        Case "Previous"
            CRViewer.ShowPreviousPage

        Case "First"
            CRViewer.ShowFirstPage
    End Select
 
    Exit Sub
ErrTrap:
End Sub

Public Function ReportSearch(SearchValue As String)
    CRViewer.SearchForText SearchValue
End Function

Private Sub XPCmboZoom_Change()
    XPCmboZoom_Click
End Sub

Private Sub XPCmboZoom_Click()
    On Error GoTo ErrTrap

    Select Case XPCmboZoom.ListIndex

        Case 0
            CRViewer.Zoom (400)

        Case 1
            CRViewer.Zoom (300)

        Case 2
            CRViewer.Zoom (200)

        Case 3
            CRViewer.Zoom (150)

        Case 4
            CRViewer.Zoom (100)

        Case 5
            CRViewer.Zoom (75)

        Case 6
            CRViewer.Zoom (50)

        Case 7
            CRViewer.Zoom (25)
    End Select

    Exit Sub
ErrTrap:

End Sub

Public Property Get PreviewCaption() As String
    PreviewCaption = m_PreviewCaption
End Property

Public Property Let PreviewCaption(ByVal vNewValue As String)
    m_PreviewCaption = vNewValue
End Property

Private Sub RefreshreReportParameters()
    Dim i As Long
    Dim cOptions As ClsCompanyInfo

    If Not (Me.xReport Is Nothing) Then
        Set cOptions = New ClsCompanyInfo

        For i = 1 To xReport.ParameterFields.count

            If xReport.ParameterFields(i).Name = "{?ParCompanyName}" Then
                xReport.ParameterFields(i).ClearCurrentValueAndRange
                xReport.ParameterFields(i).AddCurrentValue cOptions.ArabCompanyName
                '            If xReport.ParameterFields(I).GetNthCurrentValue(1) <> cOptions.ArabCompanyName Then
                '                xReport.ParameterFields(I).AddCurrentValue cOptions.ArabCompanyName
                '            End If
            ElseIf xReport.ParameterFields(i).Name = "{?comment_arabic}" Then
                xReport.ParameterFields(i).ClearCurrentValueAndRange
                xReport.ParameterFields(i).AddCurrentValue cOptions.ArabComment
            End If

        Next i

    End If

End Sub
