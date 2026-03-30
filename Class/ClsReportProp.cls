VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsReportProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Dim M_SQLQuery As String
Dim M_Transaction_ID As String
Dim m_TransactionSerial As String
Dim M_Transaction_Date As String
Dim M_Customer_ID As String
Dim M_Customer_Name As String
Dim M_Store_ID As String
Dim M_Store_Name As String
Dim M_Item_Code As String
Dim M_Item_Name As String
Dim M_Item_Price As String
Dim M_Item_Serial As String
Dim M_Item_Count As String
Dim M_Item_Discount As String
Dim M_Item_AllPrice As String
Dim M_Invo_TotalPrice As String
Dim m_InvoID As Integer
Dim m_ItemGuarantee As String
Dim M_EmpCode As String
Dim m_EmpName As String

Public Enum Flds
    Transaction_ID
    Transaction_serial
    Transaction_Date
    Customer_ID
    Customer_Name
    Store_ID
    Store_Name
    item_code
    item_name
    Item_Price
    Item_Serial
    Item_Count
    Item_Discount
    Item_AllPrice
    Invo_TotalPrice
    emp_code
    emp_name
    Item_Gguarantee
End Enum
Dim m_Openfile As String
    
Dim LstObj As ListBox

Public Function SetFldName(FldName As Flds, _
                           FieldName As String)

    Select Case FldName

        Case Transaction_ID
            M_Transaction_ID = FieldName

        Case Transaction_Date
            M_Transaction_Date = FieldName

        Case Customer_ID
            M_Customer_ID = FieldName

        Case Customer_Name
            M_Customer_Name = FieldName

        Case Store_ID
            M_Store_ID = FieldName

        Case Store_Name
            M_Store_Name = FieldName

        Case item_code
            M_Item_Code = FieldName

        Case item_name
            M_Item_Name = FieldName

        Case Item_Price
            M_Item_Price = FieldName

        Case Item_Serial
            M_Item_Serial = FieldName

        Case Item_Count
            M_Item_Count = FieldName

        Case Item_Discount
            M_Item_Discount = FieldName

        Case Item_AllPrice
            M_Item_AllPrice = FieldName

        Case Invo_TotalPrice
            M_Invo_TotalPrice = FieldName

        Case Flds.emp_code
            M_EmpCode = FieldName

        Case Flds.emp_name
            m_EmpName = FieldName

        Case Flds.Item_Gguarantee
            m_ItemGuarantee = FieldName

        Case Flds.Transaction_serial
            m_TransactionSerial = FieldName
    End Select

End Function

Public Function GetFldName(FldName As Flds) As String

    Select Case FldName

        Case Transaction_ID
            GetFldName = M_Transaction_ID

        Case Transaction_Date
            GetFldName = M_Transaction_Date

        Case Customer_ID
            GetFldName = M_Customer_ID

        Case Customer_Name
            GetFldName = M_Customer_Name

        Case Store_ID
            GetFldName = M_Store_ID

        Case Store_Name
            GetFldName = M_Store_Name

        Case item_code
            GetFldName = M_Item_Code

        Case item_name
            GetFldName = M_Item_Name

        Case Item_Price
            GetFldName = M_Item_Price

        Case Item_Serial
            GetFldName = M_Item_Serial

        Case Item_Count
            GetFldName = M_Item_Count

        Case Item_Discount
            GetFldName = M_Item_Discount

        Case Item_AllPrice
            GetFldName = M_Item_AllPrice

        Case Invo_TotalPrice
            GetFldName = M_Invo_TotalPrice

        Case Flds.Transaction_serial
            GetFldName = m_TransactionSerial

        Case Flds.emp_code
            GetFldName = M_EmpCode

        Case Flds.emp_name
            GetFldName = m_EmpName

        Case Flds.Item_Gguarantee
            GetFldName = m_ItemGuarantee
    End Select

End Function

Public Function SQLQuery(InvoID As Integer) As String
    '«·ﬂÊÌ—Ì «·„”∆Ê·Â ⁄‰ ⁄—÷ «·»Ì«‰«  ›Ì «·»—‰«„Ã  ÊÌ‘ —ÿ ›ÌÂ« «‰  Õ ÊÌ _
     ⁄·Ì «·ÕﬁÊ· «·«”«”Ì… ﬂ„« Ì‘ —ÿ ›ÌÂ« «‰  Õ ÊÌ ⁄·Ì «Ã„«·Ì «·›« Ê—… ﬂ„« Ì‘ —ÿ ›ÌÂ« «‰  Õ ÊÏ
    ' ﬂ„« Ì‘ —ÿ ›ÌÂ« «‰  Õ ÊÏ ⁄·Ì ﬁÌ„… «·’‰› »⁄œ «·Œ’„ ﬂ„« ÂÊ „»Ì‰ ›Ï «·«‰ÌÊ„Ì—« 

    SQLQuery = "select * From QryPrint where Transaction_ID=" & Me.InvoID
    '---------------------------- determine fields name -----------------
    Me.SetFldName Customer_ID, "CusID"
    Me.SetFldName Customer_Name, "CusName"
    Me.SetFldName Invo_TotalPrice, "Inv_Value"
    Me.SetFldName Item_AllPrice, "Total"
    Me.SetFldName item_code, "ItemCode"
    Me.SetFldName Item_Count, "Quantity"
    Me.SetFldName Item_Discount, "ItemDiscount"
    Me.SetFldName item_name, "ItemName"
    Me.SetFldName Item_Price, "Price"
    Me.SetFldName Item_Serial, "ItemSerial"
    Me.SetFldName Store_ID, "StoreID"
    Me.SetFldName Store_Name, "StoreName"
    Me.SetFldName Transaction_Date, "Transaction_Date"
    Me.SetFldName Transaction_ID, "Transaction_ID"
    Me.SetFldName emp_code, "Emp_Code"
    Me.SetFldName emp_name, "Emp_Name"
    Me.SetFldName Item_Gguarantee, "guaranteeTime"
    Me.SetFldName Transaction_serial, "Transaction_Serial"

End Function

Private Sub Class_Initialize()
    'SQLQuery 0
End Sub

Public Function ShowReport()
    Dim My_SQL As String
    Dim RsFiled As ADODB.Recordset
    Dim Ctrl As Control
    Dim IX As Integer
    Dim TxtTag As String
    Dim II As Long

    With FrmPreview
        Set RsFiled = New ADODB.Recordset
        My_SQL = Me.SQLQuery(Me.InvoID)
        RsFiled.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic

        If RsFiled.RecordCount > 0 Then
            .Grd(0).Rows = RsFiled.RecordCount + 1
            RsFiled.MoveFirst

            Do Until RsFiled.EOF
                II = II + 1

                For IX = 0 To .Grd(0).Cols - 1

                    If Trim(.Grd(0).ColKey(IX)) <> "" Then
                        If Trim(.Grd(0).ColKey(IX)) = "Record_Number" Then
                            .Grd(0).TextMatrix(II, IX) = II
                        Else
                            .Grd(0).TextMatrix(II, IX) = IIf(IsNull(RsFiled.Fields(.Grd(0).ColKey(IX)).value), "", RsFiled.Fields(.Grd(0).ColKey(IX)).value)
                        End If
                    End If

                Next
            
                RsFiled.MoveNext
            Loop

            RsFiled.MoveFirst

            For IX = 0 To .LblText.count - 1

                If Trim(.LblText(IX).Tag) <> "" Then
                    If .LblText(IX).Tag = "PrintDate" Then
                        .LblText(IX).Caption = Date
                    ElseIf .LblText(IX).Tag = "PrintTime" Then
                        .LblText(IX).Caption = Time
                    Else

                        Select Case Trim(.LblText(IX).Tag)

                            Case "PrintDate"
                                .LblText(IX).Caption = Format(Date, "yyyy/mm/dd")

                            Case "PrintTime"
                                .LblText(IX).Caption = Time

                            Case "PageNum", "PageOfTotal", "PageTotal"
                                GoTo EXT

                            Case Else
                                .LblText(IX).Caption = IIf(IsNull(RsFiled.Fields(Trim(.LblText(IX).Tag))), "", RsFiled.Fields(Trim(.LblText(IX).Tag)))
                        End Select

                    End If
                End If

EXT:
            Next

        End If
    
        '.Show vbModal
    End With

    Exit Function
FillOthersCtrl:

End Function

Public Function PrintBill()
End Function

Public Sub LoadFile(FilePath As String, _
                    Frm As Form)
    'load file From disk
    Dim Msg As String
    Dim FPath As String
    FPath = Dir(FilePath, vbNormal)

    If FPath = "" Then
        Msg = "⁄›Ê «·„·› €Ì— „ÊÃÊœ "
        MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Set LstObj = Frm.LstCtrl
    Dim FrFile As Integer
    Dim II As Integer
    Dim TxtSpelt() As String
    Dim Ctrl As Control
    'LoadFile = False
    FrFile = FreeFile
    II = 0
    LstObj.Clear
    Open FilePath For Random As #FrFile Len = Len(FSave)
    Get #FrFile, 1, FSave
    TxtSpelt = Split(Trim(FSave.PageSize), "~")

    With Frm

        Do Until EOF(FrFile)
            II = II + 1
            Get #FrFile, II, FSave
            LstObj.AddItem FSave.CtlProp
            On Error Resume Next
            .PicMain.Move 0, 330, val(TxtSpelt(0)), val(TxtSpelt(1))
            PicWidth = val(TxtSpelt(0))
            PicHeight = val(TxtSpelt(1))
        Loop

        Close #FrFile
    
        For II = 0 To LstObj.ListCount - 1

            If LstObj.List(II) <> "" Then
                Debug.Print LstObj.List(II)
                TxtSpelt = Split(LstObj.List(II), "**")

                If Trim(TxtSpelt(0)) <> "" Then
                    
                    If TxtSpelt(0) = "Label" Then
                        Load .LblText(.LblText.count)
                        Set Ctrl = .LblText(.LblText.count - 1)
                        SetCtrlProperty TxtSpelt(1), .LblText(.LblText.count - 1)
                    ElseIf TxtSpelt(0) = "Image" Then
                        Load .CoLog(.CoLog.count)
                        Debug.Print .CoLog.count
                        Set Ctrl = .CoLog(.CoLog.count - 1)
                        SetCtrlProperty TxtSpelt(1), Ctrl
                        
                        GoSub LoadImage
                    ElseIf TxtSpelt(0) = "Grd" Then
                        SetCtrlProperty TxtSpelt(1), .Grd(0)
                        Set Ctrl = .Grd(0)
                    End If

                    On Error Resume Next
                    '                .CoLog(0) = LoadPicture("")
                    '                .CoLog(0) = LoadPicture(Replace(FilePath, ".drp", ".img"))
                End If

                Ctrl.ZOrder 0
            End If

        Next

    End With

    Exit Sub
    '---------------------------------- Load Images to Objects --------------------------
LoadImage:
    Dim B_Read() As Byte
    Dim OutFile As String
    OutFile = Replace(FilePath, ".drp", ".dmg")
    Open OutFile For Binary Access Read As #1
    ReDim B_Read(0 To LOF(1) - 1)
    Debug.Print Ctrl.Tag
    Get #1, val(Ctrl.Tag), B_Read
    Debug.Print Seek(1)
    Close #1
    'Set Me.Picture1.Picture = Nothing
    Set Ctrl.Picture = PictureFromByteStream(B_Read)
    Return
End Sub

Public Property Get InvoID() As String
    InvoID = m_InvoID
End Property

Public Property Let InvoID(ByVal vNewValue As String)
    m_InvoID = vNewValue

End Property

Public Property Get OpenFile() As String
    OpenFile = m_Openfile
End Property

Public Property Let OpenFile(ByVal vNewValue As String)
    m_Openfile = vNewValue
End Property

