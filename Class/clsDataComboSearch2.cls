VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDCboSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Last Updated and Debuged in 2/9/2005 By Nour
Private WithEvents MDS_DCombo As MSDataListLib.DataCombo
Attribute MDS_DCombo.VB_VarHelpID = -1

Private WithEvents TxtBuddy As TextBox
Attribute TxtBuddy.VB_VarHelpID = -1

Private MDS_ADO As ADODB.Recordset  'Pointer to the ADO Data Control that the DataCombo is Linked to

Private MDS_RST As ADODB.Recordset        'Holds a clone of the recordset

Private MDS_FIELD As ADODB.Field          'Holds the field within the recordset that comprises the ListField of the DataCombo

Private MDS_STOP As Boolean         'This variable determined whether to start a new search or resume the last search

Private MDS_FIELD_NAME As String    'Holds the name of the field specified in MDS_FIELD

Private m_AllowWriting As Boolean

Private StrBuddyFieldName As String

Private Sub Class_Initialize()
    Set MDS_DCombo = Nothing
    Set MDS_RST = Nothing
    Set MDS_ADO = Nothing
    Set MDS_FIELD = Nothing
    MDS_FIELD_NAME = ""

End Sub

Private Sub Class_Terminate()
    Set MDS_DCombo = Nothing    'Free memory
    Set MDS_ADO = Nothing
    Set MDS_RST = Nothing
    Set MDS_FIELD = Nothing
End Sub

Public Property Set Client(DcbDCombo As MSDataListLib.DataCombo)

    If typename(DcbDCombo) = "DataCombo" Then   'Check that we are initialising an ADO Datacombo
        If DcbDCombo.Style <> dbcDropdownCombo Then
            Err.Raise Number:=vbObjectError + 1, Source:="clsDataCombo.Client", Description:="The Style Property Of The DataCombo " & CHR(13) & "Must be =0 <dbcDropdownCombo>"
        Else
            Set MDS_DCombo = DcbDCombo
            Set ClientRowsource = MDS_DCombo.RowSource
        End If

    Else
        Err.Raise Number:=vbObjectError + 1, Source:="clsDataCombo.Client", Description:="The object must be a datacombo."
    End If

End Property

Private Property Set ClientRowsource(AdoRS As ADODB.Recordset)
    Dim fld As ADODB.Field, blnExists As Boolean, strListfield

    If typename(AdoRS) = "Recordset" Then
        strListfield = MDS_DCombo.ListField
        Set MDS_ADO = AdoRS
        Set MDS_RST = New ADODB.Recordset
        MDS_RST.CursorLocation = adUseClient
        MDS_RST.CursorType = adOpenKeyset
        Set MDS_RST = AdoRS.Clone(adLockReadOnly)

        With MDS_RST
            blnExists = False

            For Each fld In .Fields

                If fld.Name = strListfield Then
                    MDS_FIELD_NAME = strListfield
                    blnExists = True
                    Exit For
                End If

            Next

        End With

        If Not blnExists Then
            Err.Raise Number:=vbObjectError + 3, Source:="clsDCboSearch.ClientListField", Description:="The ADO Data Control does not contain the field passed as strListField. The likely cause is that you have passed the incorrect ADODC. Ensure that the ADODC passed is the RowSource of the datacontrol."
        End If

    Else
        'err.Raise Number:=vbObjectError + 2, _
        'Source:="clsDataCombo.Client", _
        'Description:="The parameter AdoRS must point to an AdoRS."
    End If

End Property

Private Sub Highlight(ByVal lngLength As Long)

    With MDS_DCombo

        If Len(.Text) > lngLength Then
            .SelStart = lngLength
            .SelLength = Len(.Text) - lngLength
        End If

    End With

End Sub

Private Sub MDS_DCombo_Change()
On Error Resume Next
    If Not TxtBuddy Is Nothing Then
        If StrBuddyFieldName = "" Then
            TxtBuddy.Text = MDS_DCombo.BoundText
        Else
        End If
    End If

End Sub

Private Sub MDS_DCombo_Click(Area As Integer)
    MDS_DCombo_Change
End Sub

Private Sub MDS_DCombo_GotFocus()
    On Error GoTo MDS_DCombo_GotFocus_Err

    If Not MDS_ADO Is Nothing Then
        Set MDS_RST = MDS_ADO.Clone(adLockReadOnly)         'Refresh the recordset so that we have the most up-to-date data in the list
    End If

    If Not MDS_RST Is Nothing Then
        Set MDS_FIELD = MDS_RST.Fields(MDS_FIELD_NAME)
        MDS_STOP = True
    End If

    If MDS_DCombo.BoundText = "" Then
        'SendKeys "{F4}"
    End If

MDS_DCombo_GotFocus_Exit:
    Exit Sub
MDS_DCombo_GotFocus_Err:
    'Err.Raise Number:=vbObjectError + 1, _
     Source:="clsDataCombo.MDS_DCombo_GotFocus", _
     Description:="There was an error in MDS_DCombo_GotFocus. Likely cause is that the Class has not correctly been initialized."
    Resume MDS_DCombo_GotFocus_Exit
End Sub

Private Sub MDS_DCombo_KeyDown(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyDelete Then
        MDS_DCombo_KeyPress vbKeyBack
        '    If SystemOptions.SysRegisterState = DevelopVersion Then
        '        Stop
        '    End If
    End If

End Sub

Private Sub MDS_DCombo_KeyPress(KeyAscii As Integer)
    Dim StrCriteria As String
    Dim BookMk As Variant
    Dim Index As Long
    Dim DoSearch As Boolean
    Static Search As String

    On Error GoTo MDS_DCombo_KeyPress_Err

    If MDS_STOP Then Search = ""
    MDS_STOP = False

    Select Case KeyAscii

        Case 39
            KeyAscii = 0

        Case vbKeyBack

            If Len(Search) Then
                Search = left(Search, Len(Search) - 1)

                DoSearch = True
            End If

            KeyAscii = 0

        Case vbKeyReturn, vbKeyTab

            With MDS_DCombo
                .SelStart = Len(.Text)
                .SelLength = 0
            End With

            DoSearch = False
            KeyAscii = vbKeyReturn

        Case Is >= vbKeySpace
            Search = Search & UCase(CHR(KeyAscii))

            DoSearch = True
    End Select

    If (DoSearch = True) And Len(Search) > 0 Then
        If Not (MDS_RST Is Nothing) Then
            Set MDS_FIELD = MDS_RST.Fields(MDS_FIELD_NAME)
            StrCriteria = MDS_FIELD.Name & " like '" & Search & "%'"

            If Not (MDS_RST.BOF Or MDS_RST.EOF) Then

                With MDS_RST
                    BookMk = .Bookmark
                    .find StrCriteria, , , 1

                    If MDS_RST.EOF Then
                        'There is no match
                        .Bookmark = BookMk

                        If Me.AllowWriting = False Then
                            Search = left(Search, Len(Search) - 1)
                        End If

                        If Len(Search) = 0 Then
                            MDS_DCombo.Text = ""
                        End If

                    Else
                        MDS_DCombo.Text = MDS_FIELD
                    End If

                End With

            End If
        End If

        Call Highlight(Len(Search))
        KeyAscii = 0
    Else

        If Len(Search) Or KeyAscii = vbKeyReturn Then
            KeyAscii = KeyAscii
        Else
            MDS_DCombo.Text = "" 'IIf(Me.AllowWriting = True, MDS_DCombo.Text, "")
        End If
    End If

MDS_DCombo_KeyPress_Exit:
    Exit Sub
MDS_DCombo_KeyPress_Err:
    MsgBox Prompt:="Unexpected Error - " & Err.Number & " " & Err.Description, Buttons:=vbOKOnly + vbCritical, title:="clsDCboSearch.MDS_DCombo_KeyPress"
End Sub

Public Function Refresh()
    On Erro GoTo ErrTrap
    Set Me.Client = MDS_DCombo
    'MDS_RST.Requery
    Exit Function
ErrTrap:
End Function

Public Property Get AllowWriting() As Boolean
    AllowWriting = m_AllowWriting
End Property

Public Property Let AllowWriting(ByVal vNewValue As Boolean)
    m_AllowWriting = vNewValue
End Property

Private Sub TxtBuddy_GotFocus()
    SelectText TxtBuddy
End Sub

Private Sub TxtBuddy_KeyDown(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If StrBuddyFieldName <> "" Then
            'Do Code to Search
        Else
            MDS_DCombo.BoundText = val(TxtBuddy.Text)
        End If
    End If

End Sub

Public Sub SetBuddyText(Txt As TextBox, _
                        Optional StrTxtFieldName As String)
    Set TxtBuddy = Txt
    TxtBuddy.MaxLength = 4
    StrBuddyFieldName = StrTxtFieldName
End Sub

Private Sub TxtBuddy_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtBuddy.Text, 1)
End Sub
