Attribute VB_Name = "ModFgLib"
Option Explicit

Public Sub DrawFloodProgress(FG As VSFlex8UCtl.VSFlexGrid, _
                             IntFloodColIndex As Integer, _
                             Optional SngFloodColor As Single = &HC0&, _
                             Optional DblColMax As Double = 0, _
                             Optional IntValueColIndex As Integer = -1)

    'Arguments:-
    'FG ---                 The Grid Control
    'IntFloodColIndex------------·⁄„Êœ «·–Ï ÌÕ ÊÏ «·—”„ «·»Ì«‰Ï
    'SngFloodColor----------·Ê‰ «⁄„œ… «·—”„ «·»Ì«‰Ï
    'DblColMax -------------«·ﬁÌ„… «·ﬁ’ÊÏ «· Ï ”Ê› Ì „ ≈⁄ »«—Â« ›Ï «·—”„
    'IntValueColIndex ------«·⁄„Êœ «·–Ï ”Ê›  √Œ– ﬁÌ„… ›Ï «·—”„

    Dim i As Long
    Dim DblMaxValue As Double

    With FG
        .Redraw = flexRDNone
        .Cell(flexcpFloodColor, .FixedRows, IntFloodColIndex, .Rows - 1, IntFloodColIndex) = SngFloodColor

        If DblColMax = 0 Then
            DblMaxValue = .Aggregate(flexSTMax, .FixedRows, IntFloodColIndex, .Rows - 1, IntFloodColIndex)
        Else
            DblMaxValue = DblColMax
        End If

        If IntValueColIndex = -1 Then
            IntValueColIndex = IntFloodColIndex
        End If

        If DblMaxValue <> 0 Then

            For i = .FixedRows To .Rows - 1
                .Cell(flexcpFloodPercent, i, IntFloodColIndex) = 100 * val(.TextMatrix(i, IntValueColIndex)) / DblMaxValue
            Next i

        End If

        .Redraw = True
        .Refresh
    End With
   
End Sub

Public Function SetFgForNewRow(FG As VSFlex8UCtl.VSFlexGrid, _
                               IntColIndex As Long) As Long

    Dim i As Long
    Dim LngTempRow As Long

    With FG

        For i = .FixedRows - 1 To .Rows - 1

            If Trim$(.TextMatrix(i, IntColIndex)) = "" Then
                Exit For
            End If

        Next i

        If .FixedRows = .Rows Then
            .Rows = .Rows + 1
            LngTempRow = .Rows - 1
        ElseIf i - 1 = .Rows - 1 Then
            .Rows = .Rows + 1
            LngTempRow = .Rows - 1
        ElseIf i < .Rows - 1 Then
            LngTempRow = i
        Else
            .Rows = .Rows + 1
            LngTempRow = .Rows - 1
        End If

        SetFgForNewRow = LngTempRow
    End With

End Function

Public Sub LinkFgColWithDataCombo(FG As VSFlex8UCtl.VSFlexGrid, _
                                  IntColIndex As Long, _
                                  Dcombo As DataCombo)
    Dim rs As ADODB.Recordset
    Dim StrList As String

    Set rs = Dcombo.RowSource
'Dim s As String
'    s = "Select top 17000 *  from tblItems"
'    s = "SELECT ItemID,ItemName From  TblItems  Where   IsNull(IsArchive,0) <> 1 and IsNull(IsArchive ,0) <> 0 Order BY ItemName "
'
'     Set rs = New ADODB.Recordset
'     rs.Open s, Cn, adOpenStatic, adLockOptimistic
'
'
    If rs Is Nothing Then
        Exit Sub
    End If

    With FG
        StrList = .BuildComboList(rs, rs(1).Name, rs(0).Name)

        If FG.Editable <> flexEDNone Then

            'Fg is Editable
            If StrList <> "" Then
                StrList = "|" & StrList
            End If
        End If

        FG.ColComboList(IntColIndex) = StrList
    End With

End Sub

Public Function GetItemsInFg(FG As VSFlexGrid, _
                             IntColIndex As Long) As Long
    Dim i As Long
    Dim IntCount As Long

    With FG

        For i = .FixedRows To .Rows - 1

            If Trim(.TextMatrix(i, IntColIndex)) <> "" Then
                IntCount = IntCount + 1
            End If

        Next i

    End With

    GetItemsInFg = IntCount
End Function

Public Function GetNodeChildTotal(FG As VSFlex8UCtl.VSFlexGrid, _
                                  XNode As VSFlex8UCtl.VSFlexNode, _
                                  Optional IntOperation As VSFlex8UCtl.SubtotalSettings = flexSTCount, _
                                  Optional OperationalCol As Long, _
                                  Optional BolCalNodes As Boolean = False) As Double
    '-----------------------------------------------------------------------------
    '›«∆œ… Â–Â «·œ«·… ÂÊ «‰‰« ‰” Œœ„Â« ›Ï ⁄„·
    'Summation
    '»ÿ—Ìﬁ… ÌœÊÌ…
    'IntOperation=«·ÿ—Ìﬁ… «·„—«œ…
    'OperationalCol = «·⁄„Êœ «·„—«œ «·⁄„· ⁄·ÌÂ
    '-----------------------------------------------------------------------------
    Dim Y  As VSFlexNode, Z As VSFlexNode
    Dim IntCount As Long
    Dim i As Long
    Dim DblSum As Double

    Set Y = XNode.GetNode(flexNTNextSibling)
    Set Z = XNode.GetNode(flexNTLastSibling)

    '----------------------------------------------------------------
    If IntOperation = flexSTCount Then
        If Not Y Is Nothing Then

            'GetNodeChildTotal = Y.Row - (Xnode.Row + 1 + Xnode.Children)
            For i = XNode.Row To Y.Row

                If FG.IsSubtotal(i) = False Then
                    IntCount = IntCount + 1
                End If

            Next i

            GetNodeChildTotal = IntCount
        Else

            For i = XNode.Row To FG.Rows - 1

                If FG.IsSubtotal(i) = False Then
                    IntCount = IntCount + 1
                Else

                    If FG.RowOutlineLevel(i) < FG.RowOutlineLevel(XNode.Row) Then
                        Exit For
                    End If
                End If

            Next i

            GetNodeChildTotal = IntCount
        End If

    ElseIf IntOperation = flexSTSum Then
        DblSum = 0

        If Not Y Is Nothing Then

            'GetNodeChildTotal = Y.Row - (Xnode.Row + 1 + Xnode.Children)
            For i = XNode.Row To Y.Row

                If FG.IsSubtotal(i) = False Then
                    DblSum = DblSum + val(FG.TextMatrix(i, OperationalCol))
                End If

            Next i

            GetNodeChildTotal = DblSum
        Else

            For i = XNode.Row To FG.Rows - 1

                If FG.IsSubtotal(i) = False Then
                    DblSum = DblSum + val(FG.TextMatrix(i, OperationalCol))
                Else

                    If FG.RowOutlineLevel(i) < FG.RowOutlineLevel(XNode.Row) Then
                        Exit For
                    End If
                End If

            Next i

            GetNodeChildTotal = DblSum
        End If
    End If

End Function

Public Function GetFgCheckCount(FG As VSFlexGrid, _
                                IntColIndex As Long) As Long

    Dim i As Long
    Dim IntCount As Long

    With FG

        For i = .FixedRows To .Rows - 1

            If .Cell(flexcpChecked, i, IntColIndex) = flexChecked Then
                IntCount = IntCount + 1
            End If

        Next i

    End With

    GetFgCheckCount = IntCount
End Function

Public Function GetFgSortTitle(FG As VSFlexGrid, _
                               LngCol As Long, _
                               SortOrder As Integer) As String
    Dim StrTemp As String
    On Error GoTo ErrTrap
    StrTemp = ""

    With FG

        If SystemOptions.UserInterface = ArabicInterface Then
            StrTemp = " — Ì» «·»Ì«‰«  »‹ " & FG.TextMatrix(0, LngCol) & " - "

            If SortOrder = 1 Then
                StrTemp = StrTemp & " ’«⁄œÏ"
            ElseIf SortOrder = 2 Then
                StrTemp = StrTemp & " ‰«“·Ì"
            End If

        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            StrTemp = "Data Sorted By " & FG.TextMatrix(0, LngCol) & " - "

            If SortOrder = 1 Then
                StrTemp = StrTemp & "ASC"
            ElseIf SortOrder = 2 Then
                StrTemp = StrTemp & "DESC"
            End If
        End If

    End With

    GetFgSortTitle = StrTemp
    Exit Function
ErrTrap:
    GetFgSortTitle = ""
End Function

Public Sub ReSerialGrid(FG As VSFlexGrid, _
                        LngColIndex As Long)
    Dim i As Integer

    With FG

        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, LngColIndex) = i
        Next i

    End With

End Sub

Public Function ItemsInGrid(grd As VSFlexGrid, _
                            LngIDCol As Long) As Long
    Dim i As Long
    Dim BolTemp As Boolean
    On Error GoTo ErrTrap

    With grd

        If Trim(.TextMatrix(.FixedRows, LngIDCol)) = "" Then
            ItemsInGrid = -1
        Else
            ItemsInGrid = 1
        End If

    End With

    Exit Function
ErrTrap:
    ItemsInGrid = -1
End Function

