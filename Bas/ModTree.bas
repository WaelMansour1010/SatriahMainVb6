Attribute VB_Name = "ModTree"

Public Sub Make_RightToLeft(RTL_Tree As MSComctlLib.TreeView, _
                            Optional LTR As Boolean = False)

    Dim rClientRect As RECT
    Dim ReturnStyle As Long

    If LTR = False Then
        ReturnStyle = GetWindowLong(RTL_Tree.hWnd, GWL_EXSTYLE)
        SetWindowLong RTL_Tree.hWnd, GWL_EXSTYLE, ReturnStyle Or WS_EX_LAYOUTRTL
        GetClientRect RTL_Tree.hWnd, rClientRect
        InvalidateRect RTL_Tree.hWnd, rClientRect, True
    ElseIf LTR = True Then
        ReturnStyle = GetWindowLong(RTL_Tree.hWnd, GWL_EXSTYLE)
        SetWindowLong RTL_Tree.hWnd, GWL_EXSTYLE, 0
        GetClientRect RTL_Tree.hWnd, rClientRect
        InvalidateRect RTL_Tree.hWnd, rClientRect, False
    End If

End Sub

Public Sub fill_my_Children(parent_str As String, _
                            children_rs As ADODB.Recordset, _
                            MyTree As TreeView, _
                            table_name As String, _
                            parnet_field_name As String, _
                            Optional where_condtion, _
                            Optional ColName As Integer = 1, _
                            Optional StrSortBy As String = "")
    
    Dim nodX       As Node
    Dim par_rs     As New ADODB.Recordset
    Dim My_SQL     As String
    Dim StrCaption As String
    Dim i          As Long

    For i = 1 To children_rs.RecordCount
        Set nodX = MyTree.Nodes.Add(parent_str, 4, children_rs(0) & "G", IIf(IsNull(children_rs("Fullcode")), "", children_rs("Fullcode")) & " " & IIf(IsNull(children_rs(ColName)), "", children_rs(ColName)), "Closed_Node", "Open_Node")
        nodX.ExpandedImage = "Open_Node"
        nodX.Tag = "CATNode"
        My_SQL = " SELECT * "
        My_SQL = My_SQL + "  From " & table_name
        '***************
        'Add hidden node to show
'        Dim nodxtmp  As Node
'        Set nodxtmp = MyTree.Nodes.Add(parent_str, 4, "TmpNode" & children_rs(0) & "G", IIf(IsNull(children_rs("Fullcode")), "", children_rs("Fullcode")) & " " & IIf(IsNull(children_rs(ColName)), "", children_rs(ColName)), "Closed_Node", "Open_Node")
'        nodxtmp.ExpandedImage = "Open_Node"
       ' nodxtmp.Tag = "CATNode"
        
        'to expand Node
        '*************
        If IsMissing(where_condtion) Then
            My_SQL = My_SQL + " where " & parnet_field_name & "=" & children_rs(0) & "; "
        Else
            My_SQL = My_SQL + " where " & parnet_field_name & "=" & children_rs(0) & " and " & where_condtion
        End If

        If StrSortBy <> "" Then
            My_SQL = My_SQL + " Order By " & StrSortBy
        End If

        Set par_rs = Nothing
        par_rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        Call fill_my_Children(children_rs(0).value & "G", par_rs, MyTree, table_name, parnet_field_name, where_condtion, ColName)
        children_rs.MoveNext
    Next i

    Set par_rs = Nothing
End Sub
Private Sub fill_my_Children_Account(parent_str As String, _
                                     children_rs As ADODB.Recordset, _
                                     MyTree As TreeView, _
                                     table_name As String, _
                                     parnet_field_name As String, _
                                     Optional where_condtion, _
                                     Optional ColName As Integer = 1, _
                                     Optional StrSortBy As String = "")
    
    Dim nodX As Node
    Dim par_rs As New ADODB.Recordset
    Dim My_SQL As String
    Dim StrCaption As String
    Dim i As Long

    For i = 1 To children_rs.RecordCount
        Set nodX = MyTree.Nodes.Add(parent_str, 4, children_rs(1) & "G", IIf(IsNull(children_rs(ColName)), "", children_rs(ColName)), "Closed_Node", "Open_Node")
        nodX.ExpandedImage = "Open_Node"
        My_SQL = " SELECT * "
        My_SQL = My_SQL + "  From " & table_name

        If IsMissing(where_condtion) Then
        
            My_SQL = My_SQL + " where " & parnet_field_name & "='" & children_rs(1) & "'; "
        Else
            My_SQL = My_SQL + " where " & parnet_field_name & "='" & children_rs(1) & "' and " & where_condtion
        End If
        
     My_SQL = My_SQL + GetAccountByBarnchUser
     My_SQL = My_SQL + GetAccountCodeHiding
 
    
        If StrSortBy <> "" Then
            My_SQL = My_SQL + " Order By " & StrSortBy
        End If

        Set par_rs = Nothing
        par_rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        Call fill_my_Children_Account(children_rs(1).value & "G", par_rs, MyTree, table_name, parnet_field_name, where_condtion, ColName, "Account_Serial")
        children_rs.MoveNext
    Next i

    Set par_rs = Nothing
End Sub


Private Sub fill_my_Children_Accountx12102017(parent_str As String, _
                                     children_rs As ADODB.Recordset, _
                                     MyTree As TreeView, _
                                     table_name As String, _
                                     parnet_field_name As String, _
                                     Optional where_condtion, _
                                     Optional ColName As Integer = 1, _
                                     Optional StrSortBy As String = "")
    
    Dim nodX As Node
    Dim par_rs As New ADODB.Recordset
    Dim My_SQL As String
    Dim StrCaption As String
    Dim i As Long

    For i = 1 To children_rs.RecordCount
        Set nodX = MyTree.Nodes.Add(parent_str, 4, children_rs(1) & "G", IIf(IsNull(children_rs(ColName)), "", children_rs(ColName)), "Closed_Node", "Open_Node")
        nodX.ExpandedImage = "Open_Node"
        My_SQL = " SELECT * "
        My_SQL = My_SQL + "  From " & table_name

        If IsMissing(where_condtion) Then
            My_SQL = My_SQL + " where " & parnet_field_name & "='" & children_rs(1) & "'; "
        Else
            My_SQL = My_SQL + " where " & parnet_field_name & "='" & children_rs(1) & "' and " & where_condtion
        End If
        'salah
     My_SQL = My_SQL + GetAccountByBarnchUser
    
    
        If StrSortBy <> "" Then
            My_SQL = My_SQL + " Order By " & StrSortBy
        End If

        Set par_rs = Nothing
        par_rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        Call fill_my_Children_Account(children_rs(1).value & "G", par_rs, MyTree, table_name, parnet_field_name, where_condtion, ColName, "Account_Serial")
        children_rs.MoveNext
    Next i

    Set par_rs = Nothing
End Sub

Private Sub fill_my_Children_Accountx17072017(parent_str As String, _
                                     children_rs As ADODB.Recordset, _
                                     MyTree As TreeView, _
                                     table_name As String, _
                                     parnet_field_name As String, _
                                     Optional where_condtion, _
                                     Optional ColName As Integer = 1, _
                                     Optional StrSortBy As String = "")
    
    Dim nodX As Node
    Dim par_rs As New ADODB.Recordset
    Dim My_SQL As String
    Dim StrCaption As String
    Dim i As Long

    For i = 1 To children_rs.RecordCount
        Set nodX = MyTree.Nodes.Add(parent_str, 4, children_rs(1) & "G", IIf(IsNull(children_rs(ColName)), "", children_rs(ColName)), "Closed_Node", "Open_Node")
        nodX.ExpandedImage = "Open_Node"
        My_SQL = " SELECT * "
        My_SQL = My_SQL + "  From " & table_name

        If IsMissing(where_condtion) Then
            My_SQL = My_SQL + " where " & parnet_field_name & "='" & children_rs(1) & "'; "
        Else
            My_SQL = My_SQL + " where " & parnet_field_name & "='" & children_rs(1) & "' and " & where_condtion
        End If
    
        '    My_SQL = My_SQL + "  and (Branch='0' or Branch is null  or  Branch " & WhereViewString & " )"
    
        If StrSortBy <> "" Then
            My_SQL = My_SQL + " Order By " & StrSortBy
        End If

        Set par_rs = Nothing
        par_rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        Call fill_my_Children_Account(children_rs(1).value & "G", par_rs, MyTree, table_name, parnet_field_name, where_condtion, ColName, "Account_Serial")
        children_rs.MoveNext
    Next i

    Set par_rs = Nothing
End Sub
'
'Public Sub LoadTreeItems(ItemsTree As MSComctlLib.TreeView)
'
'    Dim Rs_items         As New ADODB.Recordset
'    Dim My_SQL           As String
'    Dim nodX             As Node
'    Dim nodz             As Node
'    Dim RsOptions        As New ADODB.Recordset
'    Dim my_ch_rs         As New ADODB.Recordset
'    Dim BolDisplayArabic As Boolean
'    Dim LngLoop          As Long
'    Dim Msg              As String
'    Screen.MousePointer = vbHourglass
'    On Error GoTo ErrTrap
'    '=================================E???? O??E C????C?
'    RsOptions.Open "Select OPTIONS.DisplayItemsArabic  From Options", Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    '--------------------
'    If RsOptions("DisplayItemsArabic").value = True Then
'        BolDisplayArabic = True
'        ItemsTree.Tag = "A"
'        Make_RightToLeft ItemsTree
'        '''''''''''''''''''''''''''add root
'        Set nodX = ItemsTree.Nodes.Add(, , "1G", " ?????E C????C? ", "Root")
'        ItemsTree.Nodes("1G").Expanded = True
'    Else
'        BolDisplayArabic = False
'        '''''''''''''''''''''''''''add root
'        ItemsTree.Tag = "E"
'        Set nodX = ItemsTree.Nodes.Add(, , "1G", "Groups of items", "Root")
'        ItemsTree.Nodes("1G").Expanded = True
'    End If
'
'    RsOptions.Close
'     Set RsOptions = Nothing
'    'ItemsTree.Sorted = False
'    '''''''''''''''''''''''''''' add group
'    'My_SQL = "SELECT GROUPS_OF_ITEMS.Group_Code," & _
'     "GROUPS_OF_ITEMS.Group_Name_Eng, GROUPS_OF_ITEMS.Parent_Group_Code" & _
'     " From GROUPS_OF_ITEMS "
'    'My_SQL = My_SQL + " where ([Parent_Group_Code] = 'r'); "
''    My_SQL = " SELECT GROUPS_OF_ITEMS.* "
''    My_SQL = My_SQL + "  From GROUPS_OF_ITEMS "
''    My_SQL = My_SQL + " where ([Parent_Group_Code] = 'r'); "
''    my_ch_rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly
''
''    If BolDisplayArabic = True Then
''        Call fill_my_Children("1G", my_ch_rs, ItemsTree, "GROUPS_OF_ITEMS", "Parent_Group_Code")
''    Else
''        Call fill_my_Children("1G", my_ch_rs, ItemsTree, "GROUPS_OF_ITEMS", "Parent_Group_Code", , 2)
''    End If
'
'    '''''''''''''''''''''''''''adds items
'    My_SQL = "SELECT Group_Code,Item_ID,Item_Name,Item_Name_Eng " & "FROM ITEMS ORDER BY ITEMS.Item_Name "
'    Rs_items.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If BolDisplayArabic = True Then
'
'        For LngLoop = 0 To Rs_items.RecordCount - 1
'            Set nodz = ItemsTree.Nodes.Add(Trim(Rs_items("Group_Code")), tvwChild, Rs_items("Item_ID"), Rs_items("Item_Name"), "Item")
'            nodz.Tag = "item"
'            Rs_items.MoveNext
'        Next LngLoop
'
'    Else
'
'        For LngLoop = 0 To Rs_items.RecordCount - 1
'            Set nodz = ItemsTree.Nodes.Add(Trim(Rs_items("Group_Code")), tvwChild, Rs_items("Item_ID"), IIf(IsNull(Rs_items("Item_Name_Eng")), "", Rs_items("Item_Name_Eng")), "Item")
'
'            If Trim(nodz.Text) = "" Then
'                nodz.ForeColor = vbRed
'                nodz.Bold = True
'                nodz.Text = "No Item Name"
'            End If
'
'            nodz.Tag = "item"
'            Rs_items.MoveNext
'        Next LngLoop
'
'    End If
'
'    ItemsTree.Nodes("1G").EnsureVisible
'    Rs_items.Close
'    Set Rs_items = Nothing
'    ItemsTree.Refresh
'    Screen.MousePointer = vbDefault
'    Exit Sub
'ErrTrap:
'    Screen.MousePointer = vbDefault
'    Msg = "ÕœÀ Œÿ« √À‰«¡  Õ„Ì· ‘Ã—… «·√’‰«›"
'    Msg = Msg & CHR(13) & Err.Description
'    Msg = Msg & CHR(13) & Err.Number
'    Msg = Msg & CHR(13) & Err.Source
'    Msg = Msg & CHR(13) & "ModTree:LoadTreeItems"
'    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'End Sub

Public Sub LoadGridTree(parent_str As String, _
                        children_rs As ADODB.Recordset, _
                        myGrid As VSFlexGrid, _
                        table_name As String, _
                        parnet_field_name As String, _
                        Optional where_condtion As String = "", _
                        Optional NoField_Connection As Boolean = False, _
                        Optional IntColName As Integer = 1, _
                        Optional SngNodeColor As Single = 0, _
                        Optional StrOrderBy As String = "")
    
    'By Nour 26/5/2003
    'This sub used to Fill VsFlexGrid as OutLine Tree
    'This Function is very siimialr to that "Load Tree"
    Dim i As Integer
    Dim par_rs As New ADODB.Recordset
    Dim My_SQL As String
    Dim LngParentOutLineLevel As Long
    Dim NodeX As VSFlexNode

    For i = 1 To children_rs.RecordCount

        'Set nodX = MyTree.Nodes.Add(parent_str, 4, children_rs(0), children_rs(1), 2, 3)
        With myGrid
            Debug.Print table_name
            LngParentOutLineLevel = .FindRow(parent_str, .FixedRows, , True, True)
            .AddItem CStr(children_rs(IntColName).value) 'Add the name of the account
            'add the account code in the row data
            .Rowdata(.Rows - 1) = CStr(children_rs(0).value) & "G"
            .RowOutlineLevel(.Rows - 1) = .RowOutlineLevel(LngParentOutLineLevel) + 1
            'set this row as node
            .IsSubtotal(.Rows - 1) = True
            .Cell(flexcpFontBold, .Rows - 1, 0) = True

            If SngNodeColor <> 0 Then
                .Cell(flexcpForeColor, .Rows - 1, 0) = SngNodeColor
            End If
        
            '        If NoField_Connection = False Then
            '
            '            If Not (children_rs("last_account").Value) Then
            '                .IsSubtotal(.Rows - 1) = Not (children_rs("last_account").Value)
            '                .Cell(flexcpPictureAlignment, .Rows - 1, 0) = flexPicAlignRightCenter
            '                .Cell(flexcpFontBold, .Rows - 1, 0) = Not (children_rs("last_account").Value)
            '                If SngNodeColor <> 0 Then
            '                    .Cell(flexcpForeColor, .Rows - 1, 0) = SngNodeColor
            '                End If
            '            End If
            '        Else
            '            .IsSubtotal(.Rows - 1) = True
            '            .Cell(flexcpPictureAlignment, .Rows - 1, 0) = flexPicAlignRightCenter
            '            .Cell(flexcpFontBold, .Rows - 1, 0) = True
            '            If SngNodeColor <> 0 Then
            '                .Cell(flexcpForeColor, .Rows - 1, 0) = SngNodeColor
            '            End If
            '        End If
        End With

        My_SQL = " SELECT * "
        My_SQL = My_SQL + "  From " & table_name

        If where_condtion = "" Then
            My_SQL = My_SQL + " where " & parnet_field_name & "=" & children_rs(0).value & "; "
        Else
            My_SQL = My_SQL + " where " & parnet_field_name & "=" & children_rs(0).value & " and " & where_condtion
        End If

        If StrOrderBy <> "" Then
            My_SQL = My_SQL + " ORDER By " & StrOrderBy
        End If

        Set par_rs = Nothing
        Set par_rs = New ADODB.Recordset
        par_rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (par_rs.EOF Or par_rs.BOF) Then
            Call LoadGridTree(children_rs(0).value & "G", par_rs, myGrid, table_name, parnet_field_name, where_condtion, NoField_Connection, IntColName, SngNodeColor, StrOrderBy)
        End If

        children_rs.MoveNext
    Next i

    Set par_rs = Nothing
End Sub

Public Sub LoadTreeGroups(ItemsTree As MSComctlLib.TreeView, _
                          Optional StrGroupSort As String = "", _
                          Optional StrItemsSort As String = "", _
                          Optional BolWithFullOption As Boolean = False)

    Dim Rs_items         As ADODB.Recordset
    Dim My_SQL           As String
    Dim nodX             As Node
    Dim nodz             As Node
    Dim RsOptions        As ADODB.Recordset
    Dim my_ch_rs         As ADODB.Recordset
    Dim BolDisplayArabic As Boolean
    ' Dim LngLoop          As Long

    '    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = ArabicInterface Then
        BolDisplayArabic = True
        ItemsTree.Tag = "A"
        Make_RightToLeft ItemsTree
        '''''''''''''''''''''''''''add root
        Set nodX = ItemsTree.Nodes.Add(, , "r", "  „Ã„Ê⁄«  «·«’‰«› ", "Root")
        ItemsTree.Nodes("r").Expanded = True
    Else
        BolDisplayArabic = False
        '''''''''''''''''''''''''''add root
        ItemsTree.Tag = "E"
        Set nodX = ItemsTree.Nodes.Add(, , "r", "items Groups", "Root")
        ItemsTree.Nodes("r").Expanded = True
    End If

    ItemsTree.Sorted = False
    '''''''''''''''''''''''''''' add group
    ' My_SQL = " SELECT Groups.* "
    ' My_SQL = My_SQL + "  From Groups "
    ' My_SQL = My_SQL + " Where (ParentID =1)"
    '********************************************
    '''''''''''''''''''''''''''' add group Khaled
    Dim kRS   As New ADODB.Recordset
    Dim k_SQL As String
    k_SQL = "select min(Groups.GroupID) as minGroupId from groups "
    kRS.Open k_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If kRS("minGroupId") = 0 Then
        My_SQL = " SELECT Groups.* "
        My_SQL = My_SQL + "  From Groups "
        My_SQL = My_SQL + " Where (ParentID =0)"
    Else
        My_SQL = " SELECT Groups.* "
        My_SQL = My_SQL + "  From Groups "
        My_SQL = My_SQL + " Where (ParentID =1)"
    End If
    '*************************************************************************************
  
    If StrGroupSort <> "" Then
        My_SQL = My_SQL + " Order By " & StrGroupSort
    Else
        My_SQL = My_SQL + " Order By GroupID ASC"
    End If

    Set my_ch_rs = New ADODB.Recordset
    my_ch_rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If my_ch_rs.EOF Or my_ch_rs.BOF Then
        Exit Sub
    End If

    ' BolDisplayArabic = True

    If BolDisplayArabic = True Then
        Call fill_my_Children("r", my_ch_rs, ItemsTree, "Groups", "ParentID")
    Else
        Call fill_my_Children("r", my_ch_rs, ItemsTree, "Groups", "ParentID", , 11)
    End If

    '''''''''''''''''''''''''''add items
    Dim itemFlds As String
    itemFlds = "AssbliedItem,GroupID,ItemID,ItemName,RelatedItem , ItemType "
    If StrItemsSort <> "" Then
        My_SQL = "SELECT " & itemFlds & " From TblItems Order By " & StrItemsSort & ""
    Else

        If SystemOptions.SysDataBaseType = AccessDataBase Then
            My_SQL = "SELECT " & itemFlds & " From TblItems Order By Val(ItemCode)"
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            'My_SQL = "SELECT * From TblItems Order By cast(ItemCode as bigint)"
            My_SQL = "SELECT " & itemFlds & " From TblItems Order By ItemCode "
        End If
    End If

    Set Rs_items = New ADODB.Recordset
    Rs_items.Open My_SQL, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Do While Not Rs_items.EOF
    
        'Debug.Print Rs_items("ItemID").Value
        If Rs_items("AssbliedItem").value = True Then
            Set nodz = ItemsTree.Nodes.Add(Trim(Rs_items("GroupID") & "G"), tvwChild, Rs_items("ItemID") & "I", Rs_items("ItemName"), "Assblied")
            LoadAssbliedItemNode ItemsTree, nodz, Rs_items("ItemID").value
        ElseIf Rs_items("ItemType").value = 1 Then
            'Maintenance
            '’‰› Œœ„…
            If SystemOptions.UserInterface = ArabicInterface Then
                Set nodz = ItemsTree.Nodes.Add(Trim(IIf(IsNull(Rs_items("GroupID")), "", Rs_items("GroupID")) & "G"), tvwChild, IIf(IsNull(Rs_items("ItemID")), "", Rs_items("ItemID")) & "I", IIf(IsNull(Rs_items("ItemName")), "", Rs_items("ItemName")), "Maintenance")
            Else
                Set nodz = ItemsTree.Nodes.Add(Trim(IIf(IsNull(Rs_items("GroupID")), "", Rs_items("GroupID")) & "G"), tvwChild, IIf(IsNull(Rs_items("ItemID")), "", Rs_items("ItemID")) & "I", IIf(IsNull(Rs_items("ItemNamee")), "", Rs_items("ItemNamee")), "Maintenance")
            End If
        Else
            If SystemOptions.UserInterface = ArabicInterface Then
                Set nodz = ItemsTree.Nodes.Add(Trim(IIf(IsNull(Rs_items("GroupID")), "", Rs_items("GroupID")) & "G"), tvwChild, IIf(IsNull(Rs_items("ItemID")), "", Rs_items("ItemID")) & "I", IIf(IsNull(Rs_items("ItemName")), "", Rs_items("ItemName")), "Item")
            Else
                Set nodz = ItemsTree.Nodes.Add(Trim(IIf(IsNull(Rs_items("GroupID")), "", Rs_items("GroupID")) & "G"), tvwChild, IIf(IsNull(Rs_items("ItemID")), "", Rs_items("ItemID")) & "I", IIf(IsNull(Rs_items("ItemNamee")), "", Rs_items("ItemNamee")), "Item")
            End If
        End If

        nodz.Tag = "item"

        If Rs_items("RelatedItem").value = True Then
            LoadRelatedItemNode ItemsTree, nodz, IIf(IsNull(Rs_items("ItemID")), "", Rs_items("ItemID"))
        End If

        Rs_items.MoveNext
  
    Loop
    '    If Not (Rs_items.BOF Or Rs_items.EOF) Then
    '
    '        For LngLoop = 0 To Rs_items.RecordCount - 1
    '
    '        Next LngLoop
    '
    '    End If

    'If BolDisplayArabic = True Then
    '    For LngLoop = 0 To Rs_items.RecordCount - 1
    '        Set nodz = ItemsTree.Nodes.Add(Trim(Rs_items("GroupID") & "G"), tvwChild, Rs_items("ItemID") & "I", Rs_items("ItemName"), "Item")
    '        nodz.Tag = "item"
    '        Rs_items.MoveNext
    '    Next LngLoop
    'Else
    '    For LngLoop = 0 To Rs_items.RecordCount - 1
    '        Set nodz = ItemsTree.Nodes.Add(Trim(Rs_items("Group_Code")), tvwChild, Rs_items("Item_ID"), IIf(IsNull(Rs_items("Item_Name_Eng")), "", Rs_items("Item_Name_Eng")), "Item")
    '        If Trim(nodz.Text) = "" Then
    '            nodz.ForeColor = vbRed
    '            nodz.Bold = True
    '            nodz.Text = "No Item Name"
    '        End If
    '        nodz.Tag = "item"
    '        Rs_items.MoveNext
    '    Next LngLoop
    'End If
    ItemsTree.Nodes("r").EnsureVisible

    Rs_items.Close
    Set Rs_items = Nothing
    ItemsTree.Refresh
    Exit Sub
ErrTrap:

    If SystemOptions.SysRegisterState = DevelopVersion Then
        Stop
        'Resume
    End If

    Msg = "ÕœÀ Œÿ« √À‰«¡  Õ„Ì· ‘Ã—… «·√’‰«›"
    Msg = Msg & CHR(13) & Err.Description
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & Err.Source
    Msg = Msg & CHR(13) & "ModTree:LoadTreeItems"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Public Sub LoadTreeGroups2(ItemsTree As MSComctlLib.TreeView, _
                           Optional StrGroupSort As String = "", _
                           Optional StrItemsSort As String = "", _
                           Optional BolWithFullOption As Boolean = False)

    Dim Rs_items         As ADODB.Recordset
    Dim My_SQL           As String
    Dim nodX             As Node
    Dim nodz             As Node
    Dim RsOptions        As ADODB.Recordset
    Dim my_ch_rs         As ADODB.Recordset
    Dim BolDisplayArabic As Boolean
    ' Dim LngLoop          As Long

    '    On Error GoTo ErrTrap
    ItemsTree.Nodes.Clear

    If SystemOptions.UserInterface = ArabicInterface Then
        BolDisplayArabic = True
        ItemsTree.Tag = "A"
        Make_RightToLeft ItemsTree
        '''''''''''''''''''''''''''add root
        Set nodX = ItemsTree.Nodes.Add(, , "r", "  „Ã„Ê⁄«  «·«’‰«› ", "Root")
        ItemsTree.Nodes("r").Expanded = True
    Else
        BolDisplayArabic = False
        '''''''''''''''''''''''''''add root
        ItemsTree.Tag = "E"
        Set nodX = ItemsTree.Nodes.Add(, , "r", "items Groups", "Root")
        ItemsTree.Nodes("r").Expanded = True
    End If

    ItemsTree.Sorted = False
    '''''''''''''''''''''''''''' add group
    ' My_SQL = " SELECT Groups.* "
    ' My_SQL = My_SQL + "  From Groups "
    ' My_SQL = My_SQL + " Where (ParentID =1)"
    '********************************************
    '''''''''''''''''''''''''''' add group Khaled
    '    Dim kRS   As New ADODB.Recordset
    '    Dim k_SQL As String
    '    k_SQL = "select min(Groups.GroupID) as minGroupId from groups "
    '    kRS.Open k_SQL, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    '    If kRS("minGroupId") = 0 Then
    '        My_SQL = " SELECT Groups.* "
    '        My_SQL = My_SQL + "  From Groups "
    '        My_SQL = My_SQL + " Where (ParentID =0)"
    '    Else
    '        My_SQL = " SELECT Groups.* "
    '        My_SQL = My_SQL + "  From Groups "
    '        My_SQL = My_SQL + " Where (ParentID =1)"
    '    End If
    
    '********************************
    Dim MySQL As String
    MySQL = MySQL & "SELECT Groups.GroupID, "
    MySQL = MySQL & "       Groups.GroupName" & IIf(BolDisplayArabic, "", "e") & " GroupName, "
    'MySQl = MySQl & "       Groups.GroupNamee, "
    MySQL = MySQL & "       ISNULL(Groups.ParentID, 0) ParentID,ISNULL(Fullcode,'') Fullcode ,  "
    MySQL = MySQL & "       Groups.code "
    MySQL = MySQL & "FROM dbo.Groups WHERE  Groups.ParentID  IS NOT NULL  "
    MySQL = MySQL & "ORDER BY  ISNULL(Groups.ParentID, 0) ASC, "
    MySQL = MySQL & "GroupID "
    MySQL = MySQL & ""
    
    '********************************
    '    If StrGroupSort <> "" Then
    '        My_SQL = My_SQL + " Order By " & StrGroupSort
    '    Else
    '        My_SQL = My_SQL + " Order By GroupID ASC"
    '    End If
  
    Set my_ch_rs = New ADODB.Recordset
    my_ch_rs.Open MySQL, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText
 
    If my_ch_rs.BOF Then
        Exit Sub
    End If

    Dim firestParent As Integer
    Dim crntParent   As Integer
    firestParent = -1

    Do While Not my_ch_rs.EOF

        If firestParent = -1 Then
            firestParent = val(my_ch_rs!ParentID & "")
        End If

        crntParent = val(my_ch_rs!ParentID & "")
        Set nodX = ItemsTree.Nodes.Add(IIf(crntParent = firestParent, "r", CStr(crntParent) & "G"), 4, my_ch_rs!GroupID & "G", my_ch_rs!Fullcode & " " & my_ch_rs!GroupName & "", "Closed_Node", "Open_Node")
        nodX.ExpandedImage = "Open_Node"
         
        '**********************
        Dim tmpNode As Node
        Set tmpNode = ItemsTree.Nodes.Add(nodX.Key, 4, "DElME" & my_ch_rs!GroupID & "G", "", "Closed_Node", "Open_Node")
        tmpNode.ExpandedImage = "Open_Node"
        tmpNode.Tag = "TMP"
        '***********************
        my_ch_rs.MoveNext
    Loop

    ItemsTree.Refresh
    Exit Sub
    '''''''''''''''''''''''''''add items
   
    Exit Sub
ErrTrap:

    If SystemOptions.SysRegisterState = DevelopVersion Then
        Stop
        'Resume
    End If

    Msg = "ÕœÀ Œÿ« √À‰«¡  Õ„Ì· ‘Ã—… «·√’‰«›"
    Msg = Msg & CHR(13) & Err.Description
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & Err.Source
    Msg = Msg & CHR(13) & "ModTree:LoadTreeItems"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Sub trytoremove(tree, Key)
On Error GoTo EH

    tree.Nodes.Remove Key
EH:
End Sub
Sub LoadItemNode(nodz As MSComctlLib.Node, _
                 ItemsTree As MSComctlLib.TreeView, _
                 Optional force As Boolean = False, _
                 Optional mItemId As String = "")
    On Error GoTo ErrTrap
   
    '''''''''''''''''''''''''''add items
    If (nodz.Tag = "0" And Not force) Or nodz.Tag = "item" Then Exit Sub
    nodz.Tag = "0" ' nod item loaded

    If nodz.Key = "r" Then Exit Sub
    Screen.MousePointer = vbHourglass
     
    '******Remove Tmp Nod*********
  
    '  ItemsTree.Nodes.Remove "DElME" & nodz.key
    trytoremove ItemsTree, "DElME" & nodz.Key
    '****************
    
    Dim itemFlds      As String
    Dim ItemNameFld   As String
    Dim ItemKeyFilter As String
    ItemKeyFilter = IIf(mItemId <> "", " And ItemId = " & mItemId, "")
    itemFlds = "AssbliedItem,GroupID,ItemID,ItemName" & IIf(SystemOptions.UserInterface = ArabicInterface, "", "e") & " ItemName,RelatedItem , ItemType "
      
    If StrItemsSort <> "" Then
        My_SQL = "SELECT " & itemFlds & " From TblItems WHERE GroupID = " & Replace(nodz.Key, "G", "") & ItemKeyFilter & "  Order By " & StrItemsSort & ""
    Else

        If SystemOptions.SysDataBaseType = AccessDataBase Then
            My_SQL = "SELECT " & itemFlds & " From TblItems WHERE GroupID = " & Replace(nodz.Key, "G", "") & ItemKeyFilter & "  Order By Val(ItemCode)"
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            My_SQL = "SELECT " & itemFlds & " From TblItems WHERE GroupID = " & Replace(nodz.Key, "G", "") & ItemKeyFilter & "   Order By ItemCode "
        End If
    End If

    Set Rs_items = New ADODB.Recordset
    Rs_items.Open My_SQL, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText
    Dim ItemNode As Node

    Do While Not Rs_items.EOF
       
        Dim ItemKey       As String
        Dim ItemParentKey As String
        ItemParentKey = Rs_items!GroupID & "G"
        ItemKey = Rs_items!ItemID & "I"

        'if node found Replace it<For Refresh Data After Save >
        If force Then
            trytoremove ItemsTree, ItemKey
        End If

        '''''''''''''''''''
        If Rs_items("AssbliedItem").value = True Then
            Set ItemNode = ItemsTree.Nodes.Add(ItemParentKey, tvwChild, ItemKey, Rs_items!ItemName & "", "Assblied")
            LoadAssbliedItemNode ItemsTree, ItemNode, Rs_items("ItemID").value
        ElseIf Rs_items("ItemType").value = 1 Then
            'Maintenance
            '’‰› Œœ„…
            Set ItemNode = ItemsTree.Nodes.Add(ItemParentKey, tvwChild, ItemKey, Rs_items!ItemName & "", "Maintenance")

        Else
            Set ItemNode = ItemsTree.Nodes.Add(ItemParentKey, tvwChild, ItemKey, Rs_items!ItemName & "", "Item")
        End If

        ItemNode.Tag = "item"

        If Rs_items("RelatedItem").value = True Then
            LoadRelatedItemNode ItemsTree, ItemNode, val(Rs_items!ItemID & "")
        End If

        Rs_items.MoveNext
    Loop
   
    ItemsTree.Nodes("r").EnsureVisible

    Rs_items.Close
    Set Rs_items = Nothing
    ItemsTree.Refresh
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
    Msg = "ÕœÀ Œÿ« √À‰«¡  Õ„Ì· ‘Ã—… «·√’‰«›"
    Msg = Msg & CHR(13) & Err.Description
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & Err.Source
    Msg = Msg & CHR(13) & "ModTree:LoadTreeItems"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

 Public Sub LoadTreeAccount(ChartTree As MSComctlLib.TreeView)
    'ChartTree The TreeView which we will load
    'BolRtl=Make The Tree Right to Left
    Dim My_SQL As String
    Dim nodX As MSComctlLib.Node
    Dim nodz As MSComctlLib.Node
    Dim IntBalSheetShow As Integer
    Dim RsOptions As New ADODB.Recordset
    Dim rs_chart As New ADODB.Recordset
    Dim LngLoop As Long
    Dim BolRtl As Boolean
    Dim IntColName As Integer
    Dim StrSelectedNode As String
    Dim Msg As String
     On Error GoTo ErrTrap

    '=============================== Õ„Ì· «·‘Ã—… «·„Õ«”»Ì…
    IntBalSheetShow = 1
    '--------------------
    'Â–« «·ŒÌ«— ›«∆œ …  ÕœÌœ Â· Ì „ ⁄—÷ «”„ «·Õ”«» ›ﬁÿ
    '√„
    '⁄—÷ √”„ «·Õ”«» Ê«·ﬂÊœ «·„Õ«”»Ï
    'IntBalSheetShow = IIf(IsNull(RsOptions("BalSheetShow").Value), 1, RsOptions("BalSheetShow").Value)
    'MsgBox "IN LoadTreeAccount line 3"

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True
    Else
        BolRtl = False
    End If

    '----Load Data From Data Base
    If SystemOptions.SysDataBaseType = AccessDataBase Then
        My_SQL = " SELECT ACCOUNTS.* "
        My_SQL = My_SQL + " From ACCOUNTS "
        My_SQL = My_SQL + " WHERE (((ACCOUNTS.last_account)=False)" & "and(Account_Code <> 'r' or Account_Code <> 'temp') and Parent_Account_Code = 'r');"
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        My_SQL = " SELECT ACCOUNTS.* "
        My_SQL = My_SQL + " From ACCOUNTS "
        My_SQL = My_SQL + " WHERE ( ( (ACCOUNTS.last_account)=0)" & "and(Account_Code <> 'r' or Account_Code <> 'temp') and Parent_Account_Code = 'r' ) "
'
    End If

    Set rs_chart = Nothing
    Set rs_chart = New ADODB.Recordset
    rs_chart.CursorLocation = adUseClient

    rs_chart.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'MsgBox "IN LoadTreeAccount line 6"
    '----begin to write data
    'Set The TreeView Control
 
    If ChartTree.Nodes.count > 0 Then
        If Not ChartTree.SelectedItem Is Nothing Then
            StrSelectedNode = ChartTree.SelectedItem.Key
        End If
    End If

    ChartTree.Nodes.Clear
    ChartTree.Sorted = False

    If BolRtl = True Then
        Make_RightToLeft ChartTree
        Set nodX = ChartTree.Nodes.Add(, , "r", "«·œ·Ì· «·„Õ«”»Ï ", "Root")
        IntColName = 2
    Else
        Set nodX = ChartTree.Nodes.Add(, , "r", "Chart Of Accounts", "Root")
        IntColName = 9
    End If

    Call fill_my_Children_Account("r", rs_chart, ChartTree, "ACCOUNTS", "Parent_Account_Code", " (ACCOUNTS.last_account)=0 ", IntColName, "Account_Serial")
    My_SQL = ""


    If BolRtl = True Then 'Load Arabic Name
        If IntBalSheetShow = 1 Then 'Load Arabic Name only
            My_SQL = " SELECT ACCOUNTS.* "
            My_SQL = My_SQL + " From ACCOUNTS WHERE Account_Code <> 'r' "
            '//////
            My_SQL = My_SQL + GetAccountByBarnchUser
            My_SQL = My_SQL + GetAccountCodeHiding
            ''/////////
            My_SQL = My_SQL & "   ORDER BY ACCOUNTS.Account_Serial "
            IntColName = 2
        Else 'Load Arabic Name and Account Code
            My_SQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name  +'('+ iif( ACCOUNTS.Account_Serial is null,'  ' ,  ACCOUNTS.Account_Serial)+')' as Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account, ACCOUNTS.cannot_del, ACCOUNTS.Account_Serial"
            My_SQL = My_SQL + " FROM ACCOUNTS  WHERE Account_Code <> 'r' "
            ''////////////
            My_SQL = My_SQL + GetAccountByBarnchUser
            My_SQL = My_SQL + GetAccountCodeHiding
            
           '/////
            My_SQL = My_SQL & "   ORDER BY ACCOUNTS.Account_Serial "
            IntColName = 1
        End If

    ElseIf BolRtl = False Then

        'Load English  Name
        If IntBalSheetShow = 1 Then
            'Load English Name only
            My_SQL = " SELECT ACCOUNTS.* "
            My_SQL = My_SQL + " From ACCOUNTS WHERE Account_Code <> 'r' "
            My_SQL = My_SQL + GetAccountByBarnchUser
            My_SQL = My_SQL + GetAccountCodeHiding
            My_SQL = My_SQL + " ORDER BY ACCOUNTS.Account_Serial "
            IntColName = 9
        Else
            'Load English Name and Account Code
            My_SQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_NameEng  +'('+ iif( ACCOUNTS.Account_Serial is null,'  ', ACCOUNTS.Account_Serial)+')' as Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account, ACCOUNTS.cannot_del, ACCOUNTS.Account_Serial"
            My_SQL = My_SQL + " FROM ACCOUNTS  WHERE Account_Code <> 'r'"
            My_SQL = My_SQL + GetAccountByBarnchUser
            My_SQL = My_SQL + GetAccountCodeHiding
            My_SQL = My_SQL + " ORDER BY ACCOUNTS.Account_Serial "
            IntColName = 1
        End If
    End If

    Set rs_chart = Nothing
    Set rs_chart = New ADODB.Recordset
    rs_chart.Open My_SQL, Cn, adOpenStatic, adLockReadOnly

    For LngLoop = 0 To rs_chart.RecordCount - 1
        If rs_chart("last_account") Then
            Set nodz = ChartTree.Nodes.Add(Trim(rs_chart("Parent_Account_Code").value) & "G", tvwChild, rs_chart("Account_Code").value, rs_chart(IntColName).value, "Item")
            nodz.Tag = "last"
        Else
            Set nodz = ChartTree.Nodes(CStr(rs_chart("Account_Code").value) & "G")
        End If

        If rs_chart("cannot_del") Then
            nodz.Tag = nodz.Tag + "cannot_del"
        End If
'End If
        rs_chart.MoveNext
    Next LngLoop

    rs_chart.Close
    Set rs_chart = Nothing
    'RsOptions.Close
    'Set RsOptions = Nothing
    
    On Error Resume Next
    ChartTree.Nodes("r").EnsureVisible
    StrSelectedNode = ChartTree.SelectedItem.Key

    If StrSelectedNode <> "" Then

        'StrSelectedNode = Replace$(StrSelectedNode, "G", "")
 
        '      If FrmAccountCharts.OptAccountType(1).value = True Then
        '       StrSelectedNode = StrSelectedNode & "G"
        '        Else
        '        StrSelectedNode = StrSelectedNode & "G"
        ' StrSelectedNode = ChartTree.SelectedItem.key
        '        End If
 
        If FrmAccountCharts.OptAccountType(0).value = True Then
            '    StrSelectedNode = Replace$(StrSelectedNode, "G", "")
        Else

            If right(StrSelectedNode, 1) <> "G" Then
                StrSelectedNode = StrSelectedNode & "G"
            End If

            ' StrSelectedNode = ChartTree.SelectedItem.key
        End If

        ChartTree.SetFocus
        ChartTree.Nodes(StrSelectedNode).EnsureVisible
    
        ChartTree.Nodes(StrSelectedNode).Selected = True
    
        Exit Sub
    End If
 
    Exit Sub
ErrTrap:
    'Resume
    Msg = "ÕœÀ Œÿ« √À‰«¡  Õ„Ì· «·»Ì«‰« " & CHR(13) & "»—Ã«¡ «·√ ’«· »«·‘—ﬂ… " & CHR(13) & "”Ê› Ì „ €·ﬁ «·»—‰«„Ã"
    Msg = Msg & CHR(13) & Err.Description
    Msg = Msg & CHR(13) & Err.Source
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & "Error Location LoadTreeAccount" & CStr(rs_chart("Account_Code").value) & CHR(13) & rs_chart("Account_Name").value
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Cn.Execute "delete from TblAccountBranch"
    Cn.Execute "delete from TblAccountUser"
    
     
  


End Sub

Private Sub Form_Load()

End Sub

Public Sub LoadTreeAccountx11102017(ChartTree As MSComctlLib.TreeView)
    'ChartTree The TreeView which we will load
    'BolRtl=Make The Tree Right to Left
    Dim My_SQL As String
    Dim nodX As MSComctlLib.Node
    Dim nodz As MSComctlLib.Node
    Dim IntBalSheetShow As Integer
    Dim RsOptions As New ADODB.Recordset
    Dim rs_chart As New ADODB.Recordset
    Dim LngLoop As Long
    Dim BolRtl As Boolean
    Dim IntColName As Integer
    Dim StrSelectedNode As String
    Dim Msg As String
     On Error GoTo ErrTrap

    '=============================== Õ„Ì· «·‘Ã—… «·„Õ«”»Ì…
    IntBalSheetShow = 1
    '--------------------
    'Â–« «·ŒÌ«— ›«∆œ …  ÕœÌœ Â· Ì „ ⁄—÷ «”„ «·Õ”«» ›ﬁÿ
    '√„
    '⁄—÷ √”„ «·Õ”«» Ê«·ﬂÊœ «·„Õ«”»Ï
    'IntBalSheetShow = IIf(IsNull(RsOptions("BalSheetShow").Value), 1, RsOptions("BalSheetShow").Value)
    'MsgBox "IN LoadTreeAccount line 3"

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True
    Else
        BolRtl = False
    End If

    '----Load Data From Data Base
    If SystemOptions.SysDataBaseType = AccessDataBase Then
        My_SQL = " SELECT ACCOUNTS.* "
        My_SQL = My_SQL + " From ACCOUNTS "
        My_SQL = My_SQL + " WHERE (((ACCOUNTS.last_account)=False)" & "and(Account_Code <> 'r' or Account_Code <> 'temp') and Parent_Account_Code = 'r');"
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        My_SQL = " SELECT ACCOUNTS.* "
        My_SQL = My_SQL + " From ACCOUNTS "
        My_SQL = My_SQL + " WHERE ( ( (ACCOUNTS.last_account)=0)" & "and(Account_Code <> 'r' or Account_Code <> 'temp') and Parent_Account_Code = 'r' ) "
'
    End If

    Set rs_chart = Nothing
    Set rs_chart = New ADODB.Recordset
    rs_chart.CursorLocation = adUseClient

    rs_chart.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'MsgBox "IN LoadTreeAccount line 6"
    '----begin to write data
    'Set The TreeView Control
 
    If ChartTree.Nodes.count > 0 Then
        If Not ChartTree.SelectedItem Is Nothing Then
            StrSelectedNode = ChartTree.SelectedItem.Key
        End If
    End If

    ChartTree.Nodes.Clear
    ChartTree.Sorted = False

    If BolRtl = True Then
        Make_RightToLeft ChartTree
        Set nodX = ChartTree.Nodes.Add(, , "r", "«·œ·Ì· «·„Õ«”»Ï ", "Root")
        IntColName = 2
    Else
        Set nodX = ChartTree.Nodes.Add(, , "r", "Chart Of Accounts", "Root")
        IntColName = 9
    End If

    Call fill_my_Children_Account("r", rs_chart, ChartTree, "ACCOUNTS", "Parent_Account_Code", " (ACCOUNTS.last_account)=0 ", IntColName, "Account_Serial")
    My_SQL = ""

    If BolRtl = True Then 'Load Arabic Name
        If IntBalSheetShow = 1 Then 'Load Arabic Name only
            My_SQL = " SELECT ACCOUNTS.* "
            My_SQL = My_SQL + " From ACCOUNTS WHERE Account_Code <> 'r' "
            '//////'salah
            My_SQL = My_SQL + GetAccountByBarnchUser
            ''/////////
            My_SQL = My_SQL & "   ORDER BY ACCOUNTS.Account_Serial "
            IntColName = 2
        Else 'Load Arabic Name and Account Code
            My_SQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name  +'('+ iif( ACCOUNTS.Account_Serial is null,'  ' ,  ACCOUNTS.Account_Serial)+')' as Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account, ACCOUNTS.cannot_del, ACCOUNTS.Account_Serial"
            My_SQL = My_SQL + " FROM ACCOUNTS  WHERE Account_Code <> 'r' "
            ''////////////
            My_SQL = My_SQL + GetAccountByBarnchUser
           '/////
            My_SQL = My_SQL & "   ORDER BY ACCOUNTS.Account_Serial "
            IntColName = 1
        End If

    ElseIf BolRtl = False Then

        'Load English  Name
        If IntBalSheetShow = 1 Then
            'Load English Name only
            My_SQL = " SELECT ACCOUNTS.* "
            My_SQL = My_SQL + " From ACCOUNTS WHERE Account_Code <> 'r' "
            My_SQL = My_SQL + GetAccountByBarnchUser
            My_SQL = My_SQL + " ORDER BY ACCOUNTS.Account_Serial "
            IntColName = 9
        Else
            'Load English Name and Account Code
            My_SQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_NameEng  +'('+ iif( ACCOUNTS.Account_Serial is null,'  ', ACCOUNTS.Account_Serial)+')' as Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account, ACCOUNTS.cannot_del, ACCOUNTS.Account_Serial"
            My_SQL = My_SQL + " FROM ACCOUNTS  WHERE Account_Code <> 'r'"
            My_SQL = My_SQL + GetAccountByBarnchUser
            My_SQL = My_SQL + " ORDER BY ACCOUNTS.Account_Serial "
            IntColName = 1
        End If
    End If

    Set rs_chart = Nothing
    Set rs_chart = New ADODB.Recordset
    rs_chart.Open My_SQL, Cn, adOpenStatic, adLockReadOnly

    For LngLoop = 0 To rs_chart.RecordCount - 1

        If rs_chart("last_account") Then
            Set nodz = ChartTree.Nodes.Add(Trim(rs_chart("Parent_Account_Code").value) & "G", tvwChild, rs_chart("Account_Code").value, rs_chart(IntColName).value, "Item")
            nodz.Tag = "last"
        Else
            Set nodz = ChartTree.Nodes(CStr(rs_chart("Account_Code").value) & "G")
        End If

        If rs_chart("cannot_del") Then
            nodz.Tag = nodz.Tag + "cannot_del"
        End If

        rs_chart.MoveNext
    Next LngLoop

    rs_chart.Close
    Set rs_chart = Nothing
    'RsOptions.Close
    'Set RsOptions = Nothing
    On Error Resume Next
    ChartTree.Nodes("r").EnsureVisible
    StrSelectedNode = ChartTree.SelectedItem.Key

    If StrSelectedNode <> "" Then

        'StrSelectedNode = Replace$(StrSelectedNode, "G", "")
 
        '      If FrmAccountCharts.OptAccountType(1).value = True Then
        '       StrSelectedNode = StrSelectedNode & "G"
        '        Else
        '        StrSelectedNode = StrSelectedNode & "G"
        ' StrSelectedNode = ChartTree.SelectedItem.key
        '        End If
 
        If FrmAccountCharts.OptAccountType(0).value = True Then
            '    StrSelectedNode = Replace$(StrSelectedNode, "G", "")
        Else

            If right(StrSelectedNode, 1) <> "G" Then
                StrSelectedNode = StrSelectedNode & "G"
            End If

            ' StrSelectedNode = ChartTree.SelectedItem.key
        End If

        ChartTree.SetFocus
        ChartTree.Nodes(StrSelectedNode).EnsureVisible
    
        ChartTree.Nodes(StrSelectedNode).Selected = True
    
        Exit Sub
    End If
 
    Exit Sub
ErrTrap:
    'Resume
    Msg = "ÕœÀ Œÿ« √À‰«¡  Õ„Ì· «·»Ì«‰« " & CHR(13) & "»—Ã«¡ «·√ ’«· »«·‘—ﬂ… " & CHR(13) & "”Ê› Ì „ €·ﬁ «·»—‰«„Ã"
    Msg = Msg & CHR(13) & Err.Description
    Msg = Msg & CHR(13) & Err.Source
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & "Error Location LoadTreeAccount" & CStr(rs_chart("Account_Code").value)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Public Sub LoadTreeAccountx17072017(ChartTree As MSComctlLib.TreeView)
    'ChartTree The TreeView which we will load
    'BolRtl=Make The Tree Right to Left
    Dim My_SQL As String
    Dim nodX As MSComctlLib.Node
    Dim nodz As MSComctlLib.Node
    Dim IntBalSheetShow As Integer
    Dim RsOptions As New ADODB.Recordset
    Dim rs_chart As New ADODB.Recordset
    Dim LngLoop As Long
    Dim BolRtl As Boolean
    Dim IntColName As Integer
    Dim StrSelectedNode As String
    Dim Msg As String
     On Error GoTo ErrTrap

    '=============================== Õ„Ì· «·‘Ã—… «·„Õ«”»Ì…
    IntBalSheetShow = 1
    '--------------------
    'Â–« «·ŒÌ«— ›«∆œ …  ÕœÌœ Â· Ì „ ⁄—÷ «”„ «·Õ”«» ›ﬁÿ
    '√„
    '⁄—÷ √”„ «·Õ”«» Ê«·ﬂÊœ «·„Õ«”»Ï
    'IntBalSheetShow = IIf(IsNull(RsOptions("BalSheetShow").Value), 1, RsOptions("BalSheetShow").Value)
    'MsgBox "IN LoadTreeAccount line 3"

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True
    Else
        BolRtl = False
    End If

    '----Load Data From Data Base
    If SystemOptions.SysDataBaseType = AccessDataBase Then
        My_SQL = " SELECT ACCOUNTS.* "
        My_SQL = My_SQL + " From ACCOUNTS "
        My_SQL = My_SQL + " WHERE (((ACCOUNTS.last_account)=False)" & "and(Account_Code <> 'r' or Account_Code <> 'temp') and Parent_Account_Code = 'r');"
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        My_SQL = " SELECT ACCOUNTS.* "
        My_SQL = My_SQL + " From ACCOUNTS "
        My_SQL = My_SQL + " WHERE ( ( (ACCOUNTS.last_account)=0)" & "and(Account_Code <> 'r' or Account_Code <> 'temp') and Parent_Account_Code = 'r' ) "

    End If

    Set rs_chart = Nothing
    Set rs_chart = New ADODB.Recordset
    rs_chart.CursorLocation = adUseClient

    rs_chart.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'MsgBox "IN LoadTreeAccount line 6"
    '----begin to write data
    'Set The TreeView Control
 
    If ChartTree.Nodes.count > 0 Then
        If Not ChartTree.SelectedItem Is Nothing Then
            StrSelectedNode = ChartTree.SelectedItem.Key
        End If
    End If

    ChartTree.Nodes.Clear
    ChartTree.Sorted = False

    If BolRtl = True Then
        Make_RightToLeft ChartTree
        Set nodX = ChartTree.Nodes.Add(, , "r", "«·œ·Ì· «·„Õ«”»Ï ", "Root")
        IntColName = 2
    Else
        Set nodX = ChartTree.Nodes.Add(, , "r", "Chart Of Accounts", "Root")
        IntColName = 9
    End If

    Call fill_my_Children_Account("r", rs_chart, ChartTree, "ACCOUNTS", "Parent_Account_Code", " (ACCOUNTS.last_account)=0 ", IntColName, "Account_Serial")
    My_SQL = ""

    If BolRtl = True Then 'Load Arabic Name
        If IntBalSheetShow = 1 Then 'Load Arabic Name only
            My_SQL = " SELECT ACCOUNTS.* "
            My_SQL = My_SQL + " From ACCOUNTS WHERE Account_Code <> 'r' ORDER BY ACCOUNTS.Account_Serial "
            IntColName = 2
        Else 'Load Arabic Name and Account Code
            My_SQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name  +'('+ iif( ACCOUNTS.Account_Serial is null,'  ' ,  ACCOUNTS.Account_Serial)+')' as Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account, ACCOUNTS.cannot_del, ACCOUNTS.Account_Serial"
            My_SQL = My_SQL + " FROM ACCOUNTS  WHERE Account_Code <> 'r' ORDER BY ACCOUNTS.Account_Serial "
            IntColName = 1
        End If

    ElseIf BolRtl = False Then

        'Load English  Name
        If IntBalSheetShow = 1 Then
            'Load English Name only
            My_SQL = " SELECT ACCOUNTS.* "
            My_SQL = My_SQL + " From ACCOUNTS WHERE Account_Code <> 'r' ORDER BY ACCOUNTS.Account_Serial "
            IntColName = 9
        Else
            'Load English Name and Account Code
            My_SQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_NameEng  +'('+ iif( ACCOUNTS.Account_Serial is null,'  ', ACCOUNTS.Account_Serial)+')' as Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account, ACCOUNTS.cannot_del, ACCOUNTS.Account_Serial"
            My_SQL = My_SQL + " FROM ACCOUNTS  WHERE Account_Code <> 'r' ORDER BY ACCOUNTS.Account_Serial "
            IntColName = 1
        End If
    End If

    Set rs_chart = Nothing
    Set rs_chart = New ADODB.Recordset
    rs_chart.Open My_SQL, Cn, adOpenStatic, adLockReadOnly

    For LngLoop = 0 To rs_chart.RecordCount - 1

        If rs_chart("last_account") Then
            Set nodz = ChartTree.Nodes.Add(Trim(rs_chart("Parent_Account_Code").value) & "G", tvwChild, rs_chart("Account_Code").value, rs_chart(IntColName).value, "Item")
            nodz.Tag = "last"
        Else
            Set nodz = ChartTree.Nodes(CStr(rs_chart("Account_Code").value) & "G")
        End If

        If rs_chart("cannot_del") Then
            nodz.Tag = nodz.Tag + "cannot_del"
        End If

        rs_chart.MoveNext
    Next LngLoop

    rs_chart.Close
    Set rs_chart = Nothing
    'RsOptions.Close
    'Set RsOptions = Nothing
    On Error Resume Next
    ChartTree.Nodes("r").EnsureVisible
    StrSelectedNode = ChartTree.SelectedItem.Key

    If StrSelectedNode <> "" Then

        'StrSelectedNode = Replace$(StrSelectedNode, "G", "")
 
        '      If FrmAccountCharts.OptAccountType(1).value = True Then
        '       StrSelectedNode = StrSelectedNode & "G"
        '        Else
        '        StrSelectedNode = StrSelectedNode & "G"
        ' StrSelectedNode = ChartTree.SelectedItem.key
        '        End If
 
        If FrmAccountCharts.OptAccountType(0).value = True Then
            '    StrSelectedNode = Replace$(StrSelectedNode, "G", "")
        Else

            If right(StrSelectedNode, 1) <> "G" Then
                StrSelectedNode = StrSelectedNode & "G"
            End If

            ' StrSelectedNode = ChartTree.SelectedItem.key
        End If

        ChartTree.SetFocus
        ChartTree.Nodes(StrSelectedNode).EnsureVisible
    
        ChartTree.Nodes(StrSelectedNode).Selected = True
    
        Exit Sub
    End If
 
    Exit Sub
ErrTrap:
    'Resume
    Msg = "ÕœÀ Œÿ« √À‰«¡  Õ„Ì· «·»Ì«‰« " & CHR(13) & "»—Ã«¡ «·√ ’«· »«·‘—ﬂ… " & CHR(13) & "”Ê› Ì „ €·ﬁ «·»—‰«„Ã"
    Msg = Msg & CHR(13) & Err.Description
    Msg = Msg & CHR(13) & Err.Source
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & "Error Location LoadTreeAccount" & CStr(rs_chart("Account_Code").value)
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Public Sub LoadTreeAccountBalanceSheet(ChartTree As MSComctlLib.TreeView, _
                                       Optional BalanceSheet As Boolean)
    'ChartTree The TreeView which we will load
    'BolRtl=Make The Tree Right to Left
    Dim My_SQL As String
    Dim nodX As MSComctlLib.Node
    Dim nodz As MSComctlLib.Node
    Dim IntBalSheetShow As Integer
    Dim RsOptions As New ADODB.Recordset
    Dim rs_chart As New ADODB.Recordset
    Dim LngLoop As Long
    Dim BolRtl As Boolean
    Dim IntColName As Integer
    Dim StrSelectedNode As String
    Dim Msg As String
    'On Error GoTo ErrTrap

    '=============================== Õ„Ì· «·‘Ã—… «·„Õ«”»Ì…
    IntBalSheetShow = 1
    '--------------------
    'Â–« «·ŒÌ«— ›«∆œ …  ÕœÌœ Â· Ì „ ⁄—÷ «”„ «·Õ”«» ›ﬁÿ
    '√„
    '⁄—÷ √”„ «·Õ”«» Ê«·ﬂÊœ «·„Õ«”»Ï
    'IntBalSheetShow = IIf(IsNull(RsOptions("BalSheetShow").Value), 1, RsOptions("BalSheetShow").Value)
    'MsgBox "IN LoadTreeAccount line 3"

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True
    Else
        BolRtl = False
    End If

    '----Load Data From Data Base
    If SystemOptions.SysDataBaseType = AccessDataBase Then
        My_SQL = " SELECT ACCOUNTS.* "
        My_SQL = My_SQL + " From ACCOUNTS "
        My_SQL = My_SQL + " WHERE (((ACCOUNTS.last_account)=False)" & "and(Account_Code <> 'r' or Account_Code <> 'temp') and Parent_Account_Code = 'r');"
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        My_SQL = " SELECT ACCOUNTS.* "
        My_SQL = My_SQL + " From ACCOUNTS "
        My_SQL = My_SQL + " WHERE (((ACCOUNTS.last_account)=0)" & "and(Account_Code <> 'r' or Account_Code <> 'temp') and Parent_Account_Code = 'r') "
    End If

    If BalanceSheet = True Then
        'My_SQL = My_SQL + " and (Account_Code Like 'a1%' OR Account_Code Like 'a2%'  ) "
        'My_SQL = My_SQL + "    and (ShowInBlanceSheet is null or ShowInBlanceSheet=1)"
    End If

    Set rs_chart = Nothing
    Set rs_chart = New ADODB.Recordset
    rs_chart.CursorLocation = adUseClient

    rs_chart.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    'MsgBox "IN LoadTreeAccount line 6"
    '----begin to write data
    'Set The TreeView Control
    If ChartTree.Nodes.count > 0 Then
        If Not ChartTree.SelectedItem Is Nothing Then
            StrSelectedNode = mId(ChartTree.SelectedItem.Key, 1, Len(ChartTree.SelectedItem.Key) - 1)

            If FrmAccountCharts.OptAccountType(0).value = True Then
                StrSelectedNode = ChartTree.SelectedItem.Key & "G"
            Else
                StrSelectedNode = ChartTree.SelectedItem.Key
            End If
        End If
    End If

    ChartTree.Nodes.Clear
    ChartTree.Sorted = False

    If BolRtl = True Then
        Make_RightToLeft ChartTree
        Set nodX = ChartTree.Nodes.Add(, , "r", "«·œ·Ì· «·„Õ«”»Ï ", "Root")
        IntColName = 2
    Else
        Set nodX = ChartTree.Nodes.Add(, , "r", "Chart Of Accounts", "Root")
        IntColName = 9
    End If

    Call fill_my_Children_Account("r", rs_chart, ChartTree, "ACCOUNTS", "Parent_Account_Code", " (ACCOUNTS.last_account)=0 ", IntColName)
    My_SQL = ""

    If BolRtl = True Then 'Load Arabic Name
        If IntBalSheetShow = 1 Then 'Load Arabic Name only
            My_SQL = " SELECT ACCOUNTS.* "
            My_SQL = My_SQL + " From ACCOUNTS WHERE Account_Code <> 'r' "
        
            IntColName = 2
        Else 'Load Arabic Name and Account Code
            My_SQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name  +'('+ iif( ACCOUNTS.Account_Serial is null,'  ', ACCOUNTS.Account_Serial)+')' as Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account, ACCOUNTS.cannot_del, ACCOUNTS.Account_Serial"
            My_SQL = My_SQL + " FROM ACCOUNTS  WHERE Account_Code <> 'r' "
            IntColName = 1
        End If

    ElseIf BolRtl = False Then

        'Load English  Name
        If IntBalSheetShow = 1 Then
            'Load English Name only
            My_SQL = " SELECT ACCOUNTS.* "
            My_SQL = My_SQL + " From ACCOUNTS WHERE Account_Code <> 'r'   "
            IntColName = 9
        Else
            'Load English Name and Account Code
            My_SQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_NameEng  +'('+ iif( ACCOUNTS.Account_Serial is null,'  ', ACCOUNTS.Account_Serial)+')' as Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account, ACCOUNTS.cannot_del, ACCOUNTS.Account_Serial"
            My_SQL = My_SQL + " FROM ACCOUNTS  WHERE Account_Code <> 'r'  "
            IntColName = 1
        End If
    End If

    If BalanceSheet = True Then
        My_SQL = My_SQL + " and (Account_Code Like 'a1%' OR Account_Code Like 'a2%'  ) "
        My_SQL = My_SQL + "     and (ShowInBlanceSheet is null or ShowInBlanceSheet=1)"
    End If

    My_SQL = My_SQL + " ORDER BY ACCOUNTS.Account_Name "
    Set rs_chart = Nothing
    Set rs_chart = New ADODB.Recordset
    rs_chart.Open My_SQL, Cn, adOpenStatic, adLockReadOnly

    For LngLoop = 0 To rs_chart.RecordCount - 1

        If rs_chart("last_account") Then
            Set nodz = ChartTree.Nodes.Add(Trim(rs_chart("Parent_Account_Code").value) & "G", tvwChild, rs_chart("Account_Code").value, rs_chart(IntColName).value, "Item")
            nodz.Tag = "last"
        Else
            Set nodz = ChartTree.Nodes(CStr(rs_chart("Account_Code").value) & "G")
        End If

        If rs_chart("cannot_del") Then
            nodz.Tag = nodz.Tag + "cannot_del"
        End If

        rs_chart.MoveNext
    Next LngLoop

    rs_chart.Close
    Set rs_chart = Nothing
    'RsOptions.Close
    'Set RsOptions = Nothing
    On Error Resume Next
    ChartTree.Nodes("r").EnsureVisible

    If StrSelectedNode <> "" Then
        ChartTree.SetFocus
        ChartTree.Nodes(StrSelectedNode).EnsureVisible
        ChartTree.Nodes(StrSelectedNode).Selected = True
    End If

    Exit Sub
ErrTrap:
    'Resume
    Msg = "ÕœÀ Œÿ« √À‰«¡  Õ„Ì· «·»Ì«‰« " & CHR(13) & "»—Ã«¡ «·√ ’«· »«·‘—ﬂ… " & CHR(13) & "”Ê› Ì „ €·ﬁ «·»—‰«„Ã"
    Msg = Msg & CHR(13) & Err.Description
    Msg = Msg & CHR(13) & Err.Source
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & "Error Location LoadTreeAccount"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Public Sub LoadTreeAccountBalanceSheetPrint(ChartTree As MSComctlLib.TreeView, _
                                            Optional BalanceSheet As Boolean, _
                                            Optional BalanceSheetId As Integer)
    'ChartTree The TreeView which we will load
    'BolRtl=Make The Tree Right to Left
    Dim My_SQL As String
    Dim nodX As MSComctlLib.Node
    Dim nodz As MSComctlLib.Node
    Dim IntBalSheetShow As Integer
    Dim RsOptions As New ADODB.Recordset
    Dim rs_chart As New ADODB.Recordset
    Dim LngLoop As Long
    Dim BolRtl As Boolean
    Dim IntColName As Integer
    Dim StrSelectedNode As String
    Dim Msg As String
    On Error GoTo ErrTrap

    '=============================== Õ„Ì· «·‘Ã—… «·„Õ«”»Ì…
    IntBalSheetShow = 1
    '--------------------
    'Â–« «·ŒÌ«— ›«∆œ …  ÕœÌœ Â· Ì „ ⁄—÷ «”„ «·Õ”«» ›ﬁÿ
    '√„
    '⁄—÷ √”„ «·Õ”«» Ê«·ﬂÊœ «·„Õ«”»Ï
    'IntBalSheetShow = IIf(IsNull(RsOptions("BalSheetShow").Value), 1, RsOptions("BalSheetShow").Value)
    'MsgBox "IN LoadTreeAccount line 3"

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True
    Else
        BolRtl = False
    End If

    '----Load Data From Data Base
    If SystemOptions.SysDataBaseType = AccessDataBase Then
  
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        My_SQL = " SELECT * "
        My_SQL = My_SQL + " From BalanceSheetViewAccountsQry "
        My_SQL = My_SQL + " WHERE ( last_account=0" & "and(Account_Code <> 'r' or Account_Code <> 'temp') and Parent_Account_Code = 'r' " & "and BalanceSheetId=" & BalanceSheetId & ")"
    End If
 
    Set rs_chart = Nothing
    Set rs_chart = New ADODB.Recordset
    rs_chart.CursorLocation = adUseClient

    rs_chart.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    'MsgBox "IN LoadTreeAccount line 6"
    '----begin to write data
    'Set The TreeView Control
    If ChartTree.Nodes.count > 0 Then
        If Not ChartTree.SelectedItem Is Nothing Then
            StrSelectedNode = ChartTree.SelectedItem.Key
        End If
    End If

    ChartTree.Nodes.Clear
    ChartTree.Sorted = False

    If BolRtl = True Then
        Make_RightToLeft ChartTree
        Set nodX = ChartTree.Nodes.Add(, , "r", "«·œ·Ì· «·„Õ«”»Ï ", "Root")
        IntColName = 2
    Else
        Set nodX = ChartTree.Nodes.Add(, , "r", "Chart Of Accounts", "Root")
        IntColName = 9
    End If

    Call fill_my_Children_Account("r", rs_chart, ChartTree, "BalanceSheetViewAccountsQry", "Parent_Account_Code", " (BalanceSheetViewAccountsQry.last_account)=0 ", IntColName)
    My_SQL = ""

    If BolRtl = True Then 'Load Arabic Name
        If IntBalSheetShow = 1 Then 'Load Arabic Name only
            My_SQL = " SELECT * "
            My_SQL = My_SQL + " From BalanceSheetViewAccountsQry WHERE Account_Code <> 'r' "
        
            IntColName = 2
        Else 'Load Arabic Name and Account Code
            My_SQL = "SELECT Account_Code, Account_Name  +'('+ iif( Account_Serial is null,'  ', Account_Serial)+')' as Account_Name," & "Parent_Account_Code, last_account, Account_Serial"
            My_SQL = My_SQL + " FROM BalanceSheetViewAccountsQry  WHERE Account_Code <> 'r' "
            IntColName = 1
        End If

    ElseIf BolRtl = False Then

        'Load English  Name
        If IntBalSheetShow = 1 Then
            'Load English Name only
            My_SQL = " SELECT * "
            My_SQL = My_SQL + " From BalanceSheetViewAccountsQry WHERE Account_Code <> 'r'   "
            IntColName = 9
        Else
            'Load English Name and Account Code
            My_SQL = "SELECT Account_Code, Account_NameEng  +'('+ iif( Account_Serial is null,'  ', Account_Serial)+')' as Account_Name," & "Parent_Account_Code, last_account,  Account_Serial"
            My_SQL = My_SQL + " FROM BalanceSheetViewAccountsQry  WHERE Account_Code <> 'r'  "
            IntColName = 1
        End If
    End If

    My_SQL = My_SQL + " ORDER BY  Account_Name "
    Set rs_chart = Nothing
    Set rs_chart = New ADODB.Recordset
    rs_chart.Open My_SQL, Cn, adOpenStatic, adLockReadOnly

    For LngLoop = 0 To rs_chart.RecordCount - 1

        If rs_chart("last_account") Then
            Set nodz = ChartTree.Nodes.Add(Trim(rs_chart("Parent_Account_Code").value) & "G", tvwChild, rs_chart("Account_Code").value, rs_chart(IntColName).value, "Item")
            nodz.Tag = "last"
        Else
            Set nodz = ChartTree.Nodes(CStr(rs_chart("Account_Code").value) & "G")
        End If

        If rs_chart("cannot_del") Then
            nodz.Tag = nodz.Tag + "cannot_del"
        End If

        rs_chart.MoveNext
    Next LngLoop

    rs_chart.Close
    Set rs_chart = Nothing
    'RsOptions.Close
    'Set RsOptions = Nothing
    On Error Resume Next
    ChartTree.Nodes("r").EnsureVisible

    If StrSelectedNode <> "" Then
        ChartTree.SetFocus
        ChartTree.Nodes(StrSelectedNode).EnsureVisible
        ChartTree.Nodes(StrSelectedNode).Selected = True
    End If

    Exit Sub
ErrTrap:
    'Resume
    Msg = "ÕœÀ Œÿ« √À‰«¡  Õ„Ì· «·»Ì«‰« " & CHR(13) & "»—Ã«¡ «·√ ’«· »«·‘—ﬂ… " & CHR(13) & "”Ê› Ì „ €·ﬁ «·»—‰«„Ã"
    Msg = Msg & CHR(13) & Err.Description
    Msg = Msg & CHR(13) & Err.Source
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & "Error Location LoadTreeAccount"
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub LoadAssbliedItemNode(XTree As MSComctlLib.TreeView, _
                                 XNode As MSComctlLib.Node, _
                                 LngItemID As Long)

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Dim ZNode As MSComctlLib.Node

    StrSQL = "SELECT dbo.TblItems.ItemName, dbo.TblItemsParts.TableID, dbo.TblItemsParts.ItemID," & "dbo.TblItemsParts.PartItemID, dbo.TblItemsParts.PartItemQty,"
    StrSQL = StrSQL + " dbo.TblItemsParts.PartItemPrice,dbo.TblItems.AssbliedItem"
    StrSQL = StrSQL + " FROM dbo.TblItems INNER JOIN"
    StrSQL = StrSQL + " dbo.TblItemsParts ON dbo.TblItems.ItemID = dbo.TblItemsParts.PartItemID"
    StrSQL = StrSQL + " Where dbo.TblItemsParts.ItemID=" & LngItemID & ""
    StrSQL = StrSQL + " Order BY dbo.TblItemsParts.TableID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst

        For i = 1 To rs.RecordCount

            With XTree
                Set ZNode = .Nodes.Add(XNode.Key, tvwChild, i & "-" & rs("PartItemID").value & "-" & XNode.Key, rs("ItemName").value, "ItemPart")
            End With

            ZNode.Tag = "Item"

            If rs("AssbliedItem").value = True Then
                LoadAssbliedItemNode XTree, ZNode, rs("PartItemID").value
            End If

            rs.MoveNext
        Next i

    End If

End Sub

Private Sub LoadRelatedItemNode(XTree As MSComctlLib.TreeView, _
                                XNode As MSComctlLib.Node, _
                                LngItemID As Long)
    Dim rs     As ADODB.Recordset
    Dim StrSQL As String
    Dim i      As Integer
    Dim ZNode  As MSComctlLib.Node

    'On Error Resume Next
    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT dbo.TblItemsAttach.TableID, dbo.TblItemsAttach.ItemID, dbo.TblItemsAttach.AttachItemID," & "dbo.TblItems.ItemName,    dbo.TblItems.ItemNamee,         dbo.TblItems.RelatedItem,dbo.TblItemsAttach.AttachItemQty," & "dbo.TblItemsAttach.AttachItemPrice "
        StrSQL = StrSQL + " FROM         dbo.TblItems INNER JOIN dbo.TblItemsAttach ON " & "dbo.TblItems.ItemID = dbo.TblItemsAttach.AttachItemID"
        StrSQL = StrSQL + " Where dbo.TblItemsAttach.ItemID=" & LngItemID & ""
        StrSQL = StrSQL + " Order BY dbo.TblItemsAttach.TableID"
    ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "SELECT TblItemsAttach.TableID, TblItemsAttach.ItemID, TblItemsAttach.AttachItemID," & "TblItems.ItemName, TblItems.RelatedItem,TblItemsAttach.AttachItemQty," & "TblItemsAttach.AttachItemPrice "
        StrSQL = StrSQL + " FROM   TblItems INNER JOIN TblItemsAttach ON " & "TblItems.ItemID = TblItemsAttach.AttachItemID"
        StrSQL = StrSQL + " Where TblItemsAttach.ItemID=" & LngItemID & ""
        StrSQL = StrSQL + " Order BY TblItemsAttach.TableID"
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst

        For i = 1 To rs.RecordCount

            With XTree

                '"LinkItem"
                'MsgBox XNode.Key
                If SystemOptions.UserInterface = ArabicInterface Then
                    Set ZNode = .Nodes.Add(XNode.Key, tvwChild, rs("AttachItemID").value & "-" & XNode.Key, rs("ItemName").value, "LinkItem")
                Else
                    Set ZNode = .Nodes.Add(XNode.Key, tvwChild, rs("AttachItemID").value & "-" & XNode.Key, rs("ItemNamee").value, "LinkItem")
                End If

            End With

            ZNode.Tag = "Item"

            If rs("RelatedItem").value = True Then
                LoadRelatedItemNode XTree, ZNode, rs("AttachItemID").value
            End If

            rs.MoveNext
        Next i

    End If

End Sub
