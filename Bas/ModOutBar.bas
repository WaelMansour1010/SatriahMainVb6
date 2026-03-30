Attribute VB_Name = "ModOutBar"
Option Explicit

Public Alink As dxItemLink

Public Agroup As dxGroup

Public Aitem As dxItem

Public DatName As String

Public SysOutBar As dxSideBar
Const SHGFI_ICON = &H100
Const SHGFI_LARGEICON = &H0
Const SHGFI_SMALLICON = &H1
Const MAX_PATH = 260
Const SM_CYFRAME = 33
Const SM_CXFRAME = 32
Const SM_CYCAPTION = 4

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Private Declare Function SHGetFileInfo _
                Lib "shell32.dll" _
                Alias "SHGetFileInfoA" (ByVal pszPath As String, _
                                        ByVal dwFileAttributes As Long, _
                                        psfi As SHFILEINFO, _
                                        ByVal cbFileInfo As Long, _
                                        ByVal uFlags As Long) As Long

Private Declare Function ImageList_AddIcon _
                Lib "COMCTL32" (ByVal himl As Long, _
                                ByVal hIcon As Long) As Long

Public Function UserGroup(pGroup As DXSIDEBARLibCtl.dxGroup) As Boolean

    If Not pGroup Is Nothing Then
        If pGroup.UserData >= 1 Or InStr(1, left(pGroup.ObjectName, 4), "User", vbTextCompare) <> 0 Then
            UserGroup = True
        Else
            UserGroup = False
        End If
    End If

End Function

Private Function AddIcon(ByVal FileName As String)
    Dim hList As Long
    Dim Info As SHFILEINFO
    Dim a As Long, b As Long
    hList = SHGetFileInfo(FileName, 0, Info, 352, SHGFI_ICON Or SHGFI_LARGEICON)
    a = ImageList_AddIcon(SysOutBar.GetImageListByName("dxImageList3"), Info.hIcon)
    hList = SHGetFileInfo(FileName, 0, Info, 352, SHGFI_ICON Or SHGFI_SMALLICON)
    b = ImageList_AddIcon(SysOutBar.GetImageListByName("dxImageList3"), Info.hIcon)
    AddIcon = Array(a, b)
End Function

Private Function AnyWay(ByVal s As String) As String
    Dim AppP As String
    AnyWay = s
    AppP = App.path

    If Dir(AppP & "\Account.exe") <> "" Then
        If VBA.left(s, 3) = "..\" Then
            AnyWay = VBA.left(AppP, Len(AppP) - (Len(AppP) - InStrRev(AppP, "\") + 1)) & VBA.right(s, Len(s) - 2)
        End If

        If VBA.left(s, 2) = ".\" Then
            AnyWay = AppP & VBA.right(s, Len(s) - 1)
        End If
    End If

End Function

Private Sub File_Path_Name(ByVal s As String, _
                           PathName As String, _
                           FileName As String)
    Dim a As Byte
    a = InStrRev(s, "\")
    PathName = left(s, a)
    FileName = right(s, Len(s) - a)
End Sub

Public Sub AddItem_Link()

    Dim PName As String, fname As String
    Dim NumImage
    Dim Msg As String

    On Error GoTo ErrHandler

    With mdifrmmain.cmDlg
        .Filter = "Programs|*.exe;*.com;*.bat;*.lnk|Links|*.url|Spreadsheets|*.xls|Documents|*.doc|All files|*.*"
        .Flags = cdlOFNHideReadOnly
        .ShowOpen

        If .FileName <> "" Then
            File_Path_Name .FileName, PName, fname

            If IsOutBarExistItem("User" & PName & fname) = True Then
                Msg = "ÌÊÃœ ≈Œ ’«— „ÊÃÊœ „”»ﬁ« ·Â–« «·„·› ..!!"
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            Set Aitem = SysOutBar.Items.Add
            NumImage = AddIcon(PName & fname)
            Aitem.Caption = PName & fname
            Aitem.UserData = SysOutBar.Items.count - 1
            Aitem.ItemLargeImage = NumImage(0)
            Aitem.ItemSmallImage = NumImage(1)
            Aitem.ObjectName = "User" & PName & fname
            Set Alink = Agroup.Links.Add
            Alink.Caption = fname
            Alink.DefaultCaption = False
            Set Alink.Item = Aitem
            SysOutBar.EditItemLinkCaption Alink
        End If

    End With

ErrHandler:
    Exit Sub
End Sub

Public Sub LoadOutBarData(SysOutBar As dxSideBar)
    Dim i As Long, J As Long
    Dim ICount As Integer, GCount As Integer, ActiveG As Integer
    Dim IntItemsInterval As Integer
    Dim a1, a2, a3, a4, a5, StayTop As Byte, TopHeader As Boolean
    Dim NumImage
    Dim StrLine As String
    Dim IntFreeFile As Integer
    Dim VarTemp As Variant
    On Error Resume Next
    IntFreeFile = FreeFile

    If Dir(DatName) <> "" Then
        Open DatName For Input As #IntFreeFile

        With SysOutBar
            'main
            '1-Load  Main  Info Items Count,Group Count.... ect
            Line Input #IntFreeFile, StrLine
            VarTemp = Split(StrLine, ",", , vbTextCompare)
            ICount = val(VarTemp(0))
            GCount = val(VarTemp(1))
            ActiveG = val(VarTemp(2))
            
            IntItemsInterval = val(VarTemp(3))
            .ItemInterval = IntItemsInterval
            
            'items
            If ICount > 0 Then

                For i = 0 To ICount - 1
                    '.Items(I).Caption, .IndexOf(.Items(I)), SysOutBar.Items(I).UserData, SysOutBar.Items(I).ObjectName
                    Line Input #IntFreeFile, StrLine
                    VarTemp = Split(StrLine, ",", , vbTextCompare)
                    VarTemp(0) = AnyWay(VarTemp(0))
                    Set Aitem = .Items.Add
                    Aitem.Caption = VarTemp(0)
                    Aitem.ObjectName = VarTemp(3)
                    NumImage = AddIcon(VarTemp(0))
                    Aitem.UserData = VarTemp(1)

                    If NumImage(0) = -1 Then NumImage(0) = 0
                    If NumImage(1) = -1 Then NumImage(1) = 0
                    Aitem.ItemLargeImage = NumImage(0)
                    Aitem.ItemSmallImage = NumImage(1)
                Next

            End If

            'groups
            If GCount > 0 Then

                For i = 0 To GCount - 1
                    'Write #IntFreeFile, .Groups(I).Caption, .Groups(I).Links.count - 1, _
                     .Groups(I).ItemsStyle, .Groups(I).UserData, .Groups(I).ObjectName
                    
                    Line Input #IntFreeFile, StrLine
                    VarTemp = Split(StrLine, ",", , vbTextCompare)
                    Set Agroup = .Groups.Add
                    Agroup.Caption = VarTemp(0)
                    Agroup.ItemsStyle = VarTemp(2)
                    Agroup.UserData = VarTemp(3)
                    Agroup.ObjectName = VarTemp(4)
                    Agroup.HighLightColor = &HFF0000

                    'AddBackGround Agroup, a4
                    If i = ActiveG Then Set .ActiveGroup = Agroup

                    For J = 0 To val(VarTemp(1)) - 1
                        'StrLine = .Groups(I).Links(j).Caption & "," & .IndexOf(.Groups(I).Links(j).Item) '.UserData
                        
                        Line Input #IntFreeFile, StrLine 'a3, a4
                        VarTemp = Split(StrLine, ",", , vbTextCompare)
                        Set Alink = Agroup.Links.Add
                        Alink.Caption = VarTemp(0)
                        Alink.DefaultCaption = False
                        Set Alink.Item = ItemForData(VarTemp(1))
                        Alink.FileName = Alink.Item.Caption
                    Next
                Next

            End If

        End With

        Close #IntFreeFile
    Else
        'MsgBox DatName & " is not found."
        'Set Agroup = SysOutBar.Groups.Add
        'Agroup.Caption = "FirstGroup"
        'AddBackGround Agroup, 0
    End If

End Sub

Public Sub DeleteItem(ByVal Index As Integer)
    Dim i As Integer, J As Integer, FDel As Boolean
    FDel = True

    With SysOutBar

        For i = 0 To .Groups.count - 1
            For J = 0 To .Groups(i).Links.count - 1

                If .Groups(i).Links(J).Item.UserData = Index Then
                    FDel = False
                    Exit Sub
                End If

            Next
        Next

        If FDel Then
            .Items.Remove (Index)
        End If

    End With

End Sub

Private Function ItemForData(ByVal d) As dxItem
    Dim i As Long
    Set ItemForData = Nothing

    For i = 0 To SysOutBar.Items.count - 1

        If d = SysOutBar.Items(i).UserData Then
            Set ItemForData = SysOutBar.Items(i)
            Exit Function
        End If

    Next

End Function

Public Function GetOurBarUserGroupCount() As Integer
    Dim i As Integer
    Dim IntCount As Integer

    For i = 0 To SysOutBar.Groups.count - 1

        If InStr(1, left(SysOutBar.Groups(i).ObjectName, 4), "User", vbTextCompare) <> 0 Then
            IntCount = IntCount + 1
        End If

    Next i

    GetOurBarUserGroupCount = IntCount
End Function

Private Function GetOurBarUserItemsCount() As Integer
    Dim i As Integer
    Dim IntCount As Integer

    For i = 0 To SysOutBar.Items.count - 1

        If InStr(1, left(SysOutBar.Items(i).ObjectName, 4), "user", vbTextCompare) <> 0 Then
            IntCount = IntCount + 1
        End If

    Next i

    GetOurBarUserItemsCount = IntCount
End Function

Private Function IsOutBarExistItem(StrObjectName As String) As Boolean
    Dim i As Integer

    For i = 0 To SysOutBar.Items.count - 1

        If StrComp(SysOutBar.Items(i).ObjectName, StrObjectName, vbTextCompare) = 0 Then
            IsOutBarExistItem = True
            Exit Function
        End If

    Next i

    IsOutBarExistItem = False
End Function

Public Function UserItem(xItem As DXSIDEBARLibCtl.dxItem) As Boolean

    If InStr(1, left(xItem.ObjectName, 4), "user", vbTextCompare) <> 0 Then
        UserItem = True
    Else
        UserItem = False
    End If

End Function

Public Sub SaveOurBarData()
    Dim i As Long, J As Long
    Dim g As dxGroup
    Dim l As dxItemLink
    Dim IntGCount As Integer
    Dim IntICount As Integer
    Dim IntFreeFile As Integer
    Dim StrLine As String

    On Error Resume Next

    If Dir(DatName, vbNormal) <> "" Then
        Kill DatName
    End If

    IntGCount = GetOurBarUserGroupCount
    IntICount = GetOurBarUserItemsCount
    IntFreeFile = FreeFile
    'Open StrLogFileName For Output As #IntFreeFile
    '    Print #IntFreeFile, SS
    'Close #IntFreeFile

    Open DatName For Output As #IntFreeFile

    With SysOutBar
        'main
        '1-Save Main  Info Items Count,Group Count.... ect
        StrLine = IntICount & "," & IntGCount & "," & .ActiveGroup.Index & "," & .ItemInterval
        Print #IntFreeFile, StrLine

        'items
        If IntICount > 0 Then

            For i = 0 To .Items.count - 1

                If UserItem(.Items(i)) = True Then
                    StrLine = .Items(i).Caption & "," & .IndexOf(.Items(i)) & "," & SysOutBar.Items(i).UserData & "," & SysOutBar.Items(i).ObjectName
                    Print #IntFreeFile, StrLine
                End If

            Next

        End If

        'groups
        If IntGCount > 0 Then

            For i = 0 To .Groups.count - 1

                If UserGroup(.Groups(i)) = True Then
                    StrLine = .Groups(i).Caption & ", " & .Groups(i).Links.count & "," & .Groups(i).ItemsStyle & "," & .Groups(i).UserData & "," & .Groups(i).ObjectName
                    Print #IntFreeFile, StrLine

                    For J = 0 To .Groups(i).Links.count - 1
                        StrLine = .Groups(i).Links(J).Caption & "," & .IndexOf(.Groups(i).Links(J).Item) '.UserData
                        Print #IntFreeFile, StrLine
                    Next

                End If

            Next

        End If

    End With

    Close #IntFreeFile
End Sub

Public Sub AddNewGroup()
    Dim IntNewUserGroupIndex As Integer
    IntNewUserGroupIndex = GetOurBarUserGroupCount + 1
    Set Agroup = SysOutBar.Groups.Add
    Agroup.ObjectName = "UserGroup" & IntNewUserGroupIndex
    Agroup.Caption = "„Ã„Ê⁄… ÃœÌœ…"
    Agroup.UserData = IntNewUserGroupIndex
    'AddBackGround Agroup, 0
    SysOutBar.EditGroupCaption Agroup
End Sub

Public Sub EditGroup()
    SysOutBar.EditGroupCaption Agroup
End Sub

Public Sub DeleteGroup()
    Dim Msg As String
    Dim IntMsgRes As Integer
    Dim IntGroupLinks As Integer
    Dim xTemp As dxItem

    Msg = "”Ê› Ì „ Õ–› Â–Â «·„Ã„Ê⁄… ..."
    Msg = Msg & Chr(13) & Agroup.Caption
    Msg = Msg & Chr(13) & "Â· «‰  „ «ﬂœ „‰ ⁄„·Ì… «·Õ–›"
    IntMsgRes = MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)

    If IntMsgRes = vbYes Then
        IntGroupLinks = Agroup.Links.count

        Do While Agroup.Links.count > 0
            Set xTemp = Agroup.Links(0).Item
            Agroup.Links.Remove (Agroup.IndexOf(Agroup.Links(0)))
            SysOutBar.Items.Remove SysOutBar.IndexOf(xTemp)
        Loop

    End If

    SysOutBar.Groups.Remove (Agroup.Index)
End Sub

Public Sub EditItemLink()
    SysOutBar.EditItemLinkCaption Alink
End Sub

Public Sub RemoveItemLink()
    Dim xTemp As dxItem
    Set xTemp = Alink.Item
    Agroup.Links.Remove (Agroup.IndexOf(Alink))
    SysOutBar.Items.Remove SysOutBar.IndexOf(xTemp)
End Sub
