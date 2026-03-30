VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPaneTree 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   0  'None
   Caption         =   "ÔÌÑÉ ÇáÃÕäÇÝ"
   ClientHeight    =   6255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3540
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.TreeView TrvItems 
      Height          =   5145
      Left            =   240
      TabIndex        =   0
      Top             =   90
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   9075
      _Version        =   393217
      Indentation     =   18
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "FrmPaneTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim XTip As FrmItemTreeTip

Private m_GroupsSort As String

Private m_ItemsSort As String

Private WithEvents m_MenuRefresh As Menu
Attribute m_MenuRefresh.VB_VarHelpID = -1

Private Sub HideComment()

    If Not XTip Is Nothing Then
        Unload XTip
        Set XTip = Nothing
        'm_QtyToolTipShown = False
    End If

End Sub

Private Sub Form_Activate()
    'PutFormOnTop Me.hwnd
End Sub

Private Sub Form_Load()
    Dim StrTemp As String
    Dim intIndex As Integer

    With Me.TrvItems
        .LineStyle = tvwRootLines
        .Sorted = False
        .LabelEdit = tvwManual
        .OLEDragMode = ccOLEDragAutomatic
        .OLEDropMode = ccOLEDropNone
        Set .ImageList = mdifrmmain.ImgLstTree
    End With

    '------------------------------------------
    intIndex = GetSetting(SystemOptions.SysRegsAppPath & "\DockingPanes\" & Me.name, "SortSetting", "GroupsSortIndex", 6)
    mdifrmmain.MPITP_GSort_Option(intIndex).Checked = True

    intIndex = GetSetting(SystemOptions.SysRegsAppPath & "\DockingPanes\" & Me.name, "SortSetting", "ItemsSortIndex", 6)
    mdifrmmain.MPITP_ISort_Option(intIndex).Checked = True
    StrTemp = GetSetting(SystemOptions.SysRegsAppPath & "\DockingPanes\" & Me.name, "SortSetting", "ItemsSort", " ItemName ASC")
    Me.ItemsSort = StrTemp

    StrTemp = GetSetting(SystemOptions.SysRegsAppPath & "\DockingPanes\" & Me.name, "SortSetting", "GroupsSort", " GroupName ASC ")
    Me.GroupsSort = StrTemp

    '----------------------------------------
    If SystemOptions.UserInterface = ArabicInterface Then
        Make_RightToLeft Me.TrvItems
    End If

    LoadData Me.GroupsSort, Me.ItemsSort
End Sub

Private Sub Form_Resize()

    With Me.TrvItems
        .Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight
    End With

End Sub

Public Sub LoadData(Optional StrGroupsSort As String = "", _
                    Optional StrItemsSort As String = "")
    Dim BolSaved As Boolean

    If Me.TrvItems.Nodes.count > 0 Then
        SaveNodesStatus
        BolSaved = True
    End If

    Me.TrvItems.Nodes.Clear
    ModTree.LoadTreeGroups Me.TrvItems, StrGroupsSort, StrItemsSort

    If BolSaved = True Then
        LoadNodesStatus
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    HideComment
    SaveSetting SystemOptions.SysRegsAppPath & "\DockingPanes\" & Me.name, "SortSetting", "ItemsSort", Me.ItemsSort
    SaveSetting SystemOptions.SysRegsAppPath & "\DockingPanes\" & Me.name, "SortSetting", "GroupsSort", Me.GroupsSort

    For i = mdifrmmain.MPITP_GSort_Option.LBound To mdifrmmain.MPITP_GSort_Option.UBound

        If mdifrmmain.MPITP_GSort_Option(i).Checked = True Then
            Exit For
        End If

    Next i

    SaveSetting SystemOptions.SysRegsAppPath & "\DockingPanes\" & Me.name, "SortSetting", "GroupsSortIndex", i

    For i = mdifrmmain.MPITP_ISort_Option.LBound To mdifrmmain.MPITP_ISort_Option.UBound

        If mdifrmmain.MPITP_ISort_Option(i).Checked = True Then
            Exit For
        End If

    Next i

    SaveSetting SystemOptions.SysRegsAppPath & "\DockingPanes\" & Me.name, "SortSetting", "ItemsSortIndex", i

End Sub

Private Sub m_MenuRefresh_Click()
    LoadData Me.GroupsSort, Me.ItemsSort
End Sub

Private Sub TrvItems_Collapse(ByVal Node As MSComctlLib.Node)

    If Not Node Is Nothing Then
        Node.ForeColor = vbBlack
        Node.Bold = False
    End If

End Sub

Private Sub TrvItems_Expand(ByVal Node As MSComctlLib.Node)

    If Not Node Is Nothing Then
        Node.ForeColor = vbRed
        Node.Bold = True
    End If

End Sub

Private Sub TrvItems_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               Y As Single)
    Dim Node As MSComctlLib.Node
    Dim uPoint As POINTAPI
    Dim lLeft  As Single
    Dim LTop As Single
    Dim LngTaskBarHeight As Long
    Dim LngItemID As Long

    If Button <> vbLeftButton Then
        Exit Sub
    End If

    Set Node = TrvItems.HitTest(x, Y)

    If Not Node Is Nothing Then
        If right(Node.key, 1) = "I" Then
            LngItemID = val(Node.key)
            'ClientToScreen TrvItems.hwnd, uPoint
            HideComment
            GetCursorPos uPoint
            lLeft = (uPoint.x) * Screen.TwipsPerPixelX

            If SystemOptions.UserInterface = ArabicInterface Then
                lLeft = lLeft + x
            Else
                lLeft = lLeft + (TrvItems.Width - x)
            End If

            Set XTip = New FrmItemTreeTip

            If (lLeft + XTip.Width) > Screen.Width Then
                lLeft = lLeft - (XTip.Width + TrvItems.Width)
            End If

            '======================================
            LTop = uPoint.Y * Screen.TwipsPerPixelY
            '--------------------------------------
            LngTaskBarHeight = GetTaskBarHeight

            '--------------------------------------
            If (LTop + XTip.Height + LngTaskBarHeight) > Screen.Height Then
                LTop = Screen.Height - (XTip.Height + LngTaskBarHeight)
            End If

            XTip.left = lLeft
            XTip.top = LTop
        
            XTip.LoadData LngItemID
            XTip.DialogAction fadeInOut
            ShowWindow XTip.hWnd, SW_SHOWNA
        Else
        End If
    End If

End Sub

Private Sub TrvItems_MouseUp(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             Y As Single)
    Dim XNode  As MSComctlLib.Node
    Dim LngItemID As Long

    If Button = vbRightButton Then
        Set XNode = TrvItems.SelectedItem

        If XNode Is Nothing Then
            Exit Sub
        Else
            LngItemID = val(XNode.key)
            SetItemReportsMenu LngItemID, Me
        End If
    End If

End Sub

Public Property Get GroupsSort() As String
    GroupsSort = m_GroupsSort
End Property

Public Property Let GroupsSort(ByVal vNewValue As String)
    m_GroupsSort = vNewValue
End Property

Public Property Get ItemsSort() As String
    ItemsSort = m_ItemsSort
End Property

Public Property Let ItemsSort(ByVal vNewValue As String)
    m_ItemsSort = vNewValue
End Property

Private Sub SaveNodesStatus()
    Dim StrFileName As String
    Dim IntFreeFile As Integer
    Dim StrLine As String
    Dim i As Long

    StrFileName = App.path & "\TempNodeStatus.txt"

    If Dir(StrFileName) <> "" Then
        Kill StrFileName
    End If

    IntFreeFile = FreeFile

    Open StrFileName For Append As #IntFreeFile
    StrLine = Me.TrvItems.Nodes("r").key & "," & Me.TrvItems.Nodes("r").Expanded
    Print #IntFreeFile, StrLine

    For i = 1 To Me.TrvItems.Nodes.count

        'Save Groups only
        If InStr(1, Me.TrvItems.Nodes(i).key, "G", vbTextCompare) <> 0 Then
            StrLine = Me.TrvItems.Nodes(i).key & "," & Me.TrvItems.Nodes(i).Expanded
            Print #IntFreeFile, StrLine
        End If

    Next i

    Close #IntFreeFile
End Sub

Private Sub LoadNodesStatus()
    Dim StrFileName As String
    Dim IntFreeFile As Integer
    Dim StrLine As String
    Dim i As Long
    Dim VarTemp As Variant

    StrFileName = App.path & "\TempNodeStatus.txt"

    If Dir(StrFileName) = "" Then
        Exit Sub
    End If

    IntFreeFile = FreeFile
    Open StrFileName For Input As #IntFreeFile

    Do While Not EOF(IntFreeFile)
        Line Input #1, StrLine

        If Trim$(StrLine) <> "" Then
            VarTemp = Split(StrLine, ",", , vbTextCompare)
            Me.TrvItems.Nodes(VarTemp(0)).Expanded = VarTemp(1)
        End If

    Loop

    Close #IntFreeFile

End Sub
