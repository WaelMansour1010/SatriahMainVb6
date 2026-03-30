Attribute VB_Name = "ModFile"
 Option Explicit

Private Declare Function CreateStreamOnHGlobal _
                Lib "ole32" (ByVal hGlobal As Long, _
                             ByVal fDeleteOnRelease As Long, _
                             ppstm As Any) As Long

Private Declare Function OleLoadPicture _
                Lib "olepro32" (pStream As Any, _
                                ByVal lSize As Long, _
                                ByVal fRunmode As Long, _
                                riid As Any, _
                                ppvObj As Any) As Long

Private Declare Function CLSIDFromString _
                Lib "ole32" (ByVal lpsz As Any, _
                             pclsid As Any) As Long

Private Declare Function GlobalAlloc _
                Lib "kernel32" (ByVal uFlags As Long, _
                                ByVal dwBytes As Long) As Long

Private Declare Function GlobalLock _
                Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function GlobalUnlock _
                Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Sub MoveMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (pDest As Any, _
                                       pSource As Any, _
                                       ByVal dwLength As Long)

Type FileComp
    
    RecNum As Long
    'CtlType As String * 10
    CtlProp As String * 1000
    PageSize As String * 20
End Type
Global FSave As FileComp

Public ColObj As New Collection

Public m_PagSize As String

Public PicWidth As Single

Public PicHeight As Single

Public crep As ClsReportProp
Dim LstObj As ListBox


' Clipboard routines.
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData _
                Lib "user32" (ByVal wFormat As Long, _
                              ByVal hMem As Long) As Long
Private Declare Function DragQueryFile _
                Lib "shell32.dll" _
                Alias "DragQueryFileA" (ByVal drop_handle As Long, _
                                        ByVal UINT As Long, _
                                        ByVal lpStr As String, _
                                        ByVal Ch As Long) As Long

' File list clipboard format code.
Private Const CF_HDROP = 15

' DROPFILES data structure.
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type DROPFILES
    pFiles As Long
    pt As POINTAPI
    fNC As Long
    fWide As Long
End Type

' Global memory routines.
                                
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
'Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMem _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (Destination As Any, _
                                       Source As Any, _
                                       ByVal Length As Long)

' Global Memory Flags
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)


'*********************
'************************
Private Const LOCALE_SDATE                 As Long = &H1D    'date separator
Private Const LOCALE_STIME                 As Long = &H1E    'time separator
Private Const LOCALE_SSHORTDATE            As Long = &H1F    'short date format string
Private Const LOCALE_SLONGDATE             As Long = &H20    'long date format string
Private Const LOCALE_STIMEFORMAT           As Long = &H1003  'time format string
Private Const LOCALE_IDATE                 As Long = &H21    'short date format ordering
Private Const LOCALE_ILDATE                As Long = &H22    'long date format ordering
Private Const LOCALE_ITIME                 As Long = &H23    'time format specifier
 
Private Declare Function SetLocaleInfo& Lib "kernel32" Alias "SetLocaleInfoA" (ByVal _
Locale As Long, ByVal LCType As Long, ByVal lpLCData As String)
'************************
'**********************



Public Function ClipboardSetFile(file_name As String) As Boolean
    Dim file_string    As String
    Dim drop_files     As DROPFILES
    Dim memory_handle  As Long
    Dim memory_pointer As Long

    '  √ﬂœ „‰ √‰ «·Õ«›Ÿ… ›«—€…
    Clipboard.Clear

    ' › Õ «·Õ«›Ÿ…
    If OpenClipboard(0) Then
        
        ' ≈⁄œ«œ ”·”·… «·„·› » ‰”Ìﬁ „‰«”» „⁄ NULL ›Ì «·‰Â«Ì…
        file_string = file_name & vbNullChar & vbNullChar
    
        ' ≈⁄œ«œ DROPFILES structure
        drop_files.pFiles = Len(drop_files)
        drop_files.fWide = 0 ' ANSI format, 0 ·€Ì— UTF-16

        '  Œ’Ì’ «·–«ﬂ—… «··«“„…
        memory_handle = GlobalAlloc(GHND, Len(drop_files) + Len(file_string))
        If memory_handle Then

            ' ﬁ›· «·–«ﬂ—… ··Õ’Ê· ⁄·Ï „ƒ‘— ··–«ﬂ—… «·„Œ’’…
            memory_pointer = GlobalLock(memory_handle)

            ' ‰”Œ DROPFILES structure ≈·Ï «·–«ﬂ—…
            CopyMem ByVal memory_pointer, drop_files, Len(drop_files)
            ' ‰”Œ „”«— «·„·› ≈·Ï «·–«ﬂ—… »⁄œ DROPFILES structure
            CopyMem ByVal memory_pointer + Len(drop_files), ByVal file_string, Len(file_string)

            ' ›ﬂ ﬁ›· «·–«ﬂ—…
            GlobalUnlock memory_handle

            ' ‰”Œ «·»Ì«‰«  ≈·Ï «·Õ«›Ÿ…
            SetClipboardData CF_HDROP, memory_handle
            ClipboardSetFile = True
        End If
        
        ' ≈€·«ﬁ «·Õ«›Ÿ…
        CloseClipboard
    End If
End Function

Public Function ClipboardSetFiles(file_names() As String) As Boolean
    Dim file_string    As String
    Dim drop_files     As DROPFILES
    Dim memory_handle  As Long
    Dim memory_pointer As Long
    Dim i              As Long
   
    Clipboard.Clear
   
    If OpenClipboard(0) Then
       
        For i = LBound(file_names) To UBound(file_names)
            file_string = file_string & file_names(i) & vbNullChar
        Next
        file_string = file_string & vbNullChar
    
        drop_files.pFiles = Len(drop_files)
        drop_files.fWide = 0
        drop_files.fNC = 0

        memory_handle = GlobalAlloc(GHND, Len(drop_files) + Len(file_string))
        If memory_handle Then

            memory_pointer = GlobalLock(memory_handle)

            CopyMem ByVal memory_pointer, drop_files, Len(drop_files)
            CopyMem ByVal memory_pointer + Len(drop_files), ByVal file_string, Len(file_string)
            GlobalUnlock memory_handle

            ' Copy the data to the clipboard.
            SetClipboardData CF_HDROP, memory_handle
            ClipboardSetFiles = True
        End If
      
        CloseClipboard
    End If
End Function
' Copy the file names to the clipboard.



Public Function SaveToFile(ByVal FilePath As String, _
                           CtrlLst As ListBox) As Boolean
    ' save data to file on disk
    On Error GoTo ErrTrap
    Dim FrFile As Integer
    Dim ii As Integer
    Dim FileExist As String
    SaveToFile = False
    FrFile = FreeFile
    FileExist = Dir(FilePath, vbNormal)

    If FileExist <> "" Then
        Kill FilePath
    End If

    Open FilePath For Random As #FrFile Len = Len(FSave)

    For ii = 0 To CtrlLst.ListCount - 1
        FSave.RecNum = ii
        FSave.PageSize = m_PagSize
        FSave.CtlProp = Trim(CtrlLst.List(ii))
    
        Put #FrFile, ii + 1, FSave
    Next

    Close #FrFile
    Exit Function
ErrTrap:
End Function

Public Function LoadFile(ByVal FilePath As String, _
                         Frm As Form)
    'load file From disk
    On Error Resume Next
    Set LstObj = Frm.LstCtrl
    Dim FrFile As Integer
    Dim ii As Integer
    Dim TxtSpelt() As String
    Dim Ctrl As Control
    LoadFile = False
    FrFile = FreeFile
    ii = 0
    LstObj.Clear
    Open FilePath For Random As #FrFile Len = Len(FSave)
    Get #FrFile, 1, FSave
    TxtSpelt = Split(Trim(FSave.PageSize), "~")

    Do Until EOF(FrFile)
        ii = ii + 1
        Get #FrFile, ii, FSave
        LstObj.AddItem FSave.CtlProp
        On Error Resume Next
        Frm.PicMain.Move 300, 300, val(TxtSpelt(0)), val(TxtSpelt(1))
        PicWidth = val(TxtSpelt(0))
        PicHeight = val(TxtSpelt(1))
    Loop

    Close #FrFile

    For ii = 0 To LstObj.ListCount - 1

        If LstObj.List(ii) <> "" Then
            TxtSpelt = Split(LstObj.List(ii), "**")

            If Trim(TxtSpelt(0)) <> "" Then

                With Frm
                
                    If TxtSpelt(0) = "Label" Then
                        Load .LblText(.LblText.count)
                        Set Ctrl = .LblText(.LblText.count - 1)
                        SetCtrlProperty TxtSpelt(1), .LblText(.LblText.count - 1)
                    ElseIf TxtSpelt(0) = "Image" Then
                        Load .CoLog(.CoLog.count)
                        Set Ctrl = .CoLog(.CoLog.count - 1)
                        SetCtrlProperty TxtSpelt(1), Ctrl
                    
                        GoSub LoadImage
                    ElseIf TxtSpelt(0) = "Grd" Then
                        SetCtrlProperty TxtSpelt(1), .grd(0)
                        Set Ctrl = .grd(0)
                    End If

                    On Error Resume Next

                    '                .CoLog(0) = LoadPicture("")
                    '                .CoLog(0) = LoadPicture(Replace(FilePath, ".drp", ".img"))
                    If Frm.Name = "FrmDesigner" Then
                        Frm.LstZorder.AddItem Ctrl.Name & "," & Ctrl.index
                    End If

                    Ctrl.ZOrder 0
                End With

            End If
        End If

    Next

    Exit Function
    '---------------------------------- Load Images to Objects --------------------------
LoadImage:
    Dim B_Read() As Byte
    Dim OutFile As String
    OutFile = Replace(FilePath, ".drp", ".dmg")
    Open OutFile For Binary Access Read As #1
    ReDim B_Read(0 To LOF(1) - 1)
    Get #1, val(Ctrl.Tag), B_Read
    Debug.Print Seek(1)
    Close #1
    'Set Me.Picture1.Picture = Nothing
    Set Ctrl.Picture = PictureFromByteStream(B_Read)
    Return
End Function

Public Function SetCtrlProperty(Prop As String, _
                                Optional Objct As Control)
    On Error Resume Next
    '·«⁄«œ… «·Œ’«∆’ ··«œ«…
    'Dim Propertz As Variant
    Dim ColInx As Integer
    Dim Frm As Form
    Set Frm = FrmPreview
    Dim CtrlProp() As String
    Dim Obj As Control
    Dim IX As Integer
    CtrlProp = Split(Prop, "~^")
    '                   0      ,1     ,  2    ,  3   ,   4    ,   5    ,    6
    'Propertz = Array("Item", "Index", "Left", "Top", "Width", "Height", "Visible", _
     "AutoSize", "Caption", "BackColor", "ForeColor", "FontName", _
     "FontBold", "FontItalic", "FontSize", "FontUnderline", "Alignment", _
     "RightToLeft", "BackStyle", "TextShadow", "TextShadowDepth", _
     "ColorTextShadow", "RotationAngle", "ColorShadow", _
     , "Zorder", "Tag")

    If Objct Is Nothing Then
        Set Obj = ColObj.Item(Int(CtrlProp(UBound(CtrlProp))))
    Else
        Set Obj = Objct
    End If

    Obj.Visible = False
    Obj.left = CtrlProp(2)
    Obj.top = CtrlProp(3)
    Obj.Width = CtrlProp(4)
    Obj.Height = CtrlProp(5)

    If TypeOf Obj Is ISAniLabel Then
        Obj.AutoSize = CtrlProp(7)
        Obj.Caption = CtrlProp(8)
        Obj.backcolor = CtrlProp(9)
        Obj.ForeColor = CtrlProp(10)
        Obj.FontName = CtrlProp(11)
        Obj.FontBold = CtrlProp(12)
        Obj.FontItalic = CtrlProp(13)
        Obj.fontsize = CtrlProp(14)
        Obj.FontUnderline = CtrlProp(15)
        Obj.Alignment = CtrlProp(16)
        Obj.RightToLeft = CtrlProp(17)
        Obj.BackStyle = CtrlProp(18)
        Obj.TextShadow = CtrlProp(19)
        Obj.TextShadowDepth = CtrlProp(20)
        Obj.ColorTextShadow = CtrlProp(21)
        Obj.RotationAngle = CtrlProp(22)
        Obj.ColorShadow = CtrlProp(23)
        Obj.ZOrder val(Trim(CtrlProp(24))) - 55
        Obj.Tag = CtrlProp(25)
    ElseIf TypeOf Obj Is Image Then
        Obj.Tag = CtrlProp(7)
    ElseIf TypeOf Obj Is VSFlex8Ctl.VSFlexGrid Then
        Obj.Cols = CtrlProp(7)
        Obj.Appearance = CtrlProp(8)
        Obj.FontName = CtrlProp(9)
        Obj.fontsize = CtrlProp(10)
        Obj.FontBold = CtrlProp(11)
        Obj.FontItalic = CtrlProp(12)
        Frm.GridColor.ForeColor = CtrlProp(13)
        Obj.ForeColor = FrmPreview.GridColor.ForeColor
        Frm.GridColor.ForeColor = CtrlProp(14)
        Obj.ForeColorFixed = FrmPreview.GridColor.ForeColor
        Frm.GridColor.ForeColor = CtrlProp(15)
        Obj.backcolor = FrmPreview.GridColor.ForeColor
        Frm.GridColor.ForeColor = CtrlProp(16)
        Obj.BackColorAlternate = FrmPreview.GridColor.ForeColor
        Frm.GridColor.ForeColor = CtrlProp(17)
        Obj.BackColorFixed = FrmPreview.GridColor.ForeColor
        Frm.GridColor.ForeColor = CtrlProp(18)
        Obj.GridColor = FrmPreview.GridColor.ForeColor
        Frm.GridColor.ForeColor = CtrlProp(19)
        Obj.GridColorFixed = FrmPreview.GridColor.ForeColor
        Frm.GridColor.ForeColor = CtrlProp(20)
        Obj.SheetBorder = FrmPreview.GridColor.ForeColor
        Obj.BorderStyle = CtrlProp(21)
        Obj.AutoSizeMode = CtrlProp(22)
        Obj.GridLines = CtrlProp(23)
        Obj.GridLinesFixed = CtrlProp(24)
        Obj.TextStyleFixed = CtrlProp(25)
        Obj.TextStyle = CtrlProp(26)
        Obj.GridLineWidth = CtrlProp(27)
        ColInx = 0

        For IX = 0 To Obj.Cols - 1
            Obj.ColKey(IX) = CtrlProp(29 + ColInx)
            Obj.ColWidth(IX) = CtrlProp(30 + ColInx)
            Obj.TextMatrix(0, IX) = CtrlProp(31 + ColInx)
            Obj.FixedAlignment(IX) = CtrlProp(32 + ColInx)
            Obj.ColAlignment(IX) = CtrlProp(33 + ColInx)
            ColInx = ColInx + 6
        Next

    End If
    
    Obj.Visible = Trim(CtrlProp(6))
End Function

Public Function GetCtrlProperty(Obj As Control) As String
    On Error Resume Next
    '·„⁄—›… Œ’«∆’ «·«œ«… ----------------------------------------------------
    Dim IX As Integer

    If Not (Obj Is Nothing) Then
        GetCtrlProperty = Obj.Name
        GetCtrlProperty = GetCtrlProperty & "~^" & Obj.index
        GetCtrlProperty = GetCtrlProperty & "~^" & Obj.left
        GetCtrlProperty = GetCtrlProperty & "~^" & Obj.top
        GetCtrlProperty = GetCtrlProperty & "~^" & Obj.Width
        GetCtrlProperty = GetCtrlProperty & "~^" & Obj.Height
        GetCtrlProperty = GetCtrlProperty & "~^" & Obj.Visible

        If TypeOf Obj Is ISAniLabel Then
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.AutoSize
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.Caption
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.backcolor
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.ForeColor
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.FontName
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.FontBold
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.FontItalic
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.fontsize
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.FontUnderline
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.Alignment
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.RightToLeft
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.BackStyle
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.TextShadow
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.TextShadowDepth
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.ColorTextShadow
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.RotationAngle
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.ColorShadow
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.interval
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.Tag
        ElseIf TypeOf Obj Is Image Then
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.Tag
        ElseIf TypeOf Obj Is VSFlexGrid Then
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.Cols
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.Appearance
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.FontName
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.fontsize
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.FontBold
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.FontItalic
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.ForeColor
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.ForeColorFixed
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.backcolor
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.BackColorAlternate
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.BackColorFixed
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.GridColor
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.GridColorFixed
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.SheetBorder
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.BorderStyle
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.AutoSizeMode
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.GridLines
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.GridLinesFixed
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.TextStyleFixed
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.TextStyle
            GetCtrlProperty = GetCtrlProperty & "~^" & Obj.GridLineWidth

            For IX = 0 To Obj.Cols - 1
                GetCtrlProperty = GetCtrlProperty & "~^" & Obj.Col
                GetCtrlProperty = GetCtrlProperty & "~^" & Obj.ColKey(IX)
                GetCtrlProperty = GetCtrlProperty & "~^" & Obj.ColWidth(IX)
                GetCtrlProperty = GetCtrlProperty & "~^" & Obj.TextMatrix(0, IX)
                GetCtrlProperty = GetCtrlProperty & "~^" & Obj.FixedAlignment(IX)
                GetCtrlProperty = GetCtrlProperty & "~^" & Obj.ColAlignment(IX)
                Debug.Print Obj.FixedAlignment(IX)
            Next

        End If
    End If

    'Debug.Print GetCtrlProperty

End Function

Public Function PictureFromByteStream(b() As Byte) As IPicture
    Dim LowerBound As Long
    Dim ByteCount  As Long
    Dim hMem  As Long
    Dim lpMem  As Long
    Dim IID_IPicture(15)
    Dim istm As stdole.StdPicture

    On Error GoTo ErrTrap

    If UBound(b, 1) < 0 Then
        Exit Function
    End If
    
    LowerBound = LBound(b)
    ByteCount = (UBound(b) - LowerBound) + 1
    hMem = GlobalAlloc(&H2, ByteCount)

    If hMem <> 0 Then
        lpMem = GlobalLock(hMem)

        If lpMem <> 0 Then
            MoveMemory ByVal lpMem, b(LowerBound), ByteCount
            Call GlobalUnlock(hMem)

            If CreateStreamOnHGlobal(hMem, 1, istm) = 0 Then
                If CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture(0)) = 0 Then
                    Call OleLoadPicture(ByVal ObjPtr(istm), ByteCount, 0, IID_IPicture(0), PictureFromByteStream)
                End If
            End If
        End If
    End If
    
    Exit Function
    
ErrTrap:

    If Err.Number = 9 Then
        'Uninitialized array
        'MsgBox "You must pass a non-empty byte array to this function!"
        'MsgBox "⁄›Ê« „·› «·’Ê— €Ì— „ÊÃÊœ", vbOKOnly, App.Title
    Else
    End If

End Function

