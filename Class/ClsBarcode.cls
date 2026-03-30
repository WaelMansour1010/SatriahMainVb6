VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsBarcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_PaperWidth As Single
Dim m_PaperHeight  As String
Dim m_StickerWidth As Single
Dim m_StickerHeight As Single
Dim m_StickerCount As Integer
Dim m_PaperCols  As Integer
Dim m_PaperRows As Integer
Dim SngFactorMM_Twip As Integer

Private FrstItemOfPage As Integer

Private CurrPage As Integer
Dim m_ShowUpText As Boolean
Dim m_Show As Boolean
Dim m_BackColor As OLE_COLOR
Dim m_ForeColor As OLE_COLOR
Dim m_BarCodeTyp As Integer
Dim m_ShowDNText As Boolean
Dim Items As New Collection

Dim m_Font As StdFont

Private Type StkSize
    SkHieght As Single
    SkWidth  As Single
End Type

Private Type PapSize
    PaperHieght As Single
    PaperWidth  As Single
End Type

Public Function PaperSize(psize As PaperSiz)
    On Error GoTo ErrTrap

    ' «·ﬁÌ„ „Õœœ… »«·”‰ Ì„ — --------------------------
    Dim PaperSizex As PapSize

    With PaperSizex

        Select Case psize

            Case 1  'A3---------------

            Case 2  'A4 --------------
                .PaperWidth = 21
                .PaperHieght = 29.7

            Case 3  'A5 -----------
                .PaperWidth = 14.8
                .PaperHieght = 21

            Case 4  'A6 ---------
                .PaperWidth = 10.5
                .PaperHieght = 14.8
        End Select

        ' ÕÊÌ· «·ﬁÌ„ ·· mm
  
        m_PaperWidth = .PaperWidth * SngFactorMM_Twip
        m_PaperHeight = .PaperHieght * SngFactorMM_Twip
        SaveSetting "BarCode", "BarCode", "PaperSize", psize

    End With

ErrTrap:
End Function

Public Property Get StickerSize() As TypNum
    On Error GoTo ErrTrap

    StickerSize = IIf(GetSetting("BarCode", "BarCode", "StickerSize") = "", TypNum.Stk72, GetSetting("BarCode", "BarCode", "StickerSize"))
    Dim StickerSizex As StkSize

    With StickerSizex
        m_StickerCount = StickerSize

        Select Case StickerSize

            Case 1
                .SkWidth = 21
                .SkHieght = 29.7
                m_PaperCols = 1

            Case 2
                .SkWidth = 21
                .SkHieght = 14.85
                m_PaperCols = 1

            Case 4
                .SkWidth = 10.5
                .SkHieght = 14.85
                m_PaperCols = 2

            Case 8
                .SkWidth = 10.5
                .SkHieght = 7.4
                m_PaperCols = 2

            Case 10
                m_PaperCols = 2
                .SkWidth = 10.5
                .SkHieght = 5.94

            Case 12
                m_PaperCols = 2
                .SkWidth = 10.5
                .SkHieght = 4.95

            Case 14
                m_PaperCols = 2
                .SkWidth = 10.5
                .SkHieght = 4.24

            Case 16
                m_PaperCols = 2
                .SkWidth = 10.5
                .SkHieght = 3.7

            Case 17
                m_PaperCols = 4
                .SkWidth = 5.25
                .SkHieght = 7.42
                m_StickerCount = 16

            Case 21
                m_PaperCols = 3
                .SkWidth = 7
                .SkHieght = 4.24

            Case 24
                m_PaperCols = 3
                .SkWidth = 7
                .SkHieght = 3.7

            Case 25
                m_PaperCols = 4
                .SkWidth = 5.25
                .SkHieght = 4.95
                m_StickerCount = 24

            Case 28
                m_PaperCols = 4
                .SkWidth = 5.25
                .SkHieght = 4.24

            Case 36
                m_PaperCols = 3
                .SkWidth = 7
                .SkHieght = 2.47

            Case 40
                m_PaperCols = 4
                .SkWidth = 5.25
                .SkHieght = 2.97

            Case 48
                m_PaperCols = 6
                .SkWidth = 3.5
                .SkHieght = 3.7

            Case 56
                m_PaperCols = 4
                .SkWidth = 5.25
                .SkHieght = 2.1

            Case 72
                m_PaperCols = 6
                .SkWidth = 3.5
                .SkHieght = 2.47

            Case 96
                m_PaperCols = 6
                .SkWidth = 3.5
                .SkHieght = 1.85

            Case 102
                m_PaperCols = 6
                .SkWidth = 3.5
                .SkHieght = 1.74

            Case 108
                m_PaperCols = 6
                .SkWidth = 3.5
                .SkHieght = 1.48

            Case 109
                m_PaperCols = 9
                .SkWidth = 2.33
                .SkHieght = 2.47
                m_StickerCount = 108

            Case 120
                m_PaperCols = 6
                .SkWidth = 3.5
                .SkHieght = 1.48

            Case 144
                m_PaperCols = 6
                .SkWidth = 3.5
                .SkHieght = 1.2
        End Select

        m_PaperRows = StickerSize / m_PaperCols
   
        '«· ÕÊÌ· ·· mm
        m_StickerWidth = .SkWidth * SngFactorMM_Twip
        m_StickerHeight = .SkHieght * SngFactorMM_Twip
    End With

ErrTrap:
End Property

Public Property Let StickerSize(Typ As TypNum)
    On Error GoTo ErrTrap

    Select Case Typ

        Case 1
            
            m_PaperCols = 1

        Case 2
            m_PaperCols = 1

        Case 4
           
            m_PaperCols = 2

        Case 8
           
            m_PaperCols = 2

        Case 10
            m_PaperCols = 2
            
        Case 12
            m_PaperCols = 2
           
        Case 14
            m_PaperCols = 2
           
        Case 16
            m_PaperCols = 2
            
        Case 17
            m_PaperCols = 4
            
        Case 21
            m_PaperCols = 3
            
        Case 24
            m_PaperCols = 3
            
        Case 25
            m_PaperCols = 4
            
        Case 28
            m_PaperCols = 4
            
        Case 36
            m_PaperCols = 3
            
        Case 40
            m_PaperCols = 4
           
        Case 48
            m_PaperCols = 6
            
        Case 56
            m_PaperCols = 4
           
        Case 72
            m_PaperCols = 6
            
        Case 96
            m_PaperCols = 6
            
        Case 102
            m_PaperCols = 6
           
        Case 108
            m_PaperCols = 6
           
        Case 109
            m_PaperCols = 9
            
        Case 120
            m_PaperCols = 6
          
        Case 144
            m_PaperCols = 6
            
    End Select

    m_PaperRows = Typ / m_PaperCols
    m_StickerCount = Typ
    SaveSetting "BarCode", "BarCode", "StickerSize", Typ
ErrTrap:
End Property

Public Sub Preview()

    ShowBarCode FrmParent 'BarCode

End Sub

Private Sub Class_Initialize()
    SngFactorMM_Twip = 10
    'Me.BarCodeTyp = Code39
End Sub

Private Function PrintBarCode() As Boolean
    'On Error GoTo ErrTrap
    Dim Msg As String
    Dim i As Integer

    Printer.CurrentX = 0
    Printer.CurrentY = 0

    Printer.FillColor = vbWhite

    With FrmParent

        For i = 1 To .Img.count - 1
            Debug.Print .Img.count - 1
            Debug.Print .Img(i).Picture
            '      If .Img(I).Picture <> 0 Then
            Debug.Print .Img(i).Picture
            Printer.ScaleMode = 6  ' sets printer scale to mm
            Printer.PaintPicture .Img(i).Picture, .Img(i).left, .Img(i).top, .Img(1).Width, .Img(i).Height
            CtrlPrint .UpLbl(i)
            CtrlPrint .DNLbl(i)
            '       End If
        Next

    End With

    Printer.NewPage
    Printer.EndDoc
    PrintBarCode = True
    Exit Function
ErrTrap:
    Msg = "⁄›Ê«  √ﬂœ „‰  ⁄—Ì› «·ÿ«»⁄…"
    Msg = Msg & Chr(13) & "Ê’› «·‰Ÿ«„ ··Œÿ«:"
    Msg = Msg & Chr(13) & Err.Number
    Msg = Msg & Chr(13) & Err.description
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.Title
    PrintBarCode = False
End Function

'----------------------------------------------------------------------------
Private Function CtrlPrint(Ctrl As Control)
    ' Combine API Calls with the Printer object
    On Error GoTo ErrTrap

    Dim result As Long
    Dim mmToTwp As Single
    mmToTwp = 56.7
    Printer.ScaleMode = vbTwips
    'Printer.Print " "
    Printer.FillColor = Ctrl.BackColor
    Printer.ForeColor = Ctrl.ForeColor
    Set Printer.Font = Ctrl.Font
    Printer.Line (Ctrl.left * mmToTwp, Ctrl.top * mmToTwp)-((Ctrl.left + Ctrl.Width) * mmToTwp, (Ctrl.top + Ctrl.Height) * mmToTwp), Ctrl.BackColor, BF
 
    Dim xx As Single

    If Ctrl.Alignment = vbLeftJustify Then
        xx = (Ctrl.left * mmToTwp) / Printer.TwipsPerPixelX
    ElseIf Ctrl.Alignment = vbCenter Then
        xx = ((Ctrl.left * mmToTwp) + ((mmToTwp * Ctrl.Width) - Printer.TextWidth(Ctrl.Caption)) / 2) / Printer.TwipsPerPixelX

        ' XX = (Ctrl.Left + ((Ctrl.Width - Printer.TextWidth(Ctrl.Caption)) / 2)) / Printer.TwipsPerPixelX
    ElseIf Ctrl.Alignment = vbRightJustify Then
        xx = ((Ctrl.left * mmToTwp) + (((Ctrl.Width * mmToTwp) - Printer.TextWidth(Ctrl.Caption)))) / Printer.TwipsPerPixelX
    End If

    Ctrl.ZOrder 0
    result = TextOut(Printer.hDC, xx, (Ctrl.top * mmToTwp) / Printer.TwipsPerPixelY, Ctrl.Caption, Len(Ctrl.Caption))

ErrTrap:
    'Printer.EndDoc
End Function

Public Property Get BarCodeTyp() As Bartyp
    On Error GoTo ErrTrap
    BarCodeTyp = IIf(GetSetting("BarCode", "BarCode", "BarCodeTyp") = "", Bartyp.Code39, GetSetting("BarCode", "BarCode", "BarCodeTyp"))
ErrTrap:
End Property

Public Property Let BarCodeTyp(Bartyp As Bartyp)
    On Error GoTo ErrTrap
    SaveSetting "BarCode", "BarCode", "BarCodeTyp", Bartyp
ErrTrap:
End Property

'---------------FontName-----------------------------------------------------
Public Property Get Font() As StdFont
    On Error GoTo ErrTrap
    Set Font = New StdFont
    Font.name = GetSetting("BarCode", "BarCode", "FontName")
    Font.Bold = IIf(GetSetting("BarCode", "BarCode", "FontBold") = "", False, GetSetting("BarCode", "BarCode", "FontBold"))
    Font.Size = IIf(GetSetting("BarCode", "BarCode", "FontSize") = "", 9, GetSetting("BarCode", "BarCode", "FontSize"))
    Font.Italic = IIf(GetSetting("BarCode", "BarCode", "FontItalic") = "", False, GetSetting("BarCode", "BarCode", "FontItalic"))
    Font.Charset = 178
    Font.Charset = IIf(GetSetting("BarCode", "BarCode", "FontCharset") = "", 117, GetSetting("BarCode", "BarCode", "FontCharset"))
ErrTrap:
End Property

Public Property Set Font(ByVal vNewValue As StdFont)
    On Error GoTo ErrTrap
    SaveSetting "BarCode", "BarCode", "FontName", vNewValue.name
    'SaveSetting "BarCode", "BarCode", "FontCharset", vNewValue.Charset
    SaveSetting "BarCode", "BarCode", "FontBold", vNewValue.Bold
    SaveSetting "BarCode", "BarCode", "FontSize", vNewValue.Size
    SaveSetting "BarCode", "BarCode", "FontItalic", vNewValue.Italic
ErrTrap:
End Property

'----------------ShowUpText---------------------------------------------------
Public Property Get ShowUpText() As Boolean
    On Error GoTo ErrTrap
    ShowUpText = IIf(GetSetting("BarCode", "BarCode", "ShowUpText") = "", False, GetSetting("BarCode", "BarCode", "ShowUpText"))
ErrTrap:
End Property

Public Property Let ShowUpText(ByVal vNewValue As Boolean)
    On Error GoTo ErrTrap
    SaveSetting "BarCode", "BarCode", "ShowUpText", vNewValue
ErrTrap:
End Property

'----------------ShowDNText---------------------------------------------------
Public Property Get ShowDNText() As Boolean
    On Error GoTo ErrTrap
    ShowDNText = IIf(GetSetting("BarCode", "BarCode", "ShowDNText") = "", False, GetSetting("BarCode", "BarCode", "ShowDNText"))
ErrTrap:
End Property

Public Property Let ShowDNText(ByVal vNewValue As Boolean)
    On Error GoTo ErrTrap
    SaveSetting "BarCode", "BarCode", "ShowDNText", vNewValue
ErrTrap:
End Property

'------------------ForeColor---------------------------------------------------

Public Property Get ForeColor() As OLE_COLOR
    On Error GoTo ErrTrap
    Dim FrColor As Long
    FrColor = val(GetSetting("BarCode", "BarCode", "ForeColor"))

    ForeColor = IIf(FrColor = 0, vbBlack, FrColor)
ErrTrap:
End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
    On Error GoTo ErrTrap
    SaveSetting "BarCode", "BarCode", "ForeColor", vNewValue
ErrTrap:
End Property

'----------------------BackColor-----------------------------------------------
Public Property Get BackColor() As OLE_COLOR
    On Error GoTo ErrTrap
    Dim BkClor As Long
    BkClor = val(GetSetting("BarCode", "BarCode", "BackColor"))
    BackColor = IIf(BkClor = 0, vbWhite, BkClor)
ErrTrap:
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
    On Error GoTo ErrTrap
    SaveSetting "BarCode", "BarCode", "BackColor", vNewValue
ErrTrap:
End Property

'------------------------------------------------------------------------------
Private Function Setting(Frm As Form)
    On Error GoTo ErrTrap

    With Frm
        Dim StSize As Long
        .AXBar.BackColor = IIf(Me.BackColor = 0, Me.BackColor = vbWhite, Me.BackColor)
        .UpLbl(0).BackColor = .AXBar.BackColor
        .DNLbl(0).BackColor = .AXBar.BackColor
    
        .AXBar.ForeColor = Me.ForeColor
        .UpLbl(0).ForeColor = .AXBar.ForeColor
        .DNLbl(0).ForeColor = .AXBar.ForeColor
    
        Set .UpLbl(0).Font = Me.Font
        Set .DNLbl(0).Font = .UpLbl(0).Font
    
        StSize = Me.StickerSize
        .AXBar.CodeType = Me.BarCodeTyp
        Me.PaperSize a4 ' Val(GetSetting("BarCode", "BarCode", "PaperSize"))

    End With

ErrTrap:
End Function

Public Function PageCount() As Integer
    On Error GoTo ErrTrap
    Dim count As Integer
    PageCount = 1

    count = Int(Items.count / m_StickerCount)

    If Items.count Mod m_StickerCount > 0 Then
        count = count + 1
    End If

    PageCount = count
ErrTrap:
End Function

Public Function MovetoPage(Page As Integer)
    On Error GoTo ErrTrap
    'CurrPage = Page
    'CurrPage = Page
    GetFrstItemOfPage Page
    Preview
ErrTrap:
End Function

Private Function GetFrstItemOfPage(ByVal Page As Integer) As Integer
    On Error GoTo ErrTrap
    GetFrstItemOfPage = 0
    Dim i As Integer

    m_StickerCount = Me.StickerSize
    Page = Page - 1

    If Page >= Me.PageCount Then
        Page = Me.PageCount - 1
    End If

    If Page < 0 Then
        Page = 0
    End If

    CurrPage = Page + 1

    GetFrstItemOfPage = Page * m_StickerCount

    FrstItemOfPage = Page * m_StickerCount
ErrTrap:
End Function

Private Function GetFrst_ItemOfPage(ByVal Page As Integer) As Integer
    On Error GoTo ErrTrap
    GetFrst_ItemOfPage = 0
    Dim i As Integer

    m_StickerCount = Me.StickerSize
    Page = Page - 1

    If Page >= Me.PageCount Then
        Page = Me.PageCount - 1
    End If

    If Page < 0 Then
        Page = 0
    End If

    GetFrst_ItemOfPage = Page * m_StickerCount

    FrstItemOfPage = Page * m_StickerCount
ErrTrap:
End Function

Public Function AddItem(Barcode As String, _
                        Optional UpText As String, _
                        Optional DNText As String)
    On Error GoTo ErrTrap

    Dim Item As String
    Item = Barcode & "~^" & UpText & "~^" & DNText
    Items.Add Item
ErrTrap:
End Function

Public Function ClearItems()
    On Error GoTo ErrTrap

    Dim i As Integer

    For i = 1 To Items.count
        Items.Remove 1
    Next

ErrTrap:
End Function

Public Function MoveFirst()
    On Error GoTo ErrTrap
    CurrPage = 1
    MovetoPage 1
ErrTrap:
End Function

Public Function MoveLast()
    On Error GoTo ErrTrap
    MovetoPage PageCount
ErrTrap:
End Function

Public Function MoveNext()
    On Error GoTo ErrTrap
    MovetoPage CurrPage + 1
ErrTrap:
End Function

Public Function MovePrev()
    On Error GoTo ErrTrap
    MovetoPage CurrPage - 1
ErrTrap:
End Function

Public Property Get PageNumber() As Integer
    On Error GoTo ErrTrap
    PageNumber = CurrPage
ErrTrap:
End Property

Private Sub ShowBarCode(Frm As Form)
    On Error GoTo ErrTrap
    Dim i As Integer
    Dim IntRows As Integer
    Dim IntCounter As Integer
    Dim IMGCount As Integer
    Dim HorMarge As Single
    Dim VertMarge As Single
    Dim IMGLeft As Single
    Dim ImgTop As Single

    Dim StrBarCode() As String
    Dim II As Integer

    Call Setting(Frm)

    With Frm

        If Frm.name = "BarCodePrint" Then
    
            .Visible = False
            '        .Hide
        Else
            .PicStikers.Width = (m_PaperWidth * 56.7)
            .PicStikers.Height = (m_PaperHeight * 56.7)

        End If
    
        For i = .Img.count To 1 Step -1

            If i - 1 = 0 Then Exit For
            Unload .Img(i - 1)
            Unload .UpLbl(i - 1)
            Unload .DNLbl(i - 1)
        Next
    
        .UpLbl(0).Height = 2.9
        .DNLbl(0).Height = 2.9
        HorMarge = 2.5
        VertMarge = 2.5
        IMGLeft = 0
        ImgTop = 0
    
        If m_StickerHeight <= 15 Then
            .UpLbl(0).Height = 2
            .DNLbl(0).Height = 2
        
            VertMarge = 0.5
        End If

        '.Left = 1800
        
        II = FrstItemOfPage
        
        For IntRows = 1 To m_PaperRows
            For IntCounter = 1 To m_PaperCols

                If II >= Items.count Then GoTo LastLine
                StrBarCode = Split(Items(II + 1), "~^")
                IMGCount = .Img.count
                Load .Img(IMGCount)
                Load .UpLbl(IMGCount)
                Load .DNLbl(IMGCount)
                .UpLbl(IMGCount).Caption = ""
                .DNLbl(IMGCount).Caption = ""
                
                .Img(IMGCount).Width = m_StickerWidth - HorMarge * 4
                .Img(IMGCount).Height = m_StickerHeight - VertMarge * 2
                .Img(IMGCount).left = IMGLeft
                IMGLeft = .Img(IMGCount).left + .Img(IMGCount).Width + (HorMarge * 4)
                .Img(IMGCount).top = ImgTop
                .AXBar.Caption = StrBarCode(0)
                .AXBar.Width = m_StickerWidth
                .AXBar.ImageWidth = m_StickerWidth
                .AXBar.ImageHeight = m_StickerHeight

                If Trim(StrBarCode(0)) <> "" Then
                    Set .Img(IMGCount).Picture = .AXBar.Picture
                Else
                    Set .Img(IMGCount).Picture = LoadPicture("")
                End If

                .Img(IMGCount).Visible = True

                If Me.ShowUpText = True Then
                    .UpLbl(IMGCount).Caption = StrBarCode(1)
                End If
                
                If Me.ShowDNText = True Then
                    .DNLbl(IMGCount).Caption = StrBarCode(2)
                End If
                
                .UpLbl(IMGCount).Move .Img(IMGCount).left, .Img(IMGCount).top, .Img(IMGCount).Width, 5
                .DNLbl(IMGCount).Move .Img(IMGCount).left, (.Img(IMGCount).top + .Img(IMGCount).Height) - .DNLbl(IMGCount).Height, .Img(IMGCount).Width, 5
                
                .UpLbl(IMGCount).Visible = True
                .DNLbl(IMGCount).Visible = True
                
                If .UpLbl(IMGCount).Caption = "" Then
                    .UpLbl(IMGCount).Visible = False
                End If
                
                If .DNLbl(IMGCount).Caption = "" Then
                    .DNLbl(IMGCount).Visible = False
                End If
                
                .UpLbl(IMGCount).ZOrder 0
                .DNLbl(IMGCount).ZOrder 0
                II = II + 1
            Next IntCounter

            IMGLeft = 0
            ImgTop = ImgTop + m_StickerHeight
        Next IntRows

LastLine:

        DoEvents

        If Frm.name = "FrmParent" Then
 
            .PicStikers.Width = (m_PaperWidth * 56.7)
            .PicStikers.Height = (m_PaperHeight * 56.7)

            If Frm.Visible = False Then
                Frm.Show vbModal
            End If

        Else
            .Visible = False
            .ZOrder 1
            '        .Hide
        End If

    End With

    Exit Sub
ErrTrap:
End Sub

Public Function PrintPage(PageNum As Integer)
    On Error GoTo ErrTrap
    Dim FristPage As Integer
    Dim EndPage As Integer
    Dim FirstItem As Integer
    Static ICount As Integer
    Dim C As Integer

    FristPage = PageNum
    EndPage = PageNum

    If PageNum = 0 Then
        FristPage = 1
        EndPage = Me.PageCount
    End If

    For ICount = FristPage To EndPage
   
        C = ICount
        FirstItem = GetFrst_ItemOfPage(C)
        Call ShowBarCode(FrmParent)

        If PrintBarCode = False Then
            Exit For
        End If

    Next

ErrTrap:
End Function

Public Property Let StickerBegin(ByVal StkNum As Integer)
    Dim i As Integer

    If StkNum = 0 Then
        StkNum = 1
    End If

    For i = 1 To StkNum - 1
        Me.AddItem ""
    Next

End Property

Public Property Get PaperCols() As Variant
    Dim xx As Integer
    xx = StickerCount
    PaperCols = m_PaperCols
End Property

Public Property Get PaperRows() As Variant
    PaperRows = m_PaperRows
End Property

Public Property Get StickerCount() As Variant
    StickerCount = m_StickerCount
End Property

