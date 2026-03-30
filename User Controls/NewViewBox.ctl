VERSION 5.00
Begin VB.UserControl NewViewBox 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1185
   ScaleHeight     =   705
   ScaleWidth      =   1185
   Begin VB.PictureBox Img 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   180
      MouseIcon       =   "NewViewBox.ctx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   345
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   270
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.HScrollBar hsl 
      Height          =   204
      LargeChange     =   10
      Left            =   72
      TabIndex        =   1
      Top             =   48
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.VScrollBar vsl 
      Height          =   588
      LargeChange     =   10
      Left            =   744
      TabIndex        =   0
      Top             =   24
      Visible         =   0   'False
      Width           =   300
   End
End
Attribute VB_Name = "NewViewBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetSystemMetrics _
                Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Private Declare Function ReleaseCapture _
                Lib "user32" () As Long

Private Const WM_NCLBUTTONDOWN = &HA1

Private Const HTCAPTION = 2

'Default Property Values:
Const m_def_View = 0

'Property Variables:
Dim m_View As ViewConstants

'Event Declarations:
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_UserMemId = -601
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp

Public Enum ViewConstants
    Normal = 0
    Stretch = 1
    AutoSize = 2
    ScrollBars = 3
    RealStrecth = 4
    HandScroll = 5
End Enum

Public Enum BorderConstants
    None = 0
    FixedSingle = 1
End Enum

Public Enum AppearanceConstants
    Flat = 0
    D3D = 1
End Enum

Dim SBS As Integer    ' Scroll Bars Size.

Private Const SM_CXVSCROLL = 2

Private Const SM_CYHSCROLL = 3

Dim sX As Integer
Dim sY As Integer

Private Sub Img_MouseDown(Button As Integer, _
                          Shift As Integer, _
                          x As Single, _
                          Y As Single)
    sX = x
    sY = Y
    RaiseEvent MouseDown(Button, Shift, x, Y)
    
End Sub

Private Sub Img_MouseMove(Button As Integer, _
                          Shift As Integer, _
                          x As Single, _
                          Y As Single)
    '    Dim L As Integer
    '    Dim T As Integer
    '
    '    L = IIf(Img.Left > 0, 0, IIf(Img.Left < ScaleWidth - Img.ScaleWidth, ScaleWidth - Img.ScaleWidth, Img.Left))
    '    T = IIf(Img.Left > 0, 0, IIf(Img.Left < ScaleHeight - Img.ScaleHeight, ScaleHeight - Img.ScaleHeight, Img.Left))
    '    Img.Move L, T
    '    Call ReleaseCapture
    '    Call SendMessage(Img.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
    '    L = IIf(Img.Left > 0, 0, IIf(Img.Left < ScaleWidth - Img.ScaleWidth, ScaleWidth - Img.ScaleWidth, Img.Left))
    '    T = IIf(Img.Left > 0, 0, IIf(Img.Left < ScaleHeight - Img.ScaleHeight, ScaleHeight - Img.ScaleHeight, Img.Left))
    '    Img.Move L, T
    
    RaiseEvent MouseMove(Button, Shift, x, Y)

    If Button <> 1 Then Exit Sub
    
    Dim dX As Integer
    Dim dY As Integer
    Dim nY As Integer
    Dim nX As Integer
    
    dX = x - sX
    dY = Y - sY
    
    nX = Img.left + dX
    nY = Img.top + dY
    
    nX = IIf(nX > 0, 0, IIf(nX < ScaleWidth - Img.Width, ScaleWidth - Img.Width, nX))
    nY = IIf(nY > 0, 0, IIf(nY < ScaleHeight - Img.Height, ScaleHeight - Img.Height, nY))
    
    nX = IIf(Img.ScaleWidth < ScaleWidth, (ScaleWidth - Img.ScaleWidth) / 2, nX)
    nY = IIf(Img.ScaleHeight < ScaleHeight, (ScaleHeight - Img.ScaleHeight) / 2, nY)
    
    Img.Move nX, nY
End Sub

Private Sub Img_MouseUp(Button As Integer, _
                        Shift As Integer, _
                        x As Single, _
                        Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

Private Sub Img_Click()
    RaiseEvent Click
End Sub

Private Sub Img_DblClick()
    RaiseEvent DblClick
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_View = m_def_View
End Sub

Private Sub UserControl_Paint()

    If Img.Picture <> 0 Then
        If m_View <> ScrollBars Then
            PrepareTheView
        Else
            DrawPicture
        End If
    End If

End Sub

'Load property values From storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", 1)
    m_View = PropBag.ReadProperty("View", m_def_View)
End Sub

Private Sub UserControl_Show()

    If Img.Picture <> 0 Then
        PrepareTheView
    Else
        Cls
    End If

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, 1)
    Call PropBag.WriteProperty("View", m_View, m_def_View)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    
    Call UserControl_Paint
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Img,Img,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
Attribute Picture.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Set Picture = Img.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set Img.Picture = New_Picture
    PropertyChanged "Picture"
    
    If Img.Picture <> 0 Then
        PrepareTheView
    Else
        Cls
    End If

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As BorderConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderConstants)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"

    Call UserControl_Paint
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub UserControl_Resize()

    If Img.Picture <> 0 Then PrepareTheView
    
    RaiseEvent Resize
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, _
                                Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, _
                              Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  Y As Single)
    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  x As Single, _
                                  Y As Single)
    RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As AppearanceConstants
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted at run time with 3-D effects."
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
    
    ' Repanit The Picture.
    Call UserControl_Paint
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get View() As ViewConstants
Attribute View.VB_ProcData.VB_Invoke_Property = ";0-Rgheed"
    View = m_View
End Property

Public Property Let View(ByVal New_View As ViewConstants)
    m_View = New_View
    PropertyChanged "View"
    
    Img.Visible = (New_View = HandScroll)
    Img.Move 0, 0

    '  ÂÌ∆… «·⁄—÷ «·„‰«”» ··’Ê—….
    If Img.Picture <> 0 Then PrepareTheView
End Property

Private Sub PrepareTheView()
    Cls
    
    vsl.Visible = False
    hsl.Visible = False
    
    Img.Visible = (m_View = HandScroll)

    Select Case m_View

        Case Normal:       PaintPicture Img.Picture, 0, 0: Exit Sub

        Case Stretch:      PaintPicture Img.Picture, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight: Exit Sub

        Case HandScroll:   Img.Move (ScaleWidth - Img.ScaleWidth) / 2, (ScaleHeight - Img.ScaleHeight) / 2: Exit Sub
    End Select
    
    If m_View = AutoSize Then
        Dim Edges As Integer
            
        Edges = UserControl.Width - UserControl.ScaleWidth
        
        UserControl.Width = Img.ScaleWidth + Edges
        UserControl.Height = Img.ScaleHeight + Edges
       
        PaintPicture Img.Picture, 0, 0
    End If
  
    If m_View = RealStrecth Then RealStretchView
    
    If m_View = ScrollBars Then
        Dim dX As Long
        Dim dY As Long
       
        SBS = GetScrollBarsSize
 
        dX = Img.ScaleWidth - UserControl.ScaleWidth
        dY = Img.ScaleHeight - UserControl.ScaleHeight
     
        If dX > 0 And dY > 0 Then
            hsl.Visible = True: hsl.Max = (dX + SBS) / 10: hsl.value = hsl.Max / 2
            vsl.Visible = True: vsl.Max = (dY + SBS) / 10: vsl.value = vsl.Max / 2
        ElseIf dX > 0 And dY <= 0 Then
            hsl.Visible = True: hsl.Max = (dX / 10): hsl.value = hsl.Max / 2
            vsl.Visible = False
        ElseIf dY > 0 And dX < 0 Then
            hsl.Visible = False
            vsl.Visible = True: vsl.Max = (dY / 10): vsl.value = vsl.Max / 2
        End If
       
        ResizeScrollBars
         
        DrawPicture
       
    End If

End Sub

Private Function GetScrollBarsSize() As Integer
    Dim W As Long
    
    W = GetSystemMetrics(SM_CXVSCROLL)
   
    GetScrollBarsSize = UserControl.ScaleX(W, vbPixels, vbTwips)
End Function

Private Sub ResizeScrollBars()
    Dim H As Integer
    Dim W As Integer
    
    H = UserControl.ScaleHeight
    W = UserControl.ScaleWidth
    
    If vsl.Visible And hsl.Visible Then
        vsl.Move W - SBS, 0, SBS, H - SBS
        hsl.Move 0, H - SBS, W - SBS, SBS
    ElseIf vsl.Visible Then
        vsl.Move W - SBS, 0, SBS, H
    ElseIf hsl.Visible Then
        hsl.Move 0, H - SBS, W, SBS
    End If

End Sub

Private Sub vsl_Change()
    DrawPicture
End Sub

Private Sub vsl_Scroll()
    DrawPicture
End Sub

Private Sub hsl_Change()
    DrawPicture
End Sub

Private Sub hsl_Scroll()
    DrawPicture
End Sub

Private Sub DrawPicture()
    Dim x As Long, Y As Long
    Dim H As Integer, W As Integer
     
    If vsl.Visible And hsl.Visible Then
        x = -hsl.value * 10
        Y = -vsl.value * 10
    ElseIf vsl.Visible Then
        x = -(Img.ScaleWidth - UserControl.ScaleWidth + SBS) / 2
        Y = -vsl.value * 10
    ElseIf hsl.Visible Then
        x = -hsl.value * 10
        Y = -(Img.ScaleHeight - UserControl.ScaleHeight + SBS) / 2
    Else
        x = -(Img.ScaleWidth - UserControl.ScaleWidth) / 2
        Y = -(Img.ScaleHeight - UserControl.ScaleHeight) / 2
    End If
     
    W = Abs(x) + UserControl.ScaleWidth - IIf(vsl.Visible, SBS, 0)
    H = Abs(Y) + UserControl.ScaleHeight - IIf(hsl.Visible, SBS, 0)
   
    PaintPicture Img.Picture, x, Y, , , 0, 0, W, H
End Sub

Private Sub RealStretchView()
    Dim PH As Long, PW As Long     ' Pic Size.
    Dim BH As Long, BW As Long     ' Box Size.
    Dim PX As Long, PY As Long
    Dim MinRate As Single
    
    BH = UserControl.ScaleHeight
    BW = UserControl.ScaleWidth
    PH = Img.ScaleHeight
    PW = Img.ScaleWidth
    
    If BH < PH Or BW < PW Then
        MinRate = IIf(BH / PH < BW / PW, BH / PH, BW / PW)
        PW = PW * MinRate
        PH = PH * MinRate
    End If

    PX = (BW - PW) / 2
    PY = (BH - PH) / 2
  
    PaintPicture Img.Picture, PX, PY, PW, PH
End Sub
