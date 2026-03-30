Attribute VB_Name = "ModPaperType"
Option Explicit

Public Const DC_ACTIVE = &H1

Public Const DC_NOTACTIVE = &H2

Public Const DC_ICON = &H4

Public Const DC_TEXT = &H8

Public Const BDR_SUNKENOUTER = &H2

Public Const BDR_RAISEDINNER = &H4

Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)

Public Const BF_BOTTOM = &H8

Public Const BF_LEFT = &H1

Public Const BF_RIGHT = &H4

Public Const BF_TOP = &H2

Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Public Const DFC_BUTTON = 4

Public Const DFC_POPUPMENU = 5            'Only Win98/2000 !!

Public Const DFCS_BUTTON3STATE = &H10

Public Const DT_CENTER = &H1

Public Const DC_GRADIENT = &H20          'Only Win98/2000 !!

Enum TypNum
    Stk1 = 1
    Stk2 = 2
    Stk4 = 4
    Stk8 = 8
    Stk10 = 10
    Stk12 = 12
    Stk14 = 14
    Stk16_2 = 16
    Stk16_4 = 17
    Stk21 = 21
    Stk24_3 = 24
    Stk24_4 = 25
    Stk28 = 28
    Stk36 = 36
    Stk40 = 40
    Stk48 = 48
    Stk56 = 56
    Stk72 = 72
    Stk96 = 96
    Stk102 = 102
    stk108_6 = 108
    Stk108_9 = 109
    Stk120 = 120
    Stk144 = 144
End Enum

Enum Bartyp
    ENA = 0
    ENA8 = 1
    ENA2 = 2
    ENA5 = 3
    UPCA = 4
    UPCE = 5
    ITF = 6
    ITF6 = 7
    Code39 = 8
    Code128 = 9
    ENA128 = 10
    B2OF5 = 11
    B12_5 = 12
    B3OF9 = 13
    CodeB = 14
    Code11 = 15
    CodaBar = 16
    MSI = 17
    ExCode39 = 18
    IPCA2 = 19
    UPCA5 = 20
    ENA82 = 21
    ENA85 = 22
    UPCE2 = 23
    UPCE5 = 24
    Telepen = 25
    TelepenA = 26
    TelepenN = 27
    PostNetA = 28
    PostNetC = 29
    PostNetCP = 30
    FIMA = 31
    FIMB = 32
    FIMC = 33
    RM4SCC = 34
    State4 = 35
    Code93 = 36
    ExCode93 = 37
    ISBN = 38
    Matrix25 = 39
    Plessey = 40
    AustraliaP = 41
    SWISS = 42
    DeutscheP = 43
    SICI = 44
    ENA14 = 45
    PLANET12 = 46
    PLANET14 = 47
    ISSN = 48
    ISMN = 49
    SSCC = 50
    KoreanPA = 51
    PosteItaline39 = 52
    PosetItaline25 = 53
    ISBN2 = 54
    ISBN5 = 55
    ISSN2 = 56
    ISSN5 = 57
    JapanPost = 58
End Enum

Enum PaperSiz
    a3 = 1
    a4 = 2
    a5 = 3
    A6 = 4
End Enum

'Public Type RECT
'    left As Long
'    top As Long
'    right As Long
'    bottom As Long
'End Type

Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    'lfFaceName As String * LF_FACESIZE
End Type
   
Public Declare Function TextOut _
               Lib "gdi32" _
               Alias "TextOutA" (ByVal hDC As Long, _
                                 ByVal x As Long, _
                                 ByVal Y As Long, _
                                 ByVal lpString As String, _
                                 ByVal nCount As Long) As Long

Public Declare Function SetParent _
               Lib "user32" (ByVal hWndChild As Long, _
                             ByVal hWndNewParent As Long) As Long

Public cBarcode As New ClsBarcode

