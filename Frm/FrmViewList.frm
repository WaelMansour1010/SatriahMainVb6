VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmViewList 
   Caption         =   "⁄—÷ «·ﬁ«∆„…"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8745
   Icon            =   "FrmViewList.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   8745
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   6990
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8745
      _cx             =   15425
      _cy             =   12330
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   2
      ChildSpacing    =   1
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   4
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmViewList.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic EleFooter 
         Height          =   525
         Left            =   30
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   6435
         Width           =   8685
         _cx             =   15319
         _cy             =   926
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   1
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Caption         =   "RTL"
            Height          =   315
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   60
            Width           =   1125
         End
         Begin VB.CheckBox Chk 
            Alignment       =   1  'Right Justify
            Caption         =   "⁄—÷ ‘Ã—… «·„Ã„Ê⁄« "
            Height          =   315
            Left            =   6600
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   30
            Width           =   2025
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   0
            Left            =   60
            TabIndex        =   3
            Top             =   60
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            Caption         =   "Œ—ÊÃ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   1
            Left            =   2760
            TabIndex        =   4
            Top             =   60
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            Caption         =   "ŒÌ«—«  «·⁄—÷"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   2
            Left            =   4080
            TabIndex        =   5
            Top             =   60
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            Caption         =   " ÕœÌœ «·»Ì«‰« "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   3
            Left            =   1410
            TabIndex        =   6
            Top             =   60
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            Caption         =   "ÿ»«⁄…"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Dynamic_Byte.vsfGroup vsfGroup1 
         Height          =   6390
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   8685
         _extentx        =   15319
         _extenty        =   11271
         separatorcolor  =   255
      End
   End
End
Attribute VB_Name = "FrmViewList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_BolRetrunOnDblClick As Boolean

Private m_RetrunFrm As Form

Private StrRetrunColKey As String

Private Sub Check1_Click()
    Me.vsfGroup1.SetRTL = IIf(Me.Check1.value = vbChecked, True, False)
End Sub

Private Sub Chk_Click()

    If Chk.value = vbChecked Then
        Me.vsfGroup1.ShowTreeGroups = True
    Else
        Me.vsfGroup1.ShowTreeGroups = False
    End If

End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
            Unload Me

        Case 1
    '        FrmViewListProperties.show vbModal

        Case 2
            ShowFilter

        Case 3
            Me.vsfGroup1.PrintData
    End Select

End Sub

Private Sub ShowFilter()
    Me.vsfGroup1.ShowFilter
End Sub

Private Sub Form_Load()

    If SystemOptions.UserInterface = ArabicInterface Then
        Me.vsfGroup1.SetRTL = True
        Me.Check1.value = vbChecked
    Else
        Me.vsfGroup1.SetRTL = False
        Me.Check1.value = vbUnchecked
    End If

    Me.vsfGroup1.ShowTreeGroups = False
    Me.Chk.value = IIf(Me.vsfGroup1.ShowTreeGroups = True, vbChecked, vbUnchecked)
    Me.Height = 9240
    Me.Width = 11100
    Resize_Form Me
End Sub

Public Function SetDblClickRetrun(m_Frm As Form, _
                                  StrColKey As String)
    '›Ï »⁄÷ «·√ÕÌ«‰ —»„« ‰—Ìœ «‰ ‰” —Ã⁄ ”Ã· „⁄Ì‰
    '≈·Ï «·‘«‘… «· Ï ŸÂ—  „‰Â« Â–Â «·›Ê—„ (
    '„‰ Œ·«· Â–Â «·œ«·… ‰Õ‰ ‰ﬁÊ„ » ÕœÌœ «·›Ê—„
    '-----------------
    'm_Frm : The Form Which We will Retrun on it
    'StrColKey :The Col Key (Must Be Use The Key NOT the Index Because We Change The Index Normally)
    '-----------------
    Set m_RetrunFrm = m_Frm
    StrRetrunColKey = StrColKey
End Function

Public Property Get BolRetrunOnDblClick() As Boolean
    BolRetrunOnDblClick = m_BolRetrunOnDblClick
End Property

Public Property Let BolRetrunOnDblClick(ByVal vNewValue As Boolean)
    m_BolRetrunOnDblClick = vNewValue
End Property

Private Sub vsfGroup1_GridDblClick(Row As Long, _
                                   Col As Long)
    Dim Lngid As Long

    If Row <= 0 Then Exit Sub
    If Col < 0 Then Exit Sub
    If Me.BolRetrunOnDblClick = True Then

        With Me.vsfGroup1.vsFlexGrid
            Lngid = val(.TextMatrix(Row, .ColIndex(StrRetrunColKey)))
        End With

        If Lngid <> 0 Then
            If Not (m_RetrunFrm) Is Nothing Then
                m_RetrunFrm.Retrive Lngid
            End If
        End If
    End If

End Sub
