VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private FrmPrg As frmProgress

Private LngHwnd As Long

Public Enum ProgressType
    Incremental = 0
    Degressmental = 1
    Waiting = 2
End Enum
Dim m_ProgressType As ProgressType

Private Sub Class_Initialize()
    Set FrmPrg = New frmProgress
    LngHwnd = FrmPrg.hWnd
End Sub

Public Sub StartProgress()

    If Me.ProgressType = Waiting Then
        FrmPrg.Timer1.Enabled = True
        FrmPrg.Timer1.interval = 6000
    '    PutFormOnTop LngHwnd
        FrmPrg.show
    End If

End Sub

Public Sub StopProgess()
On Error Resume Next
    FrmPrg.Timer1.Enabled = False
    FrmPrg.Timer1.interval = 0
    FrmPrg.PrgBar.value = FrmPrg.PrgBar.Max
    FrmPrg.ISAniLabel19.ForeFlashing = False
    FrmPrg.ISAniLabel19.Caption = " „  ⁄„·Ì… «· Õ„Ì·"
    FrmPrg.lbl1.Caption = ""

    DoEvents
    FrmPrg.Hide
End Sub

Private Sub Class_Terminate()
    Dim i As Integer
    On Error GoTo hErr

    If Not FrmPrg Is Nothing Then

        For i = 0 To Forms.count

            If Forms(i).hWnd = LngHwnd Then
                Unload Forms(i)
                Exit For
            End If

        Next i

        Set FrmPrg = Nothing
    End If

    Exit Sub
hErr:
    Set FrmPrg = Nothing
End Sub

Public Property Get ProgressType() As ProgressType
    ProgressType = m_ProgressType
End Property

Public Property Let ProgressType(ByVal vNewValue As ProgressType)
    m_ProgressType = vNewValue
End Property

Public Sub FinishProgress()
    
End Sub

Public Property Get ProgressCaption() As Variant

End Property

Public Property Let ProgressCaption(ByVal vNewValue As Variant)

End Property
