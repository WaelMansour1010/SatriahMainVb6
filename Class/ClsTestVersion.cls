VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTestVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub TestVersionType()
Dim Rs As New ADODB.Recordset
Dim Msg As String
Rs.Open "[options]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
If Rs("DemoVersion").Value = True Then
    
End If
End Sub
