VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Class Name ClsDataCombos

1- New Functions

Public Function GetItemSpecifications(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "Select ID,Name From TblSpecification Order By Name"
    Else
    StrSQL = "Select ID,Namee From TblSpecification Order By Namee"
    End If
    GetComboData MyCombo, StrSQL
End Function
 
 
 
 2- Updated Function
 
 
Public Function GetItemSGroupsupdate(MyCombo As DataCombo, _
                               Optional BolGetRoot As Boolean = True, Optional ByVal mSqlWhere As String = "")
    Dim StrSQL As String

    If BolGetRoot = True Then
        StrSQL = "SELECT GroupID,GroupName From Groups"
        If mSqlWhere <> "" Then
            StrSQL = StrSQL & " Where " & mSqlWhere
        End If
        StrSQL = StrSQL & " Order By GroupID"
    Else
        StrSQL = "SELECT GroupID,GroupName From Groups Where (GroupID <>1)"
        If mSqlWhere <> "" Then
            StrSQL = StrSQL & " And " & mSqlWhere
        End If
        StrSQL = StrSQL & " Order By GroupName "
    End If

    GetComboData MyCombo, StrSQL

End Function
 




DB_CreateField "groups", "SpecificationID", adInteger, adColNullable, , , "  ", False, True

DB_CreateField "TblItemsParts", "IsAddToPrice", adBoolean, adColNullable, , , " ", False, True
DB_CreateField "groups", "SpecificationID", adInteger, adColNullable, , , "  ", False, True

DB_CreateField "Transaction_Details", "OldshowPrice2", adCurrency, adColNullable, , , "  ", False, True

DB_CreateField "TblDefComItemDet", "OldPrice", adDouble, adColNullable, , , "    ", False, True

