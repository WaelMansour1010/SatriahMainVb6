VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

DB_CreateField "TblVATAvowal", "ChkIsFree", adBoolean, adColNullable, , , "                ", False, True
DB_CreateField "TblBranchesData", "VATItems", adInteger, adColNullable, , , "      ", False, True
TblBranchesData
VATItems
