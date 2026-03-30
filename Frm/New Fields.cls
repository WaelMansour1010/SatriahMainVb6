VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NewFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

add_record_to_table "VatTypes", "ID,VatTypeName,VatTypeNamee", "48,'«·Õ«ÊÌ« ','Containers'", "ID", 48
add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 3379 ,'«·Õ«ÊÌ«     ' ,'      Containers' ", "NotesType", 3379
add_record_to_table "VatTypes", "ID,VatTypeName,VatTypeNamee", "49,'«· ›—Ì€« ','Containers Unloading'", "ID", 49
add_record_to_table "TblNotesTypes", "NotesType,NotesTypeName,NotesTypeNamee", " 3380 ,'«· ›—Ì€«     ' ,'      Containers' ", "NotesType", 3380
      
DB_CreateField "branches", "a209", adVarWChar, adColNullable, 50, , "", False, True, , True
DB_CreateField "branches", "a210", adVarWChar, adColNullable, 50, , "", False, True, , True
      DB_CreateField "ContainerContracts", "Total", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "ContainerContracts", "Vat2", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "ContainerContracts", "Net", adDouble, adColNullable, , , "    ", False, True


DB_CreateField "ContainerContracts", "RecDate", adDBTimeStamp, adColNullable, , , "      ", False, True

DB_CreateField "ContainerContracts", "NoteID", adInteger, adColNullable, , , "    ", False, True
DB_CreateField "ContainerContracts", "NoteSerial", adVarWChar, adColNullable, 255, , "      ", False, True, , True

        DB_CreateField "ContainerContracts", "Total", adDouble, adColNullable, , , "    ", False, True
            DB_CreateField "ContainerContracts", "Vat2", adDouble, adColNullable, , , "    ", False, True
                DB_CreateField "ContainerContracts", "Net", adDouble, adColNullable, , , "    ", False, True
                 DB_CreateField "ContainerContracts", "Blanks", adDouble, adColNullable, , , "    ", False, True









DB_CreateField "ContainerContractsRecDet", "GroupID2", adInteger, adColNullable, , , " ???    ", False, True
    If DB_CreateTable("ContainerUnloading", True, "ID", False) = True Then
          DB_CreateField "ContainerUnloading", "CustID", adInteger, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "CustTel", adVarWChar, adColNullable, 50, , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "ContractNo", adInteger, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "RecordDate", adDBTimeStamp, adColNullable, , , "      ", False, True
          DB_CreateField "ContainerUnloading", "BranchID", adInteger, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "Remarks", adVarWChar, adColNullable, 4000, , "      ", False, True, , True
          DB_CreateField "ContainerUnloading", "UserID", adInteger, adColNullable, , , " ???    ", False, True
          
          DB_CreateField "ContainerUnloading", "Value", adDouble, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "Count", adDouble, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "TotalValue", adDouble, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "DiscValue", adDouble, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "DiscPercent", adDouble, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "NetBDisc", adDouble, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "Vat", adDouble, adColNullable, , , " ???    ", False, True
          DB_CreateField "ContainerUnloading", "TotalNet", adDouble, adColNullable, , , " ???    ", False, True
          
        DB_CreateField "ContainerUnloading", "CustID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "ContainerUnloading", "NoteID", adInteger, adColNullable, , , " ???    ", False, True
        DB_CreateField "ContainerUnloading", "NoteSerial", adVarWChar, adColNullable, 255, , "      ", False, True, , True
       End If

  

  
