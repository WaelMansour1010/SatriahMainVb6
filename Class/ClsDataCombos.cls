VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsDataCombos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Function GetItemSpecifications(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "Select ID,Name From TblSpecification Order By Name"
    Else
    StrSQL = "Select ID,Namee From TblSpecification Order By Namee"
    End If
    GetComboData MyCombo, StrSQL
End Function
Public Function GetProductLineData(MyCombo As DataCombo, _
                                Optional IsBase As Boolean = True)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT      id,name FROM         dbo.TblProductLine Where 1 = 1 "
    If IsBase Then
        StrSQL = StrSQL & " And IsNull(IsBasicLine,0) = 1"
    End If
      StrSQL = StrSQL & " Order By name ASC"
    Else
    StrSQL = "SELECT     id,name FROM         dbo.TblProductLine  "
    If IsBase Then
        StrSQL = StrSQL & " And IsNull(IsBasicLine,0) = 1"
    End If
    
    StrSQL = StrSQL & " Order By name ASC"
    End If
    
 
  
    GetComboData MyCombo, StrSQL
End Function



Public Function GetTripStatus(MyCombo As DataCombo)
    Dim StrSQL As String
    StrSQL = "Select ID,Name From TblTypesTripStatus Order By ID"
    GetComboData MyCombo, StrSQL
End Function
 Public Function GetFinished(MyCombo As DataCombo, _
                                Optional Lngid As Long = 0)
    Dim StrSQL As String
    StrSQL = "Select ID,Name From TblFinished "

    If Lngid <> 0 Then
        StrSQL = StrSQL + " Where id=" & Lngid
    End If

    StrSQL = StrSQL + " Order By Name"

    GetComboData MyCombo, StrSQL
End Function


Public Function GetTblView(MyCombo As DataCombo, _
                                Optional Lngid As Long = 0)
    Dim StrSQL As String
    StrSQL = "Select ID,Name From TblView "

    If Lngid <> 0 Then
        StrSQL = StrSQL + " Where id=" & Lngid
    End If

    StrSQL = StrSQL + " Order By Name"

    GetComboData MyCombo, StrSQL
End Function





 
   




Public Function GetCarsOfCustomer(MyCombo As DataCombo, Optional CusID As Double)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT id,BoardNo From TblCusCar where CustomerID =" & CusID & " "
    Else
    StrSQL = "SELECT id,nBoardNo From TblCusCar where CustomerID =" & CusID & " "
    End If
    GetComboData MyCombo, StrSQL
End Function

  
Public Function GetQuicSearch(MyCombo As DataCombo, ByVal TxtSearch As String, ByVal mTableName As String, Optional ByVal mIDField As String = "id", Optional ByVal mNameField As String = "Name", Optional ByVal mNameeField As String = "Namee", Optional ByVal mFullCode As String = "", Optional ByVal mWhere As String = "")
Dim StrSQL As String
 
 If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = " Select  " & mIDField & " ," & mNameField & "   From " & mTableName & " Where 1 = 1  "
    If mWhere <> "" Then
        StrSQL = StrSQL & mWhere
    End If
    If Trim(TxtSearch) <> "" Then
        StrSQL = StrSQL & " and( " & mNameField & " like '%" & (TxtSearch) & "%'  or " & mNameeField & " like '%" & (TxtSearch) & "%'   "
        StrSQL = StrSQL & " or " & mIDField & " like '%" & (TxtSearch) & "%'  "
        If mFullCode <> "" Then
            StrSQL = StrSQL & " or " & mFullCode & " like '%" & (TxtSearch) & "%'  "
        End If
        StrSQL = StrSQL & " )"
        
    End If
    StrSQL = StrSQL & " order by " & mNameField
 Else
    StrSQL = " Select  " & mIDField & " ," & mNameeField & "   From " & mTableName & " Where 1 = 1  "
    If Trim(TxtSearch) <> "" Then
        StrSQL = StrSQL & " and( " & mNameField & " like '%" & (TxtSearch) & "%'  or " & mNameeField & " like '%" & (TxtSearch) & "%' or " & mIDField & " like '%" & (TxtSearch) & "%'  "
        If mFullCode <> "" Then
            StrSQL = StrSQL & " or " & mFullCode & " like '%" & (TxtSearch) & "%'  "
        End If
        StrSQL = StrSQL & " )"

    End If
    StrSQL = StrSQL & " order by " & mNameeField
    
 End If
    GetComboData MyCombo, StrSQL
End Function




Public Function getLocByStore(MyCombo As DataCombo, _
                                        Optional Lngid As Long = 0)
    Dim StrSQL As String
    StrSQL = "Select locid,Name,Namee From TblstoresLocations "

    If Lngid <> 0 Then
        StrSQL = StrSQL + " Where StoreId=" & Lngid
    End If

    StrSQL = StrSQL + " Order By locid"

    GetComboData MyCombo, StrSQL
End Function

Function GetProjects2(MyCombo As DataCombo)
  Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID,Project_name From  projects "
        StrSQL = StrSQL & " where  not (Project_name is null)and Project_name<>N'""'"
    Else
        StrSQL = "SELECT ID,Project_nameE From   projects "
        StrSQL = StrSQL & " where  not (Project_nameE is null)and Project_nameE<>N'""'"
    End If
    StrSQL = StrSQL & " and (Not (Fullcode Is Null))"
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = StrSQL & " order by  Project_name"
    Else
    StrSQL = StrSQL & " order by  Project_nameE"
    End If

    GetComboData MyCombo, StrSQL
    
End Function
 
 Function GetBartCarByVonder(MyCombo As DataCombo, Optional ID As Double)
    Dim StrSQL As String
   StrSQL = " SELECT     ID, accessory"
   StrSQL = StrSQL & "       From dbo.TblVendorCars"
   StrSQL = StrSQL & "     Where (ID = " & ID & ")"
    GetComboData MyCombo, StrSQL
End Function
Public Function GetClass(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT id,name From TblClass   "
    Else
        StrSQL = "SELECT id,namee From TblClass   "
    End If
    StrSQL = StrSQL + " Order By id "
    GetComboData MyCombo, StrSQL
End Function
 Function GetTypesTransport(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT id,Name From  TblTypesTransport order by  Name"
    Else
        StrSQL = "SELECT id,NameE From   TblTypesTransport order by   NameE"
    End If

    GetComboData MyCombo, StrSQL
End Function

 Function GetHarbors(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT id,Name From  TblHarborsData order by  Name"
    Else
        StrSQL = "SELECT id,NameE From   TblHarborsData order by   NameE"
    End If

    GetComboData MyCombo, StrSQL
End Function
 Function GetShips(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT id,Name From  TblShipsData order by  Name"
    Else
        StrSQL = "SELECT id,NameE From   TblShipsData order by   NameE"
    End If

    GetComboData MyCombo, StrSQL
End Function
 Function GetCarByVonder(MyCombo As DataCombo, Optional customerid As Double)
    Dim StrSQL As String
   StrSQL = " SELECT     ID, BoardNo"
   StrSQL = StrSQL & "       From dbo.TblVendorCars"
   If customerid <> 0 Then
   StrSQL = StrSQL & "     Where (customerid = " & customerid & ")"
   End If
    GetComboData MyCombo, StrSQL
End Function

   Function GetPartCar(MyCombo As DataCombo, Optional CarID As Double)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = " SELECT     dbo.TblCarsDataDet.PartID, dbo.FixedAssets.Name"
    Else
        StrSQL = "SELECT     dbo.TblCarsDataDet.PartID, dbo.FixedAssets.namee"
    End If
        StrSQL = StrSQL & " FROM         dbo.TblCarsDataDet RIGHT OUTER JOIN"
        StrSQL = StrSQL & "              dbo.TblCarsData ON dbo.TblCarsDataDet.EqupID = dbo.TblCarsData.fixedAssetid LEFT OUTER JOIN"
        StrSQL = StrSQL & "              dbo.FixedAssets ON dbo.TblCarsDataDet.PartID = dbo.FixedAssets.id"
        StrSQL = StrSQL & " Where (dbo.TblCarsData.ID = " & CarID & ")"
    GetComboData MyCombo, StrSQL
End Function

 '####################### khaled fill combo functions ########################
Public Function GetArch(MyCombo As DataCombo, Optional ID As Integer = 0)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select Id,Name From TblXXArch where DepID =" & ID & " Order By Name"
    Else
        StrSQL = "Select Id,Namee From TblXXArch where DepID =" & ID & " Order By Namee"
    End If
    GetComboData MyCombo, StrSQL
End Function

Public Function GetRoom(MyCombo As DataCombo, Optional ID As Integer = 0)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select Id,Name From TblXXRoom where ArchID =" & ID & "Order By Name"
    Else
        StrSQL = "Select Id,Namee From TblXXRoom where ArchID =" & ID & "Order By Namee"
    End If
    GetComboData MyCombo, StrSQL
End Function
Public Function GetBox(MyCombo As DataCombo, Optional ID As Integer = 0)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select Id,Name From TblXXBox where RoomID =" & ID & " Order By Name"
    Else
        StrSQL = "Select Id,Namee From TblXXBox where RoomID =" & ID & " Order By Namee"
    End If
    GetComboData MyCombo, StrSQL
End Function
Public Function GetShelf(MyCombo As DataCombo, Optional ID As Integer = 0)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select Id,Name From TblXXShelf where BoxID =" & ID & " Order By Name"
    Else
        StrSQL = "Select Id,Namee From TblXXShelf where BoxID =" & ID & " Order By Namee"
    End If
    GetComboData MyCombo, StrSQL
End Function
Public Function GetArchDocType(MyCombo As DataCombo, Optional ID As Integer = 0)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select Id,Name From TblXXArchDocType "
    Else
        StrSQL = "Select Id,Namee From TblXXArchDocType  "
    End If
    If ID <> 0 Then
    StrSQL = StrSQL & " Where ShelfID = " & ID & ""
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = StrSQL & " Order By Name "
    Else
    StrSQL = StrSQL & " Order By Namee "
    End If
    GetComboData MyCombo, StrSQL
End Function
'############################################################################
 
 
Public Function GetProjecInvestmentSquare(MyCombo As DataCombo, Optional ProjectCode As String)
    Dim StrSQL As String
    StrSQL = "SELECT     ID, Square"
    StrSQL = StrSQL & "     From dbo.TblProjecInvestment"
    If ProjectCode <> "" Then
    StrSQL = StrSQL & "  WHERE     (ProjectCode = N'" & ProjectCode & "')"
    End If
    GetComboData MyCombo, StrSQL
End Function
Public Function GetModelInvestment(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT ID ,Name "
    Else
    StrSQL = "SELECT ID , NameE"
    End If
    StrSQL = StrSQL & "   From  dbo.TblModelInves"
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = StrSQL & " order by Name"
    Else
    StrSQL = StrSQL & " order by NameE"
    End If
    GetComboData MyCombo, StrSQL
End Function

Public Function GetProjecInvestment(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT DISTINCT ProjectCode,Name "
    Else
    StrSQL = "SELECT DISTINCT ProjectCode , NameE"
    End If
    StrSQL = StrSQL & "   From dbo.TblProjecInvestment"
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = StrSQL & " order by Name"
    Else
    StrSQL = StrSQL & " order by NameE"
    End If
    GetComboData MyCombo, StrSQL
End Function


Public Function GetEmpSalaryCode(MyCombo As DataCombo)
    Dim StrSQL As String
    StrSQL = "SELECT DISTINCT SalaryCode, SalaryCode AS Name"
    StrSQL = StrSQL & "  From dbo.TblEmployee"
    StrSQL = StrSQL & "  Where (Not (SalaryCode Is Null))"
    GetComboData MyCombo, StrSQL
End Function
Public Function GetUserComp(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID, Name From TblUserComp"
    Else
        StrSQL = "SELECT ID, NameE From TblUserComp"
    End If
    GetComboData MyCombo, StrSQL
End Function
Public Function GetTblTypePath(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT  id,name From TblHotels  order by name"
    Else
    StrSQL = "SELECT  id,NameE From  TblHotels  order by NameE"
   End If
    GetComboData MyCombo, StrSQL
End Function
Public Function GetPand(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID,Name From TblPands "
        StrSQL = StrSQL + " Order BY Name "
    Else
        StrSQL = "SELECT ID,NameE From TblPands "
        StrSQL = StrSQL + " Order BY NameE "
    End If

    GetComboData MyCombo, StrSQL
End Function

 Public Function GetLand(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = " SELECT     ID, Name"
    StrSQL = StrSQL & "          From dbo.TblBuyLanReEst"

    Else
 StrSQL = " SELECT     ID, NameE"
 StrSQL = StrSQL & " From dbo.TblBuyLanReEst"
End If
   GetComboData MyCombo, StrSQL
End Function

Public Function GetEmpVactionAreth(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
      StrSQL = "  SELECT     Vac_ID, Vac_Name"
    Else
        StrSQL = " SELECT     Vac_ID , NameE"
    End If
       StrSQL = StrSQL & "  From dbo.tblVacancy "
       StrSQL = StrSQL & " Where(VacAretha = 1)"
    GetComboData MyCombo, StrSQL
End Function

Public Function GetNewDwpartMent(MyCombo As DataCombo, _
                             Optional Typ As Boolean = False, _
                             Optional MangID As Double)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT Id,name From TblEmpDepartmentsDet Where 1=1  "
    Else
        StrSQL = "SELECT Id,namee From TblEmpDepartmentsDet Where 1=1  "
    End If

    If Typ = True Then
        StrSQL = StrSQL + " and  DeparmentID=" & MangID & ""
    End If

    StrSQL = StrSQL + " Order By name ASC"
    GetComboData MyCombo, StrSQL
End Function
Public Function GetCompany(MyCombo As DataCombo, Optional In_out As Integer = -1, Optional Hj_omr As Integer = -1)
    Dim StrSQL As String
  If SystemOptions.UserInterface = ArabicInterface Then
   StrSQL = " SELECT     dbo.TblCustemers.CusID, dbo.TblCustemers.CusName"
   Else
   StrSQL = " SELECT     dbo.TblCustemers.CusID, dbo.TblCustemers.CusNamee"
   End If
   StrSQL = StrSQL & "     FROM         dbo.TblCustemers LEFT OUTER JOIN"
   StrSQL = StrSQL & "   dbo.TblCustomerType ON dbo.TblCustemers.CustomerTypeID = dbo.TblCustomerType.id"
   StrSQL = StrSQL & " where 1=1"
   If In_out <> 2 Then
 StrSQL = StrSQL & " and  HajEnter_Out=" & In_out & ""
 Else
 StrSQL = StrSQL & " and  (HajEnter_Out=1 or HajEnter_Out=0)"
 End If
 If Hj_omr = 0 Then
StrSQL = StrSQL & " and  TypeOmrh=1"
ElseIf Hj_omr = 1 Then
StrSQL = StrSQL & " and  TypeHaj=1"
Else
StrSQL = StrSQL & " and  (TypeHaj=1 or TypeOmrh=1)"
End If
    GetComboData MyCombo, StrSQL
End Function
Public Function GetCompanyXX(MyCombo As DataCombo)
    Dim StrSQL As String
  If SystemOptions.UserInterface = ArabicInterface Then
   StrSQL = " SELECT     dbo.TblCustemers.CusID, dbo.TblCustemers.CusName"
   Else
   StrSQL = " SELECT     dbo.TblCustemers.CusID, dbo.TblCustemers.CusNamee"
   End If
   StrSQL = StrSQL & "     FROM         dbo.TblCustemers LEFT OUTER JOIN"
   StrSQL = StrSQL & "                   dbo.TblCustomerType ON dbo.TblCustemers.CustomerTypeID = dbo.TblCustomerType.id"
  StrSQL = StrSQL & " Where (dbo.TblCustomerType.TypeHajj = 1)"

    GetComboData MyCombo, StrSQL
End Function


Public Function GetDeptCars(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT     DeparmentID ,  DepartmentName"
    Else
    StrSQL = "SELECT     DeparmentID  ,  DepartmentNamee"
    End If
    StrSQL = StrSQL & " FROM  TblEmpDepartments "
  
 
    StrSQL = StrSQL & " Order By DepartmentName ASC"
    GetComboData MyCombo, StrSQL
End Function

Public Function GetTypeDependence(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT  id,name From TblTypeDependence  order by name"
    Else
    StrSQL = "SELECT  id,NameE From TblTypeDependence  order by NameE"
   End If
    GetComboData MyCombo, StrSQL
End Function
Public Function GetSickleave(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID,Name From TblSickleave order by Name"
    Else
        StrSQL = "SELECT ID,NameE From TblSickleave order by NameE "
    End If
    GetComboData MyCombo, StrSQL
End Function
Public Function GetTypeClaim(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID,Name From TblTypeClaim order by Name"
    Else
        StrSQL = "SELECT ID,NameE From TblTypeClaim order by NameE "
    End If
    GetComboData MyCombo, StrSQL
End Function
Public Function GetStudentTypeTrining(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID, Name From TblStudentTypeTrinng order by Name"
    Else
        StrSQL = "SELECT ID, NameE From TblStudentTypeTrinng order By NameE"
    End If
    GetComboData MyCombo, StrSQL
End Function
Public Function GetStudentDeploma(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID, Name From TblStudentTypeCurs  order by Name"
    Else
        StrSQL = "SELECT ID, NameE From TblStudentTypeCurs  order By NameE"
    End If
    GetComboData MyCombo, StrSQL
End Function


Public Function GetStudentClassRooms(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID, Name From TblStudentClassRooms  order by Name"
    Else
        StrSQL = "SELECT ID, NameE From TblStudentClassRooms  order By NameE"
    End If
    GetComboData MyCombo, StrSQL
End Function

 

Public Function GetItemsUnitsDetai(MyCombo As DataCombo, Optional ItemID As Double)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT dbo.TblUnites.UnitID, dbo.TblUnites.UnitName"
    Else
        StrSQL = "SELECT dbo.TblUnites.UnitID, dbo.TblUnites.UnitNamee"
    End If
    StrSQL = StrSQL + "   FROM  dbo.TblItemsUnits LEFT OUTER JOIN"
     StrSQL = StrSQL + "  dbo.TblUnites ON dbo.TblItemsUnits.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " Where dbo.TblItemsUnits.ItemID=" & ItemID & " "
    GetComboData MyCombo, StrSQL
End Function
Public Function GeInstructor(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID, Name From TblInstructors  order by Name"
    Else
        StrSQL = "SELECT ID, NameE From TblInstructors  order By NameE"
    End If
    GetComboData MyCombo, StrSQL
End Function


 


Public Function GetStudentGroup(MyCombo As DataCombo, Optional Typ As Integer = 0)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID, Name From TblStuGroup "
        If Typ = 1 Then
        StrSQL = StrSQL & " where (FlgEnd is null )or FlgEnd=2 or FlgEnd=0"
        End If
        
       StrSQL = StrSQL & "  order by Name"
    Else
        StrSQL = "SELECT ID, NameE From TblStuGroup "
        StrSQL = StrSQL & " where (FlgEnd is null )or FlgEnd=2 or FlgEnd=0"
        StrSQL = StrSQL & "  order order By NameE"
    End If
    
    GetComboData MyCombo, StrSQL
End Function

Public Function GetStudent(MyCombo As DataCombo, Optional Typ As Integer = 0, Optional CompID As Double = 0)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID, Name From TblStudent where 1<>-1"
    Else
        StrSQL = "SELECT ID, NameE From TblStudent where 1<>-1"
    End If
    If Typ = 0 Then
    StrSQL = StrSQL & " and (StutsID <> 1 or (StutsID is null)) "
    End If
     If CompID <> 0 Then
    StrSQL = StrSQL & " and CompID =" & CompID & " "
    End If
    
        If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = StrSQL & " order by name "
    Else
        StrSQL = StrSQL & " Order by NameE "
    End If
    GetComboData MyCombo, StrSQL
End Function
Public Function GetStudentx2xxxx(MyCombo As DataCombo, Optional Typ As Integer = 0)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID, Name From TblStudent "
    Else
        StrSQL = "SELECT ID, NameE From TblStudent "
    End If
    If Typ = 0 Then
    StrSQL = StrSQL & " where StutsID <> 1 or (StutsID is null) "
    End If
        If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = StrSQL & " order by name "
    Else
        StrSQL = StrSQL & " Order by NameE "
    End If
    GetComboData MyCombo, StrSQL
End Function

Public Function GetStudentxx(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID, Name From TblStudent order by Name"
    Else
        StrSQL = "SELECT ID, NameE From TblStudent order By NameE"
    End If
    GetComboData MyCombo, StrSQL
End Function

Public Function GetStudentContract(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID, Name From TblStudentContract order by Name"
    Else
        StrSQL = "SELECT ID, NameE From TblStudentContract order By NameE"
    End If
    GetComboData MyCombo, StrSQL
End Function

Public Function GetStudentCurs(MyCombo As DataCombo, Optional ID As Double)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID, Name From TblStudentCurs "
    Else
        StrSQL = "SELECT ID, NameE From TblStudentCurs "
    End If
    If ID <> 0 Then
    StrSQL = StrSQL & " where CursTypeID =" & ID & ""
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = StrSQL & "order by Name"
    Else
    StrSQL = StrSQL & "order by NameE"
    End If
    GetComboData MyCombo, StrSQL
End Function

Public Function GetStudentCursxxxx(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID, Name From TblStudentCurs order by Name"
    Else
        StrSQL = "SELECT ID, NameE From TblStudentCurs order By NameE"
    End If
    GetComboData MyCombo, StrSQL
End Function
'**************************
Public Function GetStudentTypeCurs(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "Select id,name  From TblStudentTypeCurs    Order By name"
   Else
   StrSQL = "Select id,namee  From TblStudentTypeCurs   Order By namee"
   End If
      GetComboData MyCombo, StrSQL
End Function

Public Function GetStudentQualification(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID, Name From TblStudentQualification order by Name"
    Else
        StrSQL = "SELECT ID, NameE From TblStudentQualification order By NameE"
    End If
    GetComboData MyCombo, StrSQL
End Function


Public Function GetStudentTeachers(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID, Name From TblStudentTeachers order by Name"
    Else
        StrSQL = "SELECT ID, NameE From TblStudentTeachers order By NameE"
    End If
    GetComboData MyCombo, StrSQL
End Function

 

Public Function GetTblShrines(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT  id,name From TblShrines  order by name"
    Else
    StrSQL = "SELECT  id,NameE From TblShrines  order by NameE"
   End If
    GetComboData MyCombo, StrSQL
End Function
Public Function GetPartkLand(MyCombo As DataCombo, LandID As Double, BlockID As Double)
    Dim StrSQL As String
        StrSQL = " SELECT     dbo.TblDivInvestInformation.ID, dbo.TblDivInvestInformation.BlokNo"
        StrSQL = StrSQL & "      FROM         dbo.TblDivInvesment LEFT OUTER JOIN"
        StrSQL = StrSQL & "               dbo.TblDivInvestInformation ON dbo.TblDivInvesment.ID = dbo.TblDivInvestInformation.DivID"
        StrSQL = StrSQL & "      Where (dbo.TblDivInvesment.LandID = " & LandID & ") And (dbo.TblDivInvestInformation.DivIDDet = " & BlockID & ")"
        StrSQL = StrSQL & "        GROUP BY dbo.TblDivInvestInformation.PartNo, dbo.TblDivInvestInformation.DivIDDet, dbo.TblDivInvesment.LandID, dbo.TblDivInvestInformation.ID,"
        StrSQL = StrSQL & "              dbo.TblDivInvestInformation.BlokNo"
   GetComboData MyCombo, StrSQL
End Function
Public Function GetBlockLand(MyCombo As DataCombo, LandID As Double)
    Dim StrSQL As String
        StrSQL = "SELECT     dbo.TblDivInvestInformation.DivIDDet, dbo.TblDivInvestInformation.PartNo"
        StrSQL = StrSQL & "  FROM         dbo.TblDivInvesment LEFT OUTER JOIN"
        StrSQL = StrSQL & "                      dbo.TblDivInvestInformation ON dbo.TblDivInvesment.ID = dbo.TblDivInvestInformation.DivID"
        StrSQL = StrSQL & "   Where (dbo.TblDivInvesment.LandID = " & LandID & ")"
        StrSQL = StrSQL & "  GROUP BY dbo.TblDivInvestInformation.PartNo, dbo.TblDivInvestInformation.DivIDDet, dbo.TblDivInvesment.LandID"
   GetComboData MyCombo, StrSQL
End Function
Public Function GetTblLocations(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT  id,name From TblLocations  order by name"
    Else
    StrSQL = "SELECT  id,NameE From TblLocations  order by NameE"
   End If
    GetComboData MyCombo, StrSQL
End Function
Public Function GetTblTrips(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT  id,name From TblTrips  order by name"
    Else
    StrSQL = "SELECT  id,NameE From TblTrips  order by NameE"
   End If
    GetComboData MyCombo, StrSQL
End Function
 Public Function GetTblProgrammTypes(MyCombo As DataCombo, Optional VehicleType As Integer = 0)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT  id,name From TblProgrammTypes "
    If VehicleType <> 0 Then
   StrSQL = StrSQL & " Where VehicleType = " & VehicleType & ""
   End If
   StrSQL = StrSQL & "  order by name"
    Else
    StrSQL = "SELECT  id,NameE From TblProgrammTypes "
    If VehicleType <> 0 Then
    StrSQL = StrSQL & " where VehicleType=" & VehicleType & " "
    End If
    StrSQL = StrSQL & " order by NameE"
   End If
    GetComboData MyCombo, StrSQL
End Function

Public Function GetVehicleType(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID, Name From TblVehicleType"
    Else
        StrSQL = "SELECT ID, NameE From TblVehicleType"
    End If
    GetComboData MyCombo, StrSQL
End Function

 Public Function GetLandNotActive(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = " SELECT     ID, Name"
    StrSQL = StrSQL & "          From dbo.TblBuyLanReEst"
    StrSQL = StrSQL & "         WHERE     (ID NOT IN"
    StrSQL = StrSQL & "        (SELECT     LandOwnedID"
    StrSQL = StrSQL & "          FROM         dbo.TblActivateInvestment))"
    Else
 StrSQL = " SELECT     ID, NameE"
 StrSQL = StrSQL & " From dbo.TblBuyLanReEst"
 StrSQL = StrSQL & "  WHERE     (ID NOT IN"
 StrSQL = StrSQL & "                         (SELECT     LandOwnedID"
 StrSQL = StrSQL & "                            FROM         dbo.TblActivateInvestment))"
End If
   GetComboData MyCombo, StrSQL
End Function


 

  Public Function GetTblSpreadingInvestment(MyCombo As DataCombo, Optional Typ As Integer = 0)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT id,Name From TblSpreading"
    Else
        StrSQL = "SELECT id,NameE From TblSpreading"
    End If
   If Typ = 1 Then
StrSQL = StrSQL & " where UnitMain=1  "
Else
StrSQL = StrSQL & " where UnitMain=1 or UnitDetails=1 "
End If
   GetComboData MyCombo, StrSQL
End Function



Public Function GetLandActive(MyCombo As DataCombo, Optional InvestID As Integer = 0)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "   SELECT     dbo.TblActivateInvestment.LandOwnedID, dbo.TblBuyLanReEst.Name"
    StrSQL = StrSQL & "        FROM         dbo.TblActivateInvestment LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblBuyLanReEst ON dbo.TblActivateInvestment.LandOwnedID = dbo.TblBuyLanReEst.ID"
    Else
    StrSQL = "   SELECT     dbo.TblActivateInvestment.LandOwnedID, dbo.TblBuyLanReEst.NameE"
    StrSQL = StrSQL & "        FROM         dbo.TblActivateInvestment LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblBuyLanReEst ON dbo.TblActivateInvestment.LandOwnedID = dbo.TblBuyLanReEst.ID"
    End If
StrSQL = StrSQL & " Where (dbo.TblActivateInvestment.InviseNo = " & InvestID & ")"
   GetComboData MyCombo, StrSQL
End Function
 
 Public Function GetBuyLandRealEstate(MyCombo As DataCombo, Optional Typ As Integer = 0, Optional ActivID As Double = 0)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID,Name From TblBuyLanReEst"
    Else
        StrSQL = "SELECT ID,NameE From TblBuyLanReEst"
    End If
    StrSQL = StrSQL & " where (FlgReturn is NULL)"
    If Typ = 1 Then
    StrSQL = StrSQL & " and ( ActivID <>" & ActivID & " or  (ActivID IS NULL))  "
    ElseIf Typ = 2 Then
    StrSQL = StrSQL & " and  (ActivID IS NULL)  "
    End If
    
   GetComboData MyCombo, StrSQL
End Function

 Public Function GetInvestmentActive(MyCombo As DataCombo, Optional Typ As Integer = 0)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID,Name From dbo.Tblinvestment"
    Else
        StrSQL = "SELECT ID,NameE From dbo.Tblinvestment"
    End If
    If Typ = 0 Then
StrSQL = StrSQL & " WHERE     (FlagActive = 1)"
End If
   GetComboData MyCombo, StrSQL
End Function
 

Public Function ChekClodePeriod(RecordDate As Date) As Boolean
Dim Rs7 As ADODB.Recordset
Dim sql As String
ChekClodePeriod = False
Set Rs7 = New ADODB.Recordset
sql = "SELECT     StartDate, EndDate"
sql = sql & " From dbo.TblOpenClosPeriodDet1"
sql = sql & "  WHERE     (StartDate <=" & SQLDate(RecordDate, True) & " AND (EndDate >= " & SQLDate(RecordDate, True) & " ))"
Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
ChekClodePeriod = True
Else
ChekClodePeriod = False
End If
End Function

Public Function GetInvStoreGroup(MyCombo As DataCombo)
    Dim StrSQL As String
If checkmanyBranches(StrSQL) = False Then
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID,Name From TblInvestorsGroup"
    Else
        StrSQL = "SELECT Id,NameE From TblInvestorsGroup"
    End If
End If
   GetComboData MyCombo, StrSQL
End Function

Public Function GetInvStoreType(MyCombo As DataCombo)
    Dim StrSQL As String
If checkmanyBranches(StrSQL) = False Then
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID,Name From TblInvestorType"
    Else
        StrSQL = "SELECT Id,NameE From TblInvestorType"
    End If
End If
   GetComboData MyCombo, StrSQL
End Function

Public Function GetInvestmentType(MyCombo As DataCombo)
    Dim StrSQL As String
If checkmanyBranches(StrSQL) = False Then
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID,Name From TblShareType"
    Else
        StrSQL = "SELECT Id,NameE From TblShareType"
    End If
End If
   GetComboData MyCombo, StrSQL
End Function

Public Function GetInvestmentGroup(MyCombo As DataCombo)
    Dim StrSQL As String
If checkmanyBranches(StrSQL) = False Then
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID,Name From TblSharesGroup"
    Else
        StrSQL = "SELECT Id,NameE From TblSharesGroup"
    End If
End If
   GetComboData MyCombo, StrSQL
End Function

Public Function GetDevelopProcessPand(MyCombo As DataCombo, Optional DevlOpID As Integer = 0)
    Dim StrSQL As String
If SystemOptions.UserInterface Then
   StrSQL = "SELECT Id,Des From TblProceeDevelperDet Where DevlOpID=" & DevlOpID & "  "
   StrSQL = StrSQL + " Order By Des ASC"
   Else
   StrSQL = "SELECT Id,Des From TblProceeDevelperDet Where DevlOpID=" & DevlOpID & "  "
   StrSQL = StrSQL + " Order By Des ASC"
   End If
    GetComboData MyCombo, StrSQL
End Function

 

Public Function GetDevelopProcess(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT Id,name From TblProceeDevelper Where 1=1  "
        StrSQL = StrSQL + " Order By Name ASC"
    Else
        StrSQL = "SELECT Id,namee From TblProceeDevelper Where 1=1  "
        StrSQL = StrSQL + " Order By NameE ASC"
    End If
    
    GetComboData MyCombo, StrSQL
End Function

 
Public Function GetAdvancedPartNo(MyCombo As DataCombo, Optional EmpID As Double = 0, Optional PartDate As Date)
    Dim StrSQL As String
    StrSQL = "SELECT     TOP 100 PERCENT dbo.TblEmpAdvanceDetails.TableID, dbo.TblEmpAdvanceDetails.PartNO"
    StrSQL = StrSQL & "  FROM         dbo.TblEmpAdvanceDetails RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblEmpAdvance ON dbo.TblEmpAdvanceDetails.AdvanceID = dbo.TblEmpAdvance.AdvanceID"
    StrSQL = StrSQL & "     WHERE     (dbo.TblEmpAdvanceDetails.PartDate >= " & SQLDate(PartDate, True) & ") AND "
   StrSQL = StrSQL & " (dbo.TblEmpAdvanceDetails.Payed IS NULL or dbo.TblEmpAdvanceDetails.Payed=0)and"
   StrSQL = StrSQL & " (dbo.TblEmpAdvanceDetails.Payed1 IS NULL) AND (dbo.TblEmpAdvanceDetails.StutsID IS NULL OR"
    StrSQL = StrSQL & "                   dbo.TblEmpAdvanceDetails.StutsID = 21 OR"
    StrSQL = StrSQL & "                   dbo.TblEmpAdvanceDetails.StutsID = 22 OR"
    StrSQL = StrSQL & "                   dbo.TblEmpAdvanceDetails.StutsID = 23) AND (dbo.TblEmpAdvance.Emp_ID = " & EmpID & ")"
    StrSQL = StrSQL & "    ORDER BY dbo.TblEmpAdvanceDetails.PartDate, dbo.TblEmpAdvanceDetails.PartNO"
   ' StrSQL = StrSQL & "   WHERE   (PartDate>=" & SQLDate(PartDate, True) & ")   and (Payed1 IS NULL)  AND (StutsID IS NULL or StutsID =21 or StutsID =22 or StutsID =23) and  (AdvanceID = " & advanceID & ")"

'StrSQL = StrSQL & " order by PartNo"
    GetComboData MyCombo, StrSQL
End Function
Public Function GetAdvancedPartNoxx(MyCombo As DataCombo, Optional advanceID As Integer = 0)
    Dim StrSQL As String
    StrSQL = "SELECT     id, PartNo"
    StrSQL = StrSQL & "  From dbo.TblEmpAdvanceRequestDetails"
    StrSQL = StrSQL & "   WHERE     (Payed1 IS NULL)  AND (AdvanceID = " & advanceID & ")"

StrSQL = StrSQL & " order by PartNo"
    GetComboData MyCombo, StrSQL
End Function
Public Function GetAdminSanction(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select ID,Name From TblAdminSanction Order By Name"
    Else
        StrSQL = "Select ID,NameE From TblAdminSanction Order By NameE"
    End If

    GetComboData MyCombo, StrSQL
End Function
Public Function GetInsurancesCode(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT id,name From mofrad where acc=1    order by name"
    Else
        StrSQL = "SELECT id,nameE From mofrad  where acc=1   order by nameE"
    End If

    GetComboData MyCombo, StrSQL
End Function

Public Function Getemp_Contract_type(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT id, name   From emp_contract_type   "
    Else
        StrSQL = "SELECT id, namee  From emp_contract_type  "
    End If


    StrSQL = StrSQL + " Order By id ASC"
    GetComboData MyCombo, StrSQL
End Function
'''''''''''''''//////////////
Public Function GetReturnType(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT Id,name From TblTypeDiscards Where 1=1  "
    Else
        StrSQL = "SELECT Id,namee From TblTypeDiscards Where 1=1  "
    End If

       StrSQL = StrSQL + " Order By name ASC"
    GetComboData MyCombo, StrSQL
End Function
Public Function GetAmanhNames(MyCombo As DataCombo, Optional GovernmentID As Integer = 0)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "Select ID,name From TblMunicipality Order By id"
    Else
    StrSQL = "Select ID,namee From TblMunicipality Order By id"
    End If
    If GovernmentID <> 0 Then
    StrSQL = StrSQL & " where  GovernmentID =" & GovernmentID & " "
    End If
    GetComboData MyCombo, StrSQL
End Function
Public Function GetMunicipalityNames(MyCombo As DataCombo, Optional MunicID As Integer = 0)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "Select ID,name From TblMunicipalityDet Order By id"
    Else
    StrSQL = "Select ID,namee From TblMunicipalityDet Order By id"
    End If
    If MunicID <> 0 Then
    StrSQL = StrSQL & " where  MunicID =" & MunicID & " "
    End If
    GetComboData MyCombo, StrSQL
End Function

Public Function GetShipingMathods(MyCombo As DataCombo)
    Dim StrSQL As String
     If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT Id,Name From TblShipingData"
    Else
        StrSQL = "SELECT Id,NameE From TblShipingData"
    End If
    GetComboData MyCombo, StrSQL
End Function

Public Function GetPaymentMathods(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT Id,Name From TblPaymetData"
    Else
        StrSQL = "SELECT Id,NameE From TblPaymetData"
    End If
    GetComboData MyCombo, StrSQL
End Function

Public Function GetCatalogItem(MyCombo As DataCombo, _
                           Optional ItemID As Integer = 0)
    Dim StrSQL As String
        StrSQL = "SELECT ID,CatlogName From TblItemCatalog  "
   If ItemID > 0 Then
StrSQL = StrSQL & " Where ItemID = " & ItemID & ""
End If
    GetComboData MyCombo, StrSQL
End Function

Public Function GetProcess(MyCombo As DataCombo)

    Dim StrSQL As String
If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT TblProcessDEFID,ProcessName From TblProcessDEF "
   Else
   StrSQL = "SELECT TblProcessDEFID,ProcessNameE From TblProcessDEF "
   End If

    StrSQL = StrSQL + " Order By TblProcessDEFID "
    GetComboData MyCombo, StrSQL
End Function
Public Function GetProcessOfProjedt(MyCombo As DataCombo, Optional proj As Integer = 0, Optional Pand As String, Optional pandid As Integer, Optional para As Integer = 0)
    Dim StrSQL As String
     
If para <> 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
     StrSQL = "SELECT     dbo.TblProcessDEF.TblProcessDEFID, dbo.TblProcessDEF.ProcessName"
      StrSQL = StrSQL & "           FROM         dbo.terms_operations LEFT OUTER JOIN"
      StrSQL = StrSQL & "                  dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEFID"
    Else
       StrSQL = "SELECT     dbo.TblProcessDEF.TblProcessDEFID, dbo.TblProcessDEF.ProcessNameE"
      StrSQL = StrSQL & "           FROM         dbo.terms_operations LEFT OUTER JOIN"
      StrSQL = StrSQL & "                  dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEFID"
     
    End If
    If para = 1 Then
    StrSQL = StrSQL & "   WHERE     (dbo.terms_operations.project_id = " & proj & ") AND (dbo.terms_operations.term_fullcode = '" & Pand & "')"
    Else
    StrSQL = StrSQL & "   WHERE     (dbo.terms_operations.project_id = " & proj & ") AND (dbo.terms_operations.ProjectDes_ID = " & pandid & ")"
    End If
Else
  If SystemOptions.UserInterface = ArabicInterface Then
       StrSQL = "SELECT     dbo.TblProcessDEF.TblProcessDEFID, dbo.TblProcessDEF.ProcessName"
      StrSQL = StrSQL & "           FROM         dbo.terms_operations LEFT OUTER JOIN"
      StrSQL = StrSQL & "                  dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEFID"
  
    Else
       StrSQL = "SELECT     dbo.TblProcessDEF.TblProcessDEFID, dbo.TblProcessDEF.ProcessNameE"
      StrSQL = StrSQL & "           FROM         dbo.terms_operations LEFT OUTER JOIN"
      StrSQL = StrSQL & "                  dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEFID"
     End If

End If



    GetComboData MyCombo, StrSQL
End Function

Public Function GetRentStatus(MyCombo As DataCombo)
    Dim StrSQL As String


    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT id,name From TblRentStatus"
    Else
        StrSQL = "SELECT id,namee From TblRentStatus"
    End If


    GetComboData MyCombo, StrSQL
End Function
Public Function GetRentSatus(MyCombo As DataCombo, _
                             Optional WorkState As Boolean = False, _
                             Optional DriverOnly As Boolean = False)
    Dim StrSQL As String


    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT Id,name From TblRentStatus  "
    Else
        StrSQL = "SELECT Id,namee From TblRentStatus   "
    End If
 
    StrSQL = StrSQL + " Order By Id ASC"
    GetComboData MyCombo, StrSQL
End Function


Public Function GETNationality(MyCombo As DataCombo)
    Dim StrSQL As String
     

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT id,name From Nationality order by name "
    Else
        StrSQL = "SELECT id,namee From Nationality  order by namee "
    End If




    GetComboData MyCombo, StrSQL
End Function
Public Function GetTypeExchange(MyCombo As DataCombo, _
                             Optional WorkState As Boolean = False, _
                             Optional DriverOnly As Boolean = False)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT Id,name From TblDataTypeExchange  "
    Else
        StrSQL = "SELECT Id,namee From TblDataTypeExchange   "
    End If
 
    StrSQL = StrSQL + " Order By Id ASC"
    GetComboData MyCombo, StrSQL
End Function
  Public Function GetIqarCode(Optional EmpCode As String, _
                                      Optional ByRef Emp_id As Double, _
                                      Optional Emp_id1 As Double = 0, _
                                      Optional ByRef EmpCode1 As String, Optional ownerid As Double)
            
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double

    If Emp_id1 <> 0 Then
        sql = "select * from TblAqar where Aqarid= " & Emp_id1
    Else
 
        sql = "select * from TblAqar where  aqarNo ='" & EmpCode & "'"
    End If
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Emp_id = IIf(IsNull(rs("Aqarid").value), 0, rs("Aqarid").value)
        EmpCode1 = IIf(IsNull(rs("aqarNo").value), 0, rs("aqarNo").value)
   ownerid = IIf(IsNull(rs("ownerid").value), 0, rs("ownerid").value)
   
    Else
        Emp_id = 0
    End If

    rs.Close

End Function


 Public Function GetTblCustemersCode(Optional EmpCode As String, _
                                      Optional ByRef Emp_id As Integer, _
                                      Optional Emp_id1 As Integer = 0, _
                                      Optional ByRef EmpCode1 As String, Optional Type1 As Integer = 1)
            
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double

    If Emp_id1 <> 0 Then
        sql = "select * from TblCustemers where CusID= " & Emp_id1
    Else
    sql = "select * from TblCustemers where  Fullcode ='" & EmpCode & "'"
 If Type1 = 2 Then
    sql = sql & " AND Type  = 2 "
        End If
    End If
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then

        Emp_id = IIf(IsNull(rs("CusID").value), 0, rs("CusID").value)
        EmpCode1 = IIf(IsNull(rs("Fullcode").value), 0, rs("Fullcode").value)
 
    Else
        Emp_id = 0
    End If

    rs.Close

End Function
 
 
Public Function GetFileCustomer(MyCombo As DataCombo)
    Dim StrSQL As String
    StrSQL = "Select CusID,CusName From TblCustemers Order By CusName"
    GetComboData MyCombo, StrSQL
End Function
Public Function GetTypeDevelopment(MyCombo As DataCombo, _
                             Optional WorkState As Boolean = False, _
                             Optional DriverOnly As Boolean = False)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT Id,name From TblOpDevelopment Where 1=1  "
    Else
        StrSQL = "SELECT Id,namee From TblOpDevelopment Where 1=1  "
    End If

    If WorkState = True Then
        StrSQL = StrSQL + " and  workstate=1"
    End If

    If DriverOnly = True Then
        StrSQL = StrSQL + " and  not(driverid is null)"
    End If

    StrSQL = StrSQL + " Order By Id ASC"
    GetComboData MyCombo, StrSQL
End Function
Public Function GetSection(MyCombo As DataCombo, _
                             Optional WorkState As Boolean = False, _
                             Optional DriverOnly As Boolean = False)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT Id,name From TblSection Where 1=1  "
    Else
        StrSQL = "SELECT Id,namee From TblSection Where 1=1  "
    End If

    If WorkState = True Then
        StrSQL = StrSQL + " and  workstate=1"
    End If

    If DriverOnly = True Then
        StrSQL = StrSQL + " and  not(driverid is null)"
    End If

    StrSQL = StrSQL + " Order By name ASC"
    GetComboData MyCombo, StrSQL
End Function

Public Function GetTimeDelegate(MyCombo As DataCombo, _
                             Optional WorkState As Boolean = False, _
                             Optional DriverOnly As Boolean = False)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT id, name   From TblRegTimeDelgate Where 1=1  "
    Else
        StrSQL = "SELECT id, name  From TblRegTimeDelgate Where 1=1  "
    End If

    If WorkState = True Then
        StrSQL = StrSQL + " and  workstate=1"
    End If

    If DriverOnly = True Then
        StrSQL = StrSQL + " and  not(driverid is null)"
    End If

    StrSQL = StrSQL + " Order By id ASC"
    GetComboData MyCombo, StrSQL
End Function
Public Function GetBordNo(MyCombo As DataCombo)
    Dim StrSQL As String
        StrSQL = "SELECT id, BoardNO From dbo.TblCarsData"
    GetComboData MyCombo, StrSQL
End Function

Public Function GetAssestsOfEmp(MyCombo As DataCombo, Optional EmpID As Integer)
    Dim StrSQL As String
If EmpID <> 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = " SELECT    dbo.TblAssestes.AsID, dbo.TblAssestes.AsName"
StrSQL = StrSQL & " FROM         dbo.TblAssestes INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpAsestDetails ON dbo.TblAssestes.AsID = dbo.TblEmpAsestDetails.AsID INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpAsest ON dbo.TblEmpAsestDetails.IDAseset = dbo.TblEmpAsest.EmpAsID"
StrSQL = StrSQL & " Where (dbo.TblEmpAsestDetails.FlagAs Is Null) And (dbo.TblEmpAsestDetails.EmpID =" & EmpID & ")"
StrSQL = StrSQL & "  AND  (dbo.TblAssestes.BranchID =0 or dbo.TblAssestes.BranchID  is null or         dbo.TblAssestes.BranchID  in(" & Current_branchSql & "))"
    End If
End If
    GetComboData MyCombo, StrSQL
End Function
Public Function GetAssestsOfEmpx2(MyCombo As DataCombo, Optional EmpID As Integer)
    Dim StrSQL As String
If EmpID <> 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = " SELECT    dbo.TblAssestes.AsID, dbo.TblAssestes.AsName"
StrSQL = StrSQL & " FROM         dbo.TblAssestes INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpAsestDetails ON dbo.TblAssestes.AsID = dbo.TblEmpAsestDetails.AsID INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpAsest ON dbo.TblEmpAsestDetails.IDAseset = dbo.TblEmpAsest.EmpAsID"
StrSQL = StrSQL & " Where (dbo.TblEmpAsestDetails.FlagAs Is Null) And (dbo.TblEmpAsestDetails.EmpID =" & EmpID & ")"
    
    End If
End If
    GetComboData MyCombo, StrSQL
End Function
Public Function GetTblTypeExtraExpenses(MyCombo As DataCombo, _
                                    Optional BolLoadAdmins As Boolean = True)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT id,name From TblTypeExtraExpeneses ORDER BY ID "
    Else
        StrSQL = "SELECT id,namee From TblTypeExtraExpeneses ORDER BY ID"
    End If
 
    GetComboData MyCombo, StrSQL
End Function
Public Function GetChequeBox(MyCombo As DataCombo)
    Dim StrSQL As String
        
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select BoxID,BoxName From tblBoxesData where ChequeBox=1"
                
        If SystemOptions.usertype <> UserAdminAll Then
            StrSQL = StrSQL & " and   BranchId=" & Current_branch
        End If

        StrSQL = StrSQL & "  Order By BoxName "

    Else
        StrSQL = "Select BoxID,BoxNamee From tblBoxesData  where ChequeBox=1"
            
        If SystemOptions.usertype <> UserAdminAll Then
            StrSQL = StrSQL & " and   BranchId=" & Current_branch
        End If

        StrSQL = StrSQL & "  Order By BoxNamee "
            
    End If

    GetComboData MyCombo, StrSQL
End Function
 Public Function GetBoxes(MyCombo As DataCombo, _
                         Optional bType As Integer = -1)
    Dim StrSQL As String
    If checkmanyBoxes(StrSQL) = False Then
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select BoxID,BoxName From tblBoxesData where 1=1   "
                
       
StrSQL = StrSQL & "  AND   (  BranchId=0  or      BranchId in(" & Current_branchSql & "))"
        If bType <> -1 Then
            StrSQL = StrSQL & " and   Type=" & bType
        End If
                
        StrSQL = StrSQL & "  Order By BoxName "

    Else
        StrSQL = "Select BoxID,BoxNamee From tblBoxesData where 1=1"
            
        If SystemOptions.usertype <> UserAdminAll Then
            'StrSQL = StrSQL & " and   BranchId=" & Current_branch
        End If
            StrSQL = StrSQL & "  AND (  BranchId=0  or         BranchId in(" & Current_branchSql & "))"
        If bType <> -1 Then
            StrSQL = StrSQL & " and   Type=" & bType
        End If
                
        StrSQL = StrSQL & "  Order By BoxNamee "
            
    End If
End If
    GetComboData MyCombo, StrSQL
End Function
Public Function GetExpensesParents(MyCombo As DataCombo, _
                                   Optional FIELD_no As Integer, _
                                   Optional branch_no As Integer, _
                                   Optional DEPARTEMENT As Integer)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = " SELECT     TOP 100 PERCENT Account_Code,Account_Name from dbo.ACCOUNTS WHERE     (last_account = 0) AND (Account_Code LIKE N'a3%') AND (Parent_Account_Code <> N'r') ORDER BY Parent_Account_Code, Account_Name"
    Else
        StrSQL = " SELECT     TOP 100 PERCENTAccount_Code,  Account_Nameeng  from dbo.ACCOUNTS WHERE     (last_account = 0) AND (Account_Code LIKE N'a3%') AND (Parent_Account_Code <> N'r') ORDER BY Parent_Account_Code, Account_Nameeng"
    End If

    GetComboData MyCombo, StrSQL
End Function

Public Function GetPrefix2(MyCombo As DataCombo, _
                           Optional Sanad_No As Integer, _
                           Optional branch_no As Integer, _
                           Optional DEPARTEMENT As Integer)
    Dim StrSQL As String
    StrSQL = "select id,Prefix from sanad_numbering  where  sanad_no=" & Sanad_No & "  and branch_no=" & Current_branch & " Order By Prefix"
    GetComboData MyCombo, StrSQL
End Function

Public Function GetPrefix(MyCombo As DataCombo, _
                          Optional FIELD_no As Integer, _
                          Optional branch_no As Integer, _
                          Optional DEPARTEMENT As Integer)
    Dim StrSQL As String
'    StrSQL = "select id,prifix from coding  where  FIELD_no=" & FIELD_no & "  and branch_no=" & branch_no & " Order By prifix"
StrSQL = "select id,prifix from coding  where  FIELD_no=" & FIELD_no & " Order By prifix"

    GetComboData MyCombo, StrSQL
End Function

Public Function GetFixedAssetsGroup(MyCombo As DataCombo)
    Dim StrSQL As String
 If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "Select GroupID,GroupName From FixedAssetsGroup where LastGroup=1   Order By GroupName"
  Else
  StrSQL = "Select GroupID,GroupNamee From FixedAssetsGroup where LastGroup=1   Order By GroupNamee"
  End If
    GetComboData MyCombo, StrSQL
End Function

Public Function GetShift(MyCombo As DataCombo, _
                         Optional BolLoadAdmins As Boolean = True)
    Dim StrSQL As String
    StrSQL = "SELECT SeftCode,SheftName From TbLSheft  Order By SheftName"
 
    GetComboData MyCombo, StrSQL
End Function

Public Function GetLine(MyCombo As DataCombo, _
                        Optional BolLoadAdmins As Boolean = True)
    Dim StrSQL As String
    StrSQL = "SELECT id,name From TblProductLine  Order By Name"
 
    GetComboData MyCombo, StrSQL
End Function

Public Function GetSalePriceNames(MyCombo As DataCombo, _
                                  Optional BolLoadAdmins As Boolean = True)
    Dim StrSQL As String
    StrSQL = "SELECT  id,PriceName From TblSalePriceNames"
 
    GetComboData MyCombo, StrSQL
End Function

Public Function GetTblyearsData(MyCombo As DataCombo, _
                                Optional BolLoadAdmins As Boolean = True)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT TblyearsDataid,Remarks From TblyearsData order by TblyearsDataid"
    Else
        StrSQL = "SELECT TblyearsDataid,Remarks From TblyearsData order by TblyearsDataid"
    End If
 
    GetComboData MyCombo, StrSQL
End Function

Public Function GetTblCarsDataGroup(MyCombo As DataCombo, _
                                    Optional BolLoadAdmins As Boolean = True, Optional TypeHajj As Boolean = False)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT id,name From TBLCarTypes where 1=1 "
    Else
        StrSQL = "SELECT id,namee From TBLCarTypes where 1=1   "
    End If
 
 If TypeHajj = True Then
  StrSQL = StrSQL & " and TypeHajj=1 "
 End If
 
     If SystemOptions.UserInterface = ArabicInterface Then
       StrSQL = StrSQL & "order by name "
     Else
     StrSQL = StrSQL & "order by namee "
     End If
     
    GetComboData MyCombo, StrSQL
End Function

Public Function GetCitiesDistance(MyCombo As DataCombo, _
                                  Stype As Integer)
    Dim StrSQL As String

    If Stype = 0 Then
        StrSQL = " SELECT     dbo.TBLCitiesDistance.CityFromId, dbo.TblCountriesGovernments.GovernmentName "
        StrSQL = StrSQL & " FROM         dbo.TblCountriesGovernments RIGHT OUTER JOIN"
        StrSQL = StrSQL & " dbo.TBLCitiesDistance ON dbo.TblCountriesGovernments.GovernmentID = dbo.TBLCitiesDistance.CityFromId"
        StrSQL = StrSQL & "  GROUP BY dbo.TblCountriesGovernments.GovernmentName, dbo.TBLCitiesDistance.CityFromId"

    Else
        StrSQL = "  SELECT      dbo.TBLCitiesDistance.CityToId, dbo.TblCountriesGovernments.GovernmentName"
        StrSQL = StrSQL & " FROM         dbo.TblCountriesGovernments RIGHT OUTER JOIN"
        StrSQL = StrSQL & " dbo.TBLCitiesDistance ON dbo.TblCountriesGovernments.GovernmentID = dbo.TBLCitiesDistance.CityToId"
        StrSQL = StrSQL & " GROUP BY dbo.TblCountriesGovernments.GovernmentName, dbo.TBLCitiesDistance.CityToId"
 
    End If
 
    StrSQL = StrSQL + "  ORDER BY dbo.TblCountriesGovernments.GovernmentName"
 
    GetComboData MyCombo, StrSQL
End Function
Public Function GetCarsMaintenanceTypes(MyCombo As DataCombo, _
                                        Optional BolLoadAdmins As Boolean = True, Optional typemaint As Integer = -1, Optional MainType As Integer = -1)
    Dim StrSQL As String
  If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT id,Name  "
    StrSQL = StrSQL + " FROM   TblMaintenanceType "
    If typemaint > -1 Then
    StrSQL = StrSQL + " where   typemaint= " & typemaint
    End If
    
  If MainType > -1 Then
    StrSQL = StrSQL + " where   MainType= " & MainType
    End If
      
      
    
    StrSQL = StrSQL + " Order BY  Name "
    Else
    StrSQL = "SELECT id,NameE  "
        
    StrSQL = StrSQL + " FROM   TblMaintenanceType "
If MainType > -1 Then
    StrSQL = StrSQL + " where   MainType= " & MainType
    End If
          
       If typemaint > -1 Then
    StrSQL = StrSQL + " where   typemaint= " & typemaint
    End If
       
       
    StrSQL = StrSQL + " Order BY  NameE "
    End If
    
                            
    GetComboData MyCombo, StrSQL
End Function

Public Function GetCarsMaintenanceTypesold(MyCombo As DataCombo, _
                                        Optional BolLoadAdmins As Boolean = True)
    Dim StrSQL As String
 
    StrSQL = "SELECT id,Name,km "
    StrSQL = StrSQL + " FROM   MaintenanceTypes "
    StrSQL = StrSQL + " Order BY  Name "
                            
    GetComboData MyCombo, StrSQL
End Function

'Public Function GetCars(MyCombo As DataCombo, _
'                        Optional BolLoadAdmins As Boolean = True)
'    Dim StrSQL As String
'    StrSQL = "SELECT  id,  BoardNO   From TblCarsData"
'
'    GetComboData MyCombo, StrSQL
'End Function
Public Function GetCars(MyCombo As DataCombo, _
                        Optional BolLoadAdmins As Boolean = True, Optional Typ As Integer = 0)
    Dim StrSQL As String
    StrSQL = "SELECT  id,  BoardNO   From TblCarsData"
    If Typ <> 0 Then
    StrSQL = StrSQL & " where CarsTypeId =" & Typ & ""
    End If
 
    GetComboData MyCombo, StrSQL
End Function

Public Function GetOutType(MyCombo As DataCombo, _
                         Optional BolLoadAdmins As Boolean = True)
    Dim StrSQL As String
    StrSQL = "SELECT  * From TblOutType"

 

    GetComboData MyCombo, StrSQL
End Function


Public Function GetJobEndService(MyCombo As DataCombo, _
                         Optional BolLoadAdmins As Boolean = True)
    Dim StrSQL As String
            If SystemOptions.UserInterface = ArabicInterface Then
                StrSQL = "SELECT     id,name"
            StrSQL = StrSQL & " From dbo.jopstatus"
            StrSQL = StrSQL & " WHERE     (resignationInt = 1 OR resignationInt = 2)"
                                 
            
'                If BolLoadAdmins = False Then
'                    StrSQL = "SELECT     id,namee"
'            StrSQL = StrSQL & " From dbo.jopstatus"
'            StrSQL = StrSQL & " WHERE     (resignationInt = 1 OR resignationInt = 2)"
'                End If
Else

    '        If SystemOptions.UserInterface = ArabicInterface Then
                StrSQL = "SELECT     id,namee"
            StrSQL = StrSQL & " From dbo.jopstatus"
            StrSQL = StrSQL & " WHERE     (resignationInt = 1 OR resignationInt = 2)"
                                 
            
'                If BolLoadAdmins = False Then
'                    StrSQL = "SELECT     id,namee"
'            StrSQL = StrSQL & " From dbo.jopstatus"
'            StrSQL = StrSQL & " WHERE     (resignationInt = 1 OR resignationInt = 2)"
'                End If
'
End If

    GetComboData MyCombo, StrSQL
End Function
Public Function GetUsers(MyCombo As DataCombo, _
                         Optional BolLoadAdmins As Boolean = True, Optional justmybranch As Boolean)
    Dim StrSQL As String
    StrSQL = "SELECT  * From TblUsers where 1=1"

    If BolLoadAdmins = False Then
        StrSQL = StrSQL + " and UserID <> 1"
    End If
    
                      If justmybranch <> 0 Then
     StrSQL = StrSQL & "  and      branchid in(" & Current_branchSql & ")"
        End If
        
StrSQL = StrSQL & " order by UserName"
    GetComboData MyCombo, StrSQL
End Function

Public Function GetManEmpDes(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT SupDecID, SupDecName From dbo.TblManSupDecs"
    Else
        StrSQL = "SELECT SupDecID, SupDecNamee From dbo.TblManSupDecs"
    End If

    StrSQL = StrSQL + " Where (DecType = 2) ORDER BY SupDecID"

    GetComboData MyCombo, StrSQL
End Function

Public Function GetExpensesType(MyCombo As DataCombo)
    Dim StrSQL As String
    StrSQL = "SELECT ExpensesType.ID,ExpensesType.Name From ExpensesType Order By ExpensesType.Name"
    GetComboData MyCombo, StrSQL

End Function

Public Function GetDocTypebyid(MyCombo As DataCombo, _
                               Optional ID As Integer, _
                               Optional branch_no As Integer)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT id,name From TblDoCumentsTypes"
    Else
        StrSQL = "SELECT id,namee From TblDoCumentsTypes"
    End If

    StrSQL = StrSQL & " where Doctype=" & ID

    If branch_no <> 0 Then
        StrSQL = StrSQL & " and ( branch_no is null or branch_no = " & branch_no & ")"
    End If

    GetComboData MyCombo, StrSQL
End Function
Public Function GetExpensespaidAdvanced(MyCombo As DataCombo, _
                                   Optional LastAccount As Boolean = False, _
                                   Optional BasicAccount As Boolean = False)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT id,name From TbExpensesprovided "
    Else
        StrSQL = "SELECT id,namee From TbExpensesprovided "
    End If
    If BasicAccount = True Then
        StrSQL = StrSQL + "  Where last_account=0  "
    End If
    If LastAccount = False Then
        'StrSQL = StrSQL + "  Where last_account=0  "
    Else
        StrSQL = StrSQL + "  Where last_account=1 "
    End If
    'StrSQL = StrSQL + " Where DecType=3"
    StrSQL = StrSQL + " Order By name "
    GetComboData MyCombo, StrSQL
End Function
 

Public Function GetmaintennceType(MyCombo As DataCombo)
    Dim StrSQL As String
 
 If SystemOptions.UserInterface = ArabicInterface Then
 StrSQL = " Select  id,Name  From TblMaintenanceType  order by Name"
 Else
    StrSQL = " Select  id,Namee  From TblMaintenanceType  order by Namee"
 End If
    GetComboData MyCombo, StrSQL
End Function


Public Function GetFixedAssets(MyCombo As DataCombo, Optional ISEQUP As Boolean = False, Optional Status_id As String = "", Optional checkPurchasePrice As Boolean = True)
    Dim StrSQL As String
 If checkPurchasePrice = True Then
    StrSQL = " Select  id,Name  From FixedAssets where HaveDepreciation=1   and PurchasePrice>0 "
 Else
     StrSQL = " Select  id,Name  From FixedAssets where HaveDepreciation=1   "
 End If
 If ISEQUP = True Then
 StrSQL = StrSQL & " and isequp=1"
 End If
 
 
   If SystemOptions.usertype <> UserAdminAll Then
        StrSQL = StrSQL & " AND   branch_no=" & Current_branch
    End If
If Status_id <> "" Then
StrSQL = StrSQL & " AND   Status_id IN(" & Status_id & ")"
End If
    
 StrSQL = StrSQL & "  order by Name"
    GetComboData MyCombo, StrSQL
End Function
 Public Function GetAssests(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT AsID,AsName From TblAssestes"
    Else
        StrSQL = "SELECT branch_id,branch_namee From TblAssestes"
    End If




    GetComboData MyCombo, StrSQL
End Function
 
Public Function GetDelegate(MyCombo As DataCombo, _
                             Optional WorkState As Boolean = False, _
                             Optional DriverOnly As Boolean = False)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
      '  StrSQL = "SELECT Emp_ID,Emp_Name From TblEmployee Where 1=1  "
   StrSQL = "    SELECT     dbo.TBLSalesRepData.EmpID, dbo.TblEmployee.Emp_Name"
StrSQL = StrSQL & " FROM         dbo.TBLSalesRepData LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TBLSalesRepData.EmpID = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & " where 1=1"
    Else
          StrSQL = "    SELECT     dbo.TBLSalesRepData.EmpID, dbo.TblEmployee.Emp_Namee"
StrSQL = StrSQL & " FROM         dbo.TBLSalesRepData LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TBLSalesRepData.EmpID = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & " where 1=1"
    End If

    If WorkState = True Then
        StrSQL = StrSQL + " and  workstate=1"
    End If

    If DriverOnly = True Then
        StrSQL = StrSQL + " and  not(driverid is null)"
    End If

    StrSQL = StrSQL + " Order By Emp_Name ASC"
    GetComboData MyCombo, StrSQL
End Function
Public Function GetSpeciaAsement(MyCombo As DataCombo, _
                             Optional WorkState As Boolean = False, _
                             Optional DriverOnly As Boolean = False)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT Id,name From TblSpeciaAsement Where 1=1  "
    Else
        StrSQL = "SELECT Id,namee From TblSpeciaAsement Where 1=1  "
    End If

    If WorkState = True Then
        StrSQL = StrSQL + " and  workstate=1"
    End If

    If DriverOnly = True Then
        StrSQL = StrSQL + " and  not(driverid is null)"
    End If

    StrSQL = StrSQL + " Order By name ASC"
    GetComboData MyCombo, StrSQL
End Function



Public Function GetTypeVisit(MyCombo As DataCombo, _
                             Optional WorkState As Boolean = False, _
                             Optional DriverOnly As Boolean = False)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT Id,name From TblTypeVisit Where 1=1  "
    Else
        StrSQL = "SELECT Id,namee From TblTypeVisit Where 1=1  "
    End If

    If WorkState = True Then
        StrSQL = StrSQL + " and  workstate=1"
    End If

    If DriverOnly = True Then
        StrSQL = StrSQL + " and  not(driverid is null)"
    End If

    StrSQL = StrSQL + " Order By name ASC"
    GetComboData MyCombo, StrSQL
End Function
 
 Function GetProjects(MyCombo As DataCombo)
  Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT id,Project_name From  projects "
        StrSQL = StrSQL & " where  not (Project_name is null)and Project_name<>N'""'"
    Else
        StrSQL = "SELECT id,Project_nameE From   projects "
        StrSQL = StrSQL & " where  not (Project_nameE is null)and Project_nameE<>N'""'"
    End If
    StrSQL = StrSQL & " and (Not (Fullcode Is Null))"
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = StrSQL & " order by  Project_name"
    Else
    StrSQL = StrSQL & " order by  Project_nameE"
    End If

    GetComboData MyCombo, StrSQL
    
End Function

 Function GetProjectsxx(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT id,Project_name From  projects order by  Project_name"
    Else
        StrSQL = "SELECT id,Project_name From   projects order by   Project_name"
    End If
    StrSQL = StrSQL & " Where(Not (fullcode Is Null))"

    GetComboData MyCombo, StrSQL
End Function

 Function GetProjectsxex(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT id,Project_name From  projects order by  Project_name"
    Else
        StrSQL = "SELECT id,Project_name From   projects order by   Project_name"
    End If

    GetComboData MyCombo, StrSQL
End Function
Public Function GetStoresByUser(MyCombo As DataCombo, Optional BranchID As Integer = 0, Optional UserID As Integer = 0)
    
    
    Dim StrSQL As String

    If UserID = 0 Then UserID = user_id
 
If checkmanyStores(StrSQL) = False Then
    
    
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT StoreID,StoreName From TblStore where 1=1"
 
    Else
        StrSQL = "SELECT StoreID,StoreNamee From TblStore where 1=1"

    End If
    StrSQL = StrSQL & "  AND      StoreId in(Select StoreId From TblUsersStores Where Userid = " & UserID & " )"
        If SystemOptions.usertype <> UserAdminAll Then
    '        StrSQL = StrSQL & " where   BranchId=" & Current_branch
        End If
      StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"
      
        
             If BranchID <> 0 Then
    '        StrSQL = StrSQL & " where   BranchId=" & BranchID
        End If
        
        
 

End If



    GetComboData MyCombo, StrSQL
End Function



 Function GetTypeImage(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT id,Name From  TblTypeImage order by  name"
    Else
        StrSQL = "SELECT id,namee From   TblTypeImage order by Namee   "
    End If

    GetComboData MyCombo, StrSQL
End Function


Public Function GetBranches(MyCombo As DataCombo, Optional justmybranch As Boolean)
 
     Dim StrSQL As String
     
If checkmanyBranches(StrSQL) = False Then
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT branch_id,branch_name From TblBranchesData"
    Else
        StrSQL = "SELECT branch_id,branch_namee From TblBranchesData"
    End If

        If justmybranch = True Then
     StrSQL = StrSQL & "  AND      branch_id in(" & Current_branchSql & ")"
        End If
      
      
        
        
End If



    GetComboData MyCombo, StrSQL
End Function
Public Function GetPOS(MyCombo As DataCombo, Optional justmybranch As Boolean)
 
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT BoxID, BoxName FROM Tblposdata"
    Else
        StrSQL = "SELECT BoxID, BoxNamee FROM Tblposdata"
    End If

    GetComboData MyCombo, StrSQL
End Function

Public Function GetAdvComponents(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = " SELECT   dbo.mofrdat.mofrad_code, dbo.mofrdat.mofrad_name FROM         dbo.mofrad INNER JOIN dbo.mofrdat ON dbo.mofrad.id = dbo.mofrdat.mofrad_type Where (dbo.MOFRAD.ADVView = 1) ORDER BY dbo.mofrdat.mofrad_name"
    Else
        StrSQL = " SELECT   dbo.mofrdat.mofrad_code, dbo.mofrdat.mofrad_namee FROM         dbo.mofrad INNER JOIN dbo.mofrdat ON dbo.mofrad.id = dbo.mofrdat.mofrad_type Where (dbo.MOFRAD.ADVView = 1) ORDER BY dbo.mofrdat.mofrad_namee"
    End If

    GetComboData MyCombo, StrSQL
End Function

Public Function GetYearlyComponents(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT mofrad_code,mofrad_name From mofrdat where Monthly=0  order by mofrad_name"
    Else
        StrSQL = "SELECT mofrad_code,mofrad_name From mofrdat  where Monthly=0  order by mofrad_namee"
    End If

    GetComboData MyCombo, StrSQL
End Function

Public Function GetLOTNO(MyCombo As DataCombo, _
                         Optional Item_ID As Long = 0, _
                         Optional StoreID As Long = 0)
    Dim StrSQL As String
 
    StrSQL = " SELECT DISTINCT dbo.Transaction_Details.LotNO AS id, dbo.Transaction_Details.LotNO"
    StrSQL = StrSQL & " FROM         dbo.Transaction_Details INNER JOIN"
    StrSQL = StrSQL & " dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
    StrSQL = StrSQL & " WHERE  1=1    "
        
    If Item_ID <> 0 Then
        StrSQL = StrSQL & "  and  (dbo.Transaction_Details.Item_ID = " & Item_ID & ") "
    End If

    If StoreID <> 0 Then
        StrSQL = StrSQL & "AND (dbo.Transactions.StoreID = " & StoreID & ")"
    End If

    GetComboData MyCombo, StrSQL
End Function

Public Function GetStores(MyCombo As DataCombo, Optional BranchID As Integer = 0, Optional IsPump As Integer = 0, Optional IsAll As Integer = 0)
    
    
    Dim StrSQL As String

    
 
If checkmanyStores(StrSQL) = False Then
    
    
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT StoreID,StoreName From TblStore where 1=1"
 
    Else
        StrSQL = "SELECT StoreID,StoreNamee From TblStore where 1=1"

    End If
    
        If IsPump Then
            StrSQL = StrSQL & "  AND      StoreID in(select tblPumpType.StoreID from tblPumpType )"
        Else
            If Not IsAll Then
                StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"
            Else
            End If
        End If
    
        
             If BranchID <> 0 Then
    '        StrSQL = StrSQL & " where   BranchId=" & BranchID
        End If
        
        
 

End If



    GetComboData MyCombo, StrSQL
End Function

Public Function GetIntervalsData(MyCombo As DataCombo, _
                                 Optional WorkState As Boolean = False)
    Dim StrSQL As String
    StrSQL = " SELECT     TOP 100 PERCENT dbo.TblAccountIntervals.AccountIntervalID, dbo.TblAccountIntervals.Comment"
    StrSQL = StrSQL & " FROM         dbo.TblyearsData INNER JOIN"
    StrSQL = StrSQL & " dbo.TblAccountIntervals ON dbo.TblyearsData.TblyearsDataid = dbo.TblAccountIntervals.TblyearsDataid"
    StrSQL = StrSQL & " WHERE   CurrentYear=1"

    'StrSQL = StrSQL & " WHERE     (dbo.TblAccountIntervals.TblyearsDataid ="
    ' StrSQL = StrSQL & " (SELECT     MAX(TblyearsDataid)"
    'StrSQL = StrSQL & " FROM         TblyearsData))"

    StrSQL = StrSQL & " ORDER BY dbo.TblAccountIntervals.AccountIntervalID"
    GetComboData MyCombo, StrSQL
End Function
Public Function GetSalesRepDatapurchase(MyCombo As DataCombo, _
                                Optional WorkState As Boolean = False)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT     dbo.TBLSalesRepData2.EmpID ,  dbo.TblEmployee.emp_name"
    Else
    StrSQL = "SELECT     dbo.TBLSalesRepDat2a.EmpID ,  dbo.TblEmployee.emp_nameE"
    End If
    StrSQL = StrSQL & " FROM         dbo.TblEmployee RIGHT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TBLSalesRepData2 ON dbo.TblEmployee.Emp_ID = dbo.TBLSalesRepData2.EmpID where 1=1"
 
 StrSQL = StrSQL & " and (  TBLSalesRepData2.BranchId=0   or   TBLSalesRepData2.BranchId is null or   TBLSalesRepData2.BranchId in (  " & Current_branchSql & "))"
 
    StrSQL = StrSQL & " Order By Emp_Name ASC"
    GetComboData MyCombo, StrSQL
End Function
Public Function GetProductionData(MyCombo As DataCombo, _
                                Optional WorkState As Boolean = False)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT      id,name FROM         dbo.tblProductionType "
      StrSQL = StrSQL & " Order By name ASC"
    Else
    StrSQL = "SELECT     id,namee FROM         dbo.tblProductionType  "
    StrSQL = StrSQL & " Order By namee ASC"
    End If
    
 
  
    GetComboData MyCombo, StrSQL
End Function
Public Function GetAlarm(MyCombo As DataCombo)
                            
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT id,name From TblAlarmType   "
    Else
        StrSQL = "SELECT id,namee From TblAlarmType   "
    End If

 

    StrSQL = StrSQL + " Order By id "
    GetComboData MyCombo, StrSQL
End Function
 
 Public Function GetSalesRepData(MyCombo As DataCombo, _
                                Optional WorkState As Boolean = False, Optional GroupID As Integer = 0, Optional BranchID As Integer = 0)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT     dbo.TBLSalesRepData.EmpID ,  dbo.TblEmployee.emp_name"
    Else
    StrSQL = "SELECT     dbo.TBLSalesRepData.EmpID ,  dbo.TblEmployee.emp_nameE"
    End If
    
    
    StrSQL = StrSQL & " FROM         dbo.TblEmployee RIGHT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TBLSalesRepData ON dbo.TblEmployee.Emp_ID = dbo.TBLSalesRepData.EmpID where 1=1"
    If GroupID <> 0 Then
      StrSQL = StrSQL & " and dbo.TBLSalesRepData.GroupID =" & GroupID & ""
    End If
    
 
'   If BranchID <> 0 Then
'            StrSQL = StrSQL & " where   TBLSalesRepData.BranchId=" & BranchID
'        End If
        
StrSQL = StrSQL & " and (  TBLSalesRepData.BranchId=0 or  TBLSalesRepData.BranchId is null or   TBLSalesRepData.BranchId in (  " & Current_branchSql & "))"

    StrSQL = StrSQL & " Order By Emp_Name ASC"
    
    GetComboData MyCombo, StrSQL
End Function


Public Function GetWorkShopemployee(MyCombo As DataCombo, _
                             Optional WorkState As Boolean = False, _
                             Optional GroupID As Integer = -1)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
   StrSQL = " SELECT     Emp_ID, Emp_Name FROM         dbo.TblEmployee WHERE     (WorkShop_Job = 2) "
    Else
 StrSQL = " SELECT     Emp_ID, Emp_Namee FROM         dbo.TblEmployee WHERE     (WorkShop_Job = 2) "
    End If

  If SystemOptions.UserInterface = ArabicInterface Then

    StrSQL = StrSQL + " Order By Emp_Name ASC"
    Else
    StrSQL = StrSQL + " Order By Emp_Namee ASC"
    End If
    
    GetComboData MyCombo, StrSQL
End Function
 

Public Function GetClientName(MyCombo As DataCombo, _
                             Optional WorkState As Boolean = False, _
                             Optional DriverOnly As Boolean = False)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ClientCode,ClientName From TblCardAuthorizationReform Where 1=1  "
    Else
        StrSQL = "SELECT ClientCode,ClientName From TblCardAuthorizationReform Where 1=1  "
    End If

    If WorkState = True Then
        StrSQL = StrSQL + " and  workstate=1"
    End If

    If DriverOnly = True Then
        StrSQL = StrSQL + " and  not(driverid is null)"
    End If

    StrSQL = StrSQL + " Order By ClientName ASC"
    GetComboData MyCombo, StrSQL
End Function


Public Function GetEmpDepartmentCar(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = " SELECT     DeparmentID, DepartmentName "
StrSQL = StrSQL & " From  TblEmpDepartments"
If Not SystemOptions.IsMaintItemMode Then
StrSQL = StrSQL & " WHERE     (Dpeterial IS NOT NULL)and (Dpeterial>=0) Order By DepartmentName"
End If
    Else
     StrSQL = " SELECT     DeparmentID, DepartmentNamee "
StrSQL = StrSQL & " From  TblEmpDepartments"
StrSQL = StrSQL & " WHERE     (Dpeterial IS NOT NULL)and (Dpeterial>=0) Order By DepartmentNamee"
      
    End If

    GetComboData MyCombo, StrSQL
End Function
Public Function GetFitte(MyCombo As DataCombo, _
                             Optional WorkState As Boolean = False, _
                             Optional DriverOnly As Boolean = False)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
     StrSQL = "   SELECT DISTINCT dbo.TblCardAuthorizationReformDetails.EmpID, dbo.TblEmployee.Emp_Name"
StrSQL = StrSQL & " FROM         dbo.TblCardAuthorizationReformDetails LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblEmployee ON dbo.TblCardAuthorizationReformDetails.EmpID = dbo.TblEmployee.Emp_ID"
  StrSQL = StrSQL & "  Where dbo.TblCardAuthorizationReformDetails.EmpID > 0  "
    Else
            StrSQL = "   SELECT DISTINCT dbo.TblCardAuthorizationReformDetails.EmpID, dbo.TblEmployee.Emp_Namee"
StrSQL = StrSQL & " FROM         dbo.TblCardAuthorizationReformDetails LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblEmployee ON dbo.TblCardAuthorizationReformDetails.EmpID = dbo.TblEmployee.Emp_ID"
   StrSQL = StrSQL & "  Where dbo.TblCardAuthorizationReformDetails.EmpID > 0  "
    End If

    If WorkState = True Then
        StrSQL = StrSQL + " and  workstate=1"
    End If

    If DriverOnly = True Then
        StrSQL = StrSQL + " and  not(driverid is null)"
    End If

    StrSQL = StrSQL + " Order By EmpID ASC"
    GetComboData MyCombo, StrSQL
End Function
Public Function GetSuperVisor(MyCombo As DataCombo, _
                             Optional WorkState As Boolean = False, _
                             Optional DriverOnly As Boolean = False)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
     StrSQL = " SELECT DISTINCT dbo.TblCardAuthorizationReformDetails.empsuper, dbo.TblEmployee.Emp_Name "
StrSQL = StrSQL & " FROM         dbo.TblCardAuthorizationReformDetails LEFT OUTER JOIN"
  StrSQL = StrSQL & " dbo.TblEmployee ON dbo.TblCardAuthorizationReformDetails.empsuper = dbo.TblEmployee.Emp_ID"
 StrSQL = StrSQL & " Where (dbo.TblCardAuthorizationReformDetails.empsuper > 0)"
    Else
             StrSQL = " SELECT DISTINCT dbo.TblCardAuthorizationReformDetails.empsuper, dbo.TblEmployee.Emp_Namee "
StrSQL = StrSQL & " FROM         dbo.TblCardAuthorizationReformDetails LEFT OUTER JOIN"
  StrSQL = StrSQL & " dbo.TblEmployee ON dbo.TblCardAuthorizationReformDetails.empsuper = dbo.TblEmployee.Emp_ID"
 StrSQL = StrSQL & " Where (dbo.TblCardAuthorizationReformDetails.empsuper > 0)"
    End If

    If WorkState = True Then
        StrSQL = StrSQL + " and  workstate=1"
    End If

    If DriverOnly = True Then
        StrSQL = StrSQL + " and  not(driverid is null)"
    End If

    StrSQL = StrSQL + " Order By empsuper ASC"
    GetComboData MyCombo, StrSQL
End Function

 Public Function GetYarsData(MyCombo As DataCombo)
    Dim StrSQL As String
  StrSQL = "   SELECT       TblyearsDataid, Remarks"
  StrSQL = StrSQL & "    From dbo.TblyearsData"
  StrSQL = StrSQL & " ORDER BY Remarks"
    GetComboData MyCombo, StrSQL
End Function
 Public Function GetEmployees(MyCombo As DataCombo, _
                             Optional WorkState As Boolean = False, _
                             Optional DriverOnly As Boolean = False, Optional BranchID As Integer = 0, Optional justmybranch As Boolean)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT Emp_ID,Emp_Name From TblEmployee Where 1=1  "
'        StrSQL = StrSQL + " Order By Emp_Name ASC"
    Else
        StrSQL = "SELECT Emp_ID,Emp_Namee From TblEmployee Where 1=1  "
'        StrSQL = StrSQL + " Order By Emp_Namee ASC"
    End If

    If WorkState = True Then
        StrSQL = StrSQL + " and  workstate=1"
    End If
    If DriverOnly = True Then
    
        If SystemOptions.UserInterface = ArabicInterface Then
            StrSQL = "SELECT     dbo.tblCarDrivers.EmpID, dbo.TblEmployee.Emp_Name"
            StrSQL = StrSQL & " FROM         dbo.tblCarDrivers LEFT OUTER JOIN"
            StrSQL = StrSQL & "  dbo.TblBranchesData ON dbo.tblCarDrivers.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
            StrSQL = StrSQL & "  dbo.TblEmployee ON dbo.tblCarDrivers.EmpID = dbo.TblEmployee.Emp_ID where 1=1  "
 
    Else
            StrSQL = "SELECT     dbo.tblCarDrivers.EmpID, dbo.TblEmployee.Emp_Namee"
            StrSQL = StrSQL & " FROM         dbo.tblCarDrivers LEFT OUTER JOIN"
            StrSQL = StrSQL & "  dbo.TblBranchesData ON dbo.tblCarDrivers.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
            StrSQL = StrSQL & "  dbo.TblEmployee ON dbo.tblCarDrivers.EmpID = dbo.TblEmployee.Emp_ID  where 1=1 "
    End If
    
        'StrSQL = StrSQL + " and  not(driverid is null)"
    End If
    
If SystemOptions.AllowShowAllEmployee = False Then
      
            StrSQL = StrSQL & " and (  TblEmployee.BranchId=0 or   TblEmployee.BranchId in (  " & Current_branchSql & "))"
            
  
   End If
   
     If justmybranch = True Then
     StrSQL = StrSQL & "  AND      TblEmployee.BranchId in(" & Current_branchSql & ")"
        End If
        
   If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = StrSQL + " Order By Emp_Name ASC"
   Else
      StrSQL = StrSQL + " Order By Emp_Namee  ASC"
   
   End If
   
    GetComboData MyCombo, StrSQL
End Function
 Public Function GetEmployeesx09102017(MyCombo As DataCombo, _
                             Optional WorkState As Boolean = False, _
                             Optional DriverOnly As Boolean = False, Optional BranchID As Integer = 0)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT Emp_ID,Emp_Name From TblEmployee Where 1=1  "
'        StrSQL = StrSQL + " Order By Emp_Name ASC"
    Else
        StrSQL = "SELECT Emp_ID,Emp_Namee From TblEmployee Where 1=1  "
'        StrSQL = StrSQL + " Order By Emp_Namee ASC"
    End If

    If WorkState = True Then
        StrSQL = StrSQL + " and  workstate=1"
    End If



        
        
    If DriverOnly = True Then
    
    
        If SystemOptions.UserInterface = ArabicInterface Then
            StrSQL = "SELECT     dbo.tblCarDrivers.EmpID, dbo.TblEmployee.Emp_Name"
            StrSQL = StrSQL & " FROM         dbo.tblCarDrivers LEFT OUTER JOIN"
            StrSQL = StrSQL & "  dbo.TblBranchesData ON dbo.tblCarDrivers.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
            StrSQL = StrSQL & "  dbo.TblEmployee ON dbo.tblCarDrivers.EmpID = dbo.TblEmployee.Emp_ID where 1=1  "
 
    Else
            StrSQL = "SELECT     dbo.tblCarDrivers.EmpID, dbo.TblEmployee.Emp_Namee"
            StrSQL = StrSQL & " FROM         dbo.tblCarDrivers LEFT OUTER JOIN"
            StrSQL = StrSQL & "  dbo.TblBranchesData ON dbo.tblCarDrivers.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
            StrSQL = StrSQL & "  dbo.TblEmployee ON dbo.tblCarDrivers.EmpID = dbo.TblEmployee.Emp_ID  where 1=1 "
         

    End If


    
        'StrSQL = StrSQL + " and  not(driverid is null)"
    End If



      If SystemOptions.usertype <> UserAdminAll Then
     '       StrSQL = StrSQL & " and (  TblEmployee.BranchId=0 or   TblEmployee.BranchId=" & Current_branch & ")"
            
            StrSQL = StrSQL & "  AND   (TblEmployee.BranchId =0 or     TblEmployee.BranchId in(" & Current_branchSql & "))"
        End If


             If BranchID <> 0 Then
            StrSQL = StrSQL & " and   BranchId=" & BranchID
        End If
    
   If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = StrSQL + " Order By Emp_Name ASC"
   Else
      StrSQL = StrSQL + " Order By Emp_Namee  ASC"
   
   End If
   
    GetComboData MyCombo, StrSQL
End Function
Public Function GetGroupSalesRepData(MyCombo As DataCombo, _
                                Optional WorkState As Boolean = False)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT     id ,  name"
    Else
    StrSQL = "SELECT     id ,  namee"
    End If
    StrSQL = StrSQL & " FROM  TBLSalesRepGroups "
  
 
    StrSQL = StrSQL & " Order By name ASC"
    GetComboData MyCombo, StrSQL
End Function
Public Function Getsheft(MyCombo As DataCombo, _
                         Optional WorkState As Boolean = False)
    Dim StrSQL As String
    StrSQL = "SELECT SeftCode,SheftName From TbLSheft "

    If WorkState = True Then
        StrSQL = StrSQL + " Where workstate=1"
    End If

    StrSQL = StrSQL + " Order By SheftName ASC"
    GetComboData MyCombo, StrSQL
End Function

 
Public Function GetItemsCodes(MyCombo As DataCombo, _
                              Optional IntItemType As Integer = -1, _
                              Optional IntAssbliedType As Integer = -1, _
                              Optional ItemMakingNew As Integer = -1, _
                              Optional StoreName As Integer = 0, Optional updateStatus As String, Optional ShortName As String, Optional StrItem As String)
    Dim StrSQL As String

    'IntItemType=-1 (ALL Items)
    'IntItemType=0 (Goods Items)
    'IntItemType=1 (Services Items)

    'IntAssbliedType=-1 (All Items)
    'IntAssbliedType=0 ( NOT Assblied Items)
    'IntAssbliedType=1 ( Assblied Items)

    StrSQL = "SELECT ItemID,barCodeNO From TblItems  "
    
    StrSQL = StrSQL + " Where ((IsArchive <> 1) or (IsArchive Is null))"


If SystemOptions.WorkWithLINKEDiActivity = True Then
StrSQL = GETLINKSQLByActivity(0, 1, updateStatus)

 
GoTo activity
End If

If StoreName <> 0 And StoreName <> -1 Then

StrSQL = GETLINKSQL(val(StoreName), 1, updateStatus)

 
GoTo ll
End If

 

    If IntItemType <> -1 Then
     '   StrSQL = StrSQL + " AND (ItemType=" & IntItemType & ")"
    End If
If StrItem <> "" Then
     StrSQL = StrSQL + StrItem
End If
    If ItemMakingNew <> -1 Then
        StrSQL = StrSQL + " AND (ItemMakingNew=" & ItemMakingNew & ")"
    End If

    If IntAssbliedType <> -1 Then
        If IntAssbliedType = 0 Then

            'AssbliedItem=true
            If SystemOptions.SysDataBaseType = AccessDataBase Then
                StrSQL = StrSQL + " AND (AssbliedItem=False OR AssbliedItem=0)"
            ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrSQL = StrSQL + " AND (AssbliedItem=0)"
            End If

        ElseIf IntAssbliedType = 1 Then

            If SystemOptions.SysDataBaseType = AccessDataBase Then
                StrSQL = StrSQL + " AND (AssbliedItem=True OR AssbliedItem=1)"
            ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrSQL = StrSQL + " AND (AssbliedItem=1)"
            End If
        
        End If
    End If
ll:
    
If ShortName <> "" Then
       StrSQL = StrSQL + " AND ShortName like'%" & ShortName & "%'"
End If
activity:
    StrSQL = StrSQL + " Order BY barCodeNO"
  '  If checkdataexist(StrSQL) = False Then
  '  StrSQL = "SELECT ItemID,barCodeNO From TblItems where ItemID=1"
'    End If
    
    GetComboData MyCombo, StrSQL
End Function



Public Function GetItemsNamesNew(MyCombo As Variant, _
                              Optional IntItemType As Integer = -1, _
                              Optional IntAssbliedType As Integer = -1, _
                              Optional ItemMakingNew As Integer = -1, _
                              Optional AllItems As Boolean = False, _
                              Optional StoreName As Integer = 0, _
                              Optional updateStatus As String, _
                              Optional ShortName As String, _
                              Optional StrItem As String)
    Dim StrSQL As String

    'IntItemType=-1 (ALL Items)
    'IntItemType=0 (Goods Items)
    'IntItemType=1 (Services Items)

    'IntAssbliedType=-1 (All Items)
    'IntAssbliedType=0 ( NOT Assblied Items)
    'IntAssbliedType=1 ( Assblied Items)
    If SystemOptions.UserInterface = ArabicInterface Then

        StrSQL = "SELECT ItemID,ItemName From  TblItems "
    Else
        StrSQL = "SELECT ItemID,ItemNamee From  TblItems "
    End If

    If SystemOptions.WorkWithLINKEDiActivity = True Then
        StrSQL = GETLINKSQLByActivity(0, 2, "xxx")
 
        GoTo activity
    End If

    If StoreName <> 0 And StoreName <> -1 Then
        StrSQL = GETLINKSQL(val(StoreName), 2, updateStatus)
 
        GoTo ll
    End If

    If AllItems = False Then
       ' StrSQL = StrSQL + " Where (IsNull(IsArchive,0) <> 1 and IsNull(IsArchive ,0) = 0)"
       StrSQL = StrSQL + " Where ( IsArchive IS NULL OR IsArchive = 0 )"
    End If

    If IntItemType <> -1 Then
        '        StrSQL = StrSQL + " AND (ItemType=" & IntItemType & ")"
    End If
    If StrItem <> "" Then
        StrSQL = StrSQL + StrItem
    End If
    If ItemMakingNew <> -1 Then
        StrSQL = StrSQL + " AND (ItemMakingNew=" & ItemMakingNew & ")"
    End If

    If IntAssbliedType <> -1 Then
        If IntAssbliedType = 0 Then

            'AssbliedItem=true
            If SystemOptions.SysDataBaseType = AccessDataBase Then
             '   StrSQL = StrSQL + " AND (AssbliedItem=False OR AssbliedItem=0)"
             StrSQL = StrSQL + " AND (AssbliedItem OR AssbliedItem=0)"
            ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrSQL = StrSQL + " AND (AssbliedItem=0)"
            End If

        ElseIf IntAssbliedType = 1 Then

            If SystemOptions.SysDataBaseType = AccessDataBase Then
               ' StrSQL = StrSQL + " AND (AssbliedItem=True OR AssbliedItem=1)"
               StrSQL = StrSQL + " AND (  AssbliedItem=1)"
            ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrSQL = StrSQL + " AND (AssbliedItem=1)"
            End If
        
        End If
    End If
ll:

    If ShortName <> "" Then
        StrSQL = StrSQL + " AND ShortName like'%" & ShortName & "%'"
    End If
activity:
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = StrSQL + " Order BY ItemName "
    Else
        StrSQL = StrSQL + " Order BY ItemNamee "
    End If
    
    GetComboDataNew MyCombo, StrSQL
End Function

Private Sub GetComboDataNew(My_Combo As Variant, _
                            My_SQL As String, _
                            Optional ByVal IsOtherData As Boolean = False)
    Dim rs      As ADODB.Recordset
    Dim StrTemp As String
    Dim Msg     As String
    On Error GoTo ErrorHandler

    If InStr(1, My_SQL, "SELECT", vbTextCompare) = 0 Then
        Exit Sub
    End If
    
    Set rs = New ADODB.Recordset

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        rs.CursorLocation = adUseClient
    End If
    If IsOtherData Then
        rs.Open My_SQL, POSConnection, adOpenStatic, adLockReadOnly, adCmdText
    Else
        rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    End If

    'Populate the ADO datacombo by setting its properties
    If IsArray(My_Combo) Then
        Dim i As Integer
        For i = 0 To UBound(My_Combo)
            My_Combo(i).Tag = My_SQL
    
            StrTemp = My_Combo(i).BoundText
            Set My_Combo(i).RowSource = rs
            My_Combo(i).BoundColumn = rs(0).Name
            My_Combo(i).ListField = rs(1).Name

            If Trim(StrTemp) <> "" Then
                My_Combo(i).BoundText = StrTemp
            Else
                My_Combo(i).BoundText = ""
                My_Combo(i).text = ""
            End If
    
        Next
    Else
        My_Combo.Tag = My_SQL
        StrTemp = My_Combo.BoundText
        Set My_Combo.RowSource = rs
        My_Combo.BoundColumn = rs(0).Name
        My_Combo.ListField = rs(1).Name

        If Trim(StrTemp) <> "" Then
            My_Combo.BoundText = StrTemp
        Else
            My_Combo.BoundText = ""
            My_Combo.text = ""
        End If
    End If

Exit_Sub:
    Set rs = Nothing
    Exit Sub
ErrorHandler:
    Msg = Now
    Msg = Msg & CHR(13) & "ClsDataCombos:GetComboData"
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & Err.Description
    WriteInLogFile Msg
    'MsgBox "ERROR! Err# " & Err.Number & " Desc: " & Err.Description, vbCritical + vbOKOnly
    Resume Exit_Sub
End Sub


Public Function GetItemsNames(MyCombo As DataCombo, _
                              Optional IntItemType As Integer = -1, _
                              Optional IntAssbliedType As Integer = -1, _
                              Optional ItemMakingNew As Integer = -1, _
                              Optional AllItems As Boolean = False, _
                              Optional StoreName As Integer = 0, _
                              Optional updateStatus As String, _
                              Optional ShortName As String, _
                              Optional StrItem As String)
    Dim StrSQL As String

    'IntItemType=-1 (ALL Items)
    'IntItemType=0 (Goods Items)
    'IntItemType=1 (Services Items)

    'IntAssbliedType=-1 (All Items)
    'IntAssbliedType=0 ( NOT Assblied Items)
    'IntAssbliedType=1 ( Assblied Items)
    If SystemOptions.UserInterface = ArabicInterface Then

        StrSQL = "SELECT ItemID,ItemName From  TblItems "
    Else
        StrSQL = "SELECT ItemID,ItemNamee From  TblItems "
    End If

    If SystemOptions.WorkWithLINKEDiActivity = True Then
        StrSQL = GETLINKSQLByActivity(0, 2, "xxx")
 
        GoTo activity
    End If

    If StoreName <> 0 And StoreName <> -1 Then
        StrSQL = GETLINKSQL(val(StoreName), 2, updateStatus)
 
        GoTo ll
    End If

    If AllItems = False Then
       ' StrSQL = StrSQL + " Where (IsNull(IsArchive,0) <> 1 and IsNull(IsArchive ,0) = 0)"
       StrSQL = StrSQL + " Where ( IsArchive IS NULL OR IsArchive = 0 )"
    End If

    If IntItemType <> -1 Then
        '        StrSQL = StrSQL + " AND (ItemType=" & IntItemType & ")"
    End If
    If StrItem <> "" Then
        StrSQL = StrSQL + StrItem
    End If
    If ItemMakingNew <> -1 Then
        StrSQL = StrSQL + " AND (ItemMakingNew=" & ItemMakingNew & ")"
    End If

    If IntAssbliedType <> -1 Then
        If IntAssbliedType = 0 Then

            'AssbliedItem=true
            If SystemOptions.SysDataBaseType = AccessDataBase Then
             '   StrSQL = StrSQL + " AND (AssbliedItem=False OR AssbliedItem=0)"
             StrSQL = StrSQL + " AND (AssbliedItem OR AssbliedItem=0)"
            ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrSQL = StrSQL + " AND (AssbliedItem=0)"
            End If

        ElseIf IntAssbliedType = 1 Then

            If SystemOptions.SysDataBaseType = AccessDataBase Then
               ' StrSQL = StrSQL + " AND (AssbliedItem=True OR AssbliedItem=1)"
               StrSQL = StrSQL + " AND (  AssbliedItem=1)"
            ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrSQL = StrSQL + " AND (AssbliedItem=1)"
            End If
        
        End If
    End If
ll:

    If ShortName <> "" Then
        StrSQL = StrSQL + " AND ShortName like'%" & ShortName & "%'"
    End If
activity:
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = StrSQL + " Order BY ItemName "
    Else
        StrSQL = StrSQL + " Order BY ItemNamee "
    End If
    GetComboData MyCombo, StrSQL
End Function

 Public Function GetTemplates(MyCombo As DataCombo, _
                                  Optional BolGetRoot As Boolean = True)
    Dim StrSQL As String

    If BolGetRoot = True Then
        StrSQL = "SELECT id, notes From TblEquipments1 Order By name"
    Else
        'StrSQL = "SELECT GroupID,GroupName From Groups Where (GroupID <>1) Order By GroupName "
    End If

    GetComboData MyCombo, StrSQL
End Function

Public Function GetExpensesGroups(MyCombo As DataCombo, _
                                  Optional BolGetRoot As Boolean = True)
    Dim StrSQL As String

    If BolGetRoot = True Then
        StrSQL = "SELECT distinct parent_account,parent_account From ExpensesType Order By parent_account"
    Else
        'StrSQL = "SELECT GroupID,GroupName From Groups Where (GroupID <>1) Order By GroupName "
    End If

    GetComboData MyCombo, StrSQL
End Function

Public Function GetCCTypes(MyCombo As DataCombo, _
                           Optional BolGetRoot As Boolean = True)
    Dim StrSQL As String

    If BolGetRoot = True Then
        StrSQL = "SELECT id,name From Marakes_taklefa_type Order By name"
    Else
        StrSQL = "SELECT  id,name From Marakes_taklefa_type  Where ( ID <>1) Order By name "
    End If

    GetComboData MyCombo, StrSQL
End Function

Public Function GetItemSGroups(MyCombo As DataCombo, _
                               Optional BolGetRoot As Boolean = True)
    Dim StrSQL As String

If SystemOptions.UserInterface = ArabicInterface Then
          If BolGetRoot = True Then
              StrSQL = "SELECT GroupID,GroupName From Groups Order By GroupName"
          Else
              StrSQL = "SELECT GroupID,GroupName From Groups Where (GroupID <>1) Order By GroupID "
          End If
    Else
  
            If BolGetRoot = True Then
              StrSQL = "SELECT GroupID,GroupNamee From Groups Order By GroupNamee"
          Else
              StrSQL = "SELECT GroupID,GroupNamee From Groups Where (GroupID <>1) Order By GroupID "
          End If
         


  End If

    GetComboData MyCombo, StrSQL

End Function

Public Function GetFinancialEquations(MyCombo As DataCombo)
    Dim StrSQL As String
    StrSQL = "Select FinancialEquationsId,FinancialEquationsName From FinancialEquations "
    'StrSQL = "Select Account_Code,BankName From BanksData "
    StrSQL = StrSQL + " Order BY FinancialEquationsName"
    GetComboData MyCombo, StrSQL
End Function

Public Function GetBalanceheeet(MyCombo As DataCombo)
    Dim StrSQL As String
    StrSQL = "Select BalanceSheetId,BalanceSheetName From BalanceSheetView "
    'StrSQL = "Select Account_Code,BankName From BanksData "
    StrSQL = StrSQL + " Order BY BalanceSheetName"
    GetComboData MyCombo, StrSQL
End Function
 
  Public Function GetPaymentType(MyCombo As DataCombo, Optional TypTran As Integer = 0)
    Dim My_SQL As String
   ' StrSQL = "Select PaymentID,PaymentName From TblPaymentType "

   ' StrSQL = StrSQL + " Order BY PaymentID"
   My_SQL = " Select PaymentID,PaymentName From TblPaymentType"
   My_SQL = My_SQL & " where 1=1"
   If TypTran <> -1 Then
My_SQL = My_SQL & " and (TypTran=2 or TypTran is null)  "
End If
If SystemOptions.LinkUsersWithPayment = True Then
My_SQL = My_SQL & " and PaymentID in (SELECT     PaynetID"
My_SQL = My_SQL & " From dbo.TblPaymentUser"
My_SQL = My_SQL & " Where (UserID = " & user_id & "))"
End If
My_SQL = My_SQL & " order by PaymentID"
    
    GetComboData MyCombo, My_SQL
End Function

Public Function GetCostCenter(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
 
           StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
           
        If SystemOptions.usertype <> UserAdminAll Then
            'StrSQL = StrSQL & " and  Branch_NO=0 or  Branch_NO=" & Current_branch
            
            StrSQL = StrSQL & "  AND      branch_no in(" & Current_branchSql & ")"
        End If
        
        StrSQL = StrSQL + " Order BY account_name"
    Else
         
        StrSQL = "  SELECT code ,EnglishName FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
        If SystemOptions.usertype <> UserAdminAll Then
            'StrSQL = StrSQL & " and  Branch_NO=0 or   Branch_NO=" & Current_branch
            StrSQL = StrSQL & "  AND      branch_no in(" & Current_branchSql & ")"
        End If
            
        StrSQL = StrSQL + " Order BY EnglishName"

    End If

    GetComboData MyCombo, StrSQL
End Function

Public Function GetBanks(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select BankID,BankName From BanksData where 1=1 "
         
'        If SystemOptions.usertype <> UserAdminAll Then
'            StrSQL = StrSQL & " where  BranchId=0 or  BranchId=" & Current_branch
'        End If
StrSQL = StrSQL & "  AND  ( BranchId =0 or      BranchId in(" & Current_branchSql & "))"
        StrSQL = StrSQL + " Order BY BankName"
    Else
        StrSQL = "Select BankID,BankNameE From BanksData where 1=1  "
             
'        If SystemOptions.usertype <> UserAdminAll Then
     '       StrSQL = StrSQL & " where  BranchId=0 or   BranchId=" & Current_branch
'        End If
     StrSQL = StrSQL & "  AND   (BranchId =0 or     BranchId in(" & Current_branchSql & "))"
        StrSQL = StrSQL + " Order BY BankNameE"

    End If

    GetComboData MyCombo, StrSQL
End Function
 
  Public Function GetCustomersSuppliers(IntType As Integer, _
                                      MyCombo As DataCombo, _
                                      Optional GetCashDealer As Boolean = True, Optional BranchID As Integer = 0, Optional HideLocked As Integer = 0, Optional CHkMot3ahed As Integer = 0, Optional ByVal IsOtherData As Boolean = False)
    Dim StrSQL As String

    'IntType=0 Customers And Suppliers (ALL)
    'IntType=1 Customers Only
    'IntType=2 Suppliers Only
    If SystemOptions.UserInterface = ArabicInterface Then
        If IntType = 0 Then
            StrSQL = "Select CusID,CusName From TblCustemers"
            StrSQL = StrSQL + " Where 1=1"
         '   StrSQL = StrSQL + " Where (Type=1 OR Type=2)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " AND (CusID <> 1 And CusID <>2)"
            End If


 

        ElseIf IntType = 1 Then
            StrSQL = "Select CusID,CusName From TblCustemers Where ( Type=1  or CustomerandVendor=1 )"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 2)"
            End If
      ElseIf IntType = 156 Then
            StrSQL = "Select CusID,CusName From TblCustemers Where ( Type=1 or Type=56 or CustomerandVendor=1 )"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 2)"
            End If
               ElseIf IntType = 257 Then
            StrSQL = "Select CusID,CusName From TblCustemers Where (Type=2 or Type=57 or CustomerandVendor=1)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If
            
        ElseIf IntType = 2 Then
            StrSQL = "Select CusID,CusName From TblCustemers Where (Type=2 or CustomerandVendor=1)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If
            ElseIf IntType = 257 Then
            StrSQL = "Select CusID,CusName From TblCustemers Where (Type=2 or Type=57 or CustomerandVendor=1)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If
        ElseIf IntType = 3 Then
            StrSQL = "Select CusID,CusName From TblCustemers Where (Type=2 or Type=3 or CustomerandVendor=1)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If
        ElseIf IntType = 33 Then
            StrSQL = "Select CusID,CusName From TblCustemers Where ( Type=3 )"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If
        
        
         ElseIf IntType = 55 Then
            StrSQL = "Select CusID,CusName From TblCustemers   Where Type=55 or Type=1 "
           ElseIf IntType = 56 Then
            StrSQL = "Select CusID,CusName From TblCustemers   Where Type=56 "
               ElseIf IntType = 57 Then
            StrSQL = "Select CusID,CusName From TblCustemers   Where Type=57 "
            
       ElseIf IntType = 4 Then
            StrSQL = "Select CusID,CusName From TblCustemers Where ( ( Type=2 or Type=3)  or CustomerandVendor=1)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If
            
        End If

        If SystemOptions.usertype <> UserAdminAll Then
       '     StrSQL = StrSQL & " and   ( BranchId=0 or BranchId=" & Current_branch & ")  "
                       StrSQL = StrSQL & " and ( BranchId=0  or      BranchId in(" & Current_branchSql & "))"

        End If
        
             If BranchID <> 0 Then
         '   StrSQL = StrSQL & " and   BranchId=" & BranchID
                StrSQL = StrSQL & " and ( BranchId=0  or      BranchId in(" & Current_branchSql & "))"

          

        End If
        If HideLocked = 1 Then
        StrSQL = StrSQL & " and ( locked=0  or locked is null)"
        End If
        
    If SystemOptions.FixedCustomer = 1 Then
            StrSQL = StrSQL & " and  Empid = " & GetEmployeeIDFROMUserID(user_id)
                 StrSQL = StrSQL + " or  (CusID = 2)"
                 
        End If
        
        StrSQL = StrSQL + " Order BY CusName"
        GetComboData MyCombo, StrSQL, IsOtherData
    Else

        If IntType = 0 Then
            StrSQL = "Select CusID,CusNamee From TblCustemers"
            StrSQL = StrSQL + " Where (Type=1 OR Type=2 OR Type=3)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " AND (CusID <> 1 And CusID <>2)"
            End If

        ElseIf IntType = 1 Then
            StrSQL = "Select CusID,CusNamee From TblCustemers Where Type=1 or CustomerandVendor=1"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 2)"
            End If

        ElseIf IntType = 2 Then
            StrSQL = "Select CusID,CusNamee From TblCustemers Where Type=2 or CustomerandVendor=1"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If

        ElseIf IntType = 3 Then
            StrSQL = "Select CusID,CusNamee From TblCustemers   Where Type=2 or CustomerandVendor=1"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If
        ElseIf IntType = 55 Then
            StrSQL = "Select CusID,CusNamee From TblCustemers   Where Type=55 "
        ElseIf IntType = 4 Then
            StrSQL = "Select CusID,CusNamee From TblCustemers Where ( ( Type=2 or Type=3)  or CustomerandVendor=1)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If
            
        End If

         If SystemOptions.usertype <> UserAdminAll Then
            StrSQL = StrSQL & " and (  BranchId=0 or   BranchId=" & Current_branch & ")"
        End If
        
        
          If CHkMot3ahed = 1 Then
          StrSQL = StrSQL & " and   CHkMot3ahed=1"
        End If
        
            If SystemOptions.FixedCustomer = 1 Then
            StrSQL = StrSQL & " and  Empid = " & GetEmployeeIDFROMUserID(user_id)
              StrSQL = StrSQL + " or  (CusID = 2)"
        End If
        StrSQL = StrSQL + " Order BY CusNamee"
        GetComboData MyCombo, StrSQL, IsOtherData

    End If

End Function

   Public Function GetCustomersSuppliersx08052017(IntType As Integer, _
                                      MyCombo As DataCombo, _
                                      Optional GetCashDealer As Boolean = True, Optional BranchID As Integer = 0)
    Dim StrSQL As String

    'IntType=0 Customers And Suppliers (ALL)
    'IntType=1 Customers Only
    'IntType=2 Suppliers Only
    If SystemOptions.UserInterface = ArabicInterface Then
        If IntType = 0 Then
            StrSQL = "Select CusID,CusName From TblCustemers"
            StrSQL = StrSQL + " Where (Type=1 OR Type=2)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " AND (CusID <> 1 And CusID <>2)"
            End If


 

        ElseIf IntType = 1 Then
            StrSQL = "Select CusID,CusName From TblCustemers Where ( Type=1  or CustomerandVendor=1 )"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 2)"
            End If

        ElseIf IntType = 2 Then
            StrSQL = "Select CusID,CusName From TblCustemers Where (Type=2 or CustomerandVendor=1)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If
    
        ElseIf IntType = 3 Then
            StrSQL = "Select CusID,CusName From TblCustemers Where (Type=2 or CustomerandVendor=1)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If
        
        
         ElseIf IntType = 55 Then
            StrSQL = "Select CusID,CusName From TblCustemers   Where Type=55 or Type=1 "
        
       ElseIf IntType = 4 Then
            StrSQL = "Select CusID,CusName From TblCustemers Where ( ( Type=2 or Type=3)  or CustomerandVendor=1)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If
            
        End If

        If SystemOptions.usertype <> UserAdminAll Then
       '     StrSQL = StrSQL & " and   ( BranchId=0 or BranchId=" & Current_branch & ")  "
                       StrSQL = StrSQL & " and ( BranchId=0  or      BranchId in(" & Current_branchSql & "))"

        End If
             If BranchID <> 0 Then
         '   StrSQL = StrSQL & " and   BranchId=" & BranchID
                StrSQL = StrSQL & " and ( BranchId=0  or      BranchId in(" & Current_branchSql & "))"


        End If
        
        StrSQL = StrSQL + " Order BY CusName"
        GetComboData MyCombo, StrSQL
    Else

        If IntType = 0 Then
            StrSQL = "Select CusID,CusNamee From TblCustemers"
            StrSQL = StrSQL + " Where (Type=1 OR Type=2)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " AND (CusID <> 1 And CusID <>2)"
            End If

        ElseIf IntType = 1 Then
            StrSQL = "Select CusID,CusNamee From TblCustemers Where Type=1 or CustomerandVendor=1"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 2)"
            End If

        ElseIf IntType = 2 Then
            StrSQL = "Select CusID,CusNamee From TblCustemers Where Type=2 or CustomerandVendor=1"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If

        ElseIf IntType = 3 Then
            StrSQL = "Select CusID,CusNamee From TblCustemers   Where Type=2 or CustomerandVendor=1"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If
        ElseIf IntType = 55 Then
            StrSQL = "Select CusID,CusNamee From TblCustemers   Where Type=55 "
        ElseIf IntType = 4 Then
            StrSQL = "Select CusID,CusNamee From TblCustemers Where ( ( Type=2 or Type=3)  or CustomerandVendor=1)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If
            
        End If

         If SystemOptions.usertype <> UserAdminAll Then
            StrSQL = StrSQL & " and (  BranchId=0 or   BranchId=" & Current_branch & ")"
        End If
        
        
          If BranchID <> 0 Then
'            StrSQL = StrSQL & " and   BranchId=" & BranchID
        End If
        
        
        StrSQL = StrSQL + " Order BY CusNamee"
        GetComboData MyCombo, StrSQL

    End If

End Function
Public Function GetCustomersSuppliersxxxx(IntType As Integer, _
                                      MyCombo As DataCombo, _
                                      Optional GetCashDealer As Boolean = True, Optional BranchID As Integer = 0)
    Dim StrSQL As String

    'IntType=0 Customers And Suppliers (ALL)
    'IntType=1 Customers Only
    'IntType=2 Suppliers Only
    If SystemOptions.UserInterface = ArabicInterface Then
        If IntType = 0 Then
            StrSQL = "Select CusID,CusName From TblCustemers"
            StrSQL = StrSQL + " Where (Type=1 OR Type=2)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " AND (CusID <> 1 And CusID <>2)"
            End If


 

        ElseIf IntType = 1 Then
            StrSQL = "Select CusID,CusName From TblCustemers Where ( Type=1  or CustomerandVendor=1 )"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 2)"
            End If

        ElseIf IntType = 2 Then
            StrSQL = "Select CusID,CusName From TblCustemers Where (Type=2 or CustomerandVendor=1)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If
    
        ElseIf IntType = 3 Then
            StrSQL = "Select CusID,CusName From TblCustemers Where (Type=2 or CustomerandVendor=1)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If
        
        
         ElseIf IntType = 55 Then
            StrSQL = "Select CusID,CusName From TblCustemers   Where Type=55 or CustomerandVendor=1"
        
       ElseIf IntType = 4 Then
            StrSQL = "Select CusID,CusName From TblCustemers Where ( ( Type=2 or Type=3)  or CustomerandVendor=1)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If
            
        End If

        If SystemOptions.usertype <> UserAdminAll Then
            StrSQL = StrSQL & " and   ( BranchId=0 or BranchId=" & Current_branch & ")  "
        End If
             If BranchID <> 0 Then
            StrSQL = StrSQL & " and   BranchId=" & BranchID
        End If
        
        StrSQL = StrSQL + " Order BY CusName"
        GetComboData MyCombo, StrSQL
    Else

        If IntType = 0 Then
            StrSQL = "Select CusID,CusNamee From TblCustemers"
            StrSQL = StrSQL + " Where (Type=1 OR Type=2)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " AND (CusID <> 1 And CusID <>2)"
            End If

        ElseIf IntType = 1 Then
            StrSQL = "Select CusID,CusNamee From TblCustemers Where Type=1 or CustomerandVendor=1"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 2)"
            End If

        ElseIf IntType = 2 Then
            StrSQL = "Select CusID,CusNamee From TblCustemers Where Type=2 or CustomerandVendor=1"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If

        ElseIf IntType = 3 Then
            StrSQL = "Select CusID,CusNamee From TblCustemers   Where Type=2 or CustomerandVendor=1"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If
        ElseIf IntType = 55 Then
            StrSQL = "Select CusID,CusNamee From TblCustemers   Where Type=55 or CustomerandVendor=1"
        ElseIf IntType = 4 Then
            StrSQL = "Select CusID,CusNamee From TblCustemers Where ( ( Type=2 or Type=3)  or CustomerandVendor=1)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If
            
        End If
 
 
      
            
            
         If SystemOptions.usertype <> UserAdminAll Then
            StrSQL = StrSQL & " and (  BranchId=0 or   BranchId=" & Current_branch & ")"
        End If
        
        
          If BranchID <> 0 Then
            StrSQL = StrSQL & " and   BranchId=" & BranchID
        End If
        
        
        StrSQL = StrSQL + " Order BY CusNamee"
        GetComboData MyCombo, StrSQL

    End If

End Function
Public Function GetCustomersSuppliersold(IntType As Integer, _
                                      MyCombo As DataCombo, _
                                      Optional GetCashDealer As Boolean = True, Optional BranchID As Integer = 0)
    Dim StrSQL As String

    'IntType=0 Customers And Suppliers (ALL)
    'IntType=1 Customers Only
    'IntType=2 Suppliers Only
    If SystemOptions.UserInterface = ArabicInterface Then
        If IntType = 0 Then
            StrSQL = "Select CusID,CusName From TblCustemers"
            StrSQL = StrSQL + " Where (Type=1 OR Type=2)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " AND (CusID <> 1 And CusID <>2)"
            End If


 

        ElseIf IntType = 1 Then
            StrSQL = "Select CusID,CusName From TblCustemers Where ( Type=1  or CustomerandVendor=1 )"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 2)"
            End If

        ElseIf IntType = 2 Then
            StrSQL = "Select CusID,CusName From TblCustemers Where (Type=2 or CustomerandVendor=1)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If
    
        ElseIf IntType = 3 Then
            StrSQL = "Select CusID,CusName From TblCustemers Where (Type=2 or CustomerandVendor=1)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If
        
        
        
        
       ElseIf IntType = 4 Then
            StrSQL = "Select CusID,CusName From TblCustemers Where ( ( Type=2 or Type=3)  or CustomerandVendor=1)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If
            
        End If

        If SystemOptions.usertype <> UserAdminAll Then
            StrSQL = StrSQL & " and   ( BranchId=0 or BranchId=" & Current_branch & ")  "
        End If
             If BranchID <> 0 Then
            StrSQL = StrSQL & " and   BranchId=" & BranchID
        End If
        
        StrSQL = StrSQL + " Order BY CusName"
        GetComboData MyCombo, StrSQL
    Else

        If IntType = 0 Then
            StrSQL = "Select CusID,CusNamee From TblCustemers"
            StrSQL = StrSQL + " Where (Type=1 OR Type=2)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " AND (CusID <> 1 And CusID <>2)"
            End If

        ElseIf IntType = 1 Then
            StrSQL = "Select CusID,CusNamee From TblCustemers Where Type=1 or CustomerandVendor=1"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 2)"
            End If

        ElseIf IntType = 2 Then
            StrSQL = "Select CusID,CusNamee From TblCustemers Where Type=2 or CustomerandVendor=1"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If

        ElseIf IntType = 3 Then
            StrSQL = "Select CusID,CusNamee From TblCustemers   Where Type=2 or CustomerandVendor=1"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If
        
        ElseIf IntType = 4 Then
            StrSQL = "Select CusID,CusNamee From TblCustemers Where ( ( Type=2 or Type=3)  or CustomerandVendor=1)"

            If GetCashDealer = False Then
                StrSQL = StrSQL + " And (CusID <> 1)"
            End If
            
        End If
 
 
      
            
            
         If SystemOptions.usertype <> UserAdminAll Then
            StrSQL = StrSQL & " and (  BranchId=0 or   BranchId=" & Current_branch & ")"
        End If
        
        
          If BranchID <> 0 Then
            StrSQL = StrSQL & " and   BranchId=" & BranchID
        End If
        
        
        StrSQL = StrSQL + " Order BY CusNamee"
        GetComboData MyCombo, StrSQL

    End If

End Function

Public Function GetBoxAccounts(MyCombo As DataCombo, _
                               Optional Boxtype As Integer = 0)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "select Account_Code,BoxName from TblBoxesData where Type=" & Boxtype
               StrSQL = StrSQL & " and ( BranchId=0  or      BranchId in(" & Current_branchSql & "))"
               
        StrSQL = StrSQL + " Order BY BoxName"
    Else
        StrSQL = "select Account_Code,BoxNamee from TblBoxesData where Type=" & Boxtype
               StrSQL = StrSQL & " and ( BranchId=0  or      BranchId in(" & Current_branchSql & "))"
               
        StrSQL = StrSQL + " Order BY BoxNamee"

    End If

    GetComboData MyCombo, StrSQL
End Function

Public Function ClearMyDataCombo(MyCombo As DataCombo)
    Dim StrSQL As String
    StrSQL = "Select CusID,CusName From TblCustemers"
    StrSQL = StrSQL + " Where (Type=30)"
    StrSQL = StrSQL + " Order BY CusName"
    GetComboData MyCombo, StrSQL
End Function
Public Function GetempyeesCredite(MyCombo As DataCombo, _
                             Optional WorkState As Boolean = False, _
                             Optional DriverOnly As Boolean = False)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID,NameApplicant From TblCreditFacicity Where 1=1  "
    Else
        StrSQL = "SELECT ID,NameApplicant From TblCreditFacicity Where 1=1  "
    End If

    If WorkState = True Then
        StrSQL = StrSQL + " and  workstate=1"
    End If

    If DriverOnly = True Then
        StrSQL = StrSQL + " and  not(driverid is null)"
    End If

    StrSQL = StrSQL + " Order By NameApplicant ASC"
    GetComboData MyCombo, StrSQL
End Function
Public Function GetempyeesFollow(MyCombo As DataCombo, _
                             Optional WorkState As Boolean = False, _
                             Optional DriverOnly As Boolean = False)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID,FromName From TblFollowUpAdminMeasure Where 1=1  "
    Else
        StrSQL = "SELECT ID,FromName From TblFollowUpAdminMeasure Where 1=1  "
    End If

    If WorkState = True Then
        StrSQL = StrSQL + " and  workstate=1"
    End If

    If DriverOnly = True Then
        StrSQL = StrSQL + " and  not(driverid is null)"
    End If

    StrSQL = StrSQL + " Order By FromName ASC"
    GetComboData MyCombo, StrSQL
End Function
Public Function GetClientComputerCheck(MyCombo As DataCombo, _
                             Optional WorkState As Boolean = False, _
                             Optional DriverOnly As Boolean = False)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID,ClientName From TblBillComputerChek  Where 1=1  "
    Else
        StrSQL = "SELECT ID,ClientName From TblBillComputerChek Where 1=1  "
    End If

    If WorkState = True Then
        StrSQL = StrSQL + " and  workstate=1"
    End If

    If DriverOnly = True Then
        StrSQL = StrSQL + " and  not(driverid is null)"
    End If

    StrSQL = StrSQL + " Order By ClientName ASC"
    GetComboData MyCombo, StrSQL
End Function
Public Function GetNameApplicant(MyCombo As DataCombo, _
                             Optional WorkState As Boolean = False, _
                             Optional DriverOnly As Boolean = False)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID,NameApplicant From TblCreditFacicity Where 1=1  "
    Else
        StrSQL = "SELECT ID,NameApplicant From TblCreditFacicity Where 1=1  "
    End If

    If WorkState = True Then
        StrSQL = StrSQL + " and  workstate=1"
    End If

    If DriverOnly = True Then
        StrSQL = StrSQL + " and  not(driverid is null)"
    End If

    StrSQL = StrSQL + " Order By NameApplicant ASC"
    GetComboData MyCombo, StrSQL
End Function
 Public Function GetEmployeesNationlity(MyCombo As DataCombo, _
                             Optional WorkState As Boolean = False, _
                             Optional DriverOnly As Boolean = False)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT Emp_ID,  Nationality From TblEmployee Where 1=1  "
    Else
        StrSQL = "SELECT Emp_ID,  Nationality From TblEmployee Where 1=1  "
    End If

    If WorkState = True Then
        StrSQL = StrSQL + " and  workstate=1"
    End If

    If DriverOnly = True Then
        StrSQL = StrSQL + " and  not(driverid is null)"
    End If

    StrSQL = StrSQL + " Order By Emp_Name ASC"
    GetComboData MyCombo, StrSQL
End Function


 Public Function GetPersons(MyCombo As DataCombo)
    Dim StrSQL As String
If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "Select CusID,CusName From TblCustemers"
Else
StrSQL = "Select CusID,CusNamee From TblCustemers"
End If
    StrSQL = StrSQL + " Where (Type=3)"
    StrSQL = StrSQL + " and   BranchId in(" & Current_branchSql & ") "
    StrSQL = StrSQL + " Order BY CusName"
    
    
    GetComboData MyCombo, StrSQL
End Function

Private Sub GetComboData(My_Combo As DataCombo, _
                         My_SQL As String, Optional ByVal IsOtherData As Boolean = False)
    Dim rs As ADODB.Recordset
    Dim StrTemp As String
    Dim Msg As String
    On Error GoTo ErrorHandler

    If InStr(1, My_SQL, "SELECT", vbTextCompare) = 0 Then
        Exit Sub
    End If

    My_Combo.Tag = My_SQL
    Set rs = New ADODB.Recordset

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        rs.CursorLocation = adUseClient
    End If
    If IsOtherData Then
        rs.Open My_SQL, POSConnection, adOpenStatic, adLockReadOnly, adCmdText
    Else
        rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    End If

    'Populate the ADO datacombo by setting its properties
    With My_Combo
        StrTemp = .BoundText
        Set .RowSource = rs
        .BoundColumn = rs(0).Name
        .ListField = rs(1).Name

        If Trim(StrTemp) <> "" Then
            .BoundText = StrTemp
        Else
            .BoundText = ""
            .text = ""
        End If

    End With

Exit_Sub:
    Set rs = Nothing
    Exit Sub
ErrorHandler:
    Msg = Now
    Msg = Msg & CHR(13) & "ClsDataCombos:GetComboData"
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & Err.Description
    WriteInLogFile Msg
    'MsgBox "ERROR! Err# " & Err.Number & " Desc: " & Err.Description, vbCritical + vbOKOnly
    Resume Exit_Sub
End Sub

Public Function GetManSupDecs(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT SupDecID,SupDecName From TblManSupDecs "
    Else
        StrSQL = "SELECT SupDecID,SupDecNamee From TblManSupDecs "
    End If

    StrSQL = StrSQL + " Where DecType=1"
    StrSQL = StrSQL + " Order By SupDecID"
    GetComboData MyCombo, StrSQL
End Function

Public Function GetRevenuesTypes(MyCombo As DataCombo)
    Dim StrSQL As String
    StrSQL = "SELECT RevenuesID,RevenuesName From TblRevenuesTypes "
    StrSQL = StrSQL + " Order By RevenuesName"
    GetComboData MyCombo, StrSQL
End Function

Public Sub GetScreens(MyCombo As DataCombo)
    Dim StrSQL As String
    StrSQL = "SELECT ScreenName,ScreenCaption From Screens "
    StrSQL = StrSQL + " Order By ScreenType,ScreenOrder"
    GetComboData MyCombo, StrSQL
End Sub

Public Function GetManCompanies(MyCombo As DataCombo)
    Dim StrSQL As String
    StrSQL = "Select CusID,CusName From TblCustemers"
    StrSQL = StrSQL + " Where (Type=4)"
    StrSQL = StrSQL + " Order BY CusName"
    GetComboData MyCombo, StrSQL
End Function

Public Function GetManComDecs(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT SupDecID,SupDecName From TblManSupDecs "
    Else
        StrSQL = "SELECT SupDecID,SupDecNamee From TblManSupDecs "
    End If

    StrSQL = StrSQL + " Where DecType=3"
    StrSQL = StrSQL + " Order By SupDecID"
    GetComboData MyCombo, StrSQL
End Function

Public Function GetFinTransactionsTypes(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT NotesType,NotesTypeName From TblNotesTypes "
        StrSQL = StrSQL & " where NotesType in(3,4,5,80)"
        StrSQL = StrSQL + " Order By NotesTypeName"
    Else
        StrSQL = "SELECT NotesType,NotesTypeNamee From TblNotesTypes "
        StrSQL = StrSQL & " where NotesType in(3,4,5,80)"
        StrSQL = StrSQL + " Order By NotesTypeName"
    End If

    GetComboData MyCombo, StrSQL
End Function

Public Function GetDocTypes(MyCombo As DataCombo, Optional filter As Integer = -1)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT Transaction_Type,TransactionTypeName From TransactionTypes "
        If filter <> -1 Then
        StrSQL = StrSQL & " where Transaction_Type in (21,22,5,9,19,20,27,28)"
        Else
        StrSQL = StrSQL & " where Transaction_Type >2"
            End If
        StrSQL = StrSQL + " Order By TransactionTypeName"
    
        
    Else
        StrSQL = "SELECT Transaction_Type,(TransactionEnglishName) From TransactionTypes "
               If filter <> -1 Then
        StrSQL = StrSQL & " where Transaction_Type in (21,22,5,9,19,20,27,28)"
                Else
        StrSQL = StrSQL & " where Transaction_Type >2"
        
        End If
        
        StrSQL = StrSQL + " Order By TransactionEnglishName"
    End If

    GetComboData MyCombo, StrSQL
End Function
Public Function GetMedicalContractingType(MyCombo As DataCombo)

    Dim StrSQL As String
 If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT id,name From MedicalContractingType "
  Else
  StrSQL = "SELECT id,namee From MedicalContractingType "
  End If
 
    StrSQL = StrSQL + " Order By id "
    GetComboData MyCombo, StrSQL
End Function

Public Function GetCodeing(MyCombo As DataCombo, _
                           FIELD_no As Integer, Optional ByVal mTableName As String = "", Optional ByVal mWhere As String = "")

    Dim StrSQL As String
 
    StrSQL = "SELECT FIELD_no,prifix From Coding "
    StrSQL = StrSQL + "  Where FIELD_no=" & FIELD_no & " "
 
    StrSQL = StrSQL + " Order By prifix "
    
    Dim rsDummy As New ADODB.Recordset
    rsDummy.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic
    Dim xx As Boolean
    If rsDummy.EOF Then
       xx = True
    Else
        If Trim(rsDummy!prifix & "") = "" Then
            xx = True
        End If
    End If
    If xx Then
        If mTableName <> "" Then
            If mWhere = "" Then mWhere = " 1 = 1 "
            StrSQL = " Select 0,prifix from " & mTableName & "  where IsNull(prifix,'') <> '' and " & mWhere & " group by  prifix"
            
        End If
    End If
    GetComboData MyCombo, StrSQL
End Function
 

 Public Function GetAccountingCodes(MyCombo As DataCombo, _
                                   Optional LastAccount As Boolean = False, _
                                   Optional BasicAccount As Boolean = False, Optional AccountTab As Integer = -1, Optional HideAccoun As Boolean = False, Optional UserAccount As Integer = 0)

    Dim StrSQL As String
    If UserAccount = 1 Then
    If CheckmanyAccount(StrSQL) = False Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        StrSQL = "SELECT Account_Code,Account_Name From ACCOUNTS where 1=1 "
                    Else
                        StrSQL = "SELECT Account_Code,Account_Nameeng From ACCOUNTS  where 1=1 "
                    End If

    If BasicAccount = True Then
        StrSQL = StrSQL + "  and last_account=0  "
    End If

    If LastAccount = False Then
        'StrSQL = StrSQL + "  Where last_account=0  "
    Else
        StrSQL = StrSQL + "  and last_account=1 "
    End If

If AccountTab <> -1 Then
    StrSQL = StrSQL + "  and AccountTab=" & AccountTab & " "
End If
If HideAccoun = True Then
 StrSQL = StrSQL & GetAccountCodeHiding
 End If
 End If
 Else
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT Account_Code,Account_Name From ACCOUNTS where 1=1 "
    Else
        StrSQL = "SELECT Account_Code,Account_Nameeng From ACCOUNTS  where 1=1 "
    End If
     If BasicAccount = True Then
        StrSQL = StrSQL + "  and last_account=0  "
    End If

    If LastAccount = False Then
        'StrSQL = StrSQL + "  Where last_account=0  "
    Else
        StrSQL = StrSQL + "  and last_account=1 "
    End If

If AccountTab <> -1 Then
    StrSQL = StrSQL + "  and AccountTab=" & AccountTab & " "
End If
If HideAccoun = True Then
 StrSQL = StrSQL & GetAccountCodeHiding
 End If
 End If
 
    'StrSQL = StrSQL + " Where DecType=3"
    StrSQL = StrSQL + " Order By Account_Name "
    GetComboData MyCombo, StrSQL
End Function
   Public Function GetAccountingCodesx31102017(MyCombo As DataCombo, _
                                   Optional LastAccount As Boolean = False, _
                                   Optional BasicAccount As Boolean = False, Optional AccountTab As Integer = -1)

    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT Account_Code,Account_Name From ACCOUNTS where 1=1 "
    Else
        StrSQL = "SELECT Account_Code,Account_Nameeng From ACCOUNTS  where 1=1 "
    End If

    If BasicAccount = True Then
        StrSQL = StrSQL + "  and last_account=0  "
    End If

    If LastAccount = False Then
        'StrSQL = StrSQL + "  Where last_account=0  "
    Else
        StrSQL = StrSQL + "  and last_account=1 "
    End If

If AccountTab <> -1 Then
    StrSQL = StrSQL + "  and AccountTab=" & AccountTab & " "
End If
StrSQL = StrSQL & GetAccountByBarnchUser
'StrSQL = StrSQL & GetAccountCodeHiding
    'StrSQL = StrSQL + " Where DecType=3"
    StrSQL = StrSQL + " Order By Account_Name "
    GetComboData MyCombo, StrSQL
End Function
Public Function GetAccountingCodesx17072017(MyCombo As DataCombo, _
                                   Optional LastAccount As Boolean = False, _
                                   Optional BasicAccount As Boolean = False, Optional AccountTab As Integer = -1)

    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT Account_Code,Account_Name From ACCOUNTS where 1=1 "
    Else
        StrSQL = "SELECT Account_Code,Account_Nameeng From ACCOUNTS  where 1=1 "
    End If

    If BasicAccount = True Then
        StrSQL = StrSQL + "  and last_account=0  "
    End If

    If LastAccount = False Then
        'StrSQL = StrSQL + "  Where last_account=0  "
    Else
        StrSQL = StrSQL + "  and last_account=1 "
    End If

If AccountTab <> -1 Then
    StrSQL = StrSQL + "  and AccountTab=" & AccountTab & " "
End If

    'StrSQL = StrSQL + " Where DecType=3"
    StrSQL = StrSQL + " Order By Account_Name "
    GetComboData MyCombo, StrSQL
End Function

Public Function GetAccountingCodesENg(MyCombo As DataCombo, _
                                      Optional LastAccount As Boolean = True)
    Dim StrSQL As String
    StrSQL = "SELECT Account_Code,Account_NameEng From ACCOUNTS "

    If LastAccount = False Then
        StrSQL = StrSQL + "  Where last_account=0  "
    Else
        StrSQL = StrSQL + "  Where last_account=1 "
    End If

    'StrSQL = StrSQL + " Where DecType=3"
    StrSQL = StrSQL + " Order By Account_Code "
    GetComboData MyCombo, StrSQL
End Function

Public Function GetWorkOrders(MyCombo As DataCombo)
    Dim StrSQL As String
    StrSQL = "SELECT WorkOrderID,WorkOrderName From TblWorkOrdersData "
    'StrSQL = StrSQL + " Where WorkOrderState=0"
    StrSQL = StrSQL + " Order By WorkOrderID "
    GetComboData MyCombo, StrSQL
End Function

Public Function GetManDesByID(MyCombo As DataCombo, _
                              StrWhere As String)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT SupDecID, SupDecName From dbo.TblManSupDecs"
    Else
        StrSQL = "SELECT SupDecID, SupDecNamee From dbo.TblManSupDecs"
    End If

    If StrWhere <> "" Then
        StrSQL = StrSQL + " Where " & StrWhere
    End If

    StrSQL = StrSQL + " ORDER BY SupDecID"
    GetComboData MyCombo, StrSQL
End Function


Public Function GetEmpLocations(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select GroupID,GroupName  From EmpGroupDep where LastGroup=1 Order By GroupName"
    Else
        StrSQL = "Select GroupID,GroupName From EmpGroupDep where LastGroup=1 Order By GroupName"
    End If

    GetComboData MyCombo, StrSQL
End Function
Public Function GetEmpGrades(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select gradeid,name  From TblEmpGrades Order By gradeid"
    Else
        StrSQL = "Select gradeid,namee From TblEmpGrades Order By gradeid"
    End If

    GetComboData MyCombo, StrSQL
End Function
 
Public Function GetEmpRelations(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select id,name  From TblRelations Order By name"
    Else
        StrSQL = "Select id,namee From TblRelations Order By name"
    End If

    GetComboData MyCombo, StrSQL
End Function
Public Function GetEmptblworkshop(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select DeparmentID,DepartmentName From tblworkshop Order By DepartmentName"
    Else
        StrSQL = "Select DeparmentID,DepartmentNamee From tblworkshop Order By DepartmentNamee"
    End If

    GetComboData MyCombo, StrSQL
End Function
Public Function GetEquipments(MyCombo As DataCombo, Optional insql As String)
    Dim StrSQL As String

                If SystemOptions.UserInterface = ArabicInterface Then
                '    StrSQL = " SELECT     id, Name"
                '    StrSQL = StrSQL & " from dbo.FixedAssets"
               '     StrSQL = StrSQL & " WHERE    id IN"
               '     StrSQL = StrSQL & " (SELECT     fixedAssetid"
               '     StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
               '     StrSQL = StrSQL & " or   id IN"
               '     StrSQL = StrSQL & " (SELECT     fixedAssetid"
               '     StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
               '     StrSQL = StrSQL & " order by Namee  "
               '
                     StrSQL = " SELECT     id, Name"
                    StrSQL = StrSQL & " from dbo.FixedAssets where ISEQUP = 1"
                    If insql <> "" Then
                    StrSQL = StrSQL & " and id in ( " & insql & ")"
                    End If
                    
                        StrSQL = StrSQL & " order by Name   "
                Else
                   '                     StrSQL = " SELECT     id, Namee"
                   ' StrSQL = StrSQL & " from dbo.FixedAssets "
                   ' StrSQL = StrSQL & " WHERE    id IN"
                   ' StrSQL = StrSQL & " (SELECT     fixedAssetid"
                   ' StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                   ' StrSQL = StrSQL & " or   id IN"
                   ' StrSQL = StrSQL & " (SELECT     fixedAssetid"
                   ' StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                   ' StrSQL = StrSQL & " order by Name  "
                    StrSQL = " SELECT     id, Namee"
                    StrSQL = StrSQL & " from dbo.FixedAssets where ISEQUP = 1"
                         If insql <> "" Then
                    StrSQL = StrSQL & " and id in ( " & insql & ")"
                    End If
                    
                        StrSQL = StrSQL & " order by Namee "
                        
                End If
                

    GetComboData MyCombo, StrSQL
End Function


Public Function GetEmpDepartments(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select DeparmentID,DepartmentName From TblEmpDepartments Order By DepartmentName"
    Else
        StrSQL = "Select DeparmentID,DepartmentNamee From TblEmpDepartments Order By DepartmentNamee"
    End If

    GetComboData MyCombo, StrSQL
End Function

Public Function GetEmpContractTypes(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "Select id,name From Emp_Contract_type  Order By name"
    Else
    StrSQL = "Select id,nameE From Emp_Contract_type Order By nameE"
    End If
    GetComboData MyCombo, StrSQL
End Function

Public Function GetInsuranceClass(MyCombo As DataCombo)
    Dim StrSQL As String
    StrSQL = "Select id,name From insurance_class  Order By name"
    GetComboData MyCombo, StrSQL
End Function
Public Function GetWorkShopGroup(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "Select ID,name From TBLSalesRepGroups3 Order By name"
    Else
    StrSQL = "Select ID,namee From TBLSalesRepGroups3 Order By name"
    End If
    
    GetComboData MyCombo, StrSQL
End Function

Public Function GetSalesRepGroupsPurchase(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "Select ID,name From TBLSalesRepGroups2 Order By name"
    Else
    StrSQL = "Select ID,namee From TBLSalesRepGroups2 Order By name"
    End If
    
    GetComboData MyCombo, StrSQL

End Function

Public Function GetTblCommisionTypes(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "Select ID,name From TblCommisionTypes Order By name"
    Else
    StrSQL = "Select ID,namee From TblCommisionTypes Order By name"
    End If
    
    GetComboData MyCombo, StrSQL
End Function


Public Function GetSalesRepGroups(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "Select ID,name From TBLSalesRepGroups Order By name"
    Else
    StrSQL = "Select ID,namee From TBLSalesRepGroups Order By name"
    End If
    
    GetComboData MyCombo, StrSQL
End Function
 
Public Function GetEmpJobsTypes(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select JobTypeID,JobTypeName From TblEmpJobsTypes Order By JobTypeName"
    Else
        StrSQL = "Select JobTypeID,JobTypeNamee From TblEmpJobsTypes Order By JobTypeNamee"
    End If

    GetComboData MyCombo, StrSQL
End Function

 

Public Function GetEmpSpecifications(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "Select SpecificationID,SpecificationName From TblEmpSpecifications Order By SpecificationName"
    Else
    StrSQL = "Select SpecificationID,SpecificationNameE From TblEmpSpecifications Order By SpecificationName"
    End If
    GetComboData MyCombo, StrSQL
End Function

Public Function GetIndustrialItems(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select TBLProductionElementsId,name From TBLProductionElements Order By TBLProductionElementsId"
    Else
        StrSQL = "Select TBLProductionElementsId,namee From TBLProductionElements Order By TBLProductionElementsId"
    End If

    GetComboData MyCombo, StrSQL
End Function

Public Function GetCUrrencyNames(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select id,name From currency Order By ID"
    Else
        StrSQL = "Select id,code From currency Order By ID"
    End If

    GetComboData MyCombo, StrSQL
End Function

Public Function GetLCTypesName(MyCombo As DataCombo, Optional ByVal mWhere As String = "")
    Dim StrSQL As String
    StrSQL = "Select ID,name From LCTypes "
    If mWhere <> "" Then
        StrSQL = StrSQL & "  Where 1 = 1 and " & mWhere
    End If
    StrSQL = StrSQL & "  Order By name"
    GetComboData MyCombo, StrSQL
End Function

Public Function GetExpensesNames(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "select Expensesid,account_name from Expenses_accounts"
    Else
        StrSQL = "select Expensesid,account_nameeng from Expenses_accounts_eng "
    End If

    GetComboData MyCombo, StrSQL
End Function

Public Function GetCustomerType(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select id,name From TblCustomerType Order By name"
    Else
        StrSQL = "Select id,namee From TblCustomerType Order By name"
    End If

    GetComboData MyCombo, StrSQL
End Function

Public Function GetproductionLinesNames(MyCombo As DataCombo)
    Dim StrSQL As String
    StrSQL = "Select id,name From TblProductLine Order By name"
    GetComboData MyCombo, StrSQL
End Function

Public Function GetCountriesNames(MyCombo As DataCombo)
    Dim StrSQL As String
    StrSQL = "Select CountryID,CountryName From TblCountriesData Order By CountryID"
    GetComboData MyCombo, StrSQL
End Function

Public Function getCountriesGovernments(MyCombo As DataCombo, _
                                        Optional Lngid As Long = 0)
    Dim StrSQL As String
    StrSQL = "Select GovernmentID,GovernmentName From TblCountriesGovernments "

    If Lngid <> 0 Then
        StrSQL = StrSQL + " Where CountryID=" & Lngid
    End If

    StrSQL = StrSQL + " Order By GovernmentID"

    GetComboData MyCombo, StrSQL
End Function

Public Function getCCCategory(MyCombo As DataCombo, _
                              Optional Lngid As Long = 0)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select id,name From Marakes_taklefa_type "
    Else
        StrSQL = "Select id,namee From Marakes_taklefa_type "
    End If

    If Lngid <> 0 Then
        StrSQL = StrSQL + " Where account_no=" & Lngid
    End If

    StrSQL = StrSQL + " Order By name"

    GetComboData MyCombo, StrSQL
End Function

'
'
Public Function getCC(MyCombo As DataCombo, _
                      Optional str1 As String = "", _
                      Optional str2 As String = "", _
                      Optional str3 As String = "", _
                      Optional str4 As String = "")
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "Select account_no,account_name From markaas_taklefa "
    Else
        StrSQL = "Select account_no,EnglishName From markaas_taklefa "
    End If

    If str1 <> "" And str2 <> "" Then
        StrSQL = StrSQL + " Where account_type='" & str1 & "' or account_type='" & str2 & "' or account_type='" & str3 & "' or account_type='" & str4 & "'"
    End If

        If SystemOptions.usertype <> UserAdminAll Or val(Current_branch) <> 0 Then
        StrSQL = StrSQL & " and    branch_no=" & Current_branch
    End If
    
    StrSQL = StrSQL + " Order By account_name"

    GetComboData MyCombo, StrSQL
End Function

Public Function getöaArrowsGroup(MyCombo As DataCombo, _
                                 Optional Lngid As Long = 0)
    Dim StrSQL As String
    StrSQL = "Select GroupID,GroupName From ArrowsGroup "

    If Lngid <> 0 Then
        StrSQL = StrSQL + " Where GroupID=" & Lngid
    End If

    StrSQL = StrSQL + " Order By GroupName"

    GetComboData MyCombo, StrSQL
End Function

Public Function getArrowsCompany(MyCombo As DataCombo, _
                                 Optional Lngid As Long = 0)
    Dim StrSQL As String
    StrSQL = "Select CompanyId,CompanyName From ArrowsCompanies "

    If Lngid <> 0 Then
        StrSQL = StrSQL + " Where CompanyId=" & Lngid
    End If

    StrSQL = StrSQL + " Order By CompanyName"

    GetComboData MyCombo, StrSQL
End Function

Public Function getFinMarkets(MyCombo As DataCombo, _
                              Optional Lngid As Long = 0)
    Dim StrSQL As String
    StrSQL = "Select FinancialMarketId,FinancialMarketName From FinancialMarkets "

    If Lngid <> 0 Then
        StrSQL = StrSQL + " Where FinancialMarketId=" & Lngid
    End If

    StrSQL = StrSQL + " Order By FinancialMarketName"

    GetComboData MyCombo, StrSQL
End Function
Public Function GetIqar(MyCombo As DataCombo)
                            
    Dim StrSQL As String
     StrSQL = "SELECT Aqarid,aqarname From TblAqar   where 1=1  "
   
       '     If SystemOptions.usertype = UserAdminAll Then
       '  Else
       ' StrSQL = StrSQL & " where   branchid=" & Current_branch
    'End If
    
      StrSQL = StrSQL & " and ( BranchId=0  or      BranchId in(" & Current_branchSql & "))"
    StrSQL = StrSQL + " Order By aqarname ASC"
    GetComboData MyCombo, StrSQL
End Function


Public Function GetIqarxx2312205(MyCombo As DataCombo)
                            
    Dim StrSQL As String
     StrSQL = "SELECT Aqarid,aqarname From TblAqar   "
    StrSQL = StrSQL + " Order By aqarname ASC"
    GetComboData MyCombo, StrSQL
End Function
Public Function GetExpensesAccount(MyCombo As DataCombo)
    Dim StrSQL As String
        StrSQL = "SELECT  Account_Code , Account_Name from  Expenses_accounts "
  
        StrSQL = StrSQL & " order by Account_Name "
   

    GetComboData MyCombo, StrSQL
End Function
Public Function GetIqaroldd(MyCombo As DataCombo)
                            
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT Aqarid,aqarname From TblAqar   "
    Else
        StrSQL = "SELECT Aqarid,aqarname From TblAqar   "
    End If

 

    StrSQL = StrSQL + " Order By Aqarid ASC"
    GetComboData MyCombo, StrSQL
End Function
 Public Function getAkarUnit(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "Select id,name  From TblAkarUnit   Order By name"
   Else
   StrSQL = "Select id,namee  From TblAkarUnit   Order By namee"
   End If
      GetComboData MyCombo, StrSQL
End Function
Public Function GetIqarUnit(ID As Long, unittype As Long, MyCombo As DataCombo, Optional flagt As String, Optional strsql2 As String = "")
                            
    Dim StrSQL As String
     If flagt = "R" Then
        StrSQL = "SELECT Id,unitno From TblAqarDetai  where      Aqarid = " & ID & " and unittype=" & unittype
    Else
    StrSQL = "SELECT Id,unitno From TblAqarDetai  where   ( (Status IS NULL or Status=0 or Status=2 ) and  Aqarid = " & ID & " and unittype=" & unittype & " )"
    End If
    If flagt = "E" Then
    StrSQL = StrSQL & strsql2
   End If
     If ID = -2 Then
        StrSQL = "SELECT Id,unitno From TblAqarDetai " '  where      Aqarid = " & id & " and unittype=" & unittype
     End If
    

    StrSQL = StrSQL + " Order By Aqarid ASC"
    GetComboData MyCombo, StrSQL
End Function

Public Function getAkarType(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "Select id,name  From tblAkarType    Order By name"
   Else
   StrSQL = "Select id,namee  From tblAkarType   Order By namee"
   End If
      GetComboData MyCombo, StrSQL
End Function

Public Function getSchemes(MyCombo As DataCombo)
    Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "Select id,name  From tblSchemes    Order By name"
   Else
   StrSQL = "Select id,namee  From tblSchemes   Order By namee"
   End If
      GetComboData MyCombo, StrSQL
End Function
Public Function getCustomersLocations(MyCombo As DataCombo, Optional CusID As Long = 0)
    Dim StrSQL As String
If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT     dbo.TblCustomersLocations.ID, dbo.TblCustomersLocations.Name"
  Else
  StrSQL = "SELECT     dbo.TblCustomersLocations.ID, dbo.TblCustomersLocations.Namee"
  End If
 StrSQL = StrSQL + " FROM         dbo.TblCustemers INNER JOIN"
 StrSQL = StrSQL + "                       dbo.TblCustomersLocations ON dbo.TblCustemers.CusID = dbo.TblCustomersLocations.CusId"



    If CusID <> 0 Then
        StrSQL = StrSQL + "  Where(dbo.TblCustomersLocations.CusID = " & CusID & ")"
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = StrSQL + " Order By Name"
 Else
 StrSQL = StrSQL + " Order By Namee"
 End If

    GetComboData MyCombo, StrSQL
End Function

Public Function getÇhay(MyCombo As DataCombo, Optional Lngid As Long = 0)
     Dim StrSQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
      StrSQL = "Select CityID,CityName From TblCountriesGovernmentsCities where Not(CityName is null) "
    Else
    StrSQL = "Select CityID,NameE From TblCountriesGovernmentsCities  where Not(NameE is null)"
    End If
    If Lngid <> 0 Then
        StrSQL = StrSQL + " and  CityID=" & Lngid
    End If
If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = StrSQL + " Order By CityName"
Else
     StrSQL = StrSQL + " Order By NameE"
End If
    GetComboData MyCombo, StrSQL
    
End Function

Public Function GetCountrieshay(MyCombo As DataCombo, _
                                Optional Lngid As Long = 0)
    Dim StrSQL As String
    StrSQL = "Select GovernmentID,GovernmentName From TblCountriesGovernments "

    If Lngid <> 0 Then
        StrSQL = StrSQL + " Where id=" & Lngid
    End If

    StrSQL = StrSQL + " Order By GovernmentName"

    GetComboData MyCombo, StrSQL
End Function

Public Function GethayStreets(MyCombo As DataCombo, _
                              Optional Lngid As Long = 0, _
                              Optional LngGovernmentID As Long = 0)

    Dim StrSQL As String
 
    StrSQL = "SELECT     dbo.streets.CityID AS streetid, dbo.streets.CityName as streetname"
    StrSQL = StrSQL + " FROM         dbo.TblCountriesGovernmentsCities INNER JOIN"
    StrSQL = StrSQL + "  dbo.streets ON dbo.TblCountriesGovernmentsCities.CityID = dbo.streets.GovernmentID"
    StrSQL = StrSQL + " WHERE     (dbo.TblCountriesGovernmentsCities.CityID <>0)"

    If Lngid <> 0 Then
        StrSQL = StrSQL + " and      (dbo.TblCountriesGovernmentsCities.CityID = " & Lngid & ")"
    End If

    StrSQL = StrSQL + " Order By streetname"

    GetComboData MyCombo, StrSQL

End Function


Public Function GetCountriesGovernCities(MyCombo As DataCombo, _
                                         Optional Lngid As Long = 0, _
                                         Optional LngGovernmentID As Long = 0)

    Dim StrSQL As String
    'StrSQL = "SELECT dbo.TblCountriesGovernmentsCities.CityID, dbo.TblCountriesGovernmentsCities.CityName," & _
    '"dbo.TblCountriesGovernmentsCities.GovernmentID,dbo.warsha.id "
    ''StrSQL = StrSQL + "  FROM dbo.TblCountriesGovernments INNER JOIN dbo.TblCountriesGovernmentsCities ON " & _
      "'dbo.TblCountriesGovernments.GovernmentID = dbo.TblCountriesGovernmentsCities.GovernmentID INNER JOIN "
    'StrSQL = StrSQL + " dbo.warsha ON dbo.TblCountriesGovernments.id = " & _
    '" dbo.warsha.id "

    'StrSQL = StrSQL + " Where  dbo.TblCountriesGovernmentsCities.CityID <>0"
    StrSQL = "SELECT     TOP 100 PERCENT dbo.TblCountriesGovernmentsCities.CityID, dbo.TblCountriesGovernmentsCities.CityName, "
    StrSQL = StrSQL + " dbo.TblCountriesGovernmentsCities.GovernmentID"
    StrSQL = StrSQL + " FROM         dbo.TblCountriesGovernments INNER JOIN"
    StrSQL = StrSQL + " dbo.TblCountriesGovernmentsCities ON dbo.TblCountriesGovernments.GovernmentID = dbo.TblCountriesGovernmentsCities.GovernmentID"
    StrSQL = StrSQL + " Where  dbo.TblCountriesGovernmentsCities.CityID <>0"
 
    If LngGovernmentID <> 0 Then
        ' StrSQL = StrSQL + " AND dbo.warsha.id=" & Lngid
    End If

    If Lngid <> 0 Then
        StrSQL = StrSQL + " AND dbo.TblCountriesGovernmentsCities.GovernmentID=" & LngGovernmentID
    End If

    StrSQL = StrSQL + " Order By CityName"

    GetComboData MyCombo, StrSQL

End Function

Public Function GetItemsMaking(MyCombo As DataCombo, _
                               Optional IntItemType As Integer = -1, _
                               Optional IntAssbliedType As Integer = -1)
    Dim StrSQL As String

    'IntItemType=-1 (ALL Items)
    'IntItemType=0 (Goods Items)
    'IntItemType=1 (Services Items)

    'IntAssbliedType=-1 (All Items)
    'IntAssbliedType=0 ( NOT Assblied Items)
    'IntAssbliedType=1 ( Assblied Items)

    StrSQL = "SELECT ItemID,ItemName From  TblItems "
    StrSQL = StrSQL + " Where ItemMaking=1"

    If IntItemType <> -1 Then
        StrSQL = StrSQL + " AND (ItemType=" & IntItemType & ")"
    End If

    If IntAssbliedType <> -1 Then
        If IntAssbliedType = 0 Then

            'AssbliedItem=true
            If SystemOptions.SysDataBaseType = AccessDataBase Then
                StrSQL = StrSQL + " AND (AssbliedItem=False OR AssbliedItem=0)"
            ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrSQL = StrSQL + " AND (AssbliedItem=0)"
            End If

        ElseIf IntAssbliedType = 1 Then

            If SystemOptions.SysDataBaseType = AccessDataBase Then
                StrSQL = StrSQL + " AND (AssbliedItem=True OR AssbliedItem=1)"
            ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
                '            StrSQL = StrSQL + " AND (AssbliedItem=1)"
            End If
        
        End If
    End If

    StrSQL = StrSQL + " Order BY ItemName "
    GetComboData MyCombo, StrSQL
End Function

Public Function GetItemsColors(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = " SELECT 0 AS ColorID, '' AS ColorName UNION ALL SELECT ColorID,ColorName From TblItemsColors "
    Else
        StrSQL = " SELECT 0 AS ColorID, '' AS ColorName UNION ALL SELECT ColorID,ColorName From TblItemsColors "
    End If

    StrSQL = StrSQL + " Order BY ColorName "
    GetComboData MyCombo, StrSQL
End Function

Public Function GetItemsSizes(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT 0 AS SizeId, '' AS SizeName UNION ALL SELECT SizeId,SizeName From TblItemsSizes "
    Else
        StrSQL = "SELECT 0 AS SizeId, '' AS SizeName UNION ALL SELECT SizeId,SizeName From TblItemsSizes "
    End If

    StrSQL = StrSQL + " Order BY SizeName "
    GetComboData MyCombo, StrSQL
End Function
Public Function GetItemsClasses(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "  SELECT SizeId,SizeName From TblItemsclasses "
    Else
        StrSQL = "SELECT SizeId,SizeNameE From TblItemsclasses "
    End If

    StrSQL = StrSQL + " Order BY SizeName "
    GetComboData MyCombo, StrSQL
End Function

Public Function GetItemsClassesold(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT SizeId,SizeName From TblItemsclasses "
    Else
        StrSQL = "SELECT SizeId,SizeName From TblItemsclasses "
    End If

    StrSQL = StrSQL + " Order BY SizeName "
    GetComboData MyCombo, StrSQL
End Function

Public Function GetItemsUnitsábyitemid(MyCombo As DataCombo, LngItemID As Double)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT TblItemsUnits.UnitID, TblUnites.UnitName  "
        StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits " & "ON TblUnites.UnitID = TblItemsUnits.UnitID "
        StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & LngItemID
        StrSQL = StrSQL + " Order BY TblItemsUnits.SecOrder "
 
    Else
        StrSQL = "SELECT TblItemsUnits.UnitID,  TblUnites.UnitNamee "
        StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits " & "ON TblUnites.UnitID = TblItemsUnits.UnitID "
        StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & LngItemID
        StrSQL = StrSQL + " Order BY TblItemsUnits.SecOrder "
    End If

    'StrSQL = StrSQL + " Order BY UnitID "
    GetComboData MyCombo, StrSQL
End Function

Public Function GetIProcessUnits(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT UnitID,UnitName From TblProcessUnites "
    Else
        StrSQL = "SELECT UnitID,UnitNamee From TblProcessUnites "
    End If

    StrSQL = StrSQL + " Order BY UnitID "
    GetComboData MyCombo, StrSQL
End Function

Public Function GetItemsUnits(MyCombo As DataCombo)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT UnitID,UnitName From TblUnites "
    Else
        StrSQL = "SELECT UnitID,UnitNamee From TblUnites "
    End If

    StrSQL = StrSQL + " Order BY UnitID "
    GetComboData MyCombo, StrSQL
End Function
 
'''''''///////////////////
Public Function GetTblMaintenanceWork(MyCombo As DataCombo, _
                                    Optional BolLoadAdmins As Boolean = True)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT id,name From TblMaintenanceWork ORDER BY ID "
    Else
        StrSQL = "SELECT id,namee From TblMaintenanceWork ORDER BY ID"
    End If
 
    GetComboData MyCombo, StrSQL
End Function


Public Function GetTblColor(MyCombo As DataCombo, _
                                    Optional BolLoadAdmins As Boolean = True)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT id,name From TblColor order by name"
    Else
        StrSQL = "SELECT id,namee From TblColor order by namee"
    End If
 
    GetComboData MyCombo, StrSQL
End Function
Public Function GetTblCarModels(MyCombo As DataCombo, _
                                Optional BolLoadAdmins As Boolean = True, Optional Index As Integer)
    Dim StrSQL As String

    If SystemOptions.UserInterface = ArabicInterface Then
   StrSQL = "SELECT id, Model from TblCarModels where CarID=" & Index & ""
    Else
        StrSQL = "SELECT id,Modele From TblCarModels where CarID=" & Index & ""
    End If
 
    GetComboData MyCombo, StrSQL
End Function

'********************************************************************************************
Public Function GetItemsNamesupdate(MyCombo As DataCombo, _
                              Optional IntItemType As Integer = -1, _
                              Optional IntAssbliedType As Integer = -1, _
                              Optional ItemMakingNew As Integer = -1, _
                              Optional AllItems As Boolean = False, Optional Ch As Integer = 0)
    Dim StrSQL As String

    'IntItemType=-1 (ALL Items)
    'IntItemType=0 (Goods Items)
    'IntItemType=1 (Services Items)

    'IntAssbliedType=-1 (All Items)
    'IntAssbliedType=0 ( NOT Assblied Items)
    'IntAssbliedType=1 ( Assblied Items)
    If Ch = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then

        StrSQL = "SELECT ItemID,ItemName From  TblItems "
    Else
        StrSQL = "SELECT ItemID,ItemNamee From  TblItems "
    End If

    If AllItems = False Then
        StrSQL = StrSQL + " Where ((IsArchive <> 1) or (IsArchive Is null))"
    End If

    If IntItemType <> -1 Then
        StrSQL = StrSQL + " AND (ItemType=" & IntItemType & ")"
    End If

    If ItemMakingNew <> -1 Then
        StrSQL = StrSQL + " AND (ItemMakingNew=" & ItemMakingNew & ")"
    End If

    If IntAssbliedType <> -1 Then
        If IntAssbliedType = 0 Then

            'AssbliedItem=true
            If SystemOptions.SysDataBaseType = AccessDataBase Then
                StrSQL = StrSQL + " AND (AssbliedItem=False OR AssbliedItem=0)"
            ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrSQL = StrSQL + " AND (AssbliedItem=0)"
            End If

        ElseIf IntAssbliedType = 1 Then

            If SystemOptions.SysDataBaseType = AccessDataBase Then
                StrSQL = StrSQL + " AND (AssbliedItem=True OR AssbliedItem=1)"
            ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrSQL = StrSQL + " AND (AssbliedItem=1)"
            End If
        
        End If
    End If
Else
    If SystemOptions.UserInterface = ArabicInterface Then

        StrSQL = "SELECT ItemID,ItemName From  TblItems where GroupID=" & Ch & " "
    Else
        StrSQL = "SELECT ItemID,ItemNamee From  TblItems where GroupID=" & Ch & " "
    End If
End If
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = StrSQL + " Order BY ItemName "
    Else
        StrSQL = StrSQL + " Order BY ItemNamee "
    End If

    GetComboData MyCombo, StrSQL
End Function



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
 


 

