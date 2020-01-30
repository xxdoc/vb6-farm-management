VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDaily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_ModuleName As String
Private m_ADOConn As ADODB.Connection
Private m_Rs1 As ADODB.Recordset
Private m_Rs2 As ADODB.Recordset
Private m_Rs3 As ADODB.Recordset

Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   m_ModuleName = "clsDaily"
   Set m_Rs1 = New ADODB.Recordset
   Set m_Rs2 = New ADODB.Recordset
   Set m_Rs3 = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
   If m_Rs1.State = adStateOpen Then
      m_Rs1.Close
   End If
   Set m_Rs1 = Nothing

   If m_Rs2.State = adStateOpen Then
      m_Rs2.Close
   End If
   Set m_Rs2 = Nothing

   If m_Rs3.State = adStateOpen Then
      m_Rs3.Close
   End If
   Set m_Rs3 = Nothing
End Sub

Public Sub PopulateProductType(Pt As CProductType)
Dim iCount As Long

   Call Pt.DeleteAllData
   
   Pt.AddEditMode = SHOW_ADD
   Pt.PRODUCT_TYPE_NAME = "สุกรพ่อพันธ์"
   Pt.PRODUCT_TYPE_NO = "B"
   Call Pt.AddEditData
   
   Pt.AddEditMode = SHOW_ADD
   Pt.PRODUCT_TYPE_NAME = "สุกรสำรองพ่อ"
   Pt.PRODUCT_TYPE_NO = "BT"
   Call Pt.AddEditData
   
   Pt.AddEditMode = SHOW_ADD
   Pt.PRODUCT_TYPE_NAME = "สุกรแม่คลอด"
   Pt.PRODUCT_TYPE_NO = "L"
   Call Pt.AddEditData
   
   Pt.AddEditMode = SHOW_ADD
   Pt.PRODUCT_TYPE_NAME = "สุกรแม่อุ้มท้อง"
   Pt.PRODUCT_TYPE_NO = "G"
   Call Pt.AddEditData

   Pt.AddEditMode = SHOW_ADD
   Pt.PRODUCT_TYPE_NAME = "สุกรสำรองแม่"
   Pt.PRODUCT_TYPE_NO = "R"
   Call Pt.AddEditData
   
   Pt.AddEditMode = SHOW_ADD
   Pt.PRODUCT_TYPE_NAME = "สุกรปกติ"
   Pt.PRODUCT_TYPE_NO = "N"
   Call Pt.AddEditData
   
End Sub

Private Sub PopulatePrtItemMap(Yw As CYearWeek)
Dim TempRs1 As ADODB.Recordset
Dim TempRs2 As ADODB.Recordset
Dim Pm As CPrtItemMap
Dim iCount As Long
Dim Pi As CPartItem

   Set TempRs1 = New ADODB.Recordset
   Set TempRs2 = New ADODB.Recordset
   
   Set Yw.PartItemMaps = Nothing
   Set Yw.PartItemMaps = New Collection
   
   Set Pm = New CPrtItemMap
   Pm.PRTITEM_MAP_ID = -1
   Pm.YEAR_WEEK_ID = Yw.YEAR_WEEK_ID
   Call Pm.QueryData(1, TempRs1, iCount)
   Set Pm = Nothing
   While Not TempRs1.EOF
   
      Set Pm = New CPrtItemMap
      Pm.Flag = "I"
      Call Pm.PopulateFromRS(1, TempRs1)
      Call Yw.PartItemMaps.Add(Pm)
                          
       Set Pi = Pm.PartItem
       Pi.PART_ITEM_ID = Pm.PART_ITEM_ID
       Call Pi.QueryData(1, TempRs2, iCount)
       If Not TempRs2.EOF Then
         Call Pi.PopulateFromRS(1, TempRs2)
       End If
       
      Set Pm = Nothing
      
      TempRs1.MoveNext
   Wend
   
   If TempRs1.State = adStateOpen Then
      Call TempRs1.Close
   End If
   Set TempRs1 = Nothing
   
   If TempRs2.State = adStateOpen Then
      Call TempRs2.Close
   End If
   Set TempRs2 = Nothing
End Sub

'===
Public Function QueryYearSeq(Ua As CYearSeq, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long
Dim Exp As CXItem

   RName = "QueryYearSeq"
   QueryYearSeq = False

   IsOK = True
   Call Ua.QueryData(Rs, itemcount)

   If Ua.QueryFlag = 1 Then
      'Name ++++++++++++++++++++++++++
      Dim Yw As CYearWeek
      
      Set Yw = New CYearWeek
      Yw.YEAR_WEEK_ID = -1
      Yw.YEAR_SEQ_ID = Ua.YEAR_SEQ_ID
      Yw.OrderBy = 1
      Call Yw.QueryData(1, m_Rs1, iCount)
      Set Yw = Nothing
      
      Set Ua.YearWeeks = Nothing
      Set Ua.YearWeeks = New Collection
      
      While Not m_Rs1.EOF
         Set Yw = New CYearWeek
         Call Yw.PopulateFromRS(1, m_Rs1)

         Call PopulatePrtItemMap(Yw)

         Yw.Flag = "I"
         Call Ua.YearWeeks.Add(Yw)
         
         Set Yw = Nothing
         m_Rs1.MoveNext
      Wend
   End If
   
   QueryYearSeq = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryYearSeq = False
End Function

Private Sub AddEditPartItemMap(Yw As CYearWeek)
Dim Pm As CPrtItemMap
Dim Pi As CPartItem

   For Each Pm In Yw.PartItemMaps
      Set Pi = Pm.PartItem
      Call Pi.AddEditData
      
      Pm.PART_ITEM_ID = Pi.PART_ITEM_ID
      Pm.YEAR_WEEK_ID = Yw.YEAR_WEEK_ID
      Pm.PRODUCT_TYPE_ID = Pm.PRODUCT_TYPE_ID
      Call Pm.AddEditData
   Next Pm
End Sub

Public Function AddEditYearSeq(Ug As CYearSeq, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditYearSeq"
   AddEditYearSeq = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   Dim Yw As CYearWeek
   For Each Yw In Ug.YearWeeks
      If Yw.Flag = "A" Then
         Yw.YEAR_SEQ_ID = Ug.YEAR_SEQ_ID
         Yw.AddEditMode = SHOW_ADD
         Yw.AddEditData
         
         Call AddEditPartItemMap(Yw)
      ElseIf Yw.Flag = "E" Then
         Yw.AddEditMode = SHOW_EDIT
         Yw.AddEditData
         
         Call AddEditPartItemMap(Yw)
      ElseIf Yw.Flag = "D" Then
         Call AddEditPartItemMap(Yw)
         Call Yw.DeleteData
      End If
   Next Yw

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditYearSeq = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditYearSeq = False
End Function
Public Function AddEditMemoNote(Ug As CMemoNote, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditMemoNote"
   AddEditMemoNote = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditMemoNote = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditMemoNote = False
End Function
Public Function AddEditBalanceAccum(Ug As CBalanceAccum, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditBalanceAccum"
   AddEditBalanceAccum = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.UpdatePriceAmount

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditBalanceAccum = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditBalanceAccum = False
End Function

Public Function DeleteYearSeq(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim itemcount As Long
Dim B As CYearSeq

   RName = "DeleteYearSeq"
   DeleteYearSeq = False
   HasBegin = False

   Set B = New CYearSeq

   If AutoCommit Then
      Call m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.YEAR_SEQ_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteYearSeq = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteYearSeq = True
End Function

'===
Public Function QueryPartItem(Ua As CPartItem, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryPartItem"
   QueryPartItem = False

   IsOK = True
   Call Ua.QueryData(1, Rs, itemcount)

   If Ua.QueryFlag = 1 Then
      Dim Yw As CPartLocation
      
      Set Yw = New CPartLocation
      Yw.PART_LOCATION_ID = -1
      Yw.PART_ITEM_ID = Ua.PART_ITEM_ID
      Yw.LOCATION_ID = -1
      Yw.OrderBy = 1
      Call Yw.QueryData(1, m_Rs1, iCount)
      Set Yw = Nothing
      
      Set Ua.PartLocations = Nothing
      Set Ua.PartLocations = New Collection
      While Not m_Rs1.EOF
         Set Yw = New CPartLocation
         Call Yw.PopulateFromRS(m_Rs1)
                  
         Yw.Flag = "I"
         Call Ua.PartLocations.Add(Yw)
         
         Set Yw = Nothing
         m_Rs1.MoveNext
      Wend
   End If
   
   QueryPartItem = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryPartItem = False
End Function

Public Function DeletePartItem(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim itemcount As Long
Dim B As CPartItem

   RName = "DeletePartItem"
   DeletePartItem = False
   HasBegin = False

   Set B = New CPartItem

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.PART_ITEM_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeletePartItem = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeletePartItem = True
End Function


Public Function AddEditPartItem(Ug As CPartItem, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditPartItem"
   AddEditPartItem = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditPartItem = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditPartItem = False
End Function

Public Function MigrateYearWeek(Ug As CYearWeek, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim Pt As CProductType
Dim ProductTypes As Collection
Dim Pm As CPrtItemMap

   RName = "MigrateYearWeek"
   MigrateYearWeek = False
   HasBegin = False

   Set ProductTypes = New Collection
   Call LoadProductType(Nothing, ProductTypes)
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.QueryData(1, m_Rs1, iCount)
   While Not m_Rs1.EOF
      Call Ug.PopulateFromRS(1, m_Rs1)
      
      For Each Pt In ProductTypes
         Set Pm = New CPrtItemMap
         Pm.AddEditMode = SHOW_ADD
         Pm.YEAR_WEEK_ID = Ug.YEAR_WEEK_ID
         Pm.PRODUCT_TYPE_ID = Pt.PRODUCT_TYPE_ID
         If Pt.PRODUCT_TYPE_NO = "B" Then
            Pm.PART_ITEM_ID = Ug.PART_ITEM_ID1
            If Pm.PART_ITEM_ID > 0 Then
               Call Pm.AddEditData
            End If
         ElseIf Pt.PRODUCT_TYPE_NO = "BT" Then
            Pm.PART_ITEM_ID = Ug.PART_ITEM_ID2
            If Pm.PART_ITEM_ID > 0 Then
               Call Pm.AddEditData
            End If
         ElseIf Pt.PRODUCT_TYPE_NO = "G" Then
            Pm.PART_ITEM_ID = Ug.PART_ITEM_ID3
            If Pm.PART_ITEM_ID > 0 Then
               Call Pm.AddEditData
            End If
         ElseIf Pt.PRODUCT_TYPE_NO = "L" Then
            Pm.PART_ITEM_ID = Ug.PART_ITEM_ID4
            If Pm.PART_ITEM_ID > 0 Then
               Call Pm.AddEditData
            End If
         ElseIf Pt.PRODUCT_TYPE_NO = "R" Then
            Pm.PART_ITEM_ID = Ug.PART_ITEM_ID5
            If Pm.PART_ITEM_ID > 0 Then
               Call Pm.AddEditData
            End If
         ElseIf Pt.PRODUCT_TYPE_NO = "N" Then
            Pm.PART_ITEM_ID = Ug.PART_ITEM_ID6
            If Pm.PART_ITEM_ID > 0 Then
               Call Pm.AddEditData
            End If
         End If
      
         Set Pm = Nothing
      Next Pt
      
      m_Rs1.MoveNext
   Wend
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   Set ProductTypes = Nothing
   MigrateYearWeek = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   MigrateYearWeek = False
End Function

'==
Public Function QueryEnterprise(Ua As CEnterprise, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long

   RName = "QueryEnterprise"
   QueryEnterprise = False
   
   IsOK = True
   
   Call Ua.QueryData(Rs, itemcount)
   
   Dim bIsOk As Boolean
   Dim iCount As Long
   
   If Ua.QueryFlag = 1 Then
      'Contact ++++++++++++++++++++++++++
'      Dim cCstContact As CEnterprisePerson
'      Set cCstContact = New CEnterprisePerson
'      cCstContact.ENTERPRISE_PERSON_ID = -1
'      cCstContact.ENTERPRISE_ID = Ua.ENTERPRISE_ID
'      cCstContact.OrderBy = 1
'      Call cCstContact.QueryData(m_Rs1, iCount)
'      Set cCstContact = Nothing
'
'      Set Ua.EnpPersons = Nothing
'      Set Ua.EnpPersons = New Collection
'      While Not m_Rs1.EOF
'         Set cCstContact = New CEnterprisePerson
'         Call cCstContact.PopulateFromRS(m_Rs1)
'
'         Dim N As cName
'         Set N = New cName
'         N.NAME_ID = cCstContact.NAME_ID
'         Call N.QueryData(m_Rs2, iCount)
'         Set N = Nothing
'         While Not m_Rs2.EOF
'            Set N = New cName
'            Call N.PopulateFromRS(m_Rs2)
'            N.Flag = "I"
'            Set cCstContact.Name = N
'            Set N = Nothing
'
'            m_Rs2.MoveNext
'         Wend
'
'         If cCstContact.ADDRESS_ID > 0 Then
'            Dim A As CAddress
'            Set A = New CAddress
'            A.ADDRESS_ID = cCstContact.ADDRESS_ID
'            Call A.QueryData(m_Rs2, iCount)
'            Set A = Nothing
'            If Not m_Rs2.EOF Then
'               Set A = New CAddress
'               Call A.PopulateFromRS(m_Rs2)
'               A.Flag = "I"
'               Set cCstContact.Address = A
'               Set A = Nothing
'            End If
'         End If
'
'         cCstContact.Flag = "I"
'         Call Ua.EnpPersons.Add(cCstContact)
'         Set cCstContact = Nothing
'         m_Rs1.MoveNext
'      Wend
      'Coontact ++++++++++++++++++++++++++
         
      'Name ++++++++++++++++++++++++++
      Dim cEnpname As CEnterpriseName
      Set cEnpname = New CEnterpriseName
      cEnpname.ENTERPRISE_ID = Ua.ENTERPRISE_ID
      cEnpname.OrderBy = 1
      Call cEnpname.QueryData(m_Rs1, iCount)
      Set cEnpname = Nothing
      
      Set Ua.EnpNames = Nothing
      Set Ua.EnpNames = New Collection
      While Not m_Rs1.EOF
         Set cEnpname = New CEnterpriseName
         Call cEnpname.PopulateFromRS(m_Rs1)

         Dim cName As cName
         Dim iCount2 As Long
         Set cName = New cName
         cName.NAME_ID = cEnpname.NAME_ID
         Call cName.QueryData(m_Rs2, iCount2)
         Set cName = Nothing
         While Not m_Rs2.EOF
            Set cName = New cName
            Call cName.PopulateFromRS(m_Rs2)
            cName.Flag = "I"
            Call cEnpname.Names.Add(cName)
            Set cName = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cEnpname.Flag = "I"
         Call Ua.EnpNames.Add(cEnpname)
         Set cEnpname = Nothing
         m_Rs1.MoveNext
      Wend
      'Name ++++++++++++++++++++++++++
      
      'Address ++++++++++++++++++++++++++
      Dim cEnpAddr As CEnterpriseAddress
      Set cEnpAddr = New CEnterpriseAddress
      cEnpAddr.ENTERPRISE_ID = Ua.ENTERPRISE_ID
      cEnpAddr.OrderBy = 1
      Call cEnpAddr.QueryData(m_Rs1, iCount)
      Set cEnpAddr = Nothing
      
      Set Ua.EnpAddresses = Nothing
      Set Ua.EnpAddresses = New Collection
      While Not m_Rs1.EOF
         Set cEnpAddr = New CEnterpriseAddress
         Call cEnpAddr.PopulateFromRS(m_Rs1)

         Dim cAddr As CAddress
         Dim iCount3 As Long
         Set cAddr = New CAddress
         cAddr.ADDRESS_ID = cEnpAddr.ADDRESS_ID
         Call cAddr.QueryData(m_Rs2, iCount3)
         Set cAddr = Nothing
         While Not m_Rs2.EOF
            Set cAddr = New CAddress
            Call cAddr.PopulateFromRS(m_Rs2)
            cAddr.Flag = "I"
            Call cEnpAddr.Addresses.Add(cAddr)
            Set cAddr = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cEnpAddr.Flag = "I"
         Call Ua.EnpAddresses.Add(cEnpAddr)
         Set cEnpAddr = Nothing
         m_Rs1.MoveNext
      Wend
      'Address ++++++++++++++++++++++++++
   End If

   QueryEnterprise = True
   Exit Function
   
ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryEnterprise = False
End Function

Public Function AddEditEnterprise(Ug As CEnterprise, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim Name As cName
Dim Addr1 As CAddress

   RName = "AddEditEnterprise"
   AddEditEnterprise = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
         
   Dim EnpName As CEnterpriseName
   Dim EmpAddr As CEnterpriseAddress
   Dim Addr As CAddress
   
   For Each EmpAddr In Ug.EnpAddresses
      If EmpAddr.Flag = "A" Then
         Set Addr = EmpAddr.Addresses(1)
         Addr.AddEditMode = SHOW_ADD
         Call Addr.AddEditData
                  
         EmpAddr.ENTERPRISE_ID = Ug.ENTERPRISE_ID
         EmpAddr.ADDRESS_ID = Addr.ADDRESS_ID
         EmpAddr.AddEditMode = SHOW_ADD
         EmpAddr.AddEditData
      ElseIf EmpAddr.Flag = "E" Then
         Set Addr = EmpAddr.Addresses(1)
         Addr.AddEditMode = SHOW_EDIT
         Call Addr.AddEditData
         
      ElseIf EmpAddr.Flag = "D" Then
         Call EmpAddr.DeleteData
         
         Set Addr = EmpAddr.Addresses(1)
         Call Addr.DeleteData
      End If
   Next EmpAddr
   
   For Each EnpName In Ug.EnpNames
      If EnpName.Flag = "A" Then
         Set Name = EnpName.Names(1)
         Name.AddEditMode = SHOW_ADD
         Call Name.AddEditData
                  
         EnpName.ENTERPRISE_ID = Ug.ENTERPRISE_ID
         EnpName.NAME_ID = Name.NAME_ID
         EnpName.AddEditMode = SHOW_ADD
         EnpName.AddEditData
      ElseIf EnpName.Flag = "E" Then
         Set Name = EnpName.Names(1)
         Name.AddEditMode = SHOW_EDIT
         Call Name.AddEditData
      End If
   Next EnpName
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditEnterprise = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditEnterprise = False
End Function

'==
Public Function QueryCustomer(Ua As CCustomer, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long

   RName = "QueryCustomer"
   QueryCustomer = False
   
   IsOK = True
   Call Ua.QueryData2(Rs, itemcount)
   
   Dim bIsOk As Boolean
   Dim iCount As Long
   
   If Ua.QueryFlag = 1 Then
      'Name ++++++++++++++++++++++++++
      Dim cCstName As CCustomerName
      Set cCstName = New CCustomerName
      cCstName.CUSTOMER_NAME_ID = -1
      cCstName.CUSTOMER_ID = Ua.CUSTOMER_ID
      cCstName.OrderBy = 1
      Call cCstName.QueryData(m_Rs1, iCount)
      Set cCstName = Nothing
      
      Set Ua.CstNames = Nothing
      Set Ua.CstNames = New Collection
      While Not m_Rs1.EOF
         Set cCstName = New CCustomerName
         Call cCstName.PopulateFromRS(m_Rs1)

         Dim cName As cName
         Dim iCount2 As Long
         Set cName = New cName
         cName.NAME_ID = cCstName.NAME_ID
         Call cName.QueryData(m_Rs2, iCount2)
         Set cName = Nothing
         While Not m_Rs2.EOF
            Set cName = New cName
            Call cName.PopulateFromRS(m_Rs2)
            cName.Flag = "I"
            Set cCstName.Name = cName
            Set cName = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cCstName.Flag = "I"
         Call Ua.CstNames.Add(cCstName)
         Set cCstName = Nothing
         m_Rs1.MoveNext
      Wend
      'Name ++++++++++++++++++++++++++
            
      'Address ++++++++++++++++++++++++++
      Dim cEnpAddr As CCustomerAddress
      Set cEnpAddr = New CCustomerAddress
      cEnpAddr.CUSTOMER_ID = Ua.CUSTOMER_ID
      cEnpAddr.OrderBy = 1
      Call cEnpAddr.QueryData(m_Rs1, iCount)
      Set cEnpAddr = Nothing
      
      Set Ua.CstAddr = Nothing
      Set Ua.CstAddr = New Collection
      While Not m_Rs1.EOF
         Set cEnpAddr = New CCustomerAddress
         Call cEnpAddr.PopulateFromRS(m_Rs1)

         Dim cAddr As CAddress
         Dim iCount3 As Long
         Set cAddr = New CAddress
         cAddr.ADDRESS_ID = cEnpAddr.ADDRESS_ID
         Call cAddr.QueryData(m_Rs2, iCount3)
         Set cAddr = Nothing
         While Not m_Rs2.EOF
            Set cAddr = New CAddress
            Call cAddr.PopulateFromRS(m_Rs2)
            cAddr.Flag = "I"
            cAddr.ADDRESS_TYPE = cEnpAddr.ADDRESS_TYPE
            Set cEnpAddr.Addresses = cAddr
            Set cAddr = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cEnpAddr.Flag = "I"
         Call Ua.CstAddr.Add(cEnpAddr)
         Set cEnpAddr = Nothing
         m_Rs1.MoveNext
      Wend
      'Address ++++++++++++++++++++++++++
   
'      Account ++++++++++++++++++++++++++
      Dim Acc As CAccount
      Set Acc = New CAccount
      Acc.CUSTOMER_ID = Ua.CUSTOMER_ID
      Acc.OrderBy = 1
      Call Acc.QueryData(1, m_Rs1, iCount)
      Set Acc = Nothing

      Set Ua.CstAccounts = Nothing
      Set Ua.CstAccounts = New Collection
      While Not m_Rs1.EOF
         Set Acc = New CAccount
         Call Acc.PopulateFromRS(1, m_Rs1)

         Acc.Flag = "I"
         Call Ua.CstAccounts.Add(Acc)
         
         Set Acc = Nothing
         m_Rs1.MoveNext
      Wend
'      Account ++++++++++++++++++++++++++
   
'      Pkg ++++++++++++++++++++++++++
      Dim Pkg As CCustomerPackage
      Set Pkg = New CCustomerPackage
      Pkg.CUSTOMER_ID = Ua.CUSTOMER_ID
      Pkg.OrderBy = 1
      Call Pkg.QueryData(1, m_Rs1, iCount)
      Set Pkg = Nothing

      Set Ua.CstPackage = Nothing
      Set Ua.CstPackage = New Collection
      While Not m_Rs1.EOF
         Set Pkg = New CCustomerPackage
         Call Pkg.PopulateFromRS(1, m_Rs1)

         Pkg.Flag = "I"
         Call Ua.CstPackage.Add(Pkg)
         
         Set Pkg = Nothing
         m_Rs1.MoveNext
      Wend
'      Pkg ++++++++++++++++++++++++++
   
   End If
      
   QueryCustomer = True
   Exit Function
   
ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryCustomer = False
End Function

Public Function DeleteCustomer(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cCust As CCustomer

   RName = "DeleteCustomer"
   DeleteCustomer = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Set cCust = New CCustomer
   
   cCust.CUSTOMER_ID = UID
   Call cCust.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCustomer = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False

   DeleteCustomer = True
End Function

Public Function AddEditCustomer(Ug As CCustomer, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCustomer"
   AddEditCustomer = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.AddEditData
      
   Dim CstName As CCustomerName
   Dim Name As cName
   Dim CstAddr As CCustomerAddress
   Dim Addr As CAddress
   
   For Each CstAddr In Ug.CstAddr
      If CstAddr.Flag = "A" Then
         Set Addr = CstAddr.Addresses
         Addr.AddEditMode = SHOW_ADD
         Call Addr.AddEditData

         CstAddr.CUSTOMER_ID = Ug.CUSTOMER_ID
         CstAddr.ADDRESS_ID = Addr.ADDRESS_ID
         CstAddr.ADDRESS_TYPE = Addr.ADDRESS_TYPE
         CstAddr.AddEditMode = SHOW_ADD
         CstAddr.AddEditData
      ElseIf CstAddr.Flag = "E" Then
         Set Addr = CstAddr.Addresses
         Addr.AddEditMode = SHOW_EDIT
         Call Addr.AddEditData

         CstAddr.CUSTOMER_ID = Ug.CUSTOMER_ID
         CstAddr.ADDRESS_ID = Addr.ADDRESS_ID
         CstAddr.AddEditMode = SHOW_EDIT
         CstAddr.ADDRESS_TYPE = Addr.ADDRESS_TYPE
         CstAddr.AddEditData
      ElseIf CstAddr.Flag = "D" Then
         Call CstAddr.DeleteData

         Set Addr = CstAddr.Addresses
         Call Addr.DeleteData
      End If
   Next CstAddr
   
   For Each CstName In Ug.CstNames
      If CstName.Flag = "A" Then
         Set Name = CstName.Name
         Name.AddEditMode = SHOW_ADD
         Call Name.AddEditData
                  
         CstName.CUSTOMER_ID = Ug.CUSTOMER_ID
         CstName.NAME_ID = Name.NAME_ID
         CstName.AddEditMode = SHOW_ADD
         CstName.AddEditData
      ElseIf CstName.Flag = "E" Then
         Set Name = CstName.Name
         Name.AddEditMode = SHOW_EDIT
         Call Name.AddEditData
      End If
   Next CstName
   
   Dim Acc As CAccount
   For Each Acc In Ug.CstAccounts
      If Acc.Flag = "A" Then
         Acc.CUSTOMER_ID = Ug.CUSTOMER_ID
         Acc.AddEditMode = SHOW_ADD
         Acc.AddEditData
      ElseIf Acc.Flag = "E" Then
         Acc.AddEditMode = SHOW_EDIT
         Acc.AddEditData
      ElseIf Acc.Flag = "D" Then
         Call Acc.DeleteData
      End If
   Next Acc
      
      
   Dim Pkg As CCustomerPackage
   For Each Pkg In Ug.CstPackage
      If Pkg.Flag = "A" Then
         Pkg.CUSTOMER_ID = Ug.CUSTOMER_ID
         Pkg.AddEditMode = SHOW_ADD
         Pkg.AddEditData
      ElseIf Pkg.Flag = "E" Then
         Pkg.AddEditMode = SHOW_EDIT
         Pkg.AddEditData
      ElseIf Pkg.Flag = "D" Then
         Call Pkg.DeleteData
      End If
   Next Pkg
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   Call LoadCustomerPackage(Nothing, CustomerPackage)
   Call LoadPackageDetail(Nothing, PackageDetail)
   
   HasBegin = False
   IsOK = True

   AddEditCustomer = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCustomer = False
End Function

'==
Public Function QuerySupplier(Ua As CSupplier, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long

   RName = "QuerySupplier"
   QuerySupplier = False
   
   IsOK = True
   Call Ua.QueryData2(Rs, itemcount)
   
   Dim bIsOk As Boolean
   Dim iCount As Long
   
   If Ua.QueryFlag = 1 Then
      'Name ++++++++++++++++++++++++++
      Dim cCstName As CSupplierName
      Set cCstName = New CSupplierName
      cCstName.SUPPLIER_NAME_ID = -1
      cCstName.SUPPLIER_ID = Ua.SUPPLIER_ID
      cCstName.OrderBy = 1
      Call cCstName.QueryData(m_Rs1, iCount)
      Set cCstName = Nothing
      
      Set Ua.CstNames = Nothing
      Set Ua.CstNames = New Collection
      While Not m_Rs1.EOF
         Set cCstName = New CSupplierName
         Call cCstName.PopulateFromRS(m_Rs1)

         Dim cName As cName
         Dim iCount2 As Long
         Set cName = New cName
         cName.NAME_ID = cCstName.NAME_ID
         Call cName.QueryData(m_Rs2, iCount2)
         Set cName = Nothing
         While Not m_Rs2.EOF
            Set cName = New cName
            Call cName.PopulateFromRS(m_Rs2)
            cName.Flag = "I"
            Set cCstName.Name = cName
            Set cName = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cCstName.Flag = "I"
         Call Ua.CstNames.Add(cCstName)
         Set cCstName = Nothing
         m_Rs1.MoveNext
      Wend
      'Name ++++++++++++++++++++++++++
            
      'Address ++++++++++++++++++++++++++
      Dim cEnpAddr As CSupplierAddress
      Set cEnpAddr = New CSupplierAddress
      cEnpAddr.SUPPLIER_ID = Ua.SUPPLIER_ID
      cEnpAddr.OrderBy = 1
      Call cEnpAddr.QueryData(m_Rs1, iCount)
      Set cEnpAddr = Nothing
      
      Set Ua.CstAddr = Nothing
      Set Ua.CstAddr = New Collection
      While Not m_Rs1.EOF
         Set cEnpAddr = New CSupplierAddress
         Call cEnpAddr.PopulateFromRS(m_Rs1)

         Dim cAddr As CAddress
         Dim iCount3 As Long
         Set cAddr = New CAddress
         cAddr.ADDRESS_ID = cEnpAddr.ADDRESS_ID
         Call cAddr.QueryData(m_Rs2, iCount3)
         Set cAddr = Nothing
         While Not m_Rs2.EOF
            Set cAddr = New CAddress
            Call cAddr.PopulateFromRS(m_Rs2)
            cAddr.Flag = "I"
            cAddr.ADDRESS_TYPE = cEnpAddr.ADDRESS_TYPE
            Set cEnpAddr.Addresses = cAddr
            Set cAddr = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cEnpAddr.Flag = "I"
         Call Ua.CstAddr.Add(cEnpAddr)
         Set cEnpAddr = Nothing
         m_Rs1.MoveNext
      Wend
      'Address ++++++++++++++++++++++++++
   End If
   
   QuerySupplier = True
   Exit Function
   
ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QuerySupplier = False
End Function

Public Function AddEditSupplier(Ug As CSupplier, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditSupplier"
   AddEditSupplier = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.AddEditData
   
   Dim CstName As CSupplierName
   Dim Name As cName
   Dim CstAddr As CSupplierAddress
   Dim Addr As CAddress
   
   
   For Each CstAddr In Ug.CstAddr
      If CstAddr.Flag = "A" Then
         Set Addr = CstAddr.Addresses
         Addr.AddEditMode = SHOW_ADD
         Call Addr.AddEditData

         CstAddr.SUPPLIER_ID = Ug.SUPPLIER_ID
         CstAddr.ADDRESS_ID = Addr.ADDRESS_ID
         CstAddr.ADDRESS_TYPE = Addr.ADDRESS_TYPE
         CstAddr.AddEditMode = SHOW_ADD
         CstAddr.AddEditData
      ElseIf CstAddr.Flag = "E" Then
         Set Addr = CstAddr.Addresses
         Addr.AddEditMode = SHOW_EDIT
         Call Addr.AddEditData

         CstAddr.SUPPLIER_ID = Ug.SUPPLIER_ID
         CstAddr.ADDRESS_ID = Addr.ADDRESS_ID
         CstAddr.AddEditMode = SHOW_EDIT
         CstAddr.ADDRESS_TYPE = Addr.ADDRESS_TYPE
         CstAddr.AddEditData
      ElseIf CstAddr.Flag = "D" Then
         Call CstAddr.DeleteData

         Set Addr = CstAddr.Addresses
         Call Addr.DeleteData
      End If
   Next CstAddr
   
   For Each CstName In Ug.CstNames
      If CstName.Flag = "A" Then
         Set Name = CstName.Name
         Name.AddEditMode = SHOW_ADD
         Call Name.AddEditData
                  
         CstName.SUPPLIER_ID = Ug.SUPPLIER_ID
         CstName.NAME_ID = Name.NAME_ID
         CstName.AddEditMode = SHOW_ADD
         CstName.AddEditData
      ElseIf CstName.Flag = "E" Then
         Set Name = CstName.Name
         Name.AddEditMode = SHOW_EDIT
         Call Name.AddEditData
      End If
   Next CstName
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditSupplier = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditSupplier = False
End Function

Public Function DeleteSupplier(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cCust As CSupplier

   RName = "DeleteSupplier"
   DeleteSupplier = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Set cCust = New CSupplier
   
   cCust.SUPPLIER_ID = UID
   Call cCust.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteSupplier = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False

   DeleteSupplier = True
End Function

'==
Public Function QueryEmployee(Ua As CEmployee, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

Dim Qp As CEmployee

   RName = "QueryEmployee"
   QueryEmployee = False

   IsOK = True
   Call Ua.QueryData(Rs, itemcount)

   If Ua.QueryFlag = 1 Then
      Set Ua.EmpName = Nothing
      Set Ua.EmpName = New CEmployeeName
      Ua.EmpName.EMPLOYEE_NAME_ID = -1
      Ua.EmpName.EMP_ID = Ua.EMP_ID
      Call Ua.EmpName.QueryData(m_Rs1, iCount)
      If Not m_Rs1.EOF Then
         Call Ua.EmpName.PopulateFromRS(m_Rs1)
      End If
      
      Set Ua.EName = Nothing
      Set Ua.EName = New cName
      Ua.EName.NAME_ID = Ua.EmpName.NAME_ID
      Call Ua.EName.QueryData(m_Rs1, iCount)
      If Not m_Rs1.EOF Then
         Call Ua.EName.PopulateFromRS(m_Rs1)
      End If
   End If
   
   QueryEmployee = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryEmployee = False
End Function

Public Function DeleteEmployee(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim itemcount As Long

Dim B As CEmployee

   RName = "DeleteEmployee"
   DeleteEmployee = False
   HasBegin = False

   Set B = New CEmployee

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.EMP_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteEmployee = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteEmployee = True
End Function

Public Function AddEditEmployee(Ug As CEmployee, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim Addr As CAddress

   RName = "AddEditEmployee"
   AddEditEmployee = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   Call Ug.EName.AddEditData
   
   Ug.EmpName.EMP_ID = Ug.EMP_ID
   Ug.EmpName.NAME_ID = Ug.EName.NAME_ID
   Call Ug.EmpName.AddEditData
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditEmployee = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditEmployee = False
End Function

'==
Public Function QueryInventoryDoc(Ua As CInventoryDoc, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryInventoryDoc"
   QueryInventoryDoc = False

   IsOK = True
   Call Ua.QueryData(Rs, itemcount)

   If Ua.QueryFlag = 1 Then
      Dim Yw As CImportItem
      
      Set Yw = New CImportItem
      Yw.BATCH_ID = Ua.BATCH_ID
      Yw.IMPORT_ITEM_ID = -1
      Yw.INVENTORY_DOC_ID = Ua.INVENTORY_DOC_ID
      Yw.OrderBy = 1
      Call Yw.QueryData(1, m_Rs1, iCount)
      Set Yw = Nothing
      
      Set Ua.ImportItems = Nothing
      Set Ua.ImportItems = New Collection
      
      Set Ua.ExportItems = Nothing
      Set Ua.ExportItems = New Collection
      
      Set Ua.TransferItems = Nothing
      Set Ua.TransferItems = New Collection

      Set Ua.ImportExports = Nothing
      Set Ua.ImportExports = New Collection
      While Not m_Rs1.EOF
         Set Yw = New CImportItem
         Call Yw.PopulateFromRS(1, m_Rs1)
                  
         Yw.Flag = "I"
         Call Ua.ImportExports.Add(Yw)
         
         Set Yw = Nothing
         m_Rs1.MoveNext
      Wend
   
      Dim EI As CExportItem
      
      Set EI = New CExportItem
      EI.BATCH_ID = Ua.BATCH_ID
      EI.EXPORT_ITEM_ID = -1
      EI.INVENTORY_DOC_ID = Ua.INVENTORY_DOC_ID
      EI.OrderBy = 1
      Call EI.QueryData(1, m_Rs1, iCount)
      Set EI = Nothing

      While Not m_Rs1.EOF
         Set EI = New CExportItem
         Call EI.PopulateFromRS(1, m_Rs1)

         EI.Flag = "I"
         Call Ua.ImportExports.Add(EI)

         Set EI = Nothing
         m_Rs1.MoveNext
      Wend
   End If
   
   QueryInventoryDoc = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryInventoryDoc = False
End Function

Public Function DeleteInventoryDoc(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim itemcount As Long
Dim B As CInventoryDoc

   RName = "DeleteInventoryDoc"
   DeleteInventoryDoc = False
   HasBegin = False

   Set B = New CInventoryDoc

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.INVENTORY_DOC_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True
   
   DeleteInventoryDoc = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteInventoryDoc = True
End Function

Public Function AddEditInventoryDoc1(Ug As CInventoryDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim PL As CPartLocation
Dim PL2 As CPartLocation
Dim iCount As Long
Dim NewPrice As Double
Dim O As Object
Dim GuiID As Long

   RName = "AddEditInventoryDoc1"
   AddEditInventoryDoc1 = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   Dim Yw As CImportItem
   For Each O In Ug.ImportExports
      
      If O.TX_TYPE = "I" Then 'Import
         Set Yw = O
         If Yw.Flag = "A" Then
            Set PL = New CPartLocation
            PL.PART_LOCATION_ID = -1
            PL.PART_ITEM_ID = Yw.PART_ITEM_ID
            PL.LOCATION_ID = Yw.LOCATION_ID
            Call PL.QueryData(1, m_Rs2, iCount)
            If m_Rs2.EOF Then
               PL.AddEditMode = SHOW_ADD
               Call PL.AddEditData
            Else
               Call PL.PopulateFromRS(m_Rs2)
            End If
                                       
            If (PL.CURRENT_AMOUNT + Yw.IMPORT_AMOUNT) > 0 Then
               NewPrice = (PL.CURRENT_AMOUNT * PL.AVG_PRICE + Yw.INCLUDE_UNIT_PRICE * Yw.IMPORT_AMOUNT) / (PL.CURRENT_AMOUNT + Yw.IMPORT_AMOUNT)
            Else
               NewPrice = 0
            End If

'            Yw.CURRENT_AMOUNT = PL.CURRENT_AMOUNT + Yw.IMPORT_AMOUNT
            If Yw.CALCULATE_FLAG = "Y" Then
'               Yw.NEW_PRICE = NewPrice
            Else
'               Yw.NEW_PRICE = PL.AVG_PRICE
            End If
'            Yw.PREVIOUS_AMOUNT = PL.CURRENT_AMOUNT
'            Yw.PREVIOUS_PRICE = PL.AVG_PRICE
            Yw.INVENTORY_DOC_ID = Ug.INVENTORY_DOC_ID
            Yw.GUI_ID = GuiID
            Yw.AddEditMode = SHOW_ADD
            Call Yw.AddEditData
            
            PL.AddEditMode = SHOW_EDIT
            If Yw.CALCULATE_FLAG = "Y" Then
               PL.AVG_PRICE = NewPrice
               PL.LAST_PRICE = Yw.INCLUDE_UNIT_PRICE
            End If
            PL.CURRENT_AMOUNT = PL.CURRENT_AMOUNT + Yw.IMPORT_AMOUNT
            If Ug.COMMIT_FLAG = "Y" Then
               Call PL.AddEditData
            End If
            
            Set PL = Nothing
         ElseIf Yw.Flag = "E" Then
            Set PL = New CPartLocation
            PL.PART_LOCATION_ID = -1
            PL.PART_ITEM_ID = Yw.PART_ITEM_ID
            PL.LOCATION_ID = Yw.LOCATION_ID
            Call PL.QueryData(1, m_Rs2, iCount)
            If m_Rs2.EOF Then
               PL.AddEditMode = SHOW_ADD
               Call PL.AddEditData
            Else
               Call PL.PopulateFromRS(m_Rs2)
            End If
                                       
            If (PL.CURRENT_AMOUNT + Yw.IMPORT_AMOUNT) > 0 Then
               NewPrice = (PL.CURRENT_AMOUNT * PL.AVG_PRICE + Yw.INCLUDE_UNIT_PRICE * Yw.IMPORT_AMOUNT) / (PL.CURRENT_AMOUNT + Yw.IMPORT_AMOUNT)
            Else
               NewPrice = 0
            End If

'            Yw.CURRENT_AMOUNT = PL.CURRENT_AMOUNT + Yw.IMPORT_AMOUNT
            If Yw.CALCULATE_FLAG = "Y" Then
'               Yw.NEW_PRICE = NewPrice
            Else
'               Yw.NEW_PRICE = PL.AVG_PRICE
            End If
'            Yw.PREVIOUS_AMOUNT = PL.CURRENT_AMOUNT
'            Yw.PREVIOUS_PRICE = PL.AVG_PRICE
            Yw.INVENTORY_DOC_ID = Ug.INVENTORY_DOC_ID
            Yw.AddEditMode = SHOW_EDIT
            Call Yw.AddEditData
            
            PL.AddEditMode = SHOW_EDIT
            If Yw.CALCULATE_FLAG = "Y" Then
               PL.AVG_PRICE = NewPrice
               PL.LAST_PRICE = Yw.INCLUDE_UNIT_PRICE
            End If
            PL.CURRENT_AMOUNT = PL.CURRENT_AMOUNT + Yw.IMPORT_AMOUNT
            If Ug.COMMIT_FLAG = "Y" Then
               Call PL.AddEditData
            End If
            
            Set PL = Nothing
         ElseIf Yw.Flag = "D" Then
            Call Yw.DeleteData
         End If
      ElseIf O.TX_TYPE = "E" Then 'Export
         Dim EI As CExportItem
         
         Set EI = O
         If EI.Flag = "A" Then
            Call glbDatabaseMngr.GetSeqID("GUI_ID_SEQ", GuiID, ErrorObj)
            
            Set PL2 = New CPartLocation
            If (EI.CALCULATE_FLAG = "Y") And (EI.PIG_ID > 0) Then
               PL2.PART_LOCATION_ID = -1
               PL2.PART_ITEM_ID = EI.PIG_ID
               PL2.LOCATION_ID = EI.HOUSE_ID
               Call PL2.QueryData(1, m_Rs2, iCount)
               If Not m_Rs2.EOF Then
                  Call PL2.PopulateFromRS(m_Rs2)
               End If
            Else
               EI.PIG_ID = -1
               EI.HOUSE_ID = -1
               PL2.CURRENT_AMOUNT = 0
            End If
            
            Set PL = New CPartLocation
            PL.PART_LOCATION_ID = -1
            PL.PART_ITEM_ID = EI.PART_ITEM_ID
            PL.LOCATION_ID = EI.LOCATION_ID
            Call PL.QueryData(1, m_Rs2, iCount)
            If Not m_Rs2.EOF Then
               Call PL.PopulateFromRS(m_Rs2)
            End If
   
            EI.PIG_IN_HOUSE = PL2.CURRENT_AMOUNT
'            EI.EXPORT_AVG_PRICE = PL.AVG_PRICE
'            EI.EXPORT_LAST_PRICE = PL.LAST_PRICE
'            EI.PREVIOUS_AMOUNT = PL.CURRENT_AMOUNT
'            EI.CURRENT_AMOUNT = PL.CURRENT_AMOUNT - EI.EXPORT_AMOUNT
            EI.INVENTORY_DOC_ID = Ug.INVENTORY_DOC_ID
            EI.GUI_ID = GuiID
            EI.AddEditMode = SHOW_ADD
            Call EI.AddEditData
            
            PL.AddEditMode = SHOW_EDIT
            PL.CURRENT_AMOUNT = PL.CURRENT_AMOUNT - EI.EXPORT_AMOUNT
            If Ug.COMMIT_FLAG = "Y" Then
               Call PL.AddEditData
            End If
   
            Set PL = Nothing
            Set PL2 = Nothing
         ElseIf EI.Flag = "E" Then
            Set PL2 = New CPartLocation
            If (EI.CALCULATE_FLAG = "Y") And (EI.PIG_ID > 0) Then
               PL2.PART_LOCATION_ID = -1
               PL2.PART_ITEM_ID = EI.PIG_ID
               PL2.LOCATION_ID = EI.HOUSE_ID
               Call PL2.QueryData(1, m_Rs2, iCount)
               If Not m_Rs2.EOF Then
                  Call PL2.PopulateFromRS(m_Rs2)
               End If
            Else
               EI.PIG_ID = -1
               EI.HOUSE_ID = -1
               PL2.CURRENT_AMOUNT = 0
            End If
            
            Set PL = New CPartLocation
            PL.PART_LOCATION_ID = -1
            PL.PART_ITEM_ID = EI.PART_ITEM_ID
            PL.LOCATION_ID = EI.LOCATION_ID
            Call PL.QueryData(1, m_Rs2, iCount)
            If Not m_Rs2.EOF Then
               Call PL.PopulateFromRS(m_Rs2)
            End If
   
            EI.PIG_IN_HOUSE = PL2.CURRENT_AMOUNT
'            EI.EXPORT_AVG_PRICE = PL.AVG_PRICE
'            EI.EXPORT_LAST_PRICE = PL.LAST_PRICE
'            EI.PREVIOUS_AMOUNT = PL.CURRENT_AMOUNT
'            EI.CURRENT_AMOUNT = PL.CURRENT_AMOUNT - EI.EXPORT_AMOUNT
            EI.INVENTORY_DOC_ID = Ug.INVENTORY_DOC_ID
            EI.AddEditMode = SHOW_EDIT
            Call EI.AddEditData
            
            PL.AddEditMode = SHOW_EDIT
            PL.CURRENT_AMOUNT = PL.CURRENT_AMOUNT - EI.EXPORT_AMOUNT
            If Ug.COMMIT_FLAG = "Y" Then
               Call PL.AddEditData
            End If
            
            Set PL = Nothing
            Set PL2 = Nothing
         ElseIf EI.Flag = "D" Then
            Call EI.DeleteData
         End If
      End If
   Next O

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditInventoryDoc1 = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditInventoryDoc1 = False
End Function

Public Function AddEditInventoryDoc(Ug As CInventoryDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim PL As CPartLocation
Dim PL2 As CPartLocation
Dim iCount As Long
Dim NewPrice As Double
Dim O As Object
Dim GuiID As Long

   RName = "AddEditInventoryDoc"
   AddEditInventoryDoc = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
   
   Dim Cd As CConfigDoc
   If Ug.AddEditMode = SHOW_ADD And Ug.CONFIG_DOC_TYPE > 0 Then
      Set Cd = New CConfigDoc
      Call Cd.SetFieldValue("RUNNING_NO", Ug.RUNNING_NO)
      Call Cd.SetFieldValue("LAST_NO", Ug.DOCUMENT_NO)
      Call Cd.SetFieldValue("CONFIG_DOC_TYPE", Ug.CONFIG_DOC_TYPE)
      Call Cd.UpdateRunningNo
   End If
   Set Cd = Nothing
   
   Dim Yw As CImportItem
   For Each O In Ug.ImportExports
   
      If O.TX_TYPE = "I" Then 'Import
         Set Yw = O
         If Yw.Flag = "A" Then
            Yw.INVENTORY_DOC_ID = Ug.INVENTORY_DOC_ID
            Yw.GUI_ID = GuiID
            Yw.AddEditMode = SHOW_ADD
            Call Yw.AddEditData
         ElseIf Yw.Flag = "E" Then
            Yw.INVENTORY_DOC_ID = Ug.INVENTORY_DOC_ID
            Yw.AddEditMode = SHOW_EDIT
            Call Yw.AddEditData
         ElseIf Yw.Flag = "D" Then
            Call Yw.DeleteData
         End If
      ElseIf O.TX_TYPE = "E" Then 'Export
         Dim EI As CExportItem
         
         Set EI = O
         If EI.Flag = "A" Then
            Call glbDatabaseMngr.GetSeqID("GUI_ID_SEQ", GuiID, ErrorObj)
            
            EI.INVENTORY_DOC_ID = Ug.INVENTORY_DOC_ID
            If Ug.DOCUMENT_TYPE = 2 Then
               EI.PIG_AGE = GetAge(EI.PIG_NO, Ug.DOCUMENT_DATE)
               EI.AGE_CODE = GetAgeCode(EI.PIG_AGE)
            End If
            EI.GUI_ID = GuiID
            EI.AddEditMode = SHOW_ADD
            Call EI.AddEditData
            Call EI.PatchAvgPrice(EI.EXPORT_AVG_PRICE, EI.PREVIOUS_AMOUNT, EI.CURRENT_AMOUNT, EI.EXPORT_TOTAL_PRICE)
         ElseIf EI.Flag = "E" Or (EI.Flag = "I" And Ug.DOCUMENT_TYPE = 2 And EI.PIG_AGE <> GetAge(EI.PIG_NO, Ug.DOCUMENT_DATE)) Then
            'กรณีที่ Ug.DOCUMENT_TYPE = 2 แล้วมีการ SAVE เปลี่ยนวันที่เอกสารมันจะไม่เข้ามาแก้ไข ซึ่งจะทำให้ผิด    21/06/2559
            If Ug.DOCUMENT_TYPE = 2 Then
               EI.PIG_AGE = GetAge(EI.PIG_NO, Ug.DOCUMENT_DATE)
               EI.AGE_CODE = GetAgeCode(EI.PIG_AGE)
            End If
            
            EI.INVENTORY_DOC_ID = Ug.INVENTORY_DOC_ID
            EI.AddEditMode = SHOW_EDIT
            Call EI.AddEditData
            Call EI.PatchAvgPrice(EI.EXPORT_AVG_PRICE, EI.PREVIOUS_AMOUNT, EI.CURRENT_AMOUNT, EI.EXPORT_TOTAL_PRICE)
         ElseIf EI.Flag = "D" Then
            Call EI.DeleteData
         End If
      End If
   Next O

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditInventoryDoc = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditInventoryDoc = False
End Function

'==
'ตรวจสอบยอดวัตถุดิบจากตาราง PART_LOCATION
Public Function VerifyStockBalance(Col2 As Collection, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim O As Object
Dim iCount As Long
Dim PL As CPartLocation
Dim PartLocations As Collection
Dim FoundFlag As Boolean

   RName = "VerifyStockBalance"
   VerifyStockBalance = False
   HasBegin = False

   Set PartLocations = New Collection
   
   Set PL = New CPartLocation
   PL.PART_LOCATION_ID = -1
   PL.PART_ITEM_ID = -1
   PL.LOCATION_ID = -1
   Call PL.QueryData(1, m_Rs1, iCount)
   Set PL = Nothing
   While Not m_Rs1.EOF
      Set PL = New CPartLocation
      Call PL.PopulateFromRS(m_Rs1)
      
      Call PartLocations.Add(PL)
      
      Set PL = Nothing
      m_Rs1.MoveNext
   Wend
   Set PL = Nothing
   
   FoundFlag = False
   If Not (Col2 Is Nothing) Then
      For Each O In Col2
         If O.TX_TYPE = "E" Then
            For Each PL In PartLocations
               If (PL.PART_ITEM_ID = O.PART_ITEM_ID) And (PL.LOCATION_ID = O.LOCATION_ID) Then
                  If O.TX_TYPE = "E" Then
                     PL.CURRENT_AMOUNT = PL.CURRENT_AMOUNT - O.EXPORT_AMOUNT
                  Else
                     PL.CURRENT_AMOUNT = PL.CURRENT_AMOUNT + O.IMPORT_AMOUNT
                  End If
                  FoundFlag = True
                  
                  If PL.CURRENT_AMOUNT < 0 Then
                     FoundFlag = False
                  End If
                  Exit For
               End If
            Next PL
         Else
            FoundFlag = True
         End If
                  
         If Not FoundFlag Then
            glbErrorLog.LocalErrorMsg = MapText("มีปริมาณ ") & O.PART_DESC & " " & MapText(" ไม่เพียงพอ ที่คลัง ") & O.LOCATION_NAME
            glbErrorLog.ShowUserError
            
            VerifyStockBalance = False
            Exit Function
         End If
      Next O
   End If
   
   Set PartLocations = Nothing
   VerifyStockBalance = True
   Exit Function

ErrorHandler:
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   VerifyStockBalance = False
End Function

'==
'ตรวจสอบยอดวัตถุดิบจาก เข้า - ออก (เข้าลบออก)
Public Function VerifyStockBalanceEx(Col2 As Collection, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim O As Object
Dim iCount As Long
Dim PL As CPartLocation
Dim PartLocations As Collection
Dim FoundFlag As Boolean
Dim II As CImportItem
Dim Ii2 As CImportItem
Dim EI2 As CExportItem
Dim ImportItems As Collection
Dim ExportItems As Collection

   RName = "VerifyStockBalanceEx"
   VerifyStockBalanceEx = False
   HasBegin = False

   Set PartLocations = New Collection
   Set ImportItems = New Collection
   Call LoadPigHouseStatusImportAmount(Nothing, ImportItems, , , "Y")
   Set ExportItems = New Collection
   Call LoadPigHouseStatusExportAmount(Nothing, ExportItems, , , "Y")
   
   Set II = New CImportItem
   II.IMPORT_ITEM_ID = -1
   Call II.QueryData(19, m_Rs1, iCount)
   While Not m_Rs1.EOF
      Call II.PopulateFromRS(19, m_Rs1)
      
      Set Ii2 = GetImportItem(ImportItems, II.LOCATION_ID & "-" & II.PART_ITEM_ID & "-" & II.PIG_STATUS)
      Set EI2 = GetExportItem(ExportItems, II.LOCATION_ID & "-" & II.PART_ITEM_ID & "-" & II.PIG_STATUS)
      
      Set PL = New CPartLocation
      PL.LOCATION_ID = II.LOCATION_ID
      PL.PART_ITEM_ID = II.PART_ITEM_ID
      PL.PIG_STATUS = II.PIG_STATUS
      PL.CURRENT_AMOUNT = Ii2.IMPORT_AMOUNT - EI2.EXPORT_AMOUNT
      Call PartLocations.Add(PL)
      Set PL = Nothing
      
      m_Rs1.MoveNext
   Wend
   Set II = Nothing
   
   FoundFlag = False
   If Not (Col2 Is Nothing) Then
      For Each O In Col2
         If O.TX_TYPE = "E" Then
            For Each PL In PartLocations
               If (PL.PART_ITEM_ID = O.PART_ITEM_ID) And (PL.LOCATION_ID = O.LOCATION_ID) And (PL.PIG_STATUS = O.PIG_STATUS) Then
                  If O.TX_TYPE = "E" Then
                     PL.CURRENT_AMOUNT = PL.CURRENT_AMOUNT - O.EXPORT_AMOUNT
                  Else
                     PL.CURRENT_AMOUNT = PL.CURRENT_AMOUNT + O.IMPORT_AMOUNT
                  End If
                  FoundFlag = True
                  
                  If PL.CURRENT_AMOUNT < 0 Then
                     FoundFlag = False
                  End If
                  Exit For
               End If
            Next PL
         Else
            FoundFlag = True
         End If
   
         If Not FoundFlag Then
            glbErrorLog.LocalErrorMsg = MapText("มีปริมาณ ") & O.PART_DESC & " " & MapText(" ไม่เพียงพอ ที่คลัง ") & O.LOCATION_NAME
            glbErrorLog.ShowUserError
            
            VerifyStockBalanceEx = False
            Exit Function
         End If
      Next O
   End If
   
   Set ImportItems = Nothing
   Set ExportItems = Nothing
   Set PartLocations = Nothing
   VerifyStockBalanceEx = True
   Exit Function

ErrorHandler:
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   VerifyStockBalanceEx = False
End Function

Private Function CompareKey(D1 As Object, D2 As Object) As Boolean
Dim TempResult As Boolean

   TempResult = D1.TRANSACTION_SEQ < D2.TRANSACTION_SEQ
         
   CompareKey = TempResult
End Function

Private Sub Selectionsort(List As Collection, MIN As Long, MAX As Long)
Dim I As Long
Dim j As Long
Dim best_value As Object
Dim Temp As Object
Dim best_j As Integer

    For I = MIN To MAX - 1
        Set best_value = List(I)
        best_j = I
        For j = I + 1 To MAX
            If CompareKey(List(j), best_value) Then
                Set best_value = List(j)
                best_j = j
            End If
        Next j
        
        Set Temp = List(I)
        List.Remove (best_j)
        If best_j > List.Count Then
         Call List.Add(Temp, , , best_j - 1)
      Else
         Call List.Add(Temp, , best_j)
      End If
    
        List.Remove (I)
        If I > List.Count Then
         Call List.Add(best_value, , , I - 1)
      Else
         Call List.Add(best_value, , I)
      End If
    Next I
    Set best_value = Nothing
End Sub

Public Sub MergeImportExport(Imps As Collection, Exps As Collection, ImpExps As Collection)
Dim Count As Long
Dim IP As CImportItem
Dim Exp As CExportItem
Dim O As Object
   
   For Each IP In Imps
      Call ImpExps.Add(IP)
   Next IP

   For Each Exp In Exps
      Call ImpExps.Add(Exp)
   Next Exp
   
   Call Selectionsort(ImpExps, 1, ImpExps.Count)
End Sub

Private Function CompareKeyEx(D1 As Object, D2 As Object) As Boolean
Dim TempResult As Boolean

   If D1.PART_TYPE_NO = D2.PART_TYPE_NO Then
      TempResult = D1.PART_NO < D2.PART_NO
   Else
      TempResult = D1.PART_TYPE_NO < D2.PART_TYPE_NO
   End If
         
   CompareKeyEx = TempResult
End Function

Private Function PigCompareKeyEx(D1 As Object, D2 As Object) As Boolean
Dim TempResult As Boolean

   TempResult = D1.PART_NO < D2.PART_NO
         
   PigCompareKeyEx = TempResult
End Function

Public Sub SelectionsortEx(List As Collection, MIN As Long, MAX As Long)
Dim I As Long
Dim j As Long
Dim best_value As Object
Dim Temp As Object
Dim best_j As Integer

    For I = MIN To MAX - 1
        Set best_value = List(I)
        best_j = I
        For j = I + 1 To MAX
            If CompareKeyEx(List(j), best_value) Then
                Set best_value = List(j)
                best_j = j
            End If
        Next j
        
        Set Temp = List(I)
        List.Remove (best_j)
        If best_j > List.Count Then
         Call List.Add(Temp, , , best_j - 1)
      Else
         Call List.Add(Temp, , best_j)
      End If
    
        List.Remove (I)
        If I > List.Count Then
         Call List.Add(best_value, , , I - 1)
      Else
         Call List.Add(best_value, , I)
      End If
    Next I
    Set best_value = Nothing
End Sub

Public Sub PigSelectionsortEx(List As Collection, MIN As Long, MAX As Long)
Dim I As Long
Dim j As Long
Dim best_value As Object
Dim Temp As Object
Dim best_j As Integer

    For I = MIN To MAX - 1
        Set best_value = List(I)
        best_j = I
        For j = I + 1 To MAX
            If PigCompareKeyEx(List(j), best_value) Then
                Set best_value = List(j)
                best_j = j
            End If
        Next j
        
        Set Temp = List(I)
        List.Remove (best_j)
        If best_j > List.Count Then
         Call List.Add(Temp, , , best_j - 1)
      Else
         Call List.Add(Temp, , best_j)
      End If
    
        List.Remove (I)
        If I > List.Count Then
         Call List.Add(best_value, , , I - 1)
      Else
         Call List.Add(best_value, , I)
      End If
    Next I
    Set best_value = Nothing
End Sub
Public Function DateToPartItem(D1 As Date, PigTypeID As Long) As CPartItem
Dim Count As Long
Dim Pw As CYearWeek
Dim iCount As Long
Dim p As CPartItem
Dim Pm As CPrtItemMap
   
   Set Pw = New CYearWeek
   Pw.YEAR_WEEK_ID = -1
   Pw.YEAR_SEQ_ID = -1
   
   Call Pw.QueryData(1, m_Rs1, iCount)
   Set Pw = Nothing
   While Not m_Rs1.EOF
      If (DateToStringInt(D1) >= NVLS(m_Rs1("FROM_DATE"), "")) And (DateToStringInt(D1) <= NVLS(m_Rs1("TO_DATE"), "")) Then
      
         Set Pm = New CPrtItemMap
         Pm.PRTITEM_MAP_ID = -1
         Pm.YEAR_WEEK_ID = m_Rs1("YEAR_WEEK_ID")
         Call Pm.QueryData(1, m_Rs2, iCount)
         While Not m_Rs2.EOF
         
            Call Pm.PopulateFromRS(1, m_Rs2)
            If Pm.PRODUCT_TYPE_ID = PigTypeID Then
               Pm.PartItem.PART_ITEM_ID = Pm.PART_ITEM_ID
               Call Pm.PartItem.QueryData(1, m_Rs3, iCount)
               If Not m_Rs3.EOF Then
                  Call Pm.PartItem.PopulateFromRS(1, m_Rs3)
                  Set p = Pm.PartItem
                  Set DateToPartItem = p
                  Exit Function
               Else
                  Set p = Nothing
               End If
            End If
            
            m_Rs2.MoveNext
         Wend
         Set Pm = Nothing
      End If
      
      m_Rs1.MoveNext
   Wend
   
   Set Pw = Nothing
   Set DateToPartItem = Nothing
End Function

Public Function WeekIDToPartItem(Yw As CYearWeek, PigTypeID As Long) As CPartItem
Dim Count As Long
Dim Pw As CYearWeek
Dim iCount As Long
Dim p As CPartItem
Dim Pm As CPrtItemMap

   Set Pw = New CYearWeek
   Pw.YEAR_WEEK_ID = -1
   Pw.YEAR_SEQ_ID = -1
   
   Call Pw.QueryData(1, m_Rs1, iCount)
   Set Pw = Nothing
   While Not m_Rs1.EOF
      If (NVLI(m_Rs1("YEAR_WEEK_ID"), -1) = Yw.YEAR_WEEK_ID) Then
      
         Set Pm = New CPrtItemMap
         Pm.PRTITEM_MAP_ID = -1
         Pm.YEAR_WEEK_ID = Yw.YEAR_WEEK_ID
         Call Pm.QueryData(1, m_Rs2, iCount)
         While Not m_Rs2.EOF
         
            Call Pm.PopulateFromRS(1, m_Rs2)
            If Pm.PRODUCT_TYPE_ID = PigTypeID Then
               Pm.PartItem.PART_ITEM_ID = Pm.PART_ITEM_ID
               Call Pm.PartItem.QueryData(1, m_Rs3, iCount)
               If Not m_Rs3.EOF Then
                  Call Pm.PartItem.PopulateFromRS(1, m_Rs3)
                  Set p = Pm.PartItem
                  Set WeekIDToPartItem = p
                  Exit Function
               Else
                  Set p = Nothing
               End If
            End If
            
            m_Rs2.MoveNext
         Wend
         Set Pm = Nothing
      End If
      
      m_Rs1.MoveNext
   Wend
   
   Set Pw = Nothing
   Set WeekIDToPartItem = Nothing
End Function

'==
Public Function QueryBillingDoc(Ua As CBillingDoc, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryBillingDoc"
   QueryBillingDoc = False

   IsOK = True
   Call Ua.QueryData(1, Rs, itemcount)

   If Ua.QueryFlag = 1 Then
      Dim Yw As CDoItem
      
      Set Yw = New CDoItem
      Yw.DO_ITEM_ID = -1
      Yw.DO_ID = Ua.BILLING_DOC_ID
      Yw.OrderBy = -1
      Yw.OrderType = 1
      Yw.BATCH_ID = Ua.BATCH_ID
      Call Yw.QueryData(1, m_Rs1, iCount)
      Set Yw = Nothing
      
      Set Ua.DoItems = Nothing
      Set Ua.DoItems = New Collection
      Set Ua.Revenues = Nothing
      Set Ua.Revenues = New Collection
      While Not m_Rs1.EOF
         Set Yw = New CDoItem
         Call Yw.PopulateFromRS(1, m_Rs1)
                  
         Yw.Flag = "I"
         If Yw.PART_ITEM_ID > 0 Then
            Call Ua.DoItems.Add(Yw)
         Else
            Call Ua.Revenues.Add(Yw)
         End If
         
         Set Yw = Nothing
         m_Rs1.MoveNext
      Wend
      
      '==
      Dim Ri As CROItem
      
      Set Ri = New CROItem
      Ri.RO_ITEM_ID = -1
      Ri.BILLING_DOC_ID = Ua.BILLING_DOC_ID
      Ri.FROM_DATE = -1
      Ri.TO_DATE = -1
      Ri.OrderBy = -1
      Ri.OrderType = 1
      Ri.BATCH_ID = Ua.BATCH_ID
      Call Ri.QueryData(1, m_Rs1, iCount)
         Set Ri = Nothing
      
      Set Ua.RoItems = Nothing
      Set Ua.RoItems = New Collection
      While Not m_Rs1.EOF
         Set Ri = New CROItem
         Call Ri.PopulateFromRS(1, m_Rs1)
                  
         Ri.Flag = "I"
         Call Ua.RoItems.Add(Ri)
         Call QueryExpenseRatio(Ri)
         
         Set Ri = Nothing
         m_Rs1.MoveNext
      Wend
      '==
   
      Dim Rt As CReceiptItem
      Set Rt = New CReceiptItem
      Rt.RECEIPT_ITEM_ID = -1
      Rt.BILLING_DOC_ID = Ua.BILLING_DOC_ID
      Rt.OrderBy = -1
      Rt.OrderType = 1
'      Rt.BILLING_DOC_ID  = Ua.BATCH_ID
      Call Rt.QueryData(1, m_Rs1, iCount)
      Set Rt = Nothing
      
      Set Ua.ReceiptItems = Nothing
      Set Ua.ReceiptItems = New Collection
       While Not m_Rs1.EOF
         Set Rt = New CReceiptItem
         Call Rt.PopulateFromRS(1, m_Rs1)
                  
         Rt.Flag = "I"
         Call Ua.ReceiptItems.Add(Rt)
         
         Set Rt = Nothing
         m_Rs1.MoveNext
      Wend

      '==
      Dim Cm As CCapitalMovement
      Set Cm = New CCapitalMovement
       Cm.CAPITAL_MOVEMENT_ID = -1
       Cm.BILLING_DOC_ID = Ua.BILLING_DOC_ID
      Cm.FROM_DATE = -1
      Cm.TO_DATE = -1
      Cm.OrderBy = -1
      Cm.OrderType = 1
      Cm.BATCH_ID = Ua.BATCH_ID
      Call Cm.QueryData(1, m_Rs1, iCount)
      Set Cm = Nothing
      
      Set Ua.CapitalMovements = Nothing
       Set Ua.CapitalMovements = New Collection
        While Not m_Rs1.EOF
         Set Cm = New CCapitalMovement
         Call Cm.PopulateFromRS(1, m_Rs1)
                  
         Cm.Flag = "I"
         Call Ua.CapitalMovements.Add(Cm)
         Call QueryMovementItem(Cm)
         
         Set Ri = Nothing
         m_Rs1.MoveNext
      Wend
      '==
      
      Dim Rn As CReceiptCnDn
      Set Rn = New CReceiptCnDn
      Rn.RECEIPT_CNDN_ID = -1
      Rn.BILLING_DOC_ID = Ua.BILLING_DOC_ID
      Rn.OrderBy = -1
      Rn.OrderType = 1
      Call Rn.QueryData(1, m_Rs1, iCount)
      Set Rn = Nothing
      
      Set Ua.ReceiptCnDns = Nothing
      Set Ua.ReceiptCnDns = New Collection
      While Not m_Rs1.EOF
         Set Rn = New CReceiptCnDn
         Call Rn.PopulateFromRS(1, m_Rs1)
                  
         Rn.Flag = "I"
         Call Ua.ReceiptCnDns.Add(Rn)
         
         Set Rn = Nothing
         m_Rs1.MoveNext
      Wend
      
      Dim Cq As CCheque
      Dim Ct As CCashTran
      Set Ct = New CCashTran
      Call Ct.SetFieldValue("BILLING_DOC_ID", Ua.BILLING_DOC_ID)
      Call Ct.QueryData(1, m_Rs1, iCount)
      Set Ct = Nothing

      Set Ua.Payments = Nothing
      Set Ua.Payments = New Collection

      While Not m_Rs1.EOF
         Set Ct = New CCashTran
         Call Ct.PopulateFromRS(1, m_Rs1)
         
         If Ct.GetFieldValue("CHECK_ID") > 0 Then
            Call Ct.Cheque.SetFieldValue("CHEQUE_ID", Ct.GetFieldValue("CHECK_ID"))
            Call Ct.Cheque.QueryData(1, m_Rs2, iCount)
            If Not m_Rs2.EOF Then
               Call Ct.Cheque.PopulateFromRS(1, m_Rs2)
            End If
         End If
         
         Call Ua.Payments.Add(Ct)
         Set Ct = Nothing

         m_Rs1.MoveNext
      Wend

      '===
   End If
      
   QueryBillingDoc = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryBillingDoc = False
End Function
Public Function QueryBillingDocFromNo(Ua As CBillingDoc, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryBillingDocFromNo"
   QueryBillingDocFromNo = False

   IsOK = True
   Call Ua.QueryData(1, Rs, itemcount)

   If Ua.QueryFlag = 1 Then
      Dim Yw As CDoItem
      
      Set Yw = New CDoItem
      Yw.DO_ITEM_ID = -1
      If Ua.BILLING_DOC_ID <= 0 Then
         Ua.BILLING_DOC_ID = Rs("BILLING_DOC_ID") 'สำหรับกรณีการ QUERY จาก DOCUMENT_NO
      End If
      Yw.DO_ID = Ua.BILLING_DOC_ID
      Yw.OrderBy = -1
      Yw.OrderType = 1
      Yw.BATCH_ID = Ua.BATCH_ID
      Call Yw.QueryData(1, m_Rs1, iCount)
      Set Yw = Nothing
      
      Set Ua.DoItems = Nothing
      Set Ua.DoItems = New Collection
      Set Ua.Revenues = Nothing
      Set Ua.Revenues = New Collection
      While Not m_Rs1.EOF
         Set Yw = New CDoItem
         Call Yw.PopulateFromRS(1, m_Rs1)
                  
         Yw.Flag = "I"
         If Yw.PART_ITEM_ID > 0 Then
            Call Ua.DoItems.Add(Yw)
         Else
            Call Ua.Revenues.Add(Yw)
         End If
         
         Set Yw = Nothing
         m_Rs1.MoveNext
      Wend
      
      '==
      
      '===
   End If
      
   QueryBillingDocFromNo = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryBillingDocFromNo = False
End Function

Private Sub QueryExpenseRatio(Ri As CROItem)
Dim TempRs As ADODB.Recordset
Dim Er As CExpenseRatio
Dim iCount As Long

   Set TempRs = New ADODB.Recordset
   Set Er = New CExpenseRatio
   
   Er.EXPENSE_RATIO_ID = -1
   Er.RO_ITEM_ID = Ri.RO_ITEM_ID
   Call Er.QueryData(1, TempRs, iCount)
   Set Er = Nothing
   
   Set Ri.ExpenseRatios = Nothing
   Set Ri.ExpenseRatios = New Collection
   While Not TempRs.EOF
      Set Er = New CExpenseRatio
      Call Er.PopulateFromRS(1, TempRs)
      Er.Flag = "I"
      Call Ri.ExpenseRatios.Add(Er)
      
      Set Er = Nothing
      TempRs.MoveNext
   Wend
   
   Set Er = Nothing
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
End Sub

Private Sub QueryMovementItem(Cm As CCapitalMovement)
Dim TempRs As ADODB.Recordset
Dim Er As CMovementItem
Dim iCount As Long

   Set TempRs = New ADODB.Recordset
   Set Er = New CMovementItem
   
   Er.MOVEMENT_ITEM_ID = -1
   Er.CAPITAL_MOVEMENT_ID = Cm.CAPITAL_MOVEMENT_ID
   Call Er.QueryData(1, TempRs, iCount)
   Set Er = Nothing
   
   Set Cm.MovementItems = Nothing
   Set Cm.MovementItems = New Collection
   Cm.TOTAL_CAPITAL = 0
   While Not TempRs.EOF
      Set Er = New CMovementItem
      Call Er.PopulateFromRS(1, TempRs)
      Er.Flag = "I"
      Call Cm.MovementItems.Add(Er)
      
      Cm.TOTAL_CAPITAL = Cm.TOTAL_CAPITAL + Er.CAPITAL_AMOUNT
      Set Er = Nothing
      TempRs.MoveNext
   Wend
   
   Set Er = Nothing
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
End Sub

Public Function DeleteBillingDoc(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog, Optional PaymentID As Long = -1)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim itemcount As Long
Dim B As CBillingDoc

   RName = "DeleteBillingDoc"
   DeleteBillingDoc = False
   HasBegin = False

   Set B = New CBillingDoc

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.BILLING_DOC_ID = UID
   B.PAYMENT_ID = PaymentID
   Call B.QueryData(1, m_Rs1, itemcount)
   
   If Not m_Rs1.EOF Then
      Call B.PopulateFromRS(1, m_Rs1)
   End If

   Call B.DeleteData
   If B.INVENTORY_DOC_ID > 0 Then
      Call glbDaily.DeleteInventoryDoc(B.INVENTORY_DOC_ID, IsOK, False, ErrorObj)
   End If
   
   If B.PAYMENT_ID > 0 Then
      Call glbDaily.DeletePayment(B.PAYMENT_ID, IsOK, False, ErrorObj)
   End If
   
   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteBillingDoc = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False

   DeleteBillingDoc = False
End Function

Private Function AddEditMovementItem(E As CCapitalMovement)
Dim D As CMovementItem

   For Each D In E.MovementItems
      If D.Flag = "A" Then
         D.CAPITAL_MOVEMENT_ID = E.CAPITAL_MOVEMENT_ID
         D.AddEditMode = SHOW_ADD
         D.AddEditData
      ElseIf D.Flag = "E" Then
         D.AddEditMode = SHOW_EDIT
         D.AddEditData
      ElseIf D.Flag = "D" Then
         D.DeleteData
      End If
   Next D
End Function

Public Function AddEditBillingDoc(Ug As CBillingDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim PL As CPartLocation
Dim PL2 As CPartLocation
Dim iCount As Long
Dim NewPrice As Double
Dim O As Object
Dim Rn As CReceiptCnDn
Dim PaymentType As Long
Dim OldPaymentType As Long

   RName = "AddEditBillingDoc"
   AddEditBillingDoc = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
   
   Dim Cd As CConfigDoc
   If Ug.AddEditMode = SHOW_ADD And Ug.CONFIG_DOC_TYPE > 0 Then
      Set Cd = New CConfigDoc
      Call Cd.SetFieldValue("RUNNING_NO", Ug.RUNNING_NO)
      Call Cd.SetFieldValue("LAST_NO", Ug.DOCUMENT_NO)
      Call Cd.SetFieldValue("CONFIG_DOC_TYPE", Ug.CONFIG_DOC_TYPE)
      Call Cd.UpdateRunningNo
   End If
   Set Cd = Nothing
   
   Dim CstName As CDoItem
   For Each CstName In Ug.DoItems
      If CstName.Flag = "A" Then
         CstName.DO_ID = Ug.BILLING_DOC_ID
         CstName.AddEditMode = SHOW_ADD
         CstName.AddEditData
      ElseIf CstName.Flag = "E" Then
         CstName.AddEditMode = SHOW_EDIT
         CstName.AddEditData
      ElseIf CstName.Flag = "D" Then
         CstName.DeleteData
      End If
   Next CstName

   Dim Ri As CROItem
   For Each Ri In Ug.RoItems
      If Ri.Flag = "A" Then
         Ri.BILLING_DOC_ID = Ug.BILLING_DOC_ID
         Ri.AddEditMode = SHOW_ADD
         Ri.AddEditData
         
         Call AddEditExpenseRatio(Ri)
      ElseIf Ri.Flag = "E" Then
         Ri.AddEditMode = SHOW_EDIT
         Ri.AddEditData
         
         Call AddEditExpenseRatio(Ri)
      ElseIf Ri.Flag = "D" Then
         Ri.DeleteData
      End If
   Next Ri

   Dim Rt As CReceiptItem
   For Each Rt In Ug.ReceiptItems
      If Rt.Flag = "A" Then
         Rt.BILLING_DOC_ID = Ug.BILLING_DOC_ID
         Rt.AddEditMode = SHOW_ADD
         Rt.AddEditData
      ElseIf Rt.Flag = "E" Then
         Rt.AddEditMode = SHOW_EDIT
         Rt.AddEditData
      ElseIf Rt.Flag = "D" Then
         Rt.DeleteData
      End If
   Next Rt

   Dim Cm As CCapitalMovement
   For Each Cm In Ug.CapitalMovements
      If Cm.Flag = "A" Then
         Cm.BILLING_DOC_ID = Ug.BILLING_DOC_ID
         Cm.AddEditMode = SHOW_ADD
         Cm.AddEditData
         
         Call AddEditMovementItem(Cm)
      ElseIf Cm.Flag = "E" Then
         Cm.AddEditMode = SHOW_EDIT
         Cm.AddEditData
         
         Call AddEditMovementItem(Cm)
      ElseIf Cm.Flag = "D" Then
         Cm.DeleteData
      End If
   Next Cm
   
   For Each CstName In Ug.Revenues
      If CstName.Flag = "A" Then
         CstName.DO_ID = Ug.BILLING_DOC_ID
         CstName.AddEditMode = SHOW_ADD
         CstName.AddEditData
      ElseIf CstName.Flag = "E" Then
         CstName.AddEditMode = SHOW_EDIT
         CstName.AddEditData
      ElseIf CstName.Flag = "D" Then
         CstName.DeleteData
      End If
   Next CstName
   
   For Each Rn In Ug.ReceiptCnDns
      If Rn.Flag = "A" Then
         Rn.BILLING_DOC_ID = Ug.BILLING_DOC_ID
         Rn.AddEditMode = SHOW_ADD
         Rn.AddEditData
      ElseIf Rn.Flag = "E" Then
         Rn.AddEditMode = SHOW_EDIT
         Rn.AddEditData
      ElseIf Rn.Flag = "D" Then
         Rn.DeleteData
      End If
   Next Rn
   
   Dim Ct As CCashTran
   For Each Ct In Ug.Payments
      PaymentType = Ct.GetFieldValue("PAYMENT_TYPE")
      OldPaymentType = Ct.GetFieldValue("OLD_PAYMENT_TYPE")
      Call Ct.SetFieldValue("BILLING_DOC_ID", Ug.BILLING_DOC_ID)
      Call Ct.SetFieldValue("TX_NO", Ug.DOCUMENT_NO)
      Call Ct.SetFieldValue("TX_DATE", Ug.DOCUMENT_DATE)
      Call Ct.SetFieldValue("CUSTOMER_ID", Ug.CUSTOMER_ID)
      If Ct.Flag = "A" Then
         Ct.ShowMode = SHOW_ADD
         If Ct.GetFieldValue("PAYMENT_TYPE") = 3 Then 'เช็ค
            Ct.Cheque.ShowMode = SHOW_ADD
            Call Ct.Cheque.AddEditData
         End If
         Call Ct.SetFieldValue("CHECK_ID", Ct.Cheque.GetFieldValue("CHEQUE_ID"))
         Call Ct.AddEditData
      ElseIf Ct.Flag = "E" Then
         'จากอื่น ๆ มาเป็นเช็ค -> เพิ่มเช็ค, จากเช็คเป็นอื่น ๆ -> ลบเช็ค, เป็นเช็คอยู่แล้วไม่เปลี่ยน -> แก้ไขเช็ค
         If (PaymentType <> OldPaymentType) And (PaymentType = 3) Then
            Ct.Cheque.ShowMode = SHOW_ADD
            Call Ct.Cheque.AddEditData
            
            Ct.ShowMode = SHOW_EDIT
            Call Ct.SetFieldValue("CHECK_ID", Ct.Cheque.GetFieldValue("CHEQUE_ID"))
            Call Ct.AddEditData
         ElseIf (PaymentType <> OldPaymentType) And (OldPaymentType = 3) Then
            Ct.ShowMode = SHOW_EDIT
            Call Ct.SetFieldValue("CHECK_ID", -1)
            Call Ct.AddEditData
         
            Call Ct.Cheque.DeleteData
         ElseIf (PaymentType = OldPaymentType) And (PaymentType = 3) Then
            Ct.Cheque.ShowMode = SHOW_EDIT
            Call Ct.Cheque.AddEditData
            
            Ct.ShowMode = SHOW_EDIT
            Call Ct.AddEditData
         Else
            Ct.ShowMode = SHOW_EDIT
            Call Ct.AddEditData
         End If
      ElseIf Ct.Flag = "D" Then
         Call Ct.DeleteData
         If Ct.GetFieldValue("OLD_CHECK_ID") > 0 Then 'เช็ค
            Call Ct.Cheque.DeleteData
         End If
      End If
   Next Ct

   'ทำเพื่อเซต APAR_ID ให้ตรงกันกับที่ BILLING_DOC
   For Each Ct In Ug.Payments
      If Ct.GetFieldValue("CHECK_ID") > 0 Then
         Call Ct.Cheque.SetFieldValue("CHEQUE_ID", Ct.GetFieldValue("CHECK_ID"))
         Call Ct.Cheque.SetFieldValue("APAR_ID", Ug.CUSTOMER_ID)
         Call Ct.Cheque.UpdateApArID
      End If
   Next Ct

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditBillingDoc = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditBillingDoc = False
End Function

Private Function AddEditExpenseRatio(E As CROItem)
Dim D As CExpenseRatio

   For Each D In E.ExpenseRatios
      If D.Flag = "A" Then
         D.RO_ITEM_ID = E.RO_ITEM_ID
         D.AddEditMode = SHOW_ADD
         D.AddEditData
      ElseIf D.Flag = "E" Then
         D.AddEditMode = SHOW_EDIT
         D.AddEditData
      ElseIf D.Flag = "D" Then
         D.DeleteData
      End If
   Next D
End Function

Public Function ImportLegacyUnit(Ug As CLegacy, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim U As CUnit
Dim D As CLegacy
Dim HasBegin As Boolean

   RName = "ImportLegacyUnit"
   ImportLegacyUnit = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call glbLegacy.QueryLegacyUnit(Ug, m_Rs1, iCount, IsOK, glbErrorLog)
   
   I = 0
   Set D = New CLegacy
   While Not m_Rs1.EOF
      I = I + 1
      Call D.PopulateFromRS(1, m_Rs1)
      
      Set U = New CUnit
      U.AddEditMode = SHOW_ADD
      U.UNIT_NO = Format(I, "000")
      U.UNIT_NAME = D.FIELD1
      If CheckUniqueNs(UNIT_NO, U.UNIT_NO, -1) Then
         U.AddEditData
      End If
      Set U = Nothing
      
      m_Rs1.MoveNext
   Wend
   Set D = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   ImportLegacyUnit = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   ImportLegacyUnit = False
End Function

Public Function ImportLegacyPartType(Ug As CLegacy, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim U As CPartType
Dim D As CLegacy
Dim HasBegin As Boolean

   RName = "ImportLegacyPartType"
   ImportLegacyPartType = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call glbLegacy.QueryLegacyPartType(Ug, m_Rs1, iCount, IsOK, glbErrorLog)
   
   I = 0
   Set D = New CLegacy
   While Not m_Rs1.EOF
      I = I + 1
      Call D.PopulateFromRS(2, m_Rs1)
      
      Set U = New CPartType
      U.AddEditMode = SHOW_ADD
      U.PART_TYPE_NO = D.FIELD1
      U.PART_TYPE_NAME = D.FIELD2
      U.RAW_FLAG = "Y"
      If CheckUniqueNs(PARTTYPE_NO, U.PART_TYPE_NO, -1) Then
         Call U.AddEditData
      End If
      Set U = Nothing
      
      m_Rs1.MoveNext
   Wend
   Set D = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   ImportLegacyPartType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   ImportLegacyPartType = False
End Function

Public Function ImportLegacyLocation(Ug As CLegacy, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim U As CLocation
Dim D As CLegacy
Dim HasBegin As Boolean

   RName = "ImportLegacyLocation"
   ImportLegacyLocation = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call glbLegacy.QueryLegacyLocation(Ug, m_Rs1, iCount, IsOK, glbErrorLog)
   
   I = 0
   Set D = New CLegacy
   While Not m_Rs1.EOF
      I = I + 1
      Call D.PopulateFromRS(3, m_Rs1)
      
      Set U = New CLocation
      U.AddEditMode = SHOW_ADD
      U.LOCATION_NO = D.FIELD1
      U.LOCATION_NAME = D.FIELD2
      U.SALE_FLAG = "N"
      U.LOCATION_TYPE = 2
      If CheckUniqueNs(LOCATION_NO_EX, U.LOCATION_NO, 1) Then
         Call U.AddEditData
      End If
      Set U = Nothing
      
      m_Rs1.MoveNext
   Wend
   Set D = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   ImportLegacyLocation = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   ImportLegacyLocation = False
End Function

Public Function ImportLegacyHouse(Ug As CLegacy, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim U As CLocation
Dim D As CLegacy
Dim HasBegin As Boolean

   RName = "ImportLegacyHouse"
   ImportLegacyHouse = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call glbLegacy.QueryLegacyHouse(Ug, m_Rs1, iCount, IsOK, glbErrorLog)
   
   I = 0
   Set D = New CLegacy
   While Not m_Rs1.EOF
      I = I + 1
      Call D.PopulateFromRS(3, m_Rs1)
      
      Set U = New CLocation
      U.AddEditMode = SHOW_ADD
      U.LOCATION_NO = D.FIELD1
      U.LOCATION_NAME = D.FIELD2
      If D.FIELD1 = "99" Then
         U.SALE_FLAG = "Y"
   Else
         U.SALE_FLAG = "N"
      End If
      U.LOCATION_TYPE = 1
      If CheckUniqueNs(LOCATION_NO_EX, U.LOCATION_NO, 2) Then
         Call U.AddEditData
      End If
      Set U = Nothing
      
      m_Rs1.MoveNext
   Wend
   Set D = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   ImportLegacyHouse = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   ImportLegacyHouse = False
End Function

Public Function ImportLegacyPigType(Ug As CLegacy, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim U As CProductType
Dim D As CLegacy
Dim HasBegin As Boolean

   RName = "ImportLegacyPigType"
   ImportLegacyPigType = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call glbLegacy.QueryLegacyPigType(Ug, m_Rs1, iCount, IsOK, glbErrorLog)
   
   I = 0
   Set D = New CLegacy
   While Not m_Rs1.EOF
      I = I + 1
      Call D.PopulateFromRS(4, m_Rs1)
      
      Set U = New CProductType
      U.AddEditMode = SHOW_ADD
      U.PRODUCT_TYPE_NO = D.FIELD1
      U.PRODUCT_TYPE_NAME = D.FIELD2
      If CheckUniqueNs(PRODUCTTYPE_NO, U.PRODUCT_TYPE_NO, -1) Then
         Call U.AddEditData
      End If
      Set U = Nothing
      
      m_Rs1.MoveNext
   Wend
   Set D = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   ImportLegacyPigType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   ImportLegacyPigType = False
End Function

Public Function ImportLegacyPigStatus(Ug As CLegacy, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim U As CProductStatus
Dim D As CLegacy
Dim HasBegin As Boolean

   RName = "ImportLegacyPigStatus"
   ImportLegacyPigStatus = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call glbLegacy.QueryLegacyPigStatus(Ug, m_Rs1, iCount, IsOK, glbErrorLog)
   
   I = 0
   Set D = New CLegacy
   While Not m_Rs1.EOF
      I = I + 1
      Call D.PopulateFromRS(5, m_Rs1)
      
      Set U = New CProductStatus
      U.AddEditMode = SHOW_ADD
      U.PRODUCT_STATUS_NO = D.FIELD1
      U.PRODUCT_STATUS_NAME = D.FIELD2
      If CheckUniqueNs(PRODUCTSTATUS_NO, U.PRODUCT_STATUS_NO, -1) Then
         Call U.AddEditData
      End If
      Set U = Nothing
      
      m_Rs1.MoveNext
   Wend
   Set D = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   ImportLegacyPigStatus = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   ImportLegacyPigStatus = False
End Function

Private Function LookupPartTypeID(PartTypeNO As String) As Long
Static TempCol As Collection
Dim T As CPartType
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadPartType(Nothing, TempCol)
   End If
   
   I = 0
   For Each T In TempCol
      I = I + 1
      If I = 1 Then
         TempID = T.PART_TYPE_ID
      End If
      
      If T.PART_TYPE_NO = PartTypeNO Then
            TempID = T.PART_TYPE_ID
         Exit For
      End If
   Next T
   
   LookupPartTypeID = TempID
End Function

Private Function LookupPartItemID(PartNo As String) As Long
Static TempCol As Collection
Dim T As CPartItem
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadPartItem(Nothing, TempCol)
   End If
   
   I = 0
   For Each T In TempCol
      I = I + 1
      If I = 1 Then
         TempID = T.PART_ITEM_ID
      End If
      
      If (Trim(T.PART_NO) = PartNo) Then
            TempID = T.PART_ITEM_ID
         Exit For
      End If
   Next T
   
   LookupPartItemID = TempID
End Function

Public Function LookupPartItemIDFromName(PartName As String, PartTypename As String) As Long
Static TempCol As Collection
Dim T As CPartItem
Dim TempID As Long
Dim I As Long
Dim FoundFlag As Boolean

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadPartItem(Nothing, TempCol)
   End If
   
   I = 0
   FoundFlag = False
   TempID = -1
   For Each T In TempCol
      I = I + 1
      If (Trim(T.PART_DESC) = PartName) And (Trim(T.PART_TYPE_NAME) = PartTypename) Then
            TempID = T.PART_ITEM_ID
            FoundFlag = True
         Exit For
      End If
   Next T
   
   If Not FoundFlag Then
'      ''debug.print PartTypename & "___" & PartName
   End If
   
   LookupPartItemIDFromName = TempID
End Function

Public Function LookupPigID(PigNo As String, PigType As String) As Long
Static TempCol As Collection
Dim T As CPartItem
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadPartItem(Nothing, TempCol, , "Y")
   End If
   
   I = 0
   For Each T In TempCol
      I = I + 1
      If I = 1 Then
         TempID = T.PART_ITEM_ID
      End If
      
      If (Trim(T.PART_NO) = PigNo) And (T.PIG_TYPE = PigType) Then
            TempID = T.PART_ITEM_ID
         Exit For
      End If
   Next T
   
   LookupPigID = TempID
End Function

Public Function LookupPigIDEx(PigNo As String, PigType As String) As Long
Static TempCol As Collection
Dim T As CPartItem
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadPartItem(Nothing, TempCol, , "Y")
   End If
   
   I = 0
   TempID = -1
   For Each T In TempCol
      I = I + 1
      
      If (Trim(T.PART_NO) = PigNo) And (T.PIG_TYPE = PigType) Then
            TempID = T.PART_ITEM_ID
         Exit For
      End If
   Next T
   
   LookupPigIDEx = TempID
End Function

Public Function LookupPigStatus(PigStatusNo As String) As Long
Static TempCol As Collection
Dim T As CProductStatus
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadProductStatus(Nothing, TempCol)
   End If
   
   I = 0
   For Each T In TempCol
      I = I + 1
      If I = 1 Then
         TempID = T.PRODUCT_STATUS_ID
      End If
      
      If (Trim(T.PRODUCT_STATUS_NO) = PigStatusNo) Then
            TempID = T.PRODUCT_STATUS_ID
         Exit For
      End If
   Next T
   
   LookupPigStatus = TempID
End Function

Public Function LookupPigStatusEx(PigStatusNo As String) As Long
Static TempCol As Collection
Dim T As CProductStatus
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadProductStatus(Nothing, TempCol)
   End If
   
   I = 0
   TempID = -1
   For Each T In TempCol
      I = I + 1
      
      If (Trim(T.PRODUCT_STATUS_NO) = PigStatusNo) Then
            TempID = T.PRODUCT_STATUS_ID
         Exit For
      End If
   Next T
   
   LookupPigStatusEx = TempID
End Function

Public Function LookupPartItemFromPartGroup(GroupName As String) As Long
Static TempCol As Collection
Dim T As CPartGroup
Dim TempID As Long
Dim I As Long
Dim Pi As CPartItem
Dim TempRs As ADODB.Recordset
Dim iCount As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadPartGroup(Nothing, TempCol)
   End If
   
   I = 0
   TempID = -1
   For Each T In TempCol
      I = I + 1
      
      If (Trim(T.PART_GROUP_NAME) = GroupName) Then
            TempID = T.PART_GROUP_ID
         Exit For
      End If
   Next T
   
   If TempID > 0 Then
      'เอา part_item_id ตัวแรกมามี group ตรงกับชื่อที่ต้องการ
      Set TempRs = New ADODB.Recordset
      
      Set Pi = New CPartItem
      Pi.PART_ITEM_ID = -1
      Pi.PART_DESC = ""
      Pi.PART_GROUP_ID = TempID
      Call Pi.QueryData(1, TempRs, iCount)
      If Not TempRs.EOF Then
         Call Pi.PopulateFromRS(1, TempRs)
         TempID = Pi.PART_ITEM_ID
      Else
         TempID = -1
      End If
      Set Pi = Nothing
      
      Set TempRs = Nothing
   End If
   
   LookupPartItemFromPartGroup = TempID
End Function

Public Function LookupPartItemFromPartGroupID(GroupID As Long) As Long
Static TempCol As Collection
Dim T As CPartGroup
Dim TempID As Long
Dim I As Long
Dim Pi As CPartItem
Dim TempRs As ADODB.Recordset
Dim iCount As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadPartGroup(Nothing, TempCol)
   End If
   
   I = 0
   TempID = -1
   For Each T In TempCol
      I = I + 1
      
      If (Trim(T.PART_GROUP_ID) = GroupID) Then
            TempID = T.PART_GROUP_ID
         Exit For
      End If
   Next T
   
   If TempID > 0 Then
      'เอา part_item_id ตัวแรกมามี group ตรงกับชื่อที่ต้องการ
      Set TempRs = New ADODB.Recordset
      
      Set Pi = New CPartItem
      Pi.PART_ITEM_ID = -1
      Pi.PART_DESC = ""
      Pi.PART_GROUP_ID = TempID
      Call Pi.QueryData(1, TempRs, iCount)
      If Not TempRs.EOF Then
         Call Pi.PopulateFromRS(1, TempRs)
         TempID = Pi.PART_ITEM_ID
      Else
         TempID = -1
      End If
      Set Pi = Nothing
      
      Set TempRs = Nothing
   End If
   
   LookupPartItemFromPartGroupID = TempID
End Function

Public Function LookupLocationID(LocationNo As String, Optional SaleFlag As String = "", Optional Area As Long = -1) As Long
Static TempCol As Collection
Dim T As CLocation
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadLocation(Nothing, TempCol, Area, SaleFlag)
   End If
   
   I = 0
   For Each T In TempCol
      I = I + 1
      If I = 1 Then
         TempID = T.LOCATION_ID
      End If
      
      If T.LOCATION_NO = LocationNo Then
            TempID = T.LOCATION_ID
         Exit For
      End If
   Next T
   
   LookupLocationID = TempID
End Function

Public Function LookupLocationIDName(LocationName As String, Optional SaleFlag As String = "", Optional LocationType As Long = -1) As Long
Static TempCol As Collection
Dim T As CLocation
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadLocation(Nothing, TempCol, LocationType, SaleFlag)
   End If
   
   I = 0
   For Each T In TempCol
      I = I + 1
      If I = 1 Then
         TempID = T.LOCATION_ID
      End If
      
      If T.LOCATION_NAME = LocationName Then
            TempID = T.LOCATION_ID
         Exit For
      End If
   Next T
   
   LookupLocationIDName = TempID
End Function

Public Function LookupLocationIDNameEx(LocationName As String, Optional SaleFlag As String = "", Optional LocationType As Long = -1, Optional ReturnSaleFlag As String = "") As Long
Static TempCol As Collection
Dim T As CLocation
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadLocation(Nothing, TempCol, LocationType, SaleFlag)
   End If
   
   I = 0
   TempID = -1
   For Each T In TempCol
      I = I + 1
      
      If T.LOCATION_NAME = LocationName Then
            TempID = T.LOCATION_ID
            ReturnSaleFlag = T.SALE_FLAG
         Exit For
      End If
   Next T
   
   LookupLocationIDNameEx = TempID
End Function

Public Function LookupExpenseIDCode(ExpenseNo As String) As Long
Static TempCol As Collection
Dim T As CExpenseType
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadExpenseType(Nothing, TempCol)
   End If
   
   I = 0
   TempID = -1
   For Each T In TempCol
      I = I + 1
      If T.EXPENSE_TYPE_NO = ExpenseNo Then
            TempID = T.EXPENSE_TYPE_ID
         Exit For
      End If
   Next T
   
   LookupExpenseIDCode = TempID
End Function

Public Function LookupExpenseIDName(ExpenseName As String) As Long
Static TempCol As Collection
Dim T As CExpenseType
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadExpenseType(Nothing, TempCol)
   End If
   
   I = 0
   TempID = -1
   For Each T In TempCol
      I = I + 1
      If T.EXPENSE_TYPE_NAME = ExpenseName Then
            TempID = T.EXPENSE_TYPE_ID
         Exit For
      End If
   Next T
   
   LookupExpenseIDName = TempID
End Function

Private Function LookupSaleLocationID(LocationNo As String) As Long
Static TempCol As Collection
Dim T As CLocation
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadLocation(Nothing, TempCol, , "Y")
   End If
   
   I = 0
   For Each T In TempCol
      I = I + 1
      If I = 1 Then
         TempID = T.LOCATION_ID
      End If
      
      If T.LOCATION_NO = LocationNo Then
            TempID = T.LOCATION_ID
         Exit For
      End If
   Next T
   
   LookupSaleLocationID = TempID
End Function

Private Function LookupSupplierID(SupplierNo As String) As Long
Static TempCol As Collection
Dim T As CSupplier
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadSupplier(Nothing, TempCol)
   End If
   
   I = 0
   For Each T In TempCol
      I = I + 1
      If I = 1 Then
         TempID = T.SUPPLIER_ID
      End If
      
      If T.SUPPLIER_CODE = SupplierNo Then
            TempID = T.SUPPLIER_ID
         Exit For
      End If
   Next T
   
   LookupSupplierID = TempID
End Function

Private Function LookupEmpID(EmpCode As String) As Long
Static TempCol As Collection
Dim T As CEmployee
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadEmployee(Nothing, TempCol)
   End If
   
   I = 0
   For Each T In TempCol
      I = I + 1
      If I = 1 Then
         TempID = T.EMP_ID
      End If
      
      If T.EMP_CODE = EmpCode Then
            TempID = T.EMP_ID
         Exit For
      End If
   Next T
   
   LookupEmpID = TempID
End Function

Private Function LookupPartUnitID(UnitName As String) As Long
Static TempCol As Collection
Dim T As CUnit
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadUnit(Nothing, TempCol)
   End If
   
   I = 0
   For Each T In TempCol
      I = I + 1
      If I = 1 Then
         TempID = T.UNIT_ID
      End If
      
      If T.UNIT_NAME = UnitName Then
            TempID = T.UNIT_ID
         Exit For
      End If
   Next T
   
   LookupPartUnitID = TempID
End Function

Private Function LookupGroupID(GroupNo As String) As Long
Static TempCol As Collection
Dim T As CUserGroup
Dim TempID As Long
Dim I As Long

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadUserGroup(Nothing, TempCol)
   End If
   
   I = 0
   For Each T In TempCol
      I = I + 1
      If I = 1 Then
         TempID = T.GROUP_ID
      End If
      
      If T.GROUP_NAME = GroupNo Then
            TempID = T.GROUP_ID
         Exit For
      End If
   Next T
   
   LookupGroupID = TempID
End Function

Public Function ImportLegacyPartItem(Ug As CLegacy, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim U As CPartItem
Dim D As CLegacy
Dim HasBegin As Boolean

   RName = "ImportLegacyPartItem"
   ImportLegacyPartItem = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call glbLegacy.QueryLegacyPartItem(Ug, m_Rs1, iCount, IsOK, glbErrorLog)
   
   I = 0
   Set D = New CLegacy
   While Not m_Rs1.EOF
      I = I + 1
      Call D.PopulateFromRS(6, m_Rs1)
      
      Set U = New CPartItem
      U.AddEditMode = SHOW_ADD
      U.PART_NO = D.FIELD2 & "-" & D.FIELD1
      U.PART_DESC = D.FIELD3
      U.PIG_FLAG = "N"
      U.PART_TYPE = LookupPartTypeID(D.FIELD1)
      If CheckUniqueNs(PARTNO_UNIQUE, U.PART_NO, -1) Then
         U.UNIT_COUNT = LookupPartUnitID(D.FIELD4)
         Call U.AddEditData
      End If
      Set U = Nothing
      
      m_Rs1.MoveNext
   Wend
   Set D = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   ImportLegacyPartItem = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   ImportLegacyPartItem = False
End Function

Public Function ImportLegacySupplier(Ug As CLegacy, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim U As CSupplier
Dim D As CLegacy
Dim HasBegin As Boolean

   RName = "ImportLegacySupplier"
   ImportLegacySupplier = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call glbLegacy.QueryLegacySupplier(Ug, m_Rs1, iCount, IsOK, glbErrorLog)
   
   I = 0
   Set D = New CLegacy
   While Not m_Rs1.EOF
      I = I + 1
      Call D.PopulateFromRS(7, m_Rs1)
      
      Set U = New CSupplier
      U.AddEditMode = SHOW_ADD
      U.SUPPLIER_CODE = D.FIELD1
      U.SUPPLIER_NAME = D.FIELD2
      U.SUPPLIER_GRADE = -1
      U.SUPPLIER_TYPE = -1
      U.SUPPLIER_STATUS = -1
      
      Dim CstName As CSupplierName
      Dim Name As cName
      If U.CstNames.Count <= 0 Then
         Set CstName = New CSupplierName
         CstName.Flag = "A"
         
         Set Name = CstName.Name
         Name.LONG_NAME = D.FIELD2
         Name.SHORT_NAME = ""
         Name.Flag = "A"
         
         Call U.CstNames.Add(CstName)
      End If
      
      If CheckUniqueNs(SUPPLIER_UNIQUE, U.SUPPLIER_CODE, -1) Then
         Call glbDaily.AddEditSupplier(U, IsOK, False, ErrorObj)
      End If
      Set U = Nothing
      
      m_Rs1.MoveNext
   Wend
   Set D = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   ImportLegacySupplier = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   ImportLegacySupplier = False
End Function

Public Function ImportLegacyEmployee(Ug As CLegacy, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim U As CEmployee
Dim D As CLegacy
Dim HasBegin As Boolean

   RName = "ImportLegacyEmployee"
   ImportLegacyEmployee = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call glbLegacy.QueryLegacyEmployee(Ug, m_Rs1, iCount, IsOK, glbErrorLog)
   
   I = 0
   Set D = New CLegacy
   While Not m_Rs1.EOF
      I = I + 1
      Call D.PopulateFromRS(8, m_Rs1)
      
      Set U = New CEmployee
      U.AddEditMode = SHOW_ADD
      U.EMP_CODE = D.FIELD1
      U.Name = D.FIELD2
      U.LASTNAME = ""
      U.CURRENT_POSITION = -1
      U.PASS_STATUS = "Y"
      U.EXTERNAL_FLAG = "N"
      
      U.EmpName.AddEditMode = SHOW_ADD
      U.EName.AddEditMode = SHOW_ADD
      U.EName.LONG_NAME = D.FIELD2
      U.EName.LAST_NAME = ""
         
      If CheckUniqueNs(EMPCODE_UNIQUE, U.EMP_CODE, -1) Then
         Call glbDaily.AddEditEmployee(U, IsOK, False, ErrorObj)
      End If
      Set U = Nothing
      
      m_Rs1.MoveNext
   Wend
   Set D = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   ImportLegacyEmployee = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   ImportLegacyEmployee = False
End Function

Public Function ImportLegacyUser(Ug As CLegacy, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim U As CUserAccount
Dim D As CLegacy
Dim HasBegin As Boolean

   RName = "ImportLegacyUser"
   ImportLegacyUser = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call glbLegacy.QueryLegacyUser(Ug, m_Rs1, iCount, IsOK, glbErrorLog)
   
   I = 0
   Set D = New CLegacy
   While Not m_Rs1.EOF
      I = I + 1
      Call D.PopulateFromRS(9, m_Rs1)
      
      Set U = New CUserAccount
      U.AddEditMode = SHOW_ADD
      U.USER_NAME = D.FIELD1
      U.USER_PASSWORD = D.FIELD2
      U.USER_STATUS = "Y"
      U.EXCEPTION_FLAG = "Y"
      U.GROUP_ID = LookupGroupID("")
      If CheckUniqueNs(USERNAME_UNIQUE, U.USER_NAME, -1) Then
         U.AddEditData
      End If
      Set U = Nothing
      
      m_Rs1.MoveNext
   Wend
   Set D = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   ImportLegacyUser = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   ImportLegacyUser = False
End Function

Public Function ImportLegacyCustomer(Ug As CLegacy, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim U As CCustomer
Dim D As CLegacy
Dim HasBegin As Boolean

   RName = "ImportLegacyCustomer"
   ImportLegacyCustomer = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call glbLegacy.QueryLegacyCustomer(Ug, m_Rs1, iCount, IsOK, glbErrorLog)
   
   I = 0
   Set D = New CLegacy
   While Not m_Rs1.EOF
      I = I + 1
      Call D.PopulateFromRS(10, m_Rs1)
      
      Set U = New CCustomer
      U.AddEditMode = SHOW_ADD
      U.CUSTOMER_CODE = D.FIELD1
      U.CUSTOMER_TYPE = -1
      U.CUSTOMER_NAME = D.FIELD4
      
      Dim CstName As CCustomerName
      Dim Name As cName
      If U.CstNames.Count <= 0 Then
         Set CstName = New CCustomerName
         CstName.Flag = "A"
         
         Set Name = CstName.Name
         Name.LONG_NAME = D.FIELD4
         Name.SHORT_NAME = ""
         Name.Flag = "A"
         
         Call U.CstNames.Add(CstName)
      End If
      
      If U.CstAccounts.Count <= 0 Then
         Dim Acc As CAccount
         
         Set Acc = New CAccount
         Acc.ACCOUNT_NO = U.CUSTOMER_CODE
         Acc.Flag = "A"
         
         Call U.CstAccounts.Add(Acc)
         Set Acc = Nothing
      End If
   
      If CheckUniqueNs(CUSTCODE_UNIQUE, U.CUSTOMER_CODE, -1) Then
         Call glbDaily.AddEditCustomer(U, IsOK, False, ErrorObj)
      End If
      Set U = Nothing
      
      m_Rs1.MoveNext
   Wend
   Set D = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   ImportLegacyCustomer = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   ImportLegacyCustomer = False
End Function

Public Function ClearInterim(AutoCommit As Boolean)
On Error GoTo ErrorHandler
Dim RName As String
Dim H As CLegacy_h
Dim HasBegin As Boolean

   RName = "ClearInterim"
   ClearInterim = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Set H = New CLegacy_h
   Call H.DeleteAllData
   Set H = Nothing
   
   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False

   ClearInterim = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.RoutineName = RName
   glbErrorLog.ModuleName = m_ModuleName
   glbErrorLog.ShowErrorLog (LOG_TO_FILE)

   ClearInterim = False
End Function

Public Function ImportToInterim(AutoCommit As Boolean)
On Error GoTo ErrorHandler
Dim RName As String
Dim iCount As Long
Dim I As Long
Dim H As CLegacy_h
Dim D As CLegacy_d
Dim L1 As CLegacy
Dim L2 As CLegacy
Dim HasBegin As Boolean

   RName = "ImportToInterim"
   ImportToInterim = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Set L1 = New CLegacy
   Set L2 = New CLegacy
   
   '===
   Call L1.QueryData11(m_Rs1, iCount) 'Feed stock header
   While Not m_Rs1.EOF
      Call L1.PopulateFromRS(11, m_Rs1)
      L2.FIELD1 = L1.FIELD1
''debug.print L2.FIELD1
      Set H = New CLegacy_h
      H.AddEditMode = SHOW_ADD
      H.DOCUMENT_ID = 1
      H.H_FIELD1 = L1.FIELD1
      H.H_FIELD2 = L1.FIELD2
      H.H_FIELD3 = L1.FIELD3
      H.DOCUMENT_DATE = L1.FIELD2
      Call H.AddEditData

      Call L2.QueryData12(m_Rs2, iCount)  'Feed stock
      While Not m_Rs2.EOF
         Call L2.PopulateFromRS(12, m_Rs2)

         Set D = New CLegacy_d
         D.AddEditMode = SHOW_ADD
         D.LEGACY_H_ID = H.LEGACY_H_ID
         D.D_FIELD1 = L2.FIELD1
         D.D_FIELD2 = L2.FIELD2
         D.D_FIELD3 = L2.FIELD3
         D.D_FIELD4 = L2.FIELD4
         D.D_FIELD5 = L2.FIELD5
         D.D_FIELD6 = L2.FIELD6
         D.D_FIELD7 = L2.FIELD7
         D.D_FIELD8 = L2.FIELD8
         Call D.AddEditData
         Set D = Nothing

         m_Rs2.MoveNext
      Wend

      Set H = Nothing
      m_Rs1.MoveNext
   Wend
'   ===
   
'   ===
   Call L1.QueryData13(m_Rs1, iCount) 'Bill header
   While Not m_Rs1.EOF
      Call L1.PopulateFromRS(13, m_Rs1)
      L2.FIELD1 = L1.FIELD1
''debug.print L2.FIELD1
      Set H = New CLegacy_h
      H.AddEditMode = SHOW_ADD
      H.DOCUMENT_ID = 2
      H.H_FIELD1 = L1.FIELD1
      H.H_FIELD2 = L1.FIELD2
      H.H_FIELD3 = L1.FIELD3
      H.H_FIELD4 = L1.FIELD4
      H.H_FIELD5 = L1.FIELD5
      H.H_FIELD6 = L1.FIELD6
      H.H_FIELD7 = L1.FIELD7
      H.H_FIELD8 = L1.FIELD8
      H.H_FIELD9 = L1.FIELD9
      H.H_FIELD10 = L1.FIELD10
      H.DOCUMENT_DATE = L1.FIELD4
      Call H.AddEditData

      Call L2.QueryData14(m_Rs2, iCount)  'Feed stock
      While Not m_Rs2.EOF
         Call L2.PopulateFromRS(14, m_Rs2)

         Set D = New CLegacy_d
         D.AddEditMode = SHOW_ADD
         D.LEGACY_H_ID = H.LEGACY_H_ID
         D.D_FIELD1 = L2.FIELD1
         D.D_FIELD2 = L2.FIELD2
         D.D_FIELD3 = L2.FIELD3
         D.D_FIELD4 = L2.FIELD4
         D.D_FIELD5 = L2.FIELD5
         D.D_FIELD6 = L2.FIELD6
         D.D_FIELD7 = L2.FIELD7
         D.D_FIELD8 = L2.FIELD8
         D.D_FIELD9 = L2.FIELD9
         D.D_FIELD10 = L2.FIELD10
         Call D.AddEditData
         Set D = Nothing

         m_Rs2.MoveNext
      Wend

      Set H = Nothing
      m_Rs1.MoveNext
   Wend
'   ===
   
'   ===
   Call L1.QueryData15(m_Rs1, iCount) 'Bill header
   While Not m_Rs1.EOF
      Call L1.PopulateFromRS(15, m_Rs1)
      L2.FIELD1 = L1.FIELD1
''debug.print L2.FIELD1
      Set H = New CLegacy_h
      H.AddEditMode = SHOW_ADD
      H.DOCUMENT_ID = 3
      H.H_FIELD1 = L1.FIELD1
      H.H_FIELD2 = L1.FIELD2
      H.H_FIELD3 = L1.FIELD3
      H.H_FIELD4 = L1.FIELD4
      H.H_FIELD5 = L1.FIELD5
      H.H_FIELD6 = L1.FIELD6
      H.H_FIELD7 = L1.FIELD7
      H.H_FIELD8 = L1.FIELD8
      H.H_FIELD9 = L1.FIELD9
      H.H_FIELD10 = L1.FIELD10
      H.DOCUMENT_DATE = L1.FIELD2
      Call H.AddEditData

      Call L2.QueryData16(m_Rs2, iCount)  'Feed stock
      While Not m_Rs2.EOF
         Call L2.PopulateFromRS(16, m_Rs2)

         Set D = New CLegacy_d
         D.AddEditMode = SHOW_ADD
         D.LEGACY_H_ID = H.LEGACY_H_ID
         D.D_FIELD1 = L2.FIELD1
         D.D_FIELD2 = L2.FIELD2
         D.D_FIELD3 = L2.FIELD3
         D.D_FIELD4 = L2.FIELD4
         D.D_FIELD5 = L2.FIELD5
         D.D_FIELD6 = L2.FIELD6
         D.D_FIELD7 = L2.FIELD7
         D.D_FIELD8 = L2.FIELD8
         D.D_FIELD9 = L2.FIELD9
         D.D_FIELD10 = L2.FIELD10
         Call D.AddEditData
         Set D = Nothing

         m_Rs2.MoveNext
      Wend

      Set H = Nothing
      m_Rs1.MoveNext
   Wend
'   ===
   
   Call L1.QueryData17(m_Rs1, iCount)
   While Not m_Rs1.EOF
      Call L1.PopulateFromRS(17, m_Rs1)
      L2.FIELD1 = L1.FIELD1
''debug.print L2.FIELD1
      Set H = New CLegacy_h
      H.AddEditMode = SHOW_ADD
      H.DOCUMENT_ID = 4
      H.H_FIELD1 = L1.FIELD1
      H.H_FIELD2 = L1.FIELD2
      H.H_FIELD3 = L1.FIELD3
      H.H_FIELD4 = L1.FIELD4
      H.H_FIELD5 = L1.FIELD5
      H.H_FIELD6 = L1.FIELD6
      H.H_FIELD7 = L1.FIELD7
      H.H_FIELD8 = L1.FIELD8
      H.H_FIELD9 = L1.FIELD9
      H.H_FIELD10 = L1.FIELD10
      H.DOCUMENT_DATE = L1.FIELD4
      Call H.AddEditData

      Set H = Nothing
      m_Rs1.MoveNext
   Wend
   
'   ===
   Call L1.QueryData18(m_Rs1, iCount) 'Bill header
   While Not m_Rs1.EOF
      Call L1.PopulateFromRS(18, m_Rs1)
      L2.FIELD1 = L1.FIELD1
''debug.print L2.FIELD1
      Set H = New CLegacy_h
      H.AddEditMode = SHOW_ADD
      H.DOCUMENT_ID = 5
      H.H_FIELD1 = L1.FIELD1
      H.H_FIELD2 = L1.FIELD2
      H.H_FIELD3 = L1.FIELD3
      H.H_FIELD4 = L1.FIELD4
      H.H_FIELD5 = L1.FIELD5
      H.H_FIELD6 = L1.FIELD6
      H.H_FIELD7 = L1.FIELD7
      H.H_FIELD8 = L1.FIELD8
      H.H_FIELD9 = L1.FIELD9
      H.H_FIELD10 = L1.FIELD10
      H.DOCUMENT_DATE = L1.FIELD2
      Call H.AddEditData

      Call L2.QueryData19(m_Rs2, iCount)  'Feed stock
      While Not m_Rs2.EOF
         Call L2.PopulateFromRS(19, m_Rs2)

         Set D = New CLegacy_d
         D.AddEditMode = SHOW_ADD
         D.LEGACY_H_ID = H.LEGACY_H_ID
         D.D_FIELD1 = L2.FIELD1
         D.D_FIELD2 = L2.FIELD2
         D.D_FIELD3 = L2.FIELD3
         D.D_FIELD4 = L2.FIELD4
         D.D_FIELD5 = L2.FIELD5
         D.D_FIELD6 = L2.FIELD6
         D.D_FIELD7 = L2.FIELD7
         D.D_FIELD8 = L2.FIELD8
         D.D_FIELD9 = L2.FIELD9
         D.D_FIELD10 = L2.FIELD10
         Call D.AddEditData
         Set D = Nothing

         m_Rs2.MoveNext
      Wend
      Set H = Nothing
      m_Rs1.MoveNext
   Wend

   Call L1.QueryData20(m_Rs1, iCount) 'Bill header
   While Not m_Rs1.EOF
      Call L1.PopulateFromRS(20, m_Rs1)
      L2.FIELD1 = L1.FIELD1
''debug.print L2.FIELD1
      Set H = New CLegacy_h
      H.AddEditMode = SHOW_ADD
      H.DOCUMENT_ID = 0
      H.H_FIELD1 = L1.FIELD1
      H.H_FIELD2 = L1.FIELD2
      H.H_FIELD3 = L1.FIELD3
      H.H_FIELD4 = L1.FIELD4
      H.H_FIELD5 = L1.FIELD5
      H.H_FIELD6 = L1.FIELD6
      H.H_FIELD7 = L1.FIELD7
      H.H_FIELD8 = L1.FIELD8
      H.H_FIELD9 = L1.FIELD9
      H.H_FIELD10 = L1.FIELD10
      H.DOCUMENT_DATE = L1.FIELD2
      Call H.AddEditData

      Call L2.QueryData21(m_Rs2, iCount)
      While Not m_Rs2.EOF
         Call L2.PopulateFromRS(21, m_Rs2)

         Set D = New CLegacy_d
         D.AddEditMode = SHOW_ADD
         D.LEGACY_H_ID = H.LEGACY_H_ID
         D.D_FIELD1 = L2.FIELD1
         D.D_FIELD2 = L2.FIELD2
         D.D_FIELD3 = L2.FIELD3
         D.D_FIELD4 = L2.FIELD4
         D.D_FIELD5 = L2.FIELD5
         D.D_FIELD6 = L2.FIELD6
         D.D_FIELD7 = L2.FIELD7
         D.D_FIELD8 = L2.FIELD8
         D.D_FIELD9 = L2.FIELD9
         D.D_FIELD10 = L2.FIELD10
         Call D.AddEditData
         Set D = Nothing

         m_Rs2.MoveNext
      Wend

      Set H = Nothing
      m_Rs1.MoveNext
   Wend
   Set H = New CLegacy_h
   Call H.SetDocType
   Set H = Nothing
   '===
   
   Set L2 = Nothing
   Set L1 = Nothing
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False

   ImportToInterim = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.RoutineName = RName
   glbErrorLog.ModuleName = m_ModuleName
   glbErrorLog.ShowErrorLog (LOG_TO_FILE)

   ImportToInterim = False
End Function

Public Sub TriggerCommit(Col As Collection)
Dim D As Object

   For Each D In Col
      If (D.Flag <> "D") And (D.Flag <> "A") Then
         D.Flag = "E"
      End If
   Next D
End Sub

'===
Public Function QueryLegacyH(Ua As CLegacy_h, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long
Dim Exp As CXItem

   RName = "QueryLegacyH"
   QueryLegacyH = False

   IsOK = True
   Call Ua.QueryData(Rs, itemcount)

   If Ua.QueryFlag = 1 Then
      Dim Yw As CLegacy_d
      
      Set Yw = New CLegacy_d
      Yw.LEGACY_D_ID = -1
      Yw.LEGACY_H_ID = Ua.LEGACY_H_ID
      Yw.OrderBy = 1
      Call Yw.QueryData(m_Rs1, iCount)
      Set Yw = Nothing
      
      Set Ua.LegacyDetails = Nothing
      Set Ua.LegacyDetails = New Collection
      While Not m_Rs1.EOF
         Set Yw = New CLegacy_d
         Call Yw.PopulateFromRS(1, m_Rs1)
                  
         Yw.Flag = "I"
         Call Ua.LegacyDetails.Add(Yw)
         
         Set Yw = Nothing
         m_Rs1.MoveNext
      Wend
   End If
   
   QueryLegacyH = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryLegacyH = False
End Function

Public Sub CreateInventoryDoc1(Lh As CLegacy_h, Ivd As CInventoryDoc, CommitFlag As String)  'ใบเบิกวัตถุดิบ
Dim Lhd As CLegacy_d
Dim II As CExportItem
Dim Sum As Double

   Ivd.AddEditMode = SHOW_ADD
   Ivd.DOCUMENT_NO = Lh.H_FIELD1
   Ivd.DOCUMENT_DATE = InternalDateToDateEx(Lh.H_FIELD2)
   Ivd.DELIVERY_FEE = 0
   Ivd.DELIVERY_ID = -1
   Ivd.DO_NO = ""
   Ivd.DOCUMENT_TYPE = 2
   Ivd.EMP_ID = LookupEmpID(Lh.H_FIELD3)
   Ivd.SUPPLIER_ID = -1
   Ivd.TRUCK_NO = ""
   Ivd.BILL_NO = ""
   Ivd.COMMIT_FLAG = CommitFlag
      
   For Each Lhd In Lh.LegacyDetails
      Set II = New CExportItem
      II.Flag = "A"
      II.TX_TYPE = "E" 'Export
      II.CALCULATE_FLAG = "Y"
      II.EXPORT_AMOUNT = Val(Lhd.D_FIELD6)
      II.LOCATION_ID = LookupLocationID("I-" & Lhd.D_FIELD4)
      II.HOUSE_ID = LookupLocationID("H-" & Lhd.D_FIELD5)
      II.PART_ITEM_ID = LookupPartItemID(Lhd.D_FIELD2 & "-" & Lhd.D_FIELD8)
      II.PIG_ID = LookupPigID(Lhd.D_FIELD3, "M")
      
      Call Ivd.ImportExports.Add(II)
      Set II = Nothing
   Next Lhd
End Sub

Public Sub CreateInventoryDoc2(Lh As CLegacy_h, Ivd As CInventoryDoc, CommitFlag) 'ใบนำเข้า
Dim Lhd As CLegacy_d
Dim II As CImportItem
Dim Sum As Double

   Ivd.AddEditMode = SHOW_ADD
   Ivd.DOCUMENT_NO = Lh.H_FIELD1
   Ivd.DOCUMENT_DATE = InternalDateToDateEx(Lh.H_FIELD4)
   Ivd.DELIVERY_FEE = Val(Lh.H_FIELD10)
   Ivd.DELIVERY_ID = LookupSupplierID(Lh.H_FIELD6)
   Ivd.DO_NO = Lh.H_FIELD2
   Ivd.DOCUMENT_TYPE = 1
   Ivd.EMP_ID = -1
   Ivd.SUPPLIER_ID = LookupSupplierID(Lh.H_FIELD3)
   Ivd.TRUCK_NO = Lh.H_FIELD5
   Ivd.BILL_NO = Lh.H_FIELD7
   Ivd.COMMIT_FLAG = CommitFlag
   
   Sum = 0
   For Each Lhd In Lh.LegacyDetails
      Sum = Sum + Val(Lhd.D_FIELD3)
   Next Lhd
   
   For Each Lhd In Lh.LegacyDetails
      Set II = New CImportItem
      II.Flag = "A"
      II.TX_TYPE = "I" 'Import
      II.CALCULATE_FLAG = "Y"
      
      II.ACTUAL_UNIT_PRICE = Val(Lhd.D_FIELD4)
      If Sum > 0 Then
         II.INCLUDE_UNIT_PRICE = II.ACTUAL_UNIT_PRICE + (Ivd.DELIVERY_FEE / Sum)
      Else
         II.INCLUDE_UNIT_PRICE = II.ACTUAL_UNIT_PRICE
      End If
      II.IMPORT_AMOUNT = Val(Lhd.D_FIELD3)
      II.LOCATION_ID = LookupLocationID("I-00")
      II.PART_ITEM_ID = LookupPartItemID(Lhd.D_FIELD2 & "-" & Lhd.D_FIELD7)
      
      Call Ivd.ImportExports.Add(II)
      Set II = Nothing
   Next Lhd
End Sub

Public Sub CreateInventoryDoc3(Lh As CLegacy_h, Ivd As CInventoryDoc, CommitFlag) 'ใบเโอนวัตถุดิบ
Dim Lhd As CLegacy_d
Dim II As CImportItem
Dim EI As CExportItem
Dim Sum As Double
Dim PL As CPartLocation
Dim iCount As Long

   Ivd.AddEditMode = SHOW_ADD
   Ivd.DOCUMENT_NO = Lh.H_FIELD1
   Ivd.DOCUMENT_DATE = InternalDateToDateEx(Lh.H_FIELD2)
   Ivd.DELIVERY_FEE = 0
   Ivd.DELIVERY_ID = -1
   Ivd.DO_NO = ""
   Ivd.DOCUMENT_TYPE = 3
   Ivd.EMP_ID = LookupEmpID(Lh.H_FIELD3)
   Ivd.SUPPLIER_ID = -1
   Ivd.TRUCK_NO = ""
   Ivd.BILL_NO = ""
   Ivd.COMMIT_FLAG = CommitFlag
      
   Set PL = New CPartLocation
   For Each Lhd In Lh.LegacyDetails
      Set EI = New CExportItem
      EI.Flag = "A"
      EI.TX_TYPE = "E" 'Export
      EI.CALCULATE_FLAG = "Y"
      EI.EXPORT_AMOUNT = Val(Lhd.D_FIELD5)
      EI.LOCATION_ID = LookupLocationID("I-" & Lhd.D_FIELD4)
      EI.HOUSE_ID = -1
      EI.PART_ITEM_ID = LookupPartItemID(Lhd.D_FIELD2 & "-" & Lhd.D_FIELD7)
      EI.PIG_ID = -1

      Set II = New CImportItem
      II.Flag = "A"
      II.TX_TYPE = "I" 'Import
      II.CALCULATE_FLAG = "Y"
      II.IMPORT_AMOUNT = EI.EXPORT_AMOUNT
      II.LOCATION_ID = LookupLocationID("I-" & Lhd.D_FIELD3)
      II.PART_ITEM_ID = EI.PART_ITEM_ID
'      II.ACTUAL_UNIT_PRICE
'      II.INCLUDE_UNIT_PRICE
      PL.PART_LOCATION_ID = -1
      PL.PART_ITEM_ID = EI.PART_ITEM_ID
      PL.LOCATION_ID = EI.LOCATION_ID
      Call PL.QueryData(1, m_Rs3, iCount)
      If Not m_Rs3.EOF Then
         Call PL.PopulateFromRS(m_Rs3)
         
         II.ACTUAL_UNIT_PRICE = PL.AVG_PRICE
         II.INCLUDE_UNIT_PRICE = PL.AVG_PRICE
      End If
      
      Call Ivd.ImportExports.Add(EI)
      Call Ivd.ImportExports.Add(II)
      
      Set EI = Nothing
      Set II = Nothing
   Next Lhd
   Set PL = Nothing
End Sub

Public Sub CreateInventoryDoc4(Lh As CLegacy_h, Ivd As CInventoryDoc, CommitFlag As String)  'ใบสุกรคลอด
Dim Lhd As CLegacy_d
Dim II As CImportItem
Dim Sum As Double

   Ivd.AddEditMode = SHOW_ADD
   Ivd.DOCUMENT_NO = Lh.H_FIELD1
   Ivd.DOCUMENT_DATE = InternalDateToDateEx(Lh.H_FIELD4)
   Ivd.DELIVERY_FEE = 0
   Ivd.DELIVERY_ID = -1
   Ivd.DO_NO = ""
   Ivd.DOCUMENT_TYPE = 5
   Ivd.EMP_ID = -1
   Ivd.SUPPLIER_ID = -1
   Ivd.TRUCK_NO = ""
   Ivd.BILL_NO = ""
   Ivd.COMMIT_FLAG = CommitFlag
   
   Sum = 0
'   For Each Lhd In Lh.LegacyDetails
      Set II = New CImportItem
      II.Flag = "A"
      II.TX_TYPE = "I" 'Import
      II.CALCULATE_FLAG = "N"
      
      II.ACTUAL_UNIT_PRICE = 0
      II.INCLUDE_UNIT_PRICE = 0
      II.IMPORT_AMOUNT = Val(Lh.H_FIELD7)
      II.LOCATION_ID = LookupLocationID("H-" & Lh.H_FIELD3)
      II.PART_ITEM_ID = LookupPigID(Lh.H_FIELD10, "M")
      II.FATHER_NO = Lh.H_FIELD6
      II.MOTHER_NO = Lh.H_FIELD5
      II.TOTAL_WEIGHT = Val(Lh.H_FIELD9)
      Call Ivd.ImportExports.Add(II)
      Set II = Nothing
'   Next Lhd
End Sub

Public Sub CreateInventoryDoc6(Lh As CLegacy_h, Ivd As CInventoryDoc, CommitFlag As String)  'ใบเโอนย้ายสุกร
Dim Lhd As CLegacy_d
Dim II As CImportItem
Dim EI As CExportItem
Dim Sum As Double

   Ivd.AddEditMode = SHOW_ADD
   Ivd.DOCUMENT_NO = Lh.H_FIELD1
   Ivd.DOCUMENT_DATE = InternalDateToDateEx(Lh.H_FIELD2)
   Ivd.DELIVERY_FEE = 0
   Ivd.DELIVERY_ID = -1
   Ivd.DO_NO = ""
   Ivd.DOCUMENT_TYPE = 6
   Ivd.EMP_ID = LookupEmpID(Lh.H_FIELD3)
   Ivd.SUPPLIER_ID = -1
   Ivd.TRUCK_NO = ""
   Ivd.BILL_NO = ""
   Ivd.COMMIT_FLAG = CommitFlag

   For Each Lhd In Lh.LegacyDetails
      Set EI = New CExportItem
      EI.Flag = "A"
      EI.TX_TYPE = "E" 'Export
      EI.CALCULATE_FLAG = "Y"
      EI.EXPORT_AMOUNT = Val(Lhd.D_FIELD6)
      EI.LOCATION_ID = LookupLocationID("H-" & Lhd.D_FIELD5)
      EI.HOUSE_ID = -1
      If Val(Mid(Lhd.D_FIELD2, 5)) > 0 Then
         EI.PART_ITEM_ID = LookupPigID(Lhd.D_FIELD2, "M")
      Else
         EI.PART_ITEM_ID = LookupPigID(Mid(Lhd.D_FIELD2, 1, 4) & "00", Mid(Lhd.D_FIELD2, 5))
      End If
      EI.PIG_ID = -1
      EI.TOTAL_WEIGHT = Val(Lhd.D_FIELD7)
      
      Set II = New CImportItem
      II.Flag = "A"
      II.TX_TYPE = "I" 'Import
      II.CALCULATE_FLAG = "Y"
      II.IMPORT_AMOUNT = EI.EXPORT_AMOUNT
      II.LOCATION_ID = LookupLocationID("H-" & Lhd.D_FIELD4)
      If Val(Mid(Lhd.D_FIELD3, 5)) > 0 Then
         II.PART_ITEM_ID = LookupPigID(Lhd.D_FIELD3, "M")
      Else
         II.PART_ITEM_ID = LookupPigID(Mid(Lhd.D_FIELD3, 1, 4) & "00", Mid(Lhd.D_FIELD3, 5))
      End If
      II.TOTAL_WEIGHT = EI.TOTAL_WEIGHT
      
      Call Ivd.ImportExports.Add(EI)
      Call Ivd.ImportExports.Add(II)
      
      Set EI = Nothing
      Set II = Nothing
   Next Lhd
End Sub

Public Sub CreateInventoryDoc7(Lh As CLegacy_h, Ivd As CInventoryDoc, CommitFlag As String)  'ใบโอนสุกรเข้าเรือนขาย
Dim Lhd As CLegacy_d
Dim II As CImportItem
Dim EI As CExportItem
Dim Sum As Double

   Ivd.AddEditMode = SHOW_ADD
   Ivd.DOCUMENT_NO = Lh.H_FIELD1
   Ivd.DOCUMENT_DATE = InternalDateToDateEx(Lh.H_FIELD2)
   Ivd.DELIVERY_FEE = 0
   Ivd.DELIVERY_ID = -1
   Ivd.DO_NO = ""
   Ivd.DOCUMENT_TYPE = 7
   Ivd.EMP_ID = -1
   Ivd.SUPPLIER_ID = -1
   Ivd.TRUCK_NO = ""
   Ivd.BILL_NO = ""
   Ivd.COMMIT_FLAG = CommitFlag

   For Each Lhd In Lh.LegacyDetails
      Set EI = New CExportItem
      EI.Flag = "A"
      EI.TX_TYPE = "E" 'Export
      EI.CALCULATE_FLAG = "Y"
      EI.EXPORT_AMOUNT = Val(Lhd.D_FIELD5)
      EI.LOCATION_ID = LookupLocationID("H-" & Lhd.D_FIELD2)
      EI.HOUSE_ID = -1
      If Val(Mid(Lhd.D_FIELD3, 5)) > 0 Then
         EI.PART_ITEM_ID = LookupPigID(Lhd.D_FIELD3, "M")
      Else
         EI.PART_ITEM_ID = LookupPigID(Mid(Lhd.D_FIELD3, 1, 4) & "00", Mid(Lhd.D_FIELD3, 5))
      End If
      EI.PIG_ID = -1
      EI.TOTAL_WEIGHT = Val(Lhd.D_FIELD6)
      EI.PIG_STATUS = LookupPigStatus(Lhd.D_FIELD4)
      
      Set II = New CImportItem
      II.Flag = "A"
      II.TX_TYPE = "I" 'Import
      II.CALCULATE_FLAG = "Y"
      II.IMPORT_AMOUNT = EI.EXPORT_AMOUNT
      II.LOCATION_ID = LookupSaleLocationID("H-XXX") 'First for sale flag
      II.PART_ITEM_ID = EI.PART_ITEM_ID
      II.TOTAL_WEIGHT = EI.TOTAL_WEIGHT
      
      Call Ivd.ImportExports.Add(EI)
      Call Ivd.ImportExports.Add(II)
      
      Set EI = Nothing
      Set II = Nothing
   Next Lhd
End Sub

Public Sub CreateInventoryDoc8(Lh As CLegacy_h, Ivd As CInventoryDoc, CommitFlag As String)  'ใบเปลี่ยนสถานะสุกร
Dim Lhd As CLegacy_d
Dim II As CImportItem
Dim EI As CExportItem
Dim Sum As Double

   Ivd.AddEditMode = SHOW_ADD
   Ivd.DOCUMENT_NO = Lh.H_FIELD1
   Ivd.DOCUMENT_DATE = InternalDateToDateEx(Lh.H_FIELD2)
   Ivd.DELIVERY_FEE = 0
   Ivd.DELIVERY_ID = -1
   Ivd.DO_NO = ""
   Ivd.DOCUMENT_TYPE = 8
   Ivd.EMP_ID = LookupEmpID(Lh.H_FIELD3)
   Ivd.SUPPLIER_ID = -1
   Ivd.TRUCK_NO = ""
   Ivd.BILL_NO = ""
   Ivd.COMMIT_FLAG = CommitFlag

   For Each Lhd In Lh.LegacyDetails
      Set EI = New CExportItem
      EI.Flag = "A"
      EI.TX_TYPE = "E" 'Export
      EI.CALCULATE_FLAG = "Y"
      EI.EXPORT_AMOUNT = Val(Lhd.D_FIELD6)
      EI.LOCATION_ID = LookupLocationID("H-" & Lhd.D_FIELD5)
      EI.HOUSE_ID = -1
      If Val(Mid(Lhd.D_FIELD2, 5)) > 0 Then
         EI.PART_ITEM_ID = LookupPigID(Lhd.D_FIELD2, "M")
      Else
         EI.PART_ITEM_ID = LookupPigID(Mid(Lhd.D_FIELD2, 1, 4) & "00", Mid(Lhd.D_FIELD2, 5))
      End If
      EI.PIG_ID = -1
      EI.TOTAL_WEIGHT = Val(Lhd.D_FIELD7)
      
      Set II = New CImportItem
      II.Flag = "A"
      II.TX_TYPE = "I" 'Import
      II.CALCULATE_FLAG = "Y"
      II.IMPORT_AMOUNT = EI.EXPORT_AMOUNT
      II.LOCATION_ID = LookupLocationID("H-" & Lhd.D_FIELD4)
      If Val(Mid(Lhd.D_FIELD3, 5)) > 0 Then
         II.PART_ITEM_ID = LookupPigID(Lhd.D_FIELD3, "M")
      Else
         II.PART_ITEM_ID = LookupPigID(Mid(Lhd.D_FIELD3, 1, 4) & "00", Mid(Lhd.D_FIELD3, 5))
      End If
      II.TOTAL_WEIGHT = EI.TOTAL_WEIGHT
      
      Call Ivd.ImportExports.Add(EI)
      Call Ivd.ImportExports.Add(II)
      
      Set EI = Nothing
      Set II = Nothing
   Next Lhd
End Sub
Public Sub CreateTransferItems(Ivd As CInventoryDoc)
Dim Ti As CTransferItem
Dim O As Object
Dim EI As CExportItem
Dim II As CImportItem
Dim I As Long
Dim j As Long
Dim Count1 As Long
Dim Count2 As Long
Dim ImportCount As Long
Dim ExportCount As Long

   Count1 = Ivd.ImportExports.Count \ 2
   Count2 = Ivd.ImportExports.Count
   For I = 1 To Count1
      ImportCount = 0
      ExportCount = 0
      j = 1
      While j <= Count2
         Set O = Ivd.ImportExports(j)
         If (O.TX_TYPE = "I") Then
            ImportCount = ImportCount + 1
            If ImportCount = I Then
               Set II = O
            End If
         ElseIf (O.TX_TYPE = "E") Then
            ExportCount = ExportCount + 1
            If ExportCount = I Then
               Set EI = O
            End If
         End If
         j = j + 1
      Wend
         
      Set Ti = New CTransferItem
      Set Ti.ImportItem = II
      Set Ti.ExportItem = EI
      Ti.Flag = II.Flag
      Call Ivd.TransferItems.Add(Ti)
      Set Ti = Nothing
   Next I
End Sub

Public Function AddEditReportConfig(Ug As CReportConfig, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditReportConfig"
   AddEditReportConfig = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditReportConfig = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditReportConfig = False
End Function

Public Function AddEditManagementConfig(Ug As CManagementConfig, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditManagementConfig"
   AddEditManagementConfig = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditManagementConfig = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditManagementConfig = False
End Function

Public Function QueryReportConfig(Ua As CReportConfig, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryReportConfig"
   QueryReportConfig = False

   IsOK = True
   Call Ua.QueryData(Rs, itemcount)
   
   QueryReportConfig = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryReportConfig = False
End Function
Public Function QueryMemoNote(Ua As CMemoNote, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryMemoNote"
   QueryMemoNote = False

   IsOK = True
   Call Ua.QueryData(1, Rs, itemcount)

   If Ua.QueryFlag = 1 Then
      
   End If
   
   QueryMemoNote = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryMemoNote = False
End Function

Public Function QueryManagementConfig(Ua As CManagementConfig, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryManagementConfig"
   QueryManagementConfig = False

   IsOK = True
   Call Ua.QueryData(Rs, itemcount)
   
   QueryManagementConfig = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryManagementConfig = False
End Function

Private Function GetNextTx(Rs1 As ADODB.Recordset, Rs2 As ADODB.Recordset, Rs3 As ADODB.Recordset, II As CImportItem, EI As CExportItem, Ri As CROItem) As String
Dim EofFlag1 As Boolean
Dim EofFlag2 As Boolean
Dim EofFlag3 As Boolean
Dim TempCol As Collection
Dim O As Object
Dim I As Long
Dim MinSeq As Long
Dim MinDocDate As Date
Dim MinDocNo As String
Dim MinPriority As Long
Dim MinIdx As Long
   
   Set TempCol = New Collection
   
   'Import
   EofFlag1 = Rs1.EOF
   If Not Rs1.EOF Then
      Call II.PopulateFromRS(7, Rs1)
      Call TempCol.Add(II)
   End If

   'Export
   EofFlag2 = Rs2.EOF
   If Not Rs2.EOF Then
      Call EI.PopulateFromRS(13, Rs2)
      Call TempCol.Add(EI)
   End If
      
   'Ro item
   EofFlag3 = Rs3.EOF
   If Not Rs3.EOF Then
      Call Ri.PopulateFromRS(1, Rs3)
      Call TempCol.Add(Ri)
   End If
      
   I = 0
   MinIdx = 0
   For Each O In TempCol
      I = I + 1
      If O.DOCUMENT_DATE = DateSerial(2006, 1, 30) Then
         '''debug.print (O.DOCUMENT_DATE)
      End If
      If O.DOCUMENT_DATE = DateSerial(2006, 1, 31) Then
         '''debug.print (O.DOCUMENT_DATE)
      End If
      If O.DOCUMENT_DATE = DateSerial(2006, 2, 1) Then
         '''debug.print (O.DOCUMENT_DATE)
      End If
      

      If I = 1 Then
         MinDocDate = O.DOCUMENT_DATE
         MinPriority = O.PRIORITY1
         MinDocNo = O.DOCUMENT_NO
         MinSeq = O.TRANSACTION_SEQ
         MinIdx = I
      End If

      '====
      If DateToStringInt(O.DOCUMENT_DATE) = DateToStringInt(MinDocDate) Then
         If O.PRIORITY1 = MinPriority Then
            If O.DOCUMENT_NO = MinDocNo Then
               If O.TRANSACTION_SEQ < MinSeq Then
                  MinDocDate = O.DOCUMENT_DATE
                  MinPriority = O.PRIORITY1
                  MinDocNo = O.DOCUMENT_NO
                  MinSeq = O.TRANSACTION_SEQ
                  MinIdx = I
               End If
            ElseIf O.DOCUMENT_NO < MinDocNo Then
                  MinDocDate = O.DOCUMENT_DATE
                  MinPriority = O.PRIORITY1
                  MinDocNo = O.DOCUMENT_NO
                  MinSeq = O.TRANSACTION_SEQ
                  MinIdx = I
            End If
         ElseIf O.PRIORITY1 < MinPriority Then
               MinDocDate = O.DOCUMENT_DATE
               MinPriority = O.PRIORITY1
               MinDocNo = O.DOCUMENT_NO
               MinSeq = O.TRANSACTION_SEQ
               MinIdx = I
         End If
      ElseIf DateToStringInt(O.DOCUMENT_DATE) < DateToStringInt(MinDocDate) Then
            MinDocDate = O.DOCUMENT_DATE
            MinPriority = O.PRIORITY1
            MinDocNo = O.DOCUMENT_NO
            MinSeq = O.TRANSACTION_SEQ
            MinIdx = I
      End If 'Document date
      '====
'      If O.TRANSACTION_SEQ < MinSeq Then
'         MinSeq = O.TRANSACTION_SEQ
'         MinIdx = i
'      End If
   Next O

   If MinIdx > 0 Then
      Set O = TempCol(MinIdx)
'If (O.LOCATION_ID = 324) And (O.PART_ITEM_ID = 8837) Then
''   glbErrorLog.LocalErrorMsg = DateToStringInt(O.DOCUMENT_DATE) & "|" & O.PRIORITY1 & "|" & O.DOCUMENT_NO & "|" & O.TRANSACTION_SEQ
''   glbErrorLog.ShowUserErrorEx
''   ''debug.print glbErrorLog.LocalErrorMsg
'ElseIf (O.DOCUMENT_TYPE = 2) Then
''   If (O.HOUSE_ID = 324) And (O.PIG_ID = 8837) Then
''      glbErrorLog.LocalErrorMsg = DateToStringInt(O.DOCUMENT_DATE) & "|" & O.PRIORITY1 & "|" & O.DOCUMENT_NO & "|" & O.TRANSACTION_SEQ
''      glbErrorLog.ShowUserErrorEx
''      ''debug.print glbErrorLog.LocalErrorMsg
''   End If
'End If
      GetNextTx = O.TX_TYPE
   Else
      GetNextTx = ""
   End If
   
   Set TempCol = Nothing
End Function

Public Function QueryAllTransaction(Ivd As CInventoryDoc, IsOK As Boolean, RecordCount As Long, ErrorObj As clsErrorLog, Optional ClearFlag As Boolean = False, Optional BatchID As Long = -1) As Object
Static Rs1 As ADODB.Recordset
Static Rs2 As ADODB.Recordset
Static Rs3 As ADODB.Recordset
Static II As CImportItem
Static EI As CExportItem
Static Ri As CROItem
Dim iCount1 As Long
Dim iCount2 As Long
Dim iCount3 As Long
Static TxCode As String

   If ClearFlag Then
      Set II = Nothing
      Set EI = Nothing
      Set Ri = Nothing
   End If
   
   If II Is Nothing Then
      Set II = New CImportItem
      Set Rs1 = New ADODB.Recordset
      
      II.FROM_DATE = Ivd.FROM_DATE
      II.TO_DATE = Ivd.TO_DATE
      II.COMMIT_FLAG = Ivd.COMMIT_FLAG
      II.PART_ITEM_ID = -1
      II.LOCATION_ID = -1
'II.LOCATION_ID = 542
'II.PART_ITEM_ID = 14577
      II.PIG_FLAG = ""
      II.OrderBy = 2
      II.OrderType = 1
      II.BATCH_ID = BatchID
      Call II.QueryData(7, Rs1, iCount1)
   End If
   
   If EI Is Nothing Then
      Set EI = New CExportItem
      Set Rs2 = New ADODB.Recordset
      
      EI.FROM_DATE = Ivd.FROM_DATE
      EI.TO_DATE = Ivd.TO_DATE
      EI.COMMIT_FLAG = Ivd.COMMIT_FLAG
      EI.PIG_FLAG = ""
      EI.PART_ITEM_ID = -1
      EI.LOCATION_ID = -1
      
'EI.DOCUMENT_NO = "C705/35243"
'EI.LOCATION_ID = 542
'EI.PART_ITEM_ID = 14577
'EI.PIG_ID = 14577
'EI.HOUSE_ID = 542


      EI.OrderBy = 2
      EI.OrderType = 1
      EI.BATCH_ID = BatchID
      Call EI.QueryData(13, Rs2, iCount2)
   End If
      
   If Ri Is Nothing Then
      Set Ri = New CROItem
      Set Rs3 = New ADODB.Recordset
      
      Ri.FROM_DATE = Ivd.FROM_DATE
      Ri.TO_DATE = Ivd.TO_DATE
      Ri.COMMIT_FLAG = Ivd.COMMIT_FLAG
      Ri.PART_ITEM_ID = -1
      Ri.LOCATION_ID = -1
'Ri.LOCATION_ID = 542
'Ri.PART_ITEM_ID = 14571

      Ri.OrderBy = 11
      Ri.OrderType = 1
      Ri.BATCH_ID = BatchID
      Call Ri.QueryData(1, Rs3, iCount3)
   End If
   
   If ClearFlag Then
      RecordCount = iCount1 + iCount2 + iCount3
      Exit Function
   End If
   
   TxCode = GetNextTx(Rs1, Rs2, Rs3, II, EI, Ri)
   If TxCode <> "" Then
      If TxCode = "E" Then
         Set QueryAllTransaction = EI
         Rs2.MoveNext
      ElseIf TxCode = "I" Then
         Set QueryAllTransaction = II
         Rs1.MoveNext
      ElseIf TxCode = "X" Then
         Set QueryAllTransaction = Ri
         Rs3.MoveNext
      End If
   Else
      Set QueryAllTransaction = Nothing
   End If
   
   RecordCount = iCount1 + iCount2 + iCount3
End Function

Public Function ReArrangeDocument(Ivd As CInventoryDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long

   RName = "ReArrangeDocument"
   ReArrangeDocument = False
   HasBegin = False

   IsOK = True

   ReArrangeDocument = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   ReArrangeDocument = False
End Function

Public Function PatchNewPrice(Ug As CImportItem, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim II As CImportItem
Dim NewPrice As Double

   RName = "PatchNewPrice"
   PatchNewPrice = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Ug.OrderBy = 10
   Ug.OrderType = 1
   Call Ug.QueryData(1, m_Rs1, iCount)
   While Not m_Rs1.EOF
      Set II = New CImportItem
      Call II.PopulateFromRS(1, m_Rs1)
      
      NewPrice = (((II.CURRENT_AMOUNT - II.IMPORT_AMOUNT) * II.PREVIOUS_PRICE) + II.TOTAL_INCLUDE_PRICE) / II.CURRENT_AMOUNT
      II.NEW_PRICE = NewPrice
      
      II.AddEditMode = SHOW_EDIT
      Call II.AddEditData
      
      Set II = Nothing
      m_Rs1.MoveNext
   Wend
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   PatchNewPrice = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   PatchNewPrice = False
End Function

Public Function PatchTxSeq(Ug As CImportItem, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim II As CImportItem
Dim EI As CExportItem
Dim NewPrice As Double

   RName = "PatchTxSeq"
   PatchTxSeq = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.QueryData(23, m_Rs1, iCount)
   While Not m_Rs1.EOF
      Set II = New CImportItem
      Set EI = New CExportItem
      
      EI.EXPORT_ITEM_ID = NVLI(m_Rs1("EXPORT_ITEM_ID"), -1)
      EI.TRANSACTION_SEQ = NVLI(m_Rs1("IMP_SEQ"), -1)
      Call EI.PatchTxSeq
      
      II.IMPORT_ITEM_ID = NVLI(m_Rs1("IMPORT_ITEM_ID"), -1)
      II.TRANSACTION_SEQ = NVLI(m_Rs1("EXP_SEQ"), -1)
      Call II.PatchTxSeq
      
      Set EI = Nothing
      Set II = Nothing
      m_Rs1.MoveNext
   Wend
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   PatchTxSeq = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   PatchTxSeq = False
End Function

Public Function PatchIncludePrice(Ug As CInventoryDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim II As CImportItem
Dim EI As CExportItem
Dim NewPrice As Double
Dim Ivd As CInventoryDoc
Dim TempRs1 As ADODB.Recordset
Dim TempRs2 As ADODB.Recordset

   RName = "PatchIncludePrice"
   PatchIncludePrice = False
   HasBegin = False

   Set TempRs1 = New ADODB.Recordset
   Set TempRs2 = New ADODB.Recordset
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Ug.DOCUMENT_TYPE = 1
   Ug.COMMIT_FLAG = ""
   Call QueryInventoryDoc(Ug, TempRs1, iCount, IsOK, ErrorObj)
   
   While Not TempRs1.EOF
      Call Ug.PopulateFromRS(1, TempRs1)
      
      Set Ivd = New CInventoryDoc
            
      Ivd.INVENTORY_DOC_ID = Ug.INVENTORY_DOC_ID
      Ivd.COMMIT_FLAG = ""
      Ivd.QueryFlag = 1
      Call QueryInventoryDoc(Ivd, TempRs2, iCount, IsOK, ErrorObj)
      Call Ivd.PopulateFromRS(1, TempRs2)
      Call CalculateIncludePrice(Ivd, 0, Ivd.DELIVERY_FEE)
      
      For Each II In Ivd.ImportExports
         II.AddEditMode = SHOW_EDIT
         Call II.AddEditData
      Next II
         
      Set Ivd = Nothing
      TempRs1.MoveNext
   Wend
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   Set TempRs1 = Nothing
   Set TempRs2 = Nothing
   
   HasBegin = False
   IsOK = True

   PatchIncludePrice = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   PatchIncludePrice = False
End Function

Public Function RearrangeTxSeq(Ug As CInventoryDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog, Prg As ProgressBar, txtPercent As uctlTextBox)
On Error GoTo ErrorHandler
Dim O As Object
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim NewPrice As Double
Dim Ivd As CInventoryDoc
Dim TempRs1 As ADODB.Recordset
Dim TempRs2 As ADODB.Recordset
Dim TempID As Long
Dim I As Long
Dim RCount As Long

   RName = "RearrangeTxSeq"
   RearrangeTxSeq = False
   HasBegin = False

   Set TempRs1 = New ADODB.Recordset
   Set TempRs2 = New ADODB.Recordset
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

'Ug.INVENTORY_DOC_ID = 8641
   Ug.SALE_FLAG = ""
   Ug.COMMIT_FLAG = ""
   Ug.FROM_DATE = -1
   Ug.TO_DATE = -1
   Ug.DOCUMENT_TYPE = -1
   Ug.OrderBy = 3
   Ug.OrderType = 1
   Call QueryInventoryDoc(Ug, TempRs1, RCount, IsOK, ErrorObj)
   
   I = 0
   While Not TempRs1.EOF
      I = I + 1
      If Not (Prg Is Nothing) Then
         Prg.Value = MyDiff(I, RCount) * 100
         txtPercent.Text = FormatNumber(MyDiff(I, RCount) * 100)
      End If
      
      Call Ug.PopulateFromRS(1, TempRs1)
      
      Set Ivd = New CInventoryDoc
      
      Ivd.INVENTORY_DOC_ID = Ug.INVENTORY_DOC_ID
      Ivd.QueryFlag = 1
      Call QueryInventoryDoc(Ivd, TempRs2, iCount, IsOK, ErrorObj)
DoEvents
      For Each O In Ivd.ImportExports
         Call glbDatabaseMngr.GetSeqID("TRANSACTION_SEQUENCE_SEQ", TempID, ErrorObj)
         O.TRANSACTION_SEQ = TempID
         Call O.PatchTxSeq
      Next O
         
      Set Ivd = Nothing
      TempRs1.MoveNext
   Wend
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   Prg.Value = 100
   
   Set TempRs1 = Nothing
   Set TempRs2 = Nothing
   
   HasBegin = False
   IsOK = True

   RearrangeTxSeq = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   RearrangeTxSeq = False
End Function

Public Function GetBalanceAmount(PartItemID As Long, LocationID As Long, TxSeq As Long, DocDate As Date) As Object
Dim EI As CExportItem
Dim II As CImportItem
Dim TempRs As ADODB.Recordset
Dim iCount As Long

   Set TempRs = New ADODB.Recordset
   
   Set EI = New CExportItem
   Set II = New CImportItem
   
   EI.EXPORT_ITEM_ID = -1
   EI.PIG_FLAG = "N"
   EI.PART_ITEM_ID = PartItemID
   EI.LOCATION_ID = LocationID
   EI.FROM_TX_SEQ = -1
   EI.TO_TX_SEQ = TxSeq
   EI.FROM_DATE = -1
   EI.TO_DATE = DocDate
   EI.OrderBy = 11
   EI.OrderType = 2
   Call EI.QueryData(1, TempRs, iCount)
   If Not TempRs.EOF Then
      Call EI.PopulateFromRS(1, TempRs)
   End If
   
   II.IMPORT_ITEM_ID = -1
   II.PIG_FLAG = "N"
   II.PART_ITEM_ID = PartItemID
   II.LOCATION_ID = LocationID
   II.FROM_TX_SEQ = -1
   II.TO_TX_SEQ = TxSeq
   II.FROM_DATE = -1
   II.TO_DATE = DocDate
   II.OrderBy = 12
   II.OrderType = 2
   Call II.QueryData(1, TempRs, iCount)
   If Not TempRs.EOF Then
      Call II.PopulateFromRS(1, TempRs)
   End If
   
   If EI.TRANSACTION_SEQ > II.TRANSACTION_SEQ Then
      Set GetBalanceAmount = EI
   Else
      Set GetBalanceAmount = II
   End If
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   Set EI = Nothing
   Set II = Nothing
End Function

Public Function ClearCapitalMovement(Ivd As CInventoryDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cCust As CCapitalMovement
Dim Cl As CCapitalLoss

   RName = "ClearCapitalMovement"
   ClearCapitalMovement = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Set cCust = New CCapitalMovement
   cCust.DELETE_BALANCE_FLAG = Ivd.DELETE_BALANCE_FLAG
   cCust.FROM_DATE = Ivd.FROM_DATE
   cCust.TO_DATE = Ivd.TO_DATE
   cCust.BATCH_ID = Ivd.BATCH_ID
   Call cCust.ClearData
   
   Set Cl = New CCapitalLoss
   Cl.DELETE_BALANCE_FLAG = Ivd.DELETE_BALANCE_FLAG
   Cl.FROM_DATE = Ivd.FROM_DATE
   Cl.TO_DATE = Ivd.TO_DATE
   Cl.BATCH_ID = Ivd.BATCH_ID
   Call Cl.ClearData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   ClearCapitalMovement = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False

   ClearCapitalMovement = True
End Function
Public Function ClearCostAccum(CA As CCost_Accum, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "ClearCostAccum"
   ClearCostAccum = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Call CA.ClearData
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   ClearCostAccum = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False

   ClearCostAccum = True
End Function

Public Sub StartTransaction()
   m_ADOConn.BeginTrans
End Sub

Public Sub CommitTransaction()
   m_ADOConn.CommitTrans
End Sub

Public Sub RollbackTransaction()
   m_ADOConn.RollbackTrans
End Sub

Public Function AddEditPayment(Ug As CPayment, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim PL As CPartLocation
Dim PL2 As CPartLocation
Dim iCount As Long
Dim NewPrice As Double
Dim O As Object

   RName = "AddEditPayment"
   AddEditPayment = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   Dim CstName As CPaymentItem
   For Each CstName In Ug.PaymentItems
      If CstName.Flag = "A" Then
         CstName.PAYMENT_ID = Ug.PAYMENT_ID
         CstName.AddEditMode = SHOW_ADD
         CstName.AddEditData
      ElseIf CstName.Flag = "E" Then
         CstName.AddEditMode = SHOW_EDIT
         CstName.AddEditData
      ElseIf CstName.Flag = "D" Then
         CstName.DeleteData
      End If
   Next CstName

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditPayment = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditPayment = False
End Function

Public Function DeletePayment(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim itemcount As Long
Dim B As CPayment

   RName = "DeletePayment"
   DeletePayment = False
   HasBegin = False

   Set B = New CPayment

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.PAYMENT_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeletePayment = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeletePayment = True
End Function

Public Function QueryPayment(Ua As CPayment, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryPayment"
   QueryPayment = False

   IsOK = True
   Call Ua.QueryData(1, Rs, itemcount)

   If Ua.QueryFlag = 1 Then
      Dim Yw As CPaymentItem
      
      Set Yw = New CPaymentItem
      Yw.PAYMENT_ITEM_ID = -1
      Yw.PAYMENT_ID = Ua.PAYMENT_ID
      Yw.OrderBy = -1
      Yw.OrderType = 1
      Call Yw.QueryData(1, m_Rs1, iCount)
      Set Yw = Nothing
      
      Set Ua.PaymentItems = Nothing
      Set Ua.PaymentItems = New Collection
      While Not m_Rs1.EOF
         Set Yw = New CPaymentItem
         Call Yw.PopulateFromRS(1, m_Rs1)
                  
         Yw.Flag = "I"
         Call Ua.PaymentItems.Add(Yw)
         
         Set Yw = Nothing
         m_Rs1.MoveNext
      Wend
   End If
      
   QueryPayment = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryPayment = False
End Function
Public Function DO2Payment(Bd As CBillingDoc, Pm As CPayment) As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim IsOK As Boolean
Dim Di As CPaymentItem
Dim TempDi As CDoItem
Dim TotalPrice As Double
Dim RevenueTotalPrice As Double
Dim Ri As CReceiptItem

   If (Bd.DOCUMENT_TYPE = 2) And (Bd.RECEIPT_TYPE = 1) Then 'ใบเสร็จ
      TotalPrice = 0
      For Each TempDi In Bd.DoItems
         If TempDi.Flag <> "D" Then
            TotalPrice = TotalPrice + TempDi.TOTAL_PRICE
         End If
      Next TempDi
      
      RevenueTotalPrice = 0
      For Each TempDi In Bd.Revenues
         If TempDi.Flag <> "D" Then
            RevenueTotalPrice = RevenueTotalPrice + TempDi.TOTAL_PRICE
         End If
      Next TempDi
   ElseIf (Bd.DOCUMENT_TYPE = 2) And (Bd.RECEIPT_TYPE = 3) Then 'ใบเสร็จรับชำระ
      TotalPrice = 0
      RevenueTotalPrice = 0
      For Each Ri In Bd.ReceiptItems
         If Ri.Flag <> "D" Then
            TotalPrice = TotalPrice + Ri.PAID_AMOUNT
         End If
      Next Ri
   End If
   
   Set Pm = Nothing
   Set Pm = New CPayment

   If Bd.PAYMENT_ID > 0 Then
      Set TempRs = New ADODB.Recordset
      
      Pm.PAYMENT_ID = Bd.PAYMENT_ID
      Pm.QueryFlag = 1
      Call glbDaily.QueryPayment(Pm, TempRs, iCount, IsOK, glbErrorLog)
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      
      Pm.AddEditMode = SHOW_EDIT
      Set Di = Pm.PaymentItems(1)  'รับประกันว่าต้องมี
      Di.Flag = "E"
   Else
      Pm.AddEditMode = SHOW_ADD
      Set Di = New CPaymentItem
      Di.Flag = "A"
      Call Pm.PaymentItems.Add(Di)
   End If

   If Bd.PAYMENT_TYPE = BANKTRF_PMT Then
      Pm.PAYMENT_DATE = Bd.DOCUMENT_DATE
      Pm.PAYMENT_NO = Bd.DOCUMENT_NO
      Pm.COMMIT_FLAG = Bd.COMMIT_FLAG
      Pm.TO_BANK_BRANCH = Bd.BANK_BRANCH_ID
      Pm.TO_BANK_ID = Bd.BANK_ID
      Pm.ACCOUNT_NO = Bd.CHECK_NO
      Pm.TX_TYPE = "I" 'In, มีเงินเข้ามา
      Pm.INTERNAL_FLAG = "Y"
      Pm.TOTAL_AMOUNT = TotalPrice + RevenueTotalPrice
      
      Di.PAYMENT_ID = Pm.PAYMENT_ID
      Di.PAYMENT_TYPE = Bd.PAYMENT_TYPE
      Di.BANK_BRANCH = -1
      Di.BANK_ID = -1
      Di.BANK_NOTE = Bd.BANK_NOTE
      Di.CHECK_NO = ""
      Di.PAY_AMOUNT = TotalPrice + RevenueTotalPrice
   Else
      Pm.PAYMENT_DATE = Bd.DOCUMENT_DATE
      Pm.PAYMENT_NO = Bd.DOCUMENT_NO
      Pm.COMMIT_FLAG = Bd.COMMIT_FLAG
      Pm.TO_BANK_BRANCH = -1
      Pm.TO_BANK_ID = -1
      Pm.ACCOUNT_NO = ""
      Pm.TX_TYPE = "I" 'In, มีเงินเข้ามา
      Pm.INTERNAL_FLAG = "Y"
      Pm.TOTAL_AMOUNT = TotalPrice + RevenueTotalPrice
      
      Di.PAYMENT_ID = Pm.PAYMENT_ID
      Di.PAYMENT_TYPE = Bd.PAYMENT_TYPE
      Di.BANK_BRANCH = Bd.BANK_BRANCH_ID
      Di.BANK_ID = Bd.BANK_ID
      Di.BANK_NOTE = Bd.BANK_NOTE
      Di.CHECK_NO = Bd.CHECK_NO
      Di.PAY_AMOUNT = TotalPrice + RevenueTotalPrice
   End If
End Function
Public Function DO2InventoryDoc(Bd As CBillingDoc, Ivd As CInventoryDoc, Area As Long, IvdDocType As Long) As Boolean
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim IsOK As Boolean
Dim Di As CDoItem
Dim EI As CLotItem

   Set Ivd = Nothing
   Set Ivd = New CInventoryDoc

   If Bd.INVENTORY_DOC_ID > 0 Then
      Set TempRs = New ADODB.Recordset
      
      Ivd.INVENTORY_DOC_ID = Bd.INVENTORY_DOC_ID
      Ivd.QueryFlag = 1
      Call glbDaily.QueryInventoryDoc(Ivd, TempRs, iCount, IsOK, glbErrorLog)
      
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
      
      Ivd.AddEditMode = SHOW_EDIT
   Else
      Ivd.AddEditMode = SHOW_ADD
   End If
      
   Ivd.DOCUMENT_DATE = Bd.DOCUMENT_DATE
   Ivd.DOCUMENT_NO = Bd.DOCUMENT_NO
   Ivd.COMMIT_FLAG = Bd.COMMIT_FLAG
   Ivd.DOCUMENT_TYPE = IvdDocType
   Ivd.EXCEPTION_FLAG = "N"
   
   For Each Di In Bd.DoItems
      If Di.Flag = "A" Then
         Set EI = New CLotItem
         
         If Area = 1 Then
            EI.TX_TYPE = "E"
            EI.CALCULATE_FLAG = "N"
         ElseIf Area = 2 Then
            EI.TX_TYPE = "I"
            EI.CALCULATE_FLAG = "Y"
         End If
         EI.Flag = "A"
         EI.PART_ITEM_ID = Di.PART_ITEM_ID
         EI.LOCATION_ID = Di.LOCATION_ID
         EI.LOCATION_NAME = Di.LOCATION_NAME
         EI.TX_AMOUNT = Di.ITEM_AMOUNT
         EI.TOTAL_WEIGHT = Di.TOTAL_WEIGHT
         EI.SALE_TOT_PRICE = Di.TOTAL_PRICE
'         Ei.INCLUDE_UNIT_PRICE = MyDiff(Ei.SALE_TOT_PRICE, Ei.TX_AMOUNT)
         EI.LINK_ID = Di.LINK_ID
         
         Set EI.SubLotItems = Di.SubLotItems
         If Di.PART_ITEM_ID > 0 Then 'เอาเฉพาะที่มีการเบิกวัตถุดิบด้วย
            Call Ivd.ImportExports.Add(EI)
         End If
         Set EI = Nothing
      ElseIf Di.Flag = "E" Then
         Set EI = GetExportItem(Ivd, Di.LINK_ID)
         If EI Is Nothing Then
            'มีข้อมูลใน DoItems แต่ไม่มีใน ImportExports แสดงว่าไม่ได้เบิกเป็นวัตถุดิบ
            If Di.PART_ITEM_ID > 0 Then 'ของใหม่มีการเบิกวัตถุดิบ
               Set EI = New CLotItem
               If Area = 1 Then
                  EI.TX_TYPE = "E"
                  EI.CALCULATE_FLAG = "N"
               ElseIf Area = 2 Then
                  EI.TX_TYPE = "I"
                  EI.CALCULATE_FLAG = "Y"
               End If
               EI.Flag = "E"
               EI.PART_ITEM_ID = Di.PART_ITEM_ID
               EI.LOCATION_ID = Di.LOCATION_ID
               EI.LOCATION_NAME = Di.LOCATION_NAME
               EI.TX_AMOUNT = Di.ITEM_AMOUNT
               EI.TOTAL_WEIGHT = Di.TOTAL_WEIGHT
               EI.SALE_TOT_PRICE = Di.TOTAL_PRICE
               EI.INCLUDE_UNIT_PRICE = MyDiff(EI.SALE_TOT_PRICE, EI.TX_AMOUNT)
               EI.LINK_ID = Di.LINK_ID
               Set EI.SubLotItems = Di.SubLotItems
               Call Ivd.ImportExports.Add(EI)
               Set EI = Nothing
            End If
         Else 'มีข้อมูลอยู่ใน ImportExports อยู่แล้ว
            If Di.PART_ITEM_ID > 0 Then
               'แก้ไขปกติ
               EI.Flag = Di.Flag
               EI.PART_ITEM_ID = Di.PART_ITEM_ID
               EI.LOCATION_ID = Di.LOCATION_ID
               EI.LOCATION_NAME = Di.LOCATION_NAME
               EI.TX_AMOUNT = Di.ITEM_AMOUNT
               EI.TOTAL_WEIGHT = Di.TOTAL_WEIGHT
               EI.SALE_TOT_PRICE = Di.TOTAL_PRICE
               EI.CALCULATE_FLAG = "N"
               
'               Set Ei.SubLotItems = Di.SubLotItems
            Else 'ถูกแก้จากเบิกวัตถุดิบเป็นไม่มีการเบิก
               EI.Flag = "D"
            End If
         End If
      ElseIf Di.Flag = "D" Then
         Set EI = GetExportItem(Ivd, Di.LINK_ID)
         If Not (EI Is Nothing) Then
            EI.Flag = "D"
         End If
      End If
   Next Di
End Function

Public Function PatchPigAge(Ug As CExportItem, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim EI As CExportItem
Dim PigAge As Double

   RName = "PatchPigAge"
   PatchPigAge = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Ug.DOCUMENT_TYPE = 10 'Patch แต่เฉพาะจากบิลขาย
   Ug.EXPORT_ITEM_ID = -1
   Call Ug.QueryData(1, m_Rs1, iCount)
   While Not m_Rs1.EOF
      Set EI = New CExportItem
      Call EI.PopulateFromRS(1, m_Rs1)
      
      EI.PIG_AGE = GetAge(EI.PART_NO, EI.DOCUMENT_DATE)
      Call EI.PatchPigAge
      
      Set EI = Nothing
      m_Rs1.MoveNext
   Wend
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   PatchPigAge = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   PatchPigAge = False
End Function

Public Function PatchPigAgeByExportFood(Ug As CExportItem, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim EI As CExportItem
Dim PigAge As Double

   RName = "PatchPigAgeByExportFood"
   PatchPigAgeByExportFood = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Ug.DOCUMENT_TYPE = 2 'Patch แต่เฉพาะจากเบิกอาหารให้หมู
   Ug.EXPORT_ITEM_ID = -1
   Call Ug.QueryData(1, m_Rs1, iCount)
   While Not m_Rs1.EOF
      Set EI = New CExportItem
      Call EI.PopulateFromRS(1, m_Rs1)
      
      EI.PIG_AGE = GetAge(EI.PIG_NO, EI.DOCUMENT_DATE)
      Call EI.PatchPigAge
      
      Set EI = Nothing
      m_Rs1.MoveNext
   Wend
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   PatchPigAgeByExportFood = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   PatchPigAgeByExportFood = False
End Function

Public Function PatchPigBirthPartID(Ug As CImportItem, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim II As CImportItem
Dim Pi As CPartItem
Dim TempRs As ADODB.Recordset

   RName = "PatchPigBirthPartID"
   PatchPigBirthPartID = False
   HasBegin = False

   Set TempRs = New ADODB.Recordset
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Ug.DOCUMENT_TYPE = 5 'ใบเกิดสุกร
   Ug.IMPORT_ITEM_ID = -1
   Call Ug.QueryData(1, TempRs, iCount)
   While Not TempRs.EOF
      Set II = New CImportItem
      Call II.PopulateFromRS(1, TempRs)
'If II.PART_ITEM_ID = 11505 Then
'''debug.print
'End If
      Set Pi = glbDaily.DateToPartItem(II.BIRTH_DATE, PigCodeToID(II.PIG_TYPE))
      If Not (Pi Is Nothing) Then
         II.PART_ITEM_ID = Pi.PART_ITEM_ID
         Call II.PatchPartItemID
      End If
         
      Set II = Nothing
      TempRs.MoveNext
   Wend
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   If TempRs.State = adStateOpen Then
      Set TempRs = Nothing
   End If
   Set TempRs = Nothing
   
   PatchPigBirthPartID = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   PatchPigBirthPartID = False
End Function

Public Function GetPreviousLocationID(Col As Collection, LocationType As Long) As Long
Dim O As Object
Dim TempID As Long

   TempID = -1
   For Each O In Col
      If LocationType = 1 Then
         TempID = O.LOCATION_ID
      ElseIf LocationType = 2 Then
         TempID = O.HOUSE_ID
      ElseIf LocationType = 3 Then
         TempID = O.ExportItem.LOCATION_ID
      End If
   Next O
   
   GetPreviousLocationID = TempID
End Function

Public Function QueryParameter(Ua As CParameter, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long
Dim bIsOk As Boolean

   RName = "QueryParameter"
   QueryParameter = False

   IsOK = True

   Call Ua.QueryData(1, Rs, itemcount)
   
   If Ua.QueryFlag = 1 Then
      Dim Gr As CUsedPrmItem
      Set Gr = New CUsedPrmItem
      Call Gr.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      Call Gr.QueryData(1, m_Rs1, iCount)
      Set Gr = Nothing

      Set Ua.UsedPrmItems = Nothing
      Set Ua.UsedPrmItems = New Collection

      While Not m_Rs1.EOF
         Set Gr = New CUsedPrmItem
         Call Gr.PopulateFromRS(1, m_Rs1)
         Call Ua.UsedPrmItems.Add(Gr)
         Set Gr = Nothing

         m_Rs1.MoveNext
      Wend
      
      Dim Si As CSalePrmItem
      Set Si = New CSalePrmItem
      Call Si.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      Call Si.QueryData(1, m_Rs1, iCount)
      Set Si = Nothing

      Set Ua.SalePrmItems = Nothing
      Set Ua.SalePrmItems = New Collection

      While Not m_Rs1.EOF
         Set Si = New CSalePrmItem
         Call Si.PopulateFromRS(1, m_Rs1)
         Call Ua.SalePrmItems.Add(Si)
         Set Si = Nothing

         m_Rs1.MoveNext
      Wend
   
      Dim Wi As CWeightPrmItem
      Set Wi = New CWeightPrmItem
      Call Wi.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      Call Wi.QueryData(1, m_Rs1, iCount)
      Set Wi = Nothing

      Set Ua.WeightPrmItems = Nothing
      Set Ua.WeightPrmItems = New Collection

      While Not m_Rs1.EOF
         Set Wi = New CWeightPrmItem
         Call Wi.PopulateFromRS(1, m_Rs1)
         Call Ua.WeightPrmItems.Add(Wi)
         Set Wi = Nothing

         m_Rs1.MoveNext
      Wend
   
      Dim Ti As CTrnPrmItem
      Set Ti = New CTrnPrmItem
      Call Ti.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      Call Ti.QueryData(1, m_Rs1, iCount)
      Set Ti = Nothing

      Set Ua.TrnPrmItems = Nothing
      Set Ua.TrnPrmItems = New Collection

      While Not m_Rs1.EOF
         Set Ti = New CTrnPrmItem
         Call Ti.PopulateFromRS(1, m_Rs1)
         Call Ua.TrnPrmItems.Add(Ti)
         Set Ti = Nothing

         m_Rs1.MoveNext
      Wend
      
      Dim Ci As CCostPrmItem
      Set Ci = New CCostPrmItem
      Call Ci.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      Call Ci.QueryData(1, m_Rs1, iCount)
      Set Ci = Nothing

      Set Ua.CostPrmItems = Nothing
      Set Ua.CostPrmItems = New Collection

      While Not m_Rs1.EOF
         Set Ci = New CCostPrmItem
         Call Ci.PopulateFromRS(1, m_Rs1)
         Call Ua.CostPrmItems.Add(Ci)
         Set Ci = Nothing

         m_Rs1.MoveNext
      Wend
      
      Dim Ai As CAdjPrmItem
      Set Ai = New CAdjPrmItem
      Call Ai.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      Call Ai.QueryData(1, m_Rs1, iCount)
      Set Ai = Nothing

      Set Ua.AdjPrmItems = Nothing
      Set Ua.AdjPrmItems = New Collection

      While Not m_Rs1.EOF
         Set Ai = New CAdjPrmItem
         Call Ai.PopulateFromRS(1, m_Rs1)
         Call Ua.AdjPrmItems.Add(Ai)
         Set Ai = Nothing

         m_Rs1.MoveNext
      Wend
      
      Dim Bi As CBrtPrmItem
      Set Bi = New CBrtPrmItem
      Call Bi.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      Call Bi.QueryData(1, m_Rs1, iCount)
      Set Bi = Nothing
   
      Set Ua.BrtPrmItems = Nothing
      Set Ua.BrtPrmItems = New Collection

      While Not m_Rs1.EOF
         Set Bi = New CBrtPrmItem
         Call Bi.PopulateFromRS(1, m_Rs1)
         Call Ua.BrtPrmItems.Add(Bi)
         Set Bi = Nothing

         m_Rs1.MoveNext
      Wend
   
      Dim Ri As CRvnPrmItem
      Set Ri = New CRvnPrmItem
      Call Ri.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      Call Ri.QueryData(1, m_Rs1, iCount)
      Set Ri = Nothing
   
      Set Ua.RvnPrmItems = Nothing
      Set Ua.RvnPrmItems = New Collection

      While Not m_Rs1.EOF
         Set Ri = New CRvnPrmItem
         Call Ri.PopulateFromRS(1, m_Rs1)
         Call Ua.RvnPrmItems.Add(Ri)
         Set Ri = Nothing

         m_Rs1.MoveNext
      Wend
   
      Dim Pi1 As CParamItem
      Set Pi1 = New CParamItem
      Call Pi1.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      Call Pi1.SetFieldValue("PARAM_AREA", 10)
      Call Pi1.QueryData(1, m_Rs1, iCount)
      Set Pi1 = Nothing
   
      Set Ua.CustRatioItems = Nothing
      Set Ua.CustRatioItems = New Collection

      While Not m_Rs1.EOF
         Set Pi1 = New CParamItem
         Call Pi1.PopulateFromRS(1, m_Rs1)
         Call Ua.CustRatioItems.Add(Pi1)
         Set Pi1 = Nothing

         m_Rs1.MoveNext
      Wend
   
      Dim Pi2 As CParamItem
      Set Pi2 = New CParamItem
      Call Pi2.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      Call Pi2.SetFieldValue("PARAM_AREA", 11)
      Call Pi2.QueryData(1, m_Rs1, iCount)
      Set Pi2 = Nothing
   
      Set Ua.PigStatusChangeItems = Nothing
      Set Ua.PigStatusChangeItems = New Collection

      While Not m_Rs1.EOF
         Set Pi2 = New CParamItem
         Call Pi2.PopulateFromRS(1, m_Rs1)
         Call Ua.PigStatusChangeItems.Add(Pi2)
         Set Pi2 = Nothing

         m_Rs1.MoveNext
      Wend
   
      Dim Pi3 As CParamItem
      Set Pi3 = New CParamItem
      Call Pi3.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      Call Pi3.SetFieldValue("PARAM_AREA", 12)
      Call Pi3.SetFieldValue("ORDER_BY", 3)
      Call Pi3.QueryData(1, m_Rs1, iCount)
      Set Pi3 = Nothing
   
      Set Ua.PigBuyItems = Nothing
      Set Ua.PigBuyItems = New Collection

      While Not m_Rs1.EOF
         Set Pi3 = New CParamItem
         Call Pi3.PopulateFromRS(1, m_Rs1)
         Call Ua.PigBuyItems.Add(Pi3)
         Set Pi3 = Nothing

         m_Rs1.MoveNext
      Wend
   
      Dim Pi4 As CParamItem
      Set Pi4 = New CParamItem
      Call Pi4.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      Call Pi4.SetFieldValue("PARAM_AREA", 13)
      Call Pi4.QueryData(1, m_Rs1, iCount)
      Set Pi4 = Nothing
   
      Set Ua.ExpenseSharings = Nothing
      Set Ua.ExpenseSharings = New Collection

      While Not m_Rs1.EOF
         Set Pi4 = New CParamItem
         Call Pi4.PopulateFromRS(1, m_Rs1)
         Call Ua.ExpenseSharings.Add(Pi4)
         Set Pi4 = Nothing

         m_Rs1.MoveNext
      Wend
   
      Dim Pi5 As CParamItem
      Set Pi5 = New CParamItem
      Call Pi5.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      Call Pi5.SetFieldValue("PARAM_AREA", 14)
      Call Pi5.QueryData(1, m_Rs1, iCount)
      Set Pi5 = Nothing
   
      Set Ua.PigAdjustItems = Nothing
      Set Ua.PigAdjustItems = New Collection

      While Not m_Rs1.EOF
         Set Pi5 = New CParamItem
         Call Pi5.PopulateFromRS(1, m_Rs1)
         Call Ua.PigAdjustItems.Add(Pi5)
         Set Pi5 = Nothing

         m_Rs1.MoveNext
      Wend
      
      Dim Pi6 As CParamItem
      Set Pi6 = New CParamItem
      Call Pi6.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      Call Pi6.SetFieldValue("PARAM_AREA", 15)
      Call Pi6.QueryData(1, m_Rs1, iCount)
      Set Pi6 = Nothing
   
      Set Ua.MenagementExpenses = Nothing
      Set Ua.MenagementExpenses = New Collection

      While Not m_Rs1.EOF
         Set Pi6 = New CParamItem
         Call Pi6.PopulateFromRS(1, m_Rs1)
         Call Ua.MenagementExpenses.Add(Pi6)
         Set Pi6 = Nothing

         m_Rs1.MoveNext
      Wend
      
      
      Dim GLa As CGLAgeAmount
      Set GLa = New CGLAgeAmount
      GLa.PARAM_ID = Ua.GetFieldValue("PARAM_ID")
      Call GLa.QueryData(1, m_Rs1, iCount)
      Set GLa = Nothing
   
      Set Ua.Glages = Nothing
      Set Ua.Glages = New Collection
      
      While Not m_Rs1.EOF
         Set GLa = New CGLAgeAmount
         Call GLa.PopulateFromRS(1, m_Rs1)
         Call Ua.Glages.Add(GLa)
         Set GLa = Nothing

         m_Rs1.MoveNext
      Wend
      
      Dim GLb As CGLBackAmount
      Set GLb = New CGLBackAmount
      GLb.PARAM_ID = Ua.GetFieldValue("PARAM_ID")
      Call GLb.QueryData(1, m_Rs1, iCount)
      Set GLb = Nothing
   
      Set Ua.GLbacks = Nothing
      Set Ua.GLbacks = New Collection
      
      While Not m_Rs1.EOF
         Set GLb = New CGLBackAmount
         Call GLb.PopulateFromRS(1, m_Rs1)
         Call Ua.GLbacks.Add(GLb)
         Set GLb = Nothing

         m_Rs1.MoveNext
      Wend
      
   End If
   
   QueryParameter = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryParameter = False
End Function

Public Function DeleteParameter(Ug As CParameter, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteParameter"
   DeleteParameter = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteParameter = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteParameter = True
End Function

Public Function AddEditParameter(Ua As CParameter, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditParameter"
   AddEditParameter = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   Dim Gr As CUsedPrmItem
   For Each Gr In Ua.UsedPrmItems
      Call Gr.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      If Gr.Flag = "A" Then
         Gr.ShowMode = SHOW_ADD
         Call Gr.AddEditData
      ElseIf Gr.Flag = "E" Then
         Gr.ShowMode = SHOW_EDIT
         Call Gr.AddEditData
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If
   Next Gr
   
   Dim Si As CSalePrmItem
   For Each Si In Ua.SalePrmItems
      Call Si.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      If Si.Flag = "A" Then
         Si.ShowMode = SHOW_ADD
         Call Si.AddEditData
      ElseIf Si.Flag = "E" Then
         Si.ShowMode = SHOW_EDIT
         Call Si.AddEditData
      ElseIf Si.Flag = "D" Then
         Call Si.DeleteData
      End If
   Next Si
   
   Dim Wi As CWeightPrmItem
   For Each Wi In Ua.WeightPrmItems
      Call Wi.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      If Wi.Flag = "A" Then
         Wi.ShowMode = SHOW_ADD
         Call Wi.AddEditData
      ElseIf Wi.Flag = "E" Then
         Wi.ShowMode = SHOW_EDIT
         Call Wi.AddEditData
      ElseIf Wi.Flag = "D" Then
         Call Wi.DeleteData
      End If
   Next Wi
   
   Dim Ti As CTrnPrmItem
   For Each Ti In Ua.TrnPrmItems
      Call Ti.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      If Ti.Flag = "A" Then
         Ti.ShowMode = SHOW_ADD
         Call Ti.AddEditData
      ElseIf Ti.Flag = "E" Then
         Ti.ShowMode = SHOW_EDIT
         Call Ti.AddEditData
      ElseIf Ti.Flag = "D" Then
         Call Ti.DeleteData
      End If
   Next Ti
   
   Dim Ci As CCostPrmItem
   For Each Ci In Ua.CostPrmItems
      Call Ci.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      If Ci.Flag = "A" Then
         Ci.ShowMode = SHOW_ADD
         Call Ci.AddEditData
      ElseIf Ci.Flag = "E" Then
         Ci.ShowMode = SHOW_EDIT
         Call Ci.AddEditData
      ElseIf Ci.Flag = "D" Then
         Call Ci.DeleteData
      End If
   Next Ci
   
   Dim Ai As CAdjPrmItem
   For Each Ai In Ua.AdjPrmItems
      Call Ai.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      If Ai.Flag = "A" Then
         Ai.ShowMode = SHOW_ADD
         Call Ai.AddEditData
      ElseIf Ai.Flag = "E" Then
         Ai.ShowMode = SHOW_EDIT
         Call Ai.AddEditData
      ElseIf Ai.Flag = "D" Then
         Call Ai.DeleteData
      End If
   Next Ai
   
   Dim Bi As CBrtPrmItem
   For Each Bi In Ua.BrtPrmItems
      Call Bi.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      If Bi.Flag = "A" Then
         Bi.ShowMode = SHOW_ADD
         Call Bi.AddEditData
      ElseIf Bi.Flag = "E" Then
         Bi.ShowMode = SHOW_EDIT
         Call Bi.AddEditData
      ElseIf Bi.Flag = "D" Then
         Call Bi.DeleteData
      End If
   Next Bi
   
   Dim Ri As CRvnPrmItem
   For Each Ri In Ua.RvnPrmItems
      Call Ri.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      If Ri.Flag = "A" Then
         Ri.ShowMode = SHOW_ADD
         Call Ri.AddEditData
      ElseIf Ri.Flag = "E" Then
         Ri.ShowMode = SHOW_EDIT
         Call Ri.AddEditData
      ElseIf Ri.Flag = "D" Then
         Call Ri.DeleteData
      End If
   Next Ri
      
   Dim Pi1 As CParamItem
   For Each Pi1 In Ua.CustRatioItems
      Call Pi1.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      If Pi1.Flag = "A" Then
         Pi1.ShowMode = SHOW_ADD
         Call Pi1.AddEditData
      ElseIf Pi1.Flag = "E" Then
         Pi1.ShowMode = SHOW_EDIT
         Call Pi1.AddEditData
      ElseIf Pi1.Flag = "D" Then
         Call Pi1.DeleteData
      End If
   Next Pi1
      
   Dim Pi2 As CParamItem
   For Each Pi2 In Ua.PigStatusChangeItems
      Call Pi2.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      If Pi2.Flag = "A" Then
         Pi2.ShowMode = SHOW_ADD
         Call Pi2.AddEditData
      ElseIf Pi2.Flag = "E" Then
         Pi2.ShowMode = SHOW_EDIT
         Call Pi2.AddEditData
      ElseIf Pi2.Flag = "D" Then
         Call Pi2.DeleteData
      End If
   Next Pi2

   Dim Pi3 As CParamItem
   For Each Pi3 In Ua.PigBuyItems
      Call Pi3.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      If Pi3.Flag = "A" Then
         Pi3.ShowMode = SHOW_ADD
         Call Pi3.AddEditData
      ElseIf Pi3.Flag = "E" Then
         Pi3.ShowMode = SHOW_EDIT
         Call Pi3.AddEditData
      ElseIf Pi3.Flag = "D" Then
         Call Pi3.DeleteData
      End If
   Next Pi3
      
   Dim Pi4 As CParamItem
   For Each Pi4 In Ua.ExpenseSharings
      Call Pi4.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      If Pi4.Flag = "A" Then
         Pi4.ShowMode = SHOW_ADD
         Call Pi4.AddEditData
      ElseIf Pi4.Flag = "E" Then
         Pi4.ShowMode = SHOW_EDIT
         Call Pi4.AddEditData
      ElseIf Pi4.Flag = "D" Then
         Call Pi4.DeleteData
      End If
   Next Pi4
      
   Dim Pi5 As CParamItem
   For Each Pi5 In Ua.PigAdjustItems
      Call Pi5.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      If Pi5.Flag = "A" Then
         Pi5.ShowMode = SHOW_ADD
         Call Pi5.AddEditData
      ElseIf Pi5.Flag = "E" Then
         Pi5.ShowMode = SHOW_EDIT
         Call Pi5.AddEditData
      ElseIf Pi5.Flag = "D" Then
         Call Pi5.DeleteData
      End If
   Next Pi5
      
   Dim Pi6 As CParamItem
   For Each Pi6 In Ua.MenagementExpenses
      Call Pi6.SetFieldValue("PARAM_ID", Ua.GetFieldValue("PARAM_ID"))
      If Pi6.Flag = "A" Then
         Pi6.ShowMode = SHOW_ADD
         Call Pi6.AddEditData
      ElseIf Pi6.Flag = "E" Then
         Pi6.ShowMode = SHOW_EDIT
         Call Pi6.AddEditData
      ElseIf Pi6.Flag = "D" Then
         Call Pi6.DeleteData
      End If
   Next Pi6
      
   Dim GLa As CGLAgeAmount
   For Each GLa In Ua.Glages
      GLa.PARAM_ID = Ua.GetFieldValue("PARAM_ID")
      If GLa.Flag = "A" Then
         GLa.AddEditMode = SHOW_ADD
         Call GLa.AddEditData
      ElseIf GLa.Flag = "E" Then
         GLa.AddEditMode = SHOW_EDIT
         Call GLa.AddEditData
      ElseIf GLa.Flag = "D" Then
         Call GLa.DeleteData
      End If
   Next GLa
   
   Dim GLb As CGLBackAmount
   For Each GLb In Ua.GLbacks
      GLb.PARAM_ID = Ua.GetFieldValue("PARAM_ID")
      If GLb.Flag = "A" Then
         GLb.AddEditMode = SHOW_ADD
         Call GLb.AddEditData
      ElseIf GLb.Flag = "E" Then
         GLb.AddEditMode = SHOW_EDIT
         Call GLb.AddEditData
      ElseIf GLb.Flag = "D" Then
         Call GLb.DeleteData
      End If
   Next GLb
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditParameter = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditParameter = False
End Function

Public Function QueryBatch(Ua As CBatch, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long
Dim bIsOk As Boolean
Dim I As Long
   
   RName = "QueryBatch"
   QueryBatch = False

   IsOK = True

   Call Ua.QueryData(1, Rs, itemcount)
   
   If Ua.QueryFlag = 1 Then
      Dim Gr As CBatchItem
      Set Gr = New CBatchItem
      Call Gr.SetFieldValue("BATCH_ID", Ua.GetFieldValue("BATCH_ID"))
      Call Gr.QueryData(1, m_Rs1, iCount)
      Set Gr = Nothing

      Set Ua.BatchItems = Nothing
      Set Ua.BatchItems = New Collection
      
      Set Ua.BirthItems = Nothing
      Set Ua.BirthItems = New Collection
      Set Ua.TransferItems = Nothing
      Set Ua.TransferItems = New Collection
      Set Ua.SaleItems = Nothing
      Set Ua.SaleItems = New Collection
      Set Ua.WeightItems = Nothing
      Set Ua.WeightItems = New Collection
      Set Ua.FoodItems = Nothing
      Set Ua.FoodItems = New Collection
      Set Ua.Feeds = Nothing
      Set Ua.Feeds = New Collection
      Set Ua.Balances = Nothing
      Set Ua.Balances = New Collection
      Set Ua.CustRatios = Nothing
      Set Ua.CustRatios = New Collection
      Set Ua.ChangePigTypes = Nothing
      Set Ua.ChangePigTypes = New Collection
      Set Ua.BuyItems = Nothing
      Set Ua.BuyItems = New Collection
      Set Ua.ExpenseSharingItems = Nothing
      Set Ua.ExpenseSharingItems = New Collection
      Set Ua.ManagementExpenses = Nothing
      Set Ua.ManagementExpenses = New Collection
      Set Ua.Glages = Nothing
      Set Ua.Glages = New Collection
      Set Ua.GLbacks = Nothing
      Set Ua.GLbacks = New Collection
      
      I = 0
      
      While Not m_Rs1.EOF
         Set Gr = New CBatchItem
         Call Gr.PopulateFromRS(1, m_Rs1)
         I = I + 1
         Call GetCollectionFromType(Ua, Gr.GetFieldValue("PARAM_AREA")).Add(Gr)
         Set Gr = Nothing

         m_Rs1.MoveNext
      Wend
   End If
   
   QueryBatch = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryBatch = False
End Function

Public Function AddEditBatch(Ua As CBatch, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditBatch"
   AddEditBatch = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   Dim Gr As CBatchItem
   For Each Gr In Ua.BatchItems
      Call Gr.SetFieldValue("BATCH_ID", Ua.GetFieldValue("BATCH_ID"))
      If Gr.Flag = "A" Then
         Gr.ShowMode = SHOW_ADD
         Call Gr.AddEditData
      ElseIf Gr.Flag = "E" Then
         Gr.ShowMode = SHOW_EDIT
         Call Gr.AddEditData
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If
   Next Gr
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditBatch = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditBatch = False
End Function

Public Function DeleteBatch(Ug As CBatch, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteBatch"
   DeleteBatch = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteBatch = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteBatch = True
End Function

Public Function DeleteDocument(BatchID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog, Optional DeleteAll As Boolean = False)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim Ug As CBatch

   RName = "DeleteDocument"
   DeleteDocument = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Set Ug = New CBatch
   Call Ug.SetFieldValue("BATCH_ID", BatchID)
   Call Ug.DeleteBatchDocument(DeleteAll)

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteDocument = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteDocument = True
End Function

Public Function DeletePackage(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim itemcount As Long

Dim B As CPackage

   RName = "DeletePackage"
   DeletePackage = False
   HasBegin = False

   Set B = New CPackage

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.PKG_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeletePackage = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeletePackage = True
End Function

Public Function AddEditPackage(Ug As CPackage, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditPackage"
   AddEditPackage = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   Dim CstName As CPackageDetail
   For Each CstName In Ug.PackageDetail
      If CstName.Flag = "A" Then
         CstName.PKG_ID = Ug.PKG_ID
         CstName.AddEditMode = SHOW_ADD
         CstName.AddEditData
      ElseIf CstName.Flag = "E" Then
         CstName.AddEditMode = SHOW_EDIT
         CstName.AddEditData
      ElseIf CstName.Flag = "D" Then
         CstName.DeleteData
      End If
   Next CstName
         
         
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   Call LoadCustomerPackage(Nothing, CustomerPackage)
   Call LoadPackageDetail(Nothing, PackageDetail)
   
   HasBegin = False
   IsOK = True

   AddEditPackage = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditPackage = False
End Function

Public Function QueryPackage(Ua As CPackage, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryPackage"
   QueryPackage = False

   IsOK = True
   Call Ua.QueryData(1, Rs, itemcount)

   If Ua.QueryFlag = 1 Then
      Dim Yw As CPackageDetail
      
      Set Yw = New CPackageDetail
      Yw.PKG_DETAIL_ID = -1
      Yw.PKG_ID = Ua.PKG_ID
      Yw.OrderBy = -1
      Yw.OrderType = 1
      Call Yw.QueryData(1, m_Rs1, iCount)
      Set Yw = Nothing
      
      Set Ua.PackageDetail = Nothing
      Set Ua.PackageDetail = New Collection
      While Not m_Rs1.EOF
         Set Yw = New CPackageDetail
         Call Yw.PopulateFromRS(1, m_Rs1)
        
         Yw.Flag = "I"
         
            Call Ua.PackageDetail.Add(Yw)
         
         Set Yw = Nothing
         m_Rs1.MoveNext
      Wend
      
      
      '===
   End If
      
   QueryPackage = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryPackage = False
End Function

Public Function CopyPackage(Ug As CPackage, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim itemcount As Long

   RName = "CopyPackage"
   CopyPackage = False
   HasBegin = False
   
   Set TempRs = New ADODB.Recordset

   Ug.QueryFlag = 1
   Ug.OrderType = 1
   Call QueryPackage(Ug, TempRs, iCount, IsOK, ErrorObj)
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   Ug.AddEditMode = SHOW_ADD
   Ug.PKG_ID = -1
   Ug.PKG_BASIC_FLAG = "N"
   Ug.PKG_CODE = "**" & Ug.PKG_CODE
   
   Dim PkDetail As CPackageDetail
   For Each PkDetail In Ug.PackageDetail
      PkDetail.Flag = "A"
      PkDetail.AddEditMode = SHOW_ADD
   Next PkDetail
         

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call AddEditPackage(Ug, IsOK, False, ErrorObj)
      
   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   
   HasBegin = False
   IsOK = True

   CopyPackage = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   CopyPackage = False
End Function

Public Function CopyParameter(Ug As CParameter, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim itemcount As Long

   RName = "CopyParameter"
   CopyParameter = False
   HasBegin = False
   
   Set TempRs = New ADODB.Recordset

   Ug.QueryFlag = 1
   Call QueryParameter(Ug, TempRs, iCount, IsOK, ErrorObj)
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   Ug.ShowMode = SHOW_ADD
   Call Ug.SetFieldValue("PARAM_ID", -1)
   Call Ug.SetFieldValue("COMMIT_FLAG", "N")
   Call Ug.SetFieldValue("PARAM_NO", "**" & Ug.GetFieldValue("PARAM_NO"))
            
   Dim Gr As CUsedPrmItem
   For Each Gr In Ug.UsedPrmItems
      Gr.Flag = "A"
      Gr.ShowMode = SHOW_ADD
   Next Gr
   
   Dim Si As CSalePrmItem
   For Each Si In Ug.SalePrmItems
      Si.Flag = "A"
      Si.ShowMode = SHOW_ADD
   Next Si
   
   Dim Wi As CWeightPrmItem
   For Each Wi In Ug.WeightPrmItems
      Wi.Flag = "A"
      Wi.ShowMode = SHOW_ADD
   Next Wi
   
   Dim Ti As CTrnPrmItem
   For Each Ti In Ug.TrnPrmItems
      Ti.Flag = "A"
      Ti.ShowMode = SHOW_ADD
   Next Ti
   
   Dim Ci As CCostPrmItem
   For Each Ci In Ug.CostPrmItems
      Ci.Flag = "A"
      Ci.ShowMode = SHOW_ADD
   Next Ci
   
   Dim Ai As CAdjPrmItem
   For Each Ai In Ug.AdjPrmItems
      Ai.Flag = "A"
      Ai.ShowMode = SHOW_ADD
   Next Ai
   
   Dim Bi As CBrtPrmItem
   For Each Bi In Ug.BrtPrmItems
      Bi.Flag = "A"
      Bi.ShowMode = SHOW_ADD
   Next Bi
   
   Dim Ri As CRvnPrmItem
   For Each Ri In Ug.RvnPrmItems
      Ri.Flag = "A"
      Ri.ShowMode = SHOW_ADD
   Next Ri
      
   Dim Pi1 As CParamItem
   For Each Pi1 In Ug.CustRatioItems
      Pi1.Flag = "A"
      Pi1.ShowMode = SHOW_ADD
   Next Pi1
      
   Dim Pi2 As CParamItem
   For Each Pi2 In Ug.PigStatusChangeItems
      Pi2.Flag = "A"
      Pi2.ShowMode = SHOW_ADD
   Next Pi2

   Dim Pi3 As CParamItem
   For Each Pi3 In Ug.PigBuyItems
      Pi3.Flag = "A"
      Pi3.ShowMode = SHOW_ADD
   Next Pi3
      
   Dim Pi4 As CParamItem
   For Each Pi4 In Ug.ExpenseSharings
      Pi4.Flag = "A"
      'Call Pi4.SetFieldValue("EXPENSE_DATE", DateAdd("M", 1, Pi4.GetFieldValue("EXPENSE_DATE")))
      Pi4.ShowMode = SHOW_ADD
   Next Pi4
   
   Dim Pi5 As CParamItem
   For Each Pi5 In Ug.PigAdjustItems
      Pi5.Flag = "A"
      Pi5.ShowMode = SHOW_ADD
   Next Pi5
   
   Dim Pi6 As CParamItem
   For Each Pi6 In Ug.MenagementExpenses
      Pi6.Flag = "A"
      Pi6.ShowMode = SHOW_ADD
   Next Pi6
   
   Dim GLa As CGLAgeAmount
   For Each GLa In Ug.Glages
      GLa.Flag = "A"
      GLa.AddEditMode = SHOW_ADD
   Next GLa
   
   Dim GLb As CGLBackAmount
   For Each GLb In Ug.GLbacks
      GLb.Flag = "A"
      GLb.AddEditMode = SHOW_ADD
   Next GLb
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call AddEditParameter(Ug, IsOK, False, ErrorObj)
      
   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   
   HasBegin = False
   IsOK = True

   CopyParameter = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   CopyParameter = False
End Function
Public Function CopyBillingDoc(Ug As CBillingDoc, IsOK As Boolean, AutoCommit As Boolean, Area As Long, IvdDocType As Long, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim PL As CPartLocation
Dim PL2 As CPartLocation
Dim iCount As Long
Dim NewPrice As Double
Dim O As Object
Dim TempRs As ADODB.Recordset
Dim Ivd As CInventoryDoc

   RName = "CopyBillingDoc"
   CopyBillingDoc = False
   HasBegin = False

   Set TempRs = New ADODB.Recordset
   
   Ug.QueryFlag = 1
   Call QueryBillingDoc(Ug, TempRs, iCount, IsOK, ErrorObj)
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   Ug.AddEditMode = SHOW_ADD
   Ug.BILLING_DOC_ID = -1
   Ug.INVENTORY_DOC_ID = -1
   Ug.DOCUMENT_NO = "**" & Ug.DOCUMENT_NO
   Ug.COMMIT_FLAG = "N" 'จะต้องยังไม่ commit ก่อน
   
   Dim CstName As CDoItem
   For Each CstName In Ug.DoItems
      CstName.Flag = "A"
      CstName.AddEditMode = SHOW_ADD
   Next CstName
   
   Dim Rt As CReceiptItem
   For Each Rt In Ug.ReceiptItems
      Rt.Flag = "A"
      Rt.AddEditMode = SHOW_ADD
   Next Rt
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call glbDaily.DO2InventoryDoc(Ug, Ivd, Area, IvdDocType)
   Call AddEditInventoryDoc(Ivd, IsOK, False, ErrorObj)
   Ug.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID
   Call AddEditBillingDoc(Ug, IsOK, False, ErrorObj)
      
   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   HasBegin = False
   IsOK = True

   CopyBillingDoc = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   CopyBillingDoc = False
End Function

Public Function PatchBirthItemParam(AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim Bi As CBrtPrmItem
Dim TempBi As CBrtPrmItem
Dim TempRs1 As ADODB.Recordset
Dim FromDate As Date
Dim ToDate As Date

   RName = "PatchBirthItemParam"
   PatchBirthItemParam = False
   HasBegin = False

   Set TempRs1 = New ADODB.Recordset
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Set Bi = New CBrtPrmItem
   Call Bi.SetFieldValue("BRTPRM_ITEM_ID", -1)
   Call Bi.QueryData(1, TempRs1, iCount)
   
   While Not TempRs1.EOF
      Set TempBi = New CBrtPrmItem
      Call TempBi.PopulateFromRS(1, TempRs1)
      
      TempBi.ShowMode = SHOW_EDIT
      FromDate = DateAdd("D", 115, TempBi.GetFieldValue("FROM_BREED"))
      ToDate = DateAdd("D", 115, TempBi.GetFieldValue("TO_BREED"))
      Call TempBi.SetFieldValue("FROM_BIRTH", FromDate)
      Call TempBi.SetFieldValue("TO_BIRTH", ToDate)
      Call TempBi.AddEditData
      
      Set TempBi = Nothing
      TempRs1.MoveNext
   Wend
   
   Set Bi = Nothing
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   
   Set TempRs1 = Nothing
   
   HasBegin = False

   PatchBirthItemParam = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   PatchBirthItemParam = False
End Function
Public Function QueryBatchParameter(Ua As CBatch, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long
Dim bIsOk As Boolean
Dim TempBatchRs As ADODB.Recordset
Dim TempBatchItemRs As ADODB.Recordset

   RName = "QueryBatchParameter"
   QueryBatchParameter = False

   IsOK = True

   Call Ua.QueryData(1, Rs, itemcount)
   
   If Ua.QueryFlag = 1 Then
      Dim Gr As CBatchItem
      Set Gr = New CBatchItem
      Set TempBatchRs = New ADODB.Recordset
      Call Gr.SetFieldValue("PARAM_AREA", Ua.GetFieldValue("PARAM_AREA"))
      Call Gr.SetFieldValue("BATCH_ID", Ua.GetFieldValue("BATCH_ID"))
      Call Gr.QueryData(1, TempBatchRs, iCount)
      Set Gr = Nothing

      Set Ua.BatchItems = Nothing
      Set Ua.BatchItems = New Collection
      
      Set Ua.BirthItems = Nothing
      Set Ua.BirthItems = New Collection
      Set Ua.TransferItems = Nothing
      Set Ua.TransferItems = New Collection
      Set Ua.SaleItems = Nothing
      Set Ua.SaleItems = New Collection
      Set Ua.WeightItems = Nothing
      Set Ua.WeightItems = New Collection
      Set Ua.FoodItems = Nothing
      Set Ua.FoodItems = New Collection
      Set Ua.Feeds = Nothing
      Set Ua.Feeds = New Collection
      Set Ua.Balances = Nothing
      Set Ua.Balances = New Collection
      Set Ua.CustRatios = Nothing
      Set Ua.CustRatios = New Collection
      Set Ua.ChangePigTypes = Nothing
      Set Ua.ChangePigTypes = New Collection
      Set Ua.BuyItems = Nothing
      Set Ua.BuyItems = New Collection
      Set Ua.ExpenseSharingItems = Nothing
      Set Ua.ExpenseSharingItems = New Collection
      
      Dim Pm As CParameter
      While Not TempBatchRs.EOF
         Set Gr = New CBatchItem
         Set TempBatchItemRs = New ADODB.Recordset
         Call Gr.PopulateFromRS(1, TempBatchRs)
         
         Set Pm = New CParameter
         Call Pm.SetFieldValue("PARAM_ID", Gr.GetFieldValue("PARAM_ID"))
         Pm.QueryFlag = 1
         Call QueryParameter(Pm, TempBatchItemRs, itemcount, IsOK, ErrorObj)
         
         Call Pm.PopulateFromRS(1, TempBatchItemRs)
         
         Set Gr.Parameters = Pm
         Set Pm = Nothing
         Set TempBatchItemRs = Nothing
         
         Call GetCollectionFromType(Ua, Gr.GetFieldValue("PARAM_AREA")).Add(Gr)
         Set Gr = Nothing

         TempBatchRs.MoveNext
      Wend
   End If
   
   QueryBatchParameter = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryBatchParameter = False
End Function
Public Function DeleteCashDoc(Ua As CCashDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteCashDoc"
   DeleteCashDoc = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ua.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCashDoc = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteCashDoc = True
End Function

Public Function QueryCashDoc(Ua As CCashDoc, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryCashDoc"
   QueryCashDoc = False

   IsOK = True

   Call Ua.QueryData(1, Rs, itemcount)
   
   If Ua.QueryFlag = 1 Then
'      Dim Ji As CJournalItem
'      Set Ji = New CJournalItem
'      Call Ji.SetFieldValue("CASH_DOC_ID", Ua.GetFieldValue("CASH_DOC_ID"))
'      Call Ji.QueryData(1, m_Rs1, iCount)
'      Set Ji = Nothing
'
'      Set Ua.JournalItems = Nothing
'      Set Ua.JournalItems = New Collection
'
'      While Not m_Rs1.EOF
'         Set Ji = New CJournalItem
'         Call Ji.PopulateFromRS(1, m_Rs1)
'         Call Ua.JournalItems.add(Ji)
'         Set Ji = Nothing
'
'         m_Rs1.MoveNext
'      Wend
      
      Dim Cq As CCheque
      Dim Ct As CCashTran
      Set Ct = New CCashTran
      Call Ct.SetFieldValue("CASH_DOC_ID", Ua.GetFieldValue("CASH_DOC_ID"))
      Call Ct.QueryData(1, m_Rs1, iCount)
      Set Ct = Nothing

      Set Ua.CashTranItems = Nothing
      Set Ua.CashTranItems = New Collection

      While Not m_Rs1.EOF
         Set Ct = New CCashTran
         Call Ct.PopulateFromRS(1, m_Rs1)
         
         If Ct.GetFieldValue("CHECK_ID") > 0 Then
            Call Ct.Cheque.SetFieldValue("CHEQUE_ID", Ct.GetFieldValue("CHECK_ID"))
            Call Ct.Cheque.QueryData(1, m_Rs2, iCount)
            If Not m_Rs2.EOF Then
               Call Ct.Cheque.PopulateFromRS(1, m_Rs2)
            End If
         End If
         
         Call Ua.CashTranItems.Add(Ct)
         Set Ct = Nothing

         m_Rs1.MoveNext
      Wend
      
      Dim CDP As CCashDocPost
      Set CDP = New CCashDocPost
      Call CDP.SetFieldValue("CASH_DOC_ID", Ua.GetFieldValue("CASH_DOC_ID"))
      Call CDP.QueryData(1, m_Rs1, iCount)
      Set CDP = Nothing

      Set Ua.PostItems = Nothing
      Set Ua.PostItems = New Collection

      While Not m_Rs1.EOF
         Set CDP = New CCashDocPost
         Call CDP.PopulateFromRS(1, m_Rs1)
         
         Call Ua.PostItems.Add(CDP)
         Set CDP = Nothing

         m_Rs1.MoveNext
      Wend

   End If
   
   QueryCashDoc = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryCashDoc = False
End Function
Public Function AddEditCashDoc(Ua As CCashDoc, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim DocType As CASH_DOC_TYPE

   RName = "AddEditCashDoc"
   AddEditCashDoc = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   Dim Cd As CConfigDoc
   If Ua.ShowMode = SHOW_ADD And Ua.GetFieldValue("CONFIG_DOC_TYPE") > 0 Then
      Set Cd = New CConfigDoc
      Call Cd.SetFieldValue("RUNNING_NO", Ua.GetFieldValue("RUNNING_NO"))
      Call Cd.SetFieldValue("LAST_NO", Ua.GetFieldValue("DOCUMENT_NO"))
      Call Cd.SetFieldValue("CONFIG_DOC_TYPE", Ua.GetFieldValue("CONFIG_DOC_TYPE"))
      Call Cd.UpdateRunningNo
   End If
   Set Cd = Nothing
   
'   Dim Ji As CJournalItem
'   For Each Ji In Ua.JournalItems
'      Call Ji.SetFieldValue("CASH_DOC_ID", Ua.GetFieldValue("CASH_DOC_ID"))
'      If Ji.Flag = "A" Then
'         Ji.ShowMode = SHOW_ADD
'         Call Ji.AddEditData
'      ElseIf Ji.Flag = "E" Then
'         Ji.ShowMode = SHOW_EDIT
'         Call Ji.AddEditData
'      ElseIf Ji.Flag = "D" Then
'         Call Ji.DeleteData
'      End If
'   Next Ji

   Dim Ct As CCashTran
   Dim PaymentType As Long
   Dim OldPaymentType As Long
   
   DocType = Ua.GetFieldValue("DOCUMENT_TYPE")
   
   For Each Ct In Ua.CashTranItems
      PaymentType = Ct.GetFieldValue("PAYMENT_TYPE")
      OldPaymentType = Ct.GetFieldValue("OLD_PAYMENT_TYPE")
      Call Ct.SetFieldValue("CASH_DOC_ID", Ua.GetFieldValue("CASH_DOC_ID"))
      If (Ct.Flag <> "A") And ((Ct.GetFieldValue("TX_DATE") <> Ua.GetFieldValue("DOCUMENT_DATE")) Or (Ct.GetFieldValue("TX_NO") <> Ua.GetFieldValue("DOCUMENT_NO"))) Then
         Ct.Flag = "E"
      End If
      Call Ct.SetFieldValue("TX_NO", Ua.GetFieldValue("DOCUMENT_NO"))
      Call Ct.SetFieldValue("TX_DATE", Ua.GetFieldValue("DOCUMENT_DATE"))
      Call Ct.SetFieldValue("CUSTOMER_ID", Ua.GetFieldValue("CUSTOMER_ID"))
      If Ct.Flag = "A" Then
         Ct.ShowMode = SHOW_ADD
         If DocType = CASH_DEPOSIT Then 'ฝากเช็คที่มีอยู่ในระบบแล้ว
            Call Ct.Cheque.SetFieldValue("CHEQUE_ID", Ct.GetFieldValue("CHECK_ID"))
            Call Ct.Cheque.SetFieldValue("BANK_FLAG", "Y")
            Call Ct.Cheque.SetFieldValue("POST_FLAG", "N")
            Call Ct.Cheque.UpdateBankFlag
         Else
            If Ct.GetFieldValue("PAYMENT_TYPE") = 3 Then 'เช็ค
               Ct.Cheque.ShowMode = SHOW_ADD
               Call Ct.Cheque.AddEditData
            End If
            Call Ct.SetFieldValue("CHECK_ID", Ct.Cheque.GetFieldValue("CHEQUE_ID"))
         End If
         Call Ct.AddEditData
      ElseIf Ct.Flag = "E" Then
         'เมื่อเป็นการนำฝากจะไม่มีการแก้ไขเกิดขึ้น จะต้องลบแล้วสร้างใหม่
         'จากอื่น ๆ มาเป็นเช็ค -> เพิ่มเช็ค, จากเช็คเป็นอื่น ๆ -> ลบเช็ค, เป็นเช็คอยู่แล้วไม่เปลี่ยน -> แก้ไขเช็ค
         If (PaymentType <> OldPaymentType) And (PaymentType = 3) Then
            Ct.Cheque.ShowMode = SHOW_ADD
            Call Ct.Cheque.AddEditData
            
            Ct.ShowMode = SHOW_EDIT
            Call Ct.SetFieldValue("CHECK_ID", Ct.Cheque.GetFieldValue("CHEQUE_ID"))
            Call Ct.AddEditData
         ElseIf (PaymentType <> OldPaymentType) And (OldPaymentType = 3) Then
            Ct.ShowMode = SHOW_EDIT
            Call Ct.SetFieldValue("CHECK_ID", -1)
            Call Ct.AddEditData
         
            Call Ct.Cheque.DeleteData
         ElseIf (PaymentType = OldPaymentType) And (PaymentType = 3) Then
            Ct.Cheque.ShowMode = SHOW_EDIT
            Call Ct.Cheque.AddEditData
            
            Ct.ShowMode = SHOW_EDIT
            Call Ct.AddEditData
         Else
            Ct.ShowMode = SHOW_EDIT
            Call Ct.AddEditData
         End If
      ElseIf Ct.Flag = "D" Then
         Call Ct.DeleteData
         If DocType = CASH_DEPOSIT Then
            'ถ้าเป็นการลบ item สำหรับการนำเช็คในมือไปเข้า ธนาคาร ไม่ต้องลบเช็คนั้นทิ้งเพียงแต่ set BANK_FLAG ให้เป็น 'N'
            Call Ct.Cheque.SetFieldValue("BANK_FLAG", "N")
            Call Ct.Cheque.UpdateBankFlag
         Else
            If Ct.GetFieldValue("OLD_CHECK_ID") > 0 Then 'เช็ค
               Call Ct.Cheque.DeleteData
            End If
         End If
      End If
   Next Ct
   
   For Each Ct In Ua.CashTranItems
      If (DocType = CASH_WITHDRAW) Or (DocType = CASH_WHTHDRAW2) Or (DocType = CASH_TRANSFER) Then
         If Ct.GetFieldValue("CHECK_ID") > 0 Then
            Call Ct.Cheque.SetFieldValue("CHEQUE_ID", Ct.GetFieldValue("CHECK_ID"))
            Call Ct.Cheque.SetFieldValue("BANK_ID", Ua.GetFieldValue("BANK_ID"))
            Call Ct.Cheque.SetFieldValue("BANK_BRANCH", Ua.GetFieldValue("BANK_BRANCH"))
            Call Ct.Cheque.UpdateBankBranch
         End If
         
         If Ct.GetFieldValue("TX_TYPE") = "E" Then
            Call Ct.SetFieldValue("BANK_ID", Ua.GetFieldValue("BANK_ID"))
            Call Ct.SetFieldValue("BANK_BRANCH", Ua.GetFieldValue("BANK_BRANCH"))
            Call Ct.SetFieldValue("BANK_ACCOUNT", Ua.GetFieldValue("BANK_ACCOUNT"))
            Call Ct.UpdateBankBranchAccount
         End If
      ElseIf (DocType = CASH_DEPOSIT) Or (DocType = CASH_DEPOSIT2) Then
         If Ct.GetFieldValue("TX_TYPE") = "I" Then
            Call Ct.SetFieldValue("BANK_ID", Ua.GetFieldValue("BANK_ID"))
            Call Ct.SetFieldValue("BANK_BRANCH", Ua.GetFieldValue("BANK_BRANCH"))
            Call Ct.SetFieldValue("BANK_ACCOUNT", Ua.GetFieldValue("BANK_ACCOUNT"))
            Call Ct.UpdateBankBranchAccount
         End If
      End If
   Next Ct
   
   Dim CDP As CCashDocPost
   Dim Ch As CCheque
   Set Ch = New CCheque
   For Each CDP In Ua.PostItems
      If CDP.Flag = "A" Then
         Call CDP.SetFieldValue("CASH_DOC_ID", Ua.GetFieldValue("CASH_DOC_ID"))
         CDP.ShowMode = SHOW_ADD
         Call CDP.AddEditData
   
         Call Ch.SetFieldValue("CHEQUE_ID", CDP.GetFieldValue("CHEQUE_ID"))
         Call Ch.SetFieldValue("BANK_FLAG", "Y")
         Call Ch.SetFieldValue("POST_FLAG", "Y")
         Call Ch.UpdateBankFlag
         
      ElseIf CDP.Flag = "E" Then
         CDP.ShowMode = SHOW_EDIT
         Call CDP.AddEditData
      ElseIf CDP.Flag = "D" Then
         Call CDP.DeleteData
      End If
   Next CDP
   


   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditCashDoc = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   AddEditCashDoc = False
End Function
Public Sub CreateCashTransferItems(Ivd As CCashDoc)
Dim Ti As CCashTransferItem
Dim O As Object
Dim EI As CCashTran
Dim II As CCashTran
Dim I As Long
Dim j As Long
Dim Count1 As Long
Dim Count2 As Long
Dim ImportCount As Long
Dim ExportCount As Long

   Set Ivd.TransferItems = Nothing
   Set Ivd.TransferItems = New Collection
   
   Count1 = Ivd.CashTranItems.Count \ 2
   Count2 = Ivd.CashTranItems.Count
   For I = 1 To Count1
      ImportCount = 0
      ExportCount = 0
      j = 1
      While j <= Count2
         Set O = Ivd.CashTranItems(j)
         If (O.GetFieldValue("TX_TYPE") = "I") Then
            ImportCount = ImportCount + 1
            If ImportCount = I Then
               Set II = O
            End If
         ElseIf (O.GetFieldValue("TX_TYPE") = "E") Then
            ExportCount = ExportCount + 1
            If ExportCount = I Then
               Set EI = O
            End If
         End If
         j = j + 1
      Wend
         
      Set Ti = New CCashTransferItem
      Set Ti.ImportItem = II
      Set Ti.ExportItem = EI
      Ti.Flag = II.Flag
      Call Ivd.TransferItems.Add(Ti)
      Set Ti = Nothing
   Next I
End Sub
Public Function QueryCheque(Ua As CCheque, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long

   RName = "QueryCheque"
   QueryCheque = False

   IsOK = True

   Call Ua.QueryData(1, Rs, itemcount)

   QueryCheque = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryCheque = False
End Function
Public Function AddEditCheque(Ua As CCheque, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCheque"
   AddEditCheque = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditCheque = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCheque = False
End Function
Public Function DeleteCheque(Ua As CCheque, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteCheque"
   DeleteCheque = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ua.DeleteData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCheque = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteCheque = True
End Function

Public Function PatchBAPigAge(Ug As CBalanceAccum, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim EI As CBalanceAccum
Dim PigAge As Double

   RName = "PatchBAPigAge"
   PatchBAPigAge = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Ug.PIG_FLAG = "Y"
   Ug.BALANCE_ACCUM_ID = -1
   Call Ug.QueryData(1, m_Rs1, iCount)
   While Not m_Rs1.EOF
      Set EI = New CBalanceAccum
      Call EI.PopulateFromRS(1, m_Rs1)
      
      EI.PIG_AGE = GetAge(EI.PART_NO, EI.DOCUMENT_DATE)
      EI.AddEditMode = SHOW_EDIT
      Call EI.PatchPigAge
      
      Set EI = Nothing
      m_Rs1.MoveNext
   Wend
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   PatchBAPigAge = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   PatchBAPigAge = False
End Function
Public Function QueryCustomerInfo(Ua As CCustomer, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long

   RName = "QueryCustomerInfo"
   QueryCustomerInfo = False

   IsOK = True
   Call Ua.QueryData1(Rs, itemcount)
         
   QueryCustomerInfo = True
   Exit Function
   
ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryCustomerInfo = False
End Function
Public Function DeleteRevenueCost(Ua As CRevenueCost, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean

   RName = "DeleteRevenueCost"
   DeleteRevenueCost = False
   HasBegin = False
   
   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ua.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteRevenueCost = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteRevenueCost = True
End Function

Public Function QueryRevenueCost(Ua As CRevenueCost, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryRevenueCost"
   QueryRevenueCost = False

   IsOK = True

   Call Ua.QueryData(1, Rs, itemcount)
   
   If Ua.QueryFlag = 1 Then
      Dim Ji As CRevenueCostItem
      Set Ji = New CRevenueCostItem
      Call Ji.SetFieldValue("REVENUE_COST_ID", Ua.GetFieldValue("REVENUE_COST_ID"))
      Call Ji.QueryData(1, m_Rs1, iCount)
      Set Ji = Nothing
      
      Set Ua.RevenueTypeItems = Nothing
      Set Ua.RevenueTypeItems = New Collection

      While Not m_Rs1.EOF
         Set Ji = New CRevenueCostItem
         Call Ji.PopulateFromRS(1, m_Rs1)
         Call Ua.RevenueTypeItems.Add(Ji)
         Set Ji = Nothing

         m_Rs1.MoveNext
      Wend
      
      
   End If
   
   QueryRevenueCost = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryRevenueCost = False
End Function
Public Function AddEditRevenueCost(Ua As CRevenueCost, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
   
   RName = "AddEditRevenueCost"
   AddEditRevenueCost = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData
   
   Dim Ji As CRevenueCostItem
   For Each Ji In Ua.RevenueTypeItems
      Call Ji.SetFieldValue("REVENUE_COST_ID", Ua.GetFieldValue("REVENUE_COST_ID"))
      If Ji.Flag = "A" Then
         Ji.ShowMode = SHOW_ADD
         Call Ji.AddEditData
      ElseIf Ji.Flag = "E" Then
         Ji.ShowMode = SHOW_EDIT
         Call Ji.AddEditData
      ElseIf Ji.Flag = "D" Then
         Call Ji.DeleteData
      End If
   Next Ji

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditRevenueCost = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False
   AddEditRevenueCost = False
End Function
Public Function CopyBatch(Ug As CBatch, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long
Dim TempRs As ADODB.Recordset
Dim itemcount As Long

   RName = "CopyBatch"
   CopyBatch = False
   HasBegin = False
   
   Set TempRs = New ADODB.Recordset

   Ug.QueryFlag = 1
   Call QueryBatch(Ug, TempRs, iCount, IsOK, ErrorObj)
   If Not TempRs.EOF Then
      Call Ug.PopulateFromRS(1, TempRs)
   End If
   
   Call MergeBatchItem(Ug)
   
   Ug.ShowMode = SHOW_ADD
   Call Ug.SetFieldValue("BATCH_ID", -1)
   Call Ug.SetFieldValue("BATCH_NO", "***" & Ug.GetFieldValue("BATCH_NO"))
   
   Dim BItems As CBatchItem
   For Each BItems In Ug.BatchItems
      BItems.Flag = "A"
      BItems.ShowMode = SHOW_ADD
   Next BItems
         

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call AddEditBatch(Ug, IsOK, False, ErrorObj)
      
   If AutoCommit Then
      m_ADOConn.CommitTrans
   End If
   HasBegin = False
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
   
   HasBegin = False
   IsOK = True

   CopyBatch = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   CopyBatch = False
End Function
Private Sub MergeBatchItem(B As CBatch)
Dim Bi As CBatchItem

   Set B.BatchItems = Nothing
   Set B.BatchItems = New Collection
   
   For Each Bi In B.BirthItems
      Call B.BatchItems.Add(Bi)
   Next Bi
   
   For Each Bi In B.FoodItems
      Call B.BatchItems.Add(Bi)
   Next Bi

   For Each Bi In B.SaleItems
      Call B.BatchItems.Add(Bi)
   Next Bi

   For Each Bi In B.TransferItems
      Call B.BatchItems.Add(Bi)
   Next Bi

   For Each Bi In B.WeightItems
      Call B.BatchItems.Add(Bi)
   Next Bi

   For Each Bi In B.Feeds
      Call B.BatchItems.Add(Bi)
   Next Bi
   
   For Each Bi In B.Balances
      Call B.BatchItems.Add(Bi)
   Next Bi

   For Each Bi In B.Revenues
      Call B.BatchItems.Add(Bi)
   Next Bi

   For Each Bi In B.CustRatios
      Call B.BatchItems.Add(Bi)
   Next Bi

   For Each Bi In B.ChangePigTypes
      Call B.BatchItems.Add(Bi)
   Next Bi

   For Each Bi In B.BuyItems
      Call B.BatchItems.Add(Bi)
   Next Bi

   For Each Bi In B.ExpenseSharingItems
      Call B.BatchItems.Add(Bi)
   Next Bi

   For Each Bi In B.PigAdjItems
      Call B.BatchItems.Add(Bi)
   Next Bi
End Sub
Public Function QueryPriceAdjust(Ua As CPriceAdjust, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryPriceAdjust"
   QueryPriceAdjust = False

   IsOK = True
   Call Ua.QueryData(1, Rs, itemcount)

   If Ua.QueryFlag = 1 Then
      
   End If
   
   QueryPriceAdjust = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryPriceAdjust = False
End Function

Public Function DeletePriceAdjust(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim itemcount As Long

Dim B As CPriceAdjust

   RName = "DeletePriceAdjust"
   DeletePriceAdjust = False
   HasBegin = False

   Set B = New CPriceAdjust

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.PRICE_ADJUST_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeletePriceAdjust = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeletePriceAdjust = True
End Function

Public Function AddEditPriceAdjust(Ug As CPriceAdjust, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim Addr As CAddress

   RName = "AddEditPriceAdjust"
   AddEditPriceAdjust = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditPriceAdjust = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If
   
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditPriceAdjust = False
End Function

