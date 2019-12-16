VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_ModuleName As String
Private m_ADOConn As ADODB.Connection
Private m_Rs As ADODB.Recordset

Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   m_ModuleName = "clsMaster"
   Set m_Rs = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub

'==
Public Function AddEditPartType(At As CPartType, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditPartType"
   AddEditPartType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditPartType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditPartType = False
End Function

Public Function AddEditPartGroup(At As CPartGroup, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditPartGroup"
   AddEditPartGroup = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditPartGroup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditPartGroup = False
End Function

Public Function AddEditExposeType(At As CExposeType, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditExposeType"
   AddEditExposeType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditExposeType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditExposeType = False
End Function

Public Function AddEditHouseGroup(At As CHouseGroup, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditHouseGroup"
   AddEditHouseGroup = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Dim CstName As CHGroupItem
   For Each CstName In At.HGroupItems
      If CstName.Flag = "A" Then
         CstName.HOUSE_GROUP_ID = At.HOUSE_GROUP_ID
         CstName.AddEditMode = SHOW_ADD
         CstName.AddEditData
      ElseIf CstName.Flag = "E" Then
         CstName.AddEditMode = SHOW_EDIT
         CstName.AddEditData
      ElseIf CstName.Flag = "D" Then
         CstName.DeleteData
      End If
   Next CstName
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditHouseGroup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditHouseGroup = False
End Function

Public Function AddEditStatusGroup(At As CStatusGroup, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditStatusGroup"
   AddEditStatusGroup = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Dim CstName As CSGroupItem
   For Each CstName In At.HGroupItems
      If CstName.Flag = "A" Then
         CstName.STATUS_GROUP_ID = At.STATUS_GROUP_ID
         CstName.AddEditMode = SHOW_ADD
         CstName.AddEditData
      ElseIf CstName.Flag = "E" Then
         CstName.AddEditMode = SHOW_EDIT
         CstName.AddEditData
      ElseIf CstName.Flag = "D" Then
         CstName.DeleteData
      End If
   Next CstName
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditStatusGroup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditStatusGroup = False
End Function

Public Function DeletePartType(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CPartType

   RName = "DeletePartType"
   DeletePartType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CPartType
   cData.PART_TYPE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeletePartType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeletePartType = False
End Function
Public Function DeletePartGroup(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CPartGroup

   RName = "DeletePartGroup"
   DeletePartGroup = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CPartGroup
   cData.PART_GROUP_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeletePartGroup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeletePartGroup = False
End Function
Public Function DeleteExposeType(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CExposeType

   RName = "DeleteExposeType"
   DeleteExposeType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CExposeType
   cData.EXPOSE_TYPE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteExposeType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteExposeType = False
End Function

Public Function DeleteHouseGroup(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CHouseGroup

   RName = "DeleteHouseGroup"
   DeleteHouseGroup = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CHouseGroup
   cData.HOUSE_GROUP_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteHouseGroup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteHouseGroup = False
End Function

Public Function DeleteStatusGroup(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CStatusGroup

   RName = "DeleteStatusGroup"
   DeleteStatusGroup = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CStatusGroup
   cData.STATUS_GROUP_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteStatusGroup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteStatusGroup = False
End Function

'==
Public Function AddEditLocation(At As CLocation, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditLocation"
   AddEditLocation = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditLocation = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditLocation = False
End Function

Public Function DeleteLocation(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CLocation

   RName = "DeleteLocation"
   DeleteLocation = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CLocation
   cData.LOCATION_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteLocation = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteLocation = False
End Function

'==
Public Function AddEditProductType(At As CProductType, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditProductType"
   AddEditProductType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditProductType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditProductType = False
End Function

Public Function DeleteProductType(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CProductType

   RName = "DeleteProductType"
   DeleteProductType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CProductType
   cData.PRODUCT_TYPE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteProductType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteProductType = False
End Function

'==
Public Function AddEditProductStatus(At As CProductStatus, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditProductStatus"
   AddEditProductStatus = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditProductStatus = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditProductStatus = False
End Function

Public Function DeleteProductStatus(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CProductStatus

   RName = "DeleteProductStatus"
   DeleteProductStatus = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CProductStatus
   cData.PRODUCT_STATUS_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteProductStatus = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteProductStatus = False
End Function
'==
Public Function AddEditHouse(At As CHouse, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditHouse"
   AddEditHouse = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditHouse = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditHouse = False
End Function

Public Function DeleteHouse(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CHouse

   RName = "DeleteHouse"
   DeleteHouse = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CHouse
   cData.HOUSE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteHouse = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteHouse = False
End Function
'==
Public Function AddEditCountry(At As CCountry, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCountry"
   AddEditCountry = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditCountry = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCountry = False
End Function

Public Function DeleteCountry(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CCountry

   RName = "DeleteCountry"
   DeleteCountry = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CCountry
   cData.COUNTRY_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteCountry = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteCountry = True
End Function

Public Function DeleteAgeRange(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CAgeRange

   RName = "DeleteAgeRange"
   DeleteAgeRange = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CAgeRange
   cData.AGE_RANGE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteAgeRange = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteAgeRange = True
End Function

Public Function DeleteCnDnReason(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CCnDnReason

   RName = "DeleteCnDnReason"
   DeleteCnDnReason = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CCnDnReason
   cData.REASON_ID = UID
   Call cData.DeleteData
   Set cData = Nothing

   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteCnDnReason = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteCnDnReason = True
End Function

'==
Public Function AddEditCustomerGrade(At As CCustomerGrade, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCustomerGrade"
   AddEditCustomerGrade = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditCustomerGrade = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCustomerGrade = False
End Function

Public Function DeleteCustomerGrade(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CCustomerGrade

   RName = "DeleteCustomerGrade"
   DeleteCustomerGrade = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CCustomerGrade
   cData.CSTGRADE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteCustomerGrade = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteCustomerGrade = True
End Function

'==
Public Function AddEditCustomerType(At As CCustomerType, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCustomerType"
   AddEditCustomerType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditCustomerType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCustomerType = False
End Function

Public Function DeleteCustomerType(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CCustomerType

   RName = "DeleteCustomerType"
   DeleteCustomerType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CCustomerType
   cData.CSTTYPE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteCustomerType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteCustomerType = True
End Function
'==
Public Function AddEditSupplierGrade(At As CSupplierGrade, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditSupplierGrade"
   AddEditSupplierGrade = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditSupplierGrade = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditSupplierGrade = False
End Function

Public Function DeleteSupplierGrade(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CSupplierGrade

   RName = "DeleteSupplierGrade"
   DeleteSupplierGrade = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CSupplierGrade
   cData.SUPPLIER_GRADE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteSupplierGrade = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteSupplierGrade = True
End Function
'==
Public Function AddEditSupplierType(At As CSupplierType, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditSupplierType"
   AddEditSupplierType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditSupplierType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditSupplierType = False
End Function

Public Function DeleteSupplierType(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CSupplierType

   RName = "DeleteSupplierType"
   DeleteSupplierType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CSupplierType
   cData.SUPPLIER_TYPE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteSupplierType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteSupplierType = True
End Function

'==
Public Function AddEditSupplierStatus(At As CSupplierStatus, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditSupplierStatus"
   AddEditSupplierStatus = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditSupplierStatus = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditSupplierStatus = False
End Function

Public Function DeleteSupplierStatus(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CSupplierStatus

   RName = "DeleteSupplierStatus"
   DeleteSupplierStatus = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CSupplierStatus
   cData.SUPPLIER_STATUS_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteSupplierStatus = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteSupplierStatus = True
End Function
'==
Public Function AddEditPosition(Ug As CEmpPosition, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditPosition"
   AddEditPosition = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call Ug.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditPosition = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditPosition = False
End Function

Public Function DeletePosition(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim B As CEmpPosition

   RName = "DeletePosition"
   DeletePosition = False
   HasBegin = False

   Set B = New CEmpPosition

   m_ADOConn.BeginTrans
   HasBegin = True

   B.POSITION_ID = UID
   Call B.DeleteData

   Set B = Nothing

   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeletePosition = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeletePosition = True
End Function

'==
Public Function AddEditUnit(Ug As CUnit, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditUnit"
   AddEditUnit = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call Ug.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditUnit = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditUnit = False
End Function

Public Function DeleteUnit(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim B As CUnit

   RName = "DeleteUnit"
   DeleteUnit = False
   HasBegin = False

   Set B = New CUnit

   m_ADOConn.BeginTrans
   HasBegin = True

   B.UNIT_ID = UID
   Call B.DeleteData

   Set B = Nothing

   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteUnit = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteUnit = True
End Function

Public Function QueryHouseGroup(Ua As CHouseGroup, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryHouseGroup"
   QueryHouseGroup = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)

   If Ua.QueryFlag = 1 Then
      Dim Yw As CHGroupItem
      
      Set Yw = New CHGroupItem
      Yw.HGROUP_ITEM_ID = -1
      Yw.HOUSE_GROUP_ID = Ua.HOUSE_GROUP_ID
      Yw.OrderBy = 1
      Call Yw.QueryData(1, m_Rs, iCount)
      Set Yw = Nothing
      
      Set Ua.HGroupItems = Nothing
      Set Ua.HGroupItems = New Collection
      While Not m_Rs.EOF
         Set Yw = New CHGroupItem
         Call Yw.PopulateFromRS(1, m_Rs)
                  
         Yw.Flag = "I"
         Call Ua.HGroupItems.Add(Yw)
         
         Set Yw = Nothing
         m_Rs.MoveNext
      Wend
   End If
   
   QueryHouseGroup = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryHouseGroup = False
End Function

Public Function QueryAgeRange(Ua As CAgeRange, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryAgeRange"
   QueryAgeRange = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)

   If Ua.QueryFlag = 1 Then
   End If
   
   QueryAgeRange = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryAgeRange = False
End Function

Public Function QueryStatusGroup(Ua As CStatusGroup, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryStatusGroup"
   QueryStatusGroup = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)

   If Ua.QueryFlag = 1 Then
      Dim Yw As CSGroupItem
      
      Set Yw = New CSGroupItem
      Yw.SGROUP_ITEM_ID = -1
      Yw.STATUS_GROUP_ID = Ua.STATUS_GROUP_ID
      Yw.OrderBy = 1
      Call Yw.QueryData(1, m_Rs, iCount)
      Set Yw = Nothing
      
      Set Ua.HGroupItems = Nothing
      Set Ua.HGroupItems = New Collection
      While Not m_Rs.EOF
         Set Yw = New CSGroupItem
         Call Yw.PopulateFromRS(1, m_Rs)
                  
         Yw.Flag = "I"
         Call Ua.HGroupItems.Add(Yw)
         
         Set Yw = Nothing
         m_Rs.MoveNext
      Wend
   End If
   
   QueryStatusGroup = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryStatusGroup = False
End Function

Public Function AddEditAgeRange(At As CAgeRange, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditAgeRange"
   AddEditAgeRange = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditAgeRange = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditAgeRange = False
End Function

Public Function AddEditExpenseType(At As CExpenseType, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditExpenseType"
   AddEditExpenseType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Dim CstName As CExptypeRatio
   For Each CstName In At.ExpenseRatios
      If CstName.Flag = "A" Then
         CstName.EXPENSE_TYPE_ID = At.EXPENSE_TYPE_ID
         CstName.AddEditMode = SHOW_ADD
         CstName.AddEditData
      ElseIf CstName.Flag = "E" Then
         CstName.AddEditMode = SHOW_EDIT
         CstName.AddEditData
      ElseIf CstName.Flag = "D" Then
         CstName.DeleteData
      End If
   Next CstName
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditExpenseType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditExpenseType = False
End Function

Public Function DeleteExpenseType(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CExpenseType

   RName = "DeleteExpenseType"
   DeleteExpenseType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CExpenseType
   cData.EXPENSE_TYPE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteExpenseType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteExpenseType = False
End Function

Public Function QueryExpenseType(Ua As CExpenseType, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
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

   RName = "QueryExpenseType"
   QueryExpenseType = False

   IsOK = True
   Call Ua.QueryData(Rs, ItemCount)

   If Ua.QueryFlag = 1 Then
      Dim Yw As CExptypeRatio

      Set Yw = New CExptypeRatio
      Yw.EXPTYPE_RATIO_ID = -1
      Yw.EXPENSE_TYPE_ID = Ua.EXPENSE_TYPE_ID
      Yw.OrderBy = 1
      Call Yw.QueryData(1, m_Rs, iCount)
      Set Yw = Nothing

      Set Ua.ExpenseRatios = Nothing
      Set Ua.ExpenseRatios = New Collection
      While Not m_Rs.EOF
         Set Yw = New CExptypeRatio
         Call Yw.PopulateFromRS(1, m_Rs)

         Yw.Flag = "I"
         Call Ua.ExpenseRatios.Add(Yw)

         Set Yw = Nothing
         m_Rs.MoveNext
      Wend
   End If
   
   QueryExpenseType = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryExpenseType = False
End Function

Public Function AddEditDocumentType(At As CDocumentType, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditDocumentType"
   AddEditDocumentType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditDocumentType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditDocumentType = False
End Function

Public Function DeleteDocumentType(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CDocumentType

   RName = "DeleteDocumentType"
   DeleteDocumentType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CDocumentType
   cData.DOCUMENT_TYPE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteDocumentType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteDocumentType = False
End Function

Public Function AddEditBank(At As CBank, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditBank"
   AddEditBank = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditBank = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditBank = False
End Function

Public Function DeleteBank(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CBank

   RName = "DeleteBank"
   DeleteBank = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CBank
   cData.BANK_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteBank = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteBank = True
End Function

Public Function AddEditBankBranch(At As CBankBranch, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditBankBranch"
   AddEditBankBranch = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditBankBranch = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditBankBranch = False
End Function

Public Function DeleteBankBranch(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CBankBranch

   RName = "DeleteBankBranch"
   DeleteBankBranch = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CBankBranch
   cData.BBRANCH_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteBankBranch = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteBankBranch = True
End Function

Public Function AddEditRegion(Ug As CRegion, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditRegion"
   AddEditRegion = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call Ug.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditRegion = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditRegion = False
End Function

Public Function DeleteRegion(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim B As CRegion

   RName = "DeleteRegion"
   DeleteRegion = False
   HasBegin = False

   Set B = New CRegion

   m_ADOConn.BeginTrans
   HasBegin = True

   B.REGION_ID = UID
   Call B.DeleteData

   Set B = Nothing

   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteRegion = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteRegion = True
End Function


Public Function AddEditRevenueType(Ug As CRevenueType, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditRevenueType"
   AddEditRevenueType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call Ug.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditRevenueType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditRevenueType = False
End Function

Public Function DeleteRevenueType(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim ItemCount As Long

Dim B As CRevenueType

   RName = "DeleteRevenueType"
   DeleteRevenueType = False
   HasBegin = False

   Set B = New CRevenueType

   m_ADOConn.BeginTrans
   HasBegin = True

   B.REVENUE_TYPE_ID = UID
   Call B.DeleteData

   Set B = Nothing

   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteRevenueType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteRevenueType = True
End Function

Public Function AddEditCnDnReason(At As CCnDnReason, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCnDnReason"
   AddEditCnDnReason = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData

   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditCnDnReason = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCnDnReason = False
End Function

Public Function AddEditBankAccount(At As CBankAccount, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditBankAccount"
   AddEditBankAccount = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditBankAccount = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditBankAccount = False
End Function

Public Function DeleteBankAccount(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CBankAccount

   RName = "DeleteBankAccount"
   DeleteBankAccount = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CBankAccount
   cData.BANK_ACCOUNT_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteBankAccount = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteBankAccount = True
End Function

Public Function AddEditStatusType(At As CStatusType, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditStatusType"
   AddEditStatusType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditStatusType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditStatusType = False
End Function

Public Function DeleteStatusType(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CStatusType

   RName = "DeleteStatusType"
   DeleteStatusType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CStatusType
   cData.STATUS_TYPE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing

   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteStatusType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteStatusType = False
End Function

Public Function AddEditPackageType(At As CPackageType, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditPackageType"
   AddEditPackageType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditPackageType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditPackageType = False
End Function

Public Function DeletePackageType(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CPackageType

   RName = "DeletePackageType"
   DeletePackageType = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CPackageType
   cData.PACKAGE_TYPE_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeletePackageType = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeletePackageType = True
End Function

Public Function AddEditMasterRef(Ug As CMasterRef, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditMasterRef"
   AddEditMasterRef = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

'   Dim CstName As CPrtItemSet
'   For Each CstName In Ug.PrtItemSets
'      If CstName.Flag = "A" Then
'         Call CstName.SetFieldValue("MASTER_REF_ID", Ug.KEY_ID)
'         CstName.ShowMode = SHOW_ADD
'         CstName.AddEditData
'      ElseIf CstName.Flag = "E" Then
'         CstName.ShowMode = SHOW_EDIT
'         CstName.AddEditData
'      ElseIf CstName.Flag = "D" Then
'         CstName.DeleteData
'      End If
'   Next CstName

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditMasterRef = True
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
   AddEditMasterRef = False
End Function
Public Function DeleteMasterRef(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CMasterRef

   RName = "DeleteMasterRef"
   DeleteMasterRef = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CMasterRef
   cData.KEY_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteMasterRef = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteMasterRef = True
End Function

