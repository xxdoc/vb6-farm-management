VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TableName As String
Public SequenceName As String
Public PKField As String

Public QueryFlag As Long
Public ShowMode As SHOW_MODE_TYPE
Public m_FieldList As Collection
Private m_ADOConn As ADODB.Connection

Public Flag As String
Public BatchItems As Collection
Public BirthItems As Collection
Public FoodItems As Collection
Public TransferItems As Collection
Public SaleItems As Collection
Public WeightItems As Collection
Public Feeds As Collection
Public Balances As Collection
Public Revenues As Collection
Public CustRatios As Collection
Public ChangePigTypes As Collection
Public BuyItems As Collection
Public ExpenseSharingItems As Collection
Public PigAdjItems As Collection
Public ManagementExpenses As Collection
Public Glages As Collection
Public GLbacks As Collection
Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   
   Set BatchItems = New Collection
   Set BirthItems = New Collection
   Set FoodItems = New Collection
   Set TransferItems = New Collection
   Set SaleItems = New Collection
   Set WeightItems = New Collection
   Set Feeds = New Collection
   Set Balances = New Collection
   Set Revenues = New Collection
   Set CustRatios = New Collection
   Set ChangePigTypes = New Collection
   Set BuyItems = New Collection
   Set ExpenseSharingItems = New Collection
   Set PigAdjItems = New Collection
   Set ManagementExpenses = New Collection
   Set Glages = New Collection
   Set GLbacks = New Collection
   Call RegisterFields
End Sub

Private Sub Class_Terminate()
   Set m_FieldList = Nothing
   
   Set BatchItems = Nothing
   Set BirthItems = Nothing
   Set FoodItems = Nothing
   Set TransferItems = Nothing
   Set SaleItems = Nothing
   Set WeightItems = Nothing
   Set Feeds = Nothing
   Set Balances = Nothing
   Set Revenues = Nothing
   Set CustRatios = Nothing
   Set ChangePigTypes = Nothing
   Set BuyItems = Nothing
   Set ExpenseSharingItems = Nothing
   Set PigAdjItems = Nothing
   Set ManagementExpenses = Nothing
   Set Glages = Nothing
   Set GLbacks = Nothing
End Sub

Private Sub AddField(FName As String, FType As FIELD_TYPE, FValue As Variant, FCat As FIELD_CAT)
Dim F As CTableField

   Set F = New CTableField
   Call F.InitField(FName, FType, FValue, FCat)
   Call m_FieldList.Add(F, FName)
   Set F = Nothing
End Sub

Public Function GetFieldValue(FieldName As String)
Dim Tf As CTableField

   Set Tf = m_FieldList(FieldName)
   GetFieldValue = Tf.GetValue
End Function

Public Sub SetFieldValue(FieldName As String, Value As Variant)
Dim Tf As CTableField

   Set Tf = m_FieldList(FieldName)
   Tf.SetValue (Value)
End Sub

Public Sub AddEditData()
Dim SQL As String

   Call PopulateInternalField(ShowMode, Me)
   
   If ShowMode = SHOW_ADD Then
      SQL = GenerateInsertSQL(Me)
   ElseIf ShowMode = SHOW_EDIT Then
      SQL = GenerateUpdateSQL(Me)
   End If

   Call m_ADOConn.Execute(SQL)
End Sub

Public Sub DeleteData()
Dim SQL As String
Dim Tf As CTableField

   Set Tf = m_FieldList(PKField)
      
   SQL = "DELETE FROM BATCH_ITEM WHERE " & PKField & " = " & Tf.GetValue
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM " & TableName & " WHERE " & PKField & " = " & Tf.GetValue
   Call m_ADOConn.Execute(SQL)
End Sub
Public Sub DeleteBatchDocument(Optional DeleteAll As Boolean = False)
Dim SQL As String
Dim SetStr1 As String
Dim SetStr2 As String
Dim SetStr3 As String
Dim SetStr4 As String
Dim BATCH_ID As Long
   
   BATCH_ID = GetFieldValue("BATCH_ID")
   SetStr1 = "(SELECT INVENTORY_DOC_ID FROM INVENTORY_DOC WHERE BATCH_ID = " & BATCH_ID & ")"
   SetStr2 = "(SELECT BILLING_DOC_ID FROM BILLING_DOC WHERE BATCH_ID = " & BATCH_ID & ")"
   SetStr3 = "(SELECT CAPITAL_MOVEMENT_ID FROM CAPITAL_MOVEMENT WHERE BATCH_ID = " & BATCH_ID & ")"
   
   If DeleteAll Then
      SQL = "DELETE FROM MOVEMENT_ITEM "
   Else
      SQL = "DELETE FROM MOVEMENT_ITEM WHERE CAPITAL_MOVEMENT_ID IN " & SetStr3 & ""
   End If
   Call m_ADOConn.Execute(SQL)
   
   If DeleteAll Then
      SQL = "DELETE FROM CAPITAL_MOVEMENT "
   Else
      SQL = "DELETE FROM CAPITAL_MOVEMENT WHERE BATCH_ID = " & BATCH_ID
   End If
   Call m_ADOConn.Execute(SQL)
   
   If DeleteAll Then
      SQL = "DELETE FROM DO_ITEM "
   Else
      SQL = "DELETE FROM DO_ITEM WHERE DO_ID IN " & SetStr2 & ""
   End If
   Call m_ADOConn.Execute(SQL)
   
   If DeleteAll Then
      SQL = "DELETE FROM EXPENSE_RATIO "
   Else
      SQL = "DELETE FROM EXPENSE_RATIO WHERE BATCH_ID = " & BATCH_ID & ""
   End If
   Call m_ADOConn.Execute(SQL)
   
   If DeleteAll Then
      SQL = "DELETE FROM RO_ITEM "
   Else
      SQL = "DELETE FROM RO_ITEM WHERE BILLING_DOC_ID IN " & SetStr2 & ""
   End If
   Call m_ADOConn.Execute(SQL)
   
   If DeleteAll Then
      SQL = "DELETE FROM BILLING_DOC "
   Else
      SQL = "DELETE FROM BILLING_DOC WHERE BATCH_ID = " & BATCH_ID & ""
   End If
   Call m_ADOConn.Execute(SQL)
   
   If DeleteAll Then
      SQL = "DELETE FROM IMPORT_ITEM "
   Else
      SQL = "DELETE FROM IMPORT_ITEM WHERE INVENTORY_DOC_ID IN " & SetStr1 & ""
   End If
   Call m_ADOConn.Execute(SQL)
   
   If DeleteAll Then
      SQL = "DELETE FROM EXPORT_ITEM "
   Else
      SQL = "DELETE FROM EXPORT_ITEM WHERE INVENTORY_DOC_ID IN " & SetStr1 & ""
   End If
   Call m_ADOConn.Execute(SQL)
   
   If DeleteAll Then
      SQL = "DELETE FROM INVENTORY_DOC "
   Else
      SQL = "DELETE FROM INVENTORY_DOC WHERE BATCH_ID = " & BATCH_ID & ""
   End If
   Call m_ADOConn.Execute(SQL)
   
   If DeleteAll Then
      SQL = "DELETE FROM INTAKE "
   Else
      SQL = "DELETE FROM INTAKE WHERE BATCH_ID = " & BATCH_ID & ""
   End If
   Call m_ADOConn.Execute(SQL)
   
   If DeleteAll Then
      SQL = "DELETE FROM WEIGHT_RECORD "
   Else
      SQL = "DELETE FROM WEIGHT_RECORD WHERE BATCH_ID = " & BATCH_ID & ""
   End If
   Call m_ADOConn.Execute(SQL)
   
   If DeleteAll Then
      SQL = "DELETE FROM COST_ACCUM "
   Else
      SQL = "DELETE FROM COST_ACCUM WHERE BATCH_ID = " & BATCH_ID & ""
   End If
   Call m_ADOConn.Execute(SQL)
   
   If DeleteAll Then
      SQL = "DELETE FROM MONTHLY_ACCUM "
   Else
      SQL = "DELETE FROM MONTHLY_ACCUM WHERE BATCH_ID = " & BATCH_ID & ""
   End If
   Call m_ADOConn.Execute(SQL)
End Sub

Public Sub QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim GroupByStr As String
Dim OrderByStr As String
Dim TempStr As String
Dim NewStr As String
Dim SubLen As Long
Dim TempDate As String

Dim BATCH_ID As Long
Dim FROM_DATE As Date
Dim TO_DATE As Date
Dim BATCH_NO As String
Dim COMMIT_FLAG As String
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long

   BATCH_ID = Val(GetFieldValue("BATCH_ID"))
   BATCH_NO = GetFieldValue("BATCH_NO")
   COMMIT_FLAG = GetFieldValue("COMMIT_FLAG")
   FROM_DATE = GetFieldValue("FROM_DATE")
   TO_DATE = GetFieldValue("TO_DATE")
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   If Ind = 1 Then
      SelectStr1 = "SELECT EM.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EM "
      
      WhereStr = ""
   End If
   
   If BATCH_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BATCH_ID = " & BATCH_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BATCH_ID = " & BATCH_ID & ") "
      End If
'   Else
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE ((BATCH_ID <= 0) OR (BATCH_ID IS NULL)) "
'      Else
'         WhereStr = WhereStr & "AND ((BATCH_ID <= 0) OR (BATCH_ID IS NULL)) "
'      End If
   End If
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (BATCH_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (BATCH_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
      
   If Len(Trim(BATCH_NO)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(BATCH_NO, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(BATCH_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (BATCH_NO = '" & ChangeQuote(Trim(BATCH_NO)) & "')"
      End If
   End If
   
   If Len(Trim(COMMIT_FLAG)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(COMMIT_FLAG, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(COMMIT_FLAG,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (COMMIT_FLAG = '" & ChangeQuote(Trim(COMMIT_FLAG)) & "')"
      End If
   End If
   
   If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY EM.BATCH_NO "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY EM.BATCH_DATE "
      Else
         OrderByStr = "ORDER BY EM.BATCH_NO "
      End If
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("BATCH_ID", NVLI(Rs("BATCH_ID"), -1))
      Call SetFieldValue("BATCH_NO", NVLS(Rs("BATCH_NO"), ""))
      Call SetFieldValue("BATCH_DATE", InternalDateToDate(NVLS(Rs("BATCH_DATE"), "")))
      Call SetFieldValue("EXECUTE_FROM", InternalDateToDate(NVLS(Rs("EXECUTE_FROM"), "")))
      Call SetFieldValue("EXECUTE_TO", InternalDateToDate(NVLS(Rs("EXECUTE_TO"), "")))
      Call SetFieldValue("BATCH_DESC", NVLS(Rs("BATCH_DESC"), ""))
      Call SetFieldValue("COMMIT_FLAG", NVLS(Rs("COMMIT_FLAG"), ""))
      Call SetFieldValue("BALANCE_CASH", NVLD(Rs("BALANCE_CASH"), 0))
      Call SetFieldValue("CASH_FIRST_MONTH", NVLD(Rs("CASH_FIRST_MONTH"), 0))
      Call SetFieldValue("MEDICINE_FIRST_MONTH", NVLD(Rs("MEDICINE_FIRST_MONTH"), 0))
      Call SetFieldValue("OTHER_FIRST_MONTH", NVLD(Rs("OTHER_FIRST_MONTH"), 0))
      
      Call SetFieldValue("KEY_ID", NVLS(Rs("BATCH_ID"), ""))
      Call SetFieldValue("KEY_LOOKUP", NVLS(Rs("BATCH_NO"), ""))
   End If
End Sub

Public Sub RegisterFields()
   TableName = "BATCH"
   SequenceName = "BATCH_SEQ"
   PKField = "BATCH_ID"
   
   Call AddField("BATCH_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("BATCH_NO", STRING_TYPE, "", DATA_CAT)
   Call AddField("BATCH_DATE", DATE_TYPE, "", DATA_CAT)
   Call AddField("EXECUTE_FROM", DATE_TYPE, "", DATA_CAT)
   Call AddField("EXECUTE_TO", DATE_TYPE, "", DATA_CAT)
   Call AddField("BATCH_DESC", STRING_TYPE, "", DATA_CAT)
   Call AddField("COMMIT_FLAG", STRING_TYPE, "", DATA_CAT)
   Call AddField("BALANCE_CASH", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("CASH_FIRST_MONTH", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("MEDICINE_FIRST_MONTH", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("OTHER_FIRST_MONTH", MONEY_TYPE, 0, DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("KEY_ID", INT_TYPE, "", TEMP_CAT)
   Call AddField("KEY_LOOKUP", STRING_TYPE, "", TEMP_CAT)
   
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_DATE", DATE_TYPE, -1, TEMP_CAT)
   Call AddField("TO_DATE", DATE_TYPE, -1, TEMP_CAT)
   
   Call AddField("PARAM_AREA", INT_TYPE, "", TEMP_CAT)
End Sub


