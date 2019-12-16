VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParamItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TableName As String
Public SequenceName As String
Public PKField As String
Public Flag As String

Public QueryFlag As Long
Public ShowMode As SHOW_MODE_TYPE
Public m_FieldList As Collection
Private m_ADOConn As ADODB.Connection

Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   
   Call RegisterFields
End Sub

Private Sub Class_Terminate()
   Set m_FieldList = Nothing
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
   
   If GetFieldValue("DEPRECIATION_FLAG") = "" Then
      Call SetFieldValue("DEPRECIATION_FLAG", "N")
   End If
   
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
      
   SQL = "DELETE FROM " & TableName & " WHERE " & PKField & " = " & Tf.GetValue
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

Dim PARAM_ITEM_ID As Long
Dim PARAM_ID As Long
Dim BATCH_ID As Long
Dim FROM_DATE As Date
Dim TO_DATE As Date
Dim PARAM_AREA As Long
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long
Dim DEPRECIATION_FLAG As String
   PARAM_ITEM_ID = Val(GetFieldValue("PARAM_ITEM_ID"))
   PARAM_ID = Val(GetFieldValue("PARAM_ID"))
   PARAM_AREA = Val(GetFieldValue("PARAM_AREA"))
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   BATCH_ID = Val(GetFieldValue("BATCH_ID"))
   DEPRECIATION_FLAG = GetFieldValue("DEPRECIATION_FLAG")
   
   If Ind = 1 Then
      SelectStr1 = "SELECT EM.*, CM.CUSTOMER_CODE, NM1.LONG_NAME CUSTOMER_NAME, PT.PRODUCT_TYPE_NO PIG_TYPE_NO, PT.PRODUCT_TYPE_NAME PIG_TYPE_NAME, PI.PIG_TYPE PIG_TYPE_NO2, PI.PART_NO PIG_NO, PI.PART_DESC PIG_NAME, PM.PIG_TYPE PARAM_PIG_TYPE, PM.PIG_STATUS PARAM_PIG_STATUS, AC.ACCOUNT_ID, ET.EXPENSE_TYPE_NAME "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EM "
      FromStr = FromStr & "LEFT OUTER JOIN PARAMETERS PM ON (EM.PARAM_ID = PM.PARAM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (EM.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (CN.NAME_ID = NM1.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_TYPE PT ON (EM.PIG_TYPE = PT.PRODUCT_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EM.PIG_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (CM.CUSTOMER_ID = AC.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EXPENSE_TYPE ET ON (EM.EXPENSE_TYPE = ET.EXPENSE_TYPE_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT EM.YYYYMM,SUM(EM.EXP_AMOUNT) EXP_AMOUNT "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EM "
      FromStr = FromStr & "LEFT OUTER JOIN PARAMETERS PM ON (EM.PARAM_ID = PM.PARAM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BATCH_ITEM BI ON (PM.PARAM_ID = BI.PARAM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BATCH B ON (B.BATCH_ID = BI.BATCH_ID) "
      
   End If
   
   If PARAM_ITEM_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EM.PARAM_ITEM_ID = " & PARAM_ITEM_ID & ") " & vbCrLf
   End If
   
   If PARAM_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EM.PARAM_ID = " & PARAM_ID & ") " & vbCrLf
   End If
         
   If PARAM_AREA > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (PM.PARAM_AREA = " & PARAM_AREA & ") " & vbCrLf
   End If
   
   If BATCH_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (B.BATCH_ID = " & BATCH_ID & ") " & vbCrLf
   End If
      
   If Len(DEPRECIATION_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DEPRECIATION_FLAG = '" & DEPRECIATION_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (DEPRECIATION_FLAG = '" & DEPRECIATION_FLAG & "') "
      End If
   End If
      
   If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY CM.CUSTOMER_CODE "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY CM.CUSTOMER_CODE "
      ElseIf ORDER_BY = 3 Then
         OrderByStr = "ORDER BY EM.BUY_DATE "
      Else
         OrderByStr = "ORDER BY CM.CUSTOMER_CODE "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY EM.YYYYMM  "
   End If
   
    GroupByStr = ""
   If Ind = 2 Then
      GroupByStr = "GROUP BY EM.YYYYMM "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("PARAM_ITEM_ID", NVLI(Rs("PARAM_ITEM_ID"), -1))
      Call SetFieldValue("PARAM_ID", NVLS(Rs("PARAM_ID"), ""))
      Call SetFieldValue("CUSTOMER_ID", NVLD(Rs("CUSTOMER_ID"), 0))
      Call SetFieldValue("SALE_RATIO", NVLD(Rs("SALE_RATIO"), 0))
      Call SetFieldValue("PIG_TYPE", NVLD(Rs("PIG_TYPE"), 0))
      Call SetFieldValue("TRANSFER_RATE", NVLD(Rs("TRANSFER_RATE"), 0))
      Call SetFieldValue("PIG_ID", NVLD(Rs("PIG_ID"), 0))
      Call SetFieldValue("BUY_TOTAL_PRICE", NVLD(Rs("BUY_TOTAL_PRICE"), 0))
      Call SetFieldValue("BUY_AMOUNT", NVLD(Rs("BUY_AMOUNT"), 0))
      Call SetFieldValue("BUY_AVG_PRICE", NVLD(Rs("BUY_AVG_PRICE"), 0))
      Call SetFieldValue("BUY_DATE", InternalDateToDate(NVLS(Rs("BUY_DATE"), "")))
      Call SetFieldValue("BUY_AVG_WEIGHT", NVLD(Rs("BUY_AVG_WEIGHT"), 0))
      Call SetFieldValue("EXPENSE_DATE", InternalDateToDate(NVLS(Rs("EXPENSE_DATE"), "")))
      Call SetFieldValue("EXPENSE_TYPE", NVLD(Rs("EXPENSE_TYPE"), 0))
      Call SetFieldValue("EXPENSE_NAME", NVLS(Rs("EXPENSE_NAME"), ""))
      Call SetFieldValue("EXP_AMOUNT", NVLD(Rs("EXP_AMOUNT"), 0))
      Call SetFieldValue("EXP_AVG_PRICE", NVLD(Rs("EXP_AVG_PRICE"), 0))
      Call SetFieldValue("EXP_TOTAL_PRICE", NVLD(Rs("EXP_TOTAL_PRICE"), 0))
      Call SetFieldValue("SHARE_SELL_TYPE", NVLI(Rs("SHARE_SELL_TYPE"), -1))
      Call SetFieldValue("CTRL_FROM_DATE", InternalDateToDate(NVLS(Rs("CTRL_FROM_DATE"), "")))
      Call SetFieldValue("CTRL_TO_DATE", InternalDateToDate(NVLS(Rs("CTRL_TO_DATE"), "")))
      Call SetFieldValue("PIG_STATUS_ID", NVLI(Rs("PIG_STATUS_ID"), -1))
      Call SetFieldValue("CTRL_AMOUNT", NVLD(Rs("CTRL_AMOUNT"), 0))
      Call SetFieldValue("YYYYMM", NVLS(Rs("YYYYMM"), ""))
      Call SetFieldValue("DEPRECIATION_FLAG", NVLS(Rs("DEPRECIATION_FLAG"), ""))
      
      Call SetFieldValue("PARAM_PIG_TYPE", NVLI(Rs("PARAM_PIG_TYPE"), 0))
      Call SetFieldValue("PARAM_PIG_STATUS", NVLI(Rs("PARAM_PIG_STATUS"), 0))
      Call SetFieldValue("CUSTOMER_CODE", NVLS(Rs("CUSTOMER_CODE"), ""))
      Call SetFieldValue("CUSTOMER_NAME", NVLS(Rs("CUSTOMER_NAME"), ""))
      Call SetFieldValue("PIG_TYPE_NO", NVLS(Rs("PIG_TYPE_NO"), ""))
      Call SetFieldValue("PIG_TYPE_NAME", NVLS(Rs("PIG_TYPE_NAME"), ""))
      Call SetFieldValue("PIG_TYPE_NO2", NVLS(Rs("PIG_TYPE_NO2"), ""))
      Call SetFieldValue("PIG_NO", NVLS(Rs("PIG_NO"), ""))
      Call SetFieldValue("PIG_DESC", NVLS(Rs("PIG_NAME"), ""))
      Call SetFieldValue("ACCOUNT_ID", NVLI(Rs("ACCOUNT_ID"), 0))
      Call SetFieldValue("EXPENSE_TYPE_NAME", NVLS(Rs("EXPENSE_TYPE_NAME"), ""))
      
      Call SetFieldValue("KEY_ID", NVLS(Rs("PARAM_ITEM_ID"), ""))
      Call SetFieldValue("KEY_LOOKUP", NVLS(Rs("PARAM_ID"), ""))
   ElseIf Ind = 2 Then
      Call SetFieldValue("EXP_AMOUNT", NVLD(Rs("EXP_AMOUNT"), 0))
      Call SetFieldValue("YYYYMM", NVLS(Rs("YYYYMM"), ""))
   End If
End Sub

Public Sub RegisterFields()
   TableName = "PARAM_ITEM"
   SequenceName = "PARAM_ITEM_SEQ"
   PKField = "PARAM_ITEM_ID"
   
   Call AddField("PARAM_ITEM_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("PARAM_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("CUSTOMER_ID", INT_TYPE, 0, DATA_CAT)
   Call AddField("SALE_RATIO", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("PIG_TYPE", INT_TYPE, 0, DATA_CAT)
   Call AddField("TRANSFER_RATE", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("PIG_ID", INT_TYPE, 0, DATA_CAT)
   Call AddField("BUY_DATE", DATE_TYPE, -1, DATA_CAT)
   Call AddField("BUY_AVG_PRICE", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("BUY_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("BUY_TOTAL_PRICE", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("BUY_AVG_WEIGHT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("EXPENSE_DATE", DATE_TYPE, -1, DATA_CAT)
   Call AddField("EXPENSE_TYPE", INT_TYPE, 0, DATA_CAT)
   Call AddField("EXPENSE_NAME", STRING_TYPE, "", DATA_CAT)
   Call AddField("EXP_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("EXP_TOTAL_PRICE", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("EXP_AVG_PRICE", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("SHARE_SELL_TYPE", INT_TYPE, -1, DATA_CAT)
   Call AddField("CTRL_FROM_DATE", DATE_TYPE, -1, DATA_CAT)
   Call AddField("CTRL_TO_DATE", DATE_TYPE, -1, DATA_CAT)
   Call AddField("PIG_STATUS_ID", INT_TYPE, 0, DATA_CAT)
   Call AddField("CTRL_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("YYYYMM", STRING_TYPE, "", DATA_CAT)
   Call AddField("DEPRECIATION_FLAG", STRING_TYPE, "", DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("KEY_ID", INT_TYPE, "", TEMP_CAT)
   Call AddField("KEY_LOOKUP", STRING_TYPE, "", TEMP_CAT)
   Call AddField("CUSTOMER_CODE", STRING_TYPE, "", TEMP_CAT)
   Call AddField("CUSTOMER_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PARAM_AREA", INT_TYPE, -1, TEMP_CAT)
   Call AddField("PIG_TYPE_NO", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PIG_TYPE_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PIG_NO", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PIG_DESC", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PIG_TYPE_NO2", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PARAM_PIG_TYPE", INT_TYPE, -1, TEMP_CAT)
   Call AddField("PARAM_PIG_STATUS", INT_TYPE, -1, TEMP_CAT)
   Call AddField("ACCOUNT_ID", INT_TYPE, 0, TEMP_CAT)
   Call AddField("EXPENSE_TYPE_NAME", STRING_TYPE, "", TEMP_CAT)
   
   Call AddField("BATCH_ID", INT_TYPE, 0, TEMP_CAT)
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_DATE", DATE_TYPE, -1, TEMP_CAT)
   Call AddField("TO_DATE", DATE_TYPE, -1, TEMP_CAT)
End Sub
Public Sub CopyItemCollection(TempPi As CParamItem)
   Call TempPi.SetFieldValue("PARAM_ITEM_ID", GetFieldValue("PARAM_ITEM_ID"))
   Call TempPi.SetFieldValue("PARAM_ID", GetFieldValue("PARAM_ID"))
   Call TempPi.SetFieldValue("CUSTOMER_ID", GetFieldValue("CUSTOMER_ID"))
   Call TempPi.SetFieldValue("SALE_RATIO", GetFieldValue("SALE_RATIO"))
   Call TempPi.SetFieldValue("PIG_TYPE", GetFieldValue("PIG_TYPE"))
   Call TempPi.SetFieldValue("TRANSFER_RATE", GetFieldValue("TRANSFER_RATE"))
   Call TempPi.SetFieldValue("PIG_ID", GetFieldValue("PIG_ID"))
   Call TempPi.SetFieldValue("BUY_DATE", GetFieldValue("BUY_DATE"))
   Call TempPi.SetFieldValue("BUY_AVG_PRICE", GetFieldValue("BUY_AVG_PRICE"))
   Call TempPi.SetFieldValue("BUY_AMOUNT", GetFieldValue("BUY_AMOUNT"))
   
   Call TempPi.SetFieldValue("BUY_TOTAL_PRICE", GetFieldValue("BUY_TOTAL_PRICE"))
   Call TempPi.SetFieldValue("BUY_AVG_WEIGHT", GetFieldValue("BUY_AVG_WEIGHT"))
   Call TempPi.SetFieldValue("EXPENSE_DATE", GetFieldValue("EXPENSE_DATE"))
   Call TempPi.SetFieldValue("EXPENSE_TYPE", GetFieldValue("EXPENSE_TYPE"))
   Call TempPi.SetFieldValue("EXPENSE_NAME", GetFieldValue("EXPENSE_NAME"))
   
   Call TempPi.SetFieldValue("EXP_AMOUNT", GetFieldValue("EXP_AMOUNT"))
   Call TempPi.SetFieldValue("EXP_TOTAL_PRICE", GetFieldValue("EXP_TOTAL_PRICE"))
   Call TempPi.SetFieldValue("EXP_AVG_PRICE", GetFieldValue("EXP_AVG_PRICE"))
   Call TempPi.SetFieldValue("SHARE_SELL_TYPE", GetFieldValue("SHARE_SELL_TYPE"))
   Call TempPi.SetFieldValue("CTRL_FROM_DATE", GetFieldValue("CTRL_FROM_DATE"))
   
   Call TempPi.SetFieldValue("CTRL_TO_DATE", GetFieldValue("CTRL_TO_DATE"))
   Call TempPi.SetFieldValue("PIG_STATUS_ID", GetFieldValue("PIG_STATUS_ID"))
   Call TempPi.SetFieldValue("CTRL_AMOUNT", GetFieldValue("CTRL_AMOUNT"))
   Call TempPi.SetFieldValue("YYYYMM", GetFieldValue("YYYYMM"))
   Call TempPi.SetFieldValue("DEPRECIATION_FLAG", GetFieldValue("DEPRECIATION_FLAG"))
   
   Call TempPi.SetFieldValue("CUSTOMER_CODE", GetFieldValue("CUSTOMER_CODE"))
   Call TempPi.SetFieldValue("CUSTOMER_NAME", GetFieldValue("CUSTOMER_NAME"))
   Call TempPi.SetFieldValue("PARAM_AREA", GetFieldValue("PARAM_AREA"))
   Call TempPi.SetFieldValue("PIG_TYPE_NO", GetFieldValue("PIG_TYPE_NO"))
   Call TempPi.SetFieldValue("PIG_TYPE_NAME", GetFieldValue("PIG_TYPE_NAME"))
   
   Call TempPi.SetFieldValue("PIG_NO", GetFieldValue("PIG_NO"))
   Call TempPi.SetFieldValue("PIG_DESC", GetFieldValue("PIG_DESC"))
   Call TempPi.SetFieldValue("PIG_TYPE_NO2", GetFieldValue("PIG_TYPE_NO2"))
   Call TempPi.SetFieldValue("PARAM_PIG_TYPE", GetFieldValue("PARAM_PIG_TYPE"))
   Call TempPi.SetFieldValue("PARAM_PIG_STATUS", GetFieldValue("PARAM_PIG_STATUS"))
   
   Call TempPi.SetFieldValue("ACCOUNT_ID", GetFieldValue("ACCOUNT_ID"))
   Call TempPi.SetFieldValue("EXPENSE_TYPE_NAME", GetFieldValue("EXPENSE_TYPE_NAME"))
   Call TempPi.SetFieldValue("BATCH_ID", GetFieldValue("BATCH_ID"))
   
End Sub

