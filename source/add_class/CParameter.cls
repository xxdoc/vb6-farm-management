VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParameter"
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

Public BrtPrmItems As Collection
Public UsedPrmItems As Collection
Public SalePrmItems As Collection
Public WeightPrmItems As Collection
Public TrnPrmItems As Collection
Public CostPrmItems As Collection
Public AdjPrmItems As Collection
Public RvnPrmItems As Collection
Public CustRatioItems As Collection
Public PigStatusChangeItems As Collection
Public PigBuyItems As Collection
Public ExpenseSharings As Collection
Public PigAdjustItems As Collection
Public MenagementExpenses As Collection
Public Glages As Collection
Public GLbacks   As Collection
Public Flag As String

Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   
   Set BrtPrmItems = New Collection
   Set UsedPrmItems = New Collection
   Set SalePrmItems = New Collection
   Set WeightPrmItems = New Collection
   Set TrnPrmItems = New Collection
   Set CostPrmItems = New Collection
   Set AdjPrmItems = New Collection
   Set RvnPrmItems = New Collection
   Set CustRatioItems = New Collection
   Set PigStatusChangeItems = New Collection
   Set PigBuyItems = New Collection
   Set ExpenseSharings = New Collection
   Set PigAdjustItems = New Collection
   Set MenagementExpenses = New Collection
   Set Glages = New Collection
   Set GLbacks = New Collection
   Call RegisterFields
End Sub

Private Sub Class_Terminate()
   Set m_FieldList = Nothing
   Set BrtPrmItems = Nothing
   Set UsedPrmItems = Nothing
   Set SalePrmItems = Nothing
   Set WeightPrmItems = Nothing
   Set TrnPrmItems = Nothing
   Set CostPrmItems = Nothing
   Set AdjPrmItems = Nothing
   Set RvnPrmItems = Nothing
   Set CustRatioItems = Nothing
   Set PigStatusChangeItems = Nothing
   Set PigBuyItems = Nothing
   Set ExpenseSharings = Nothing
   Set PigAdjustItems = Nothing
   Set MenagementExpenses = Nothing
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
         
   SQL = "DELETE FROM PARAM_ITEM WHERE " & PKField & " = " & Tf.GetValue
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM RVNPRM_ITEM WHERE " & PKField & " = " & Tf.GetValue
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM TRNPRM_ITEM WHERE " & PKField & " = " & Tf.GetValue
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM WEIGHTPRM_ITEM WHERE " & PKField & " = " & Tf.GetValue
   Call m_ADOConn.Execute(SQL)

   SQL = "DELETE FROM SALEPRM_ITEM WHERE " & PKField & " = " & Tf.GetValue
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM USEDPRM_ITEM WHERE " & PKField & " = " & Tf.GetValue
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM COSTPRM_ITEM WHERE " & PKField & " = " & Tf.GetValue
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM BRTPRM_ITEM WHERE " & PKField & " = " & Tf.GetValue
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM ADJPRM_ITEM WHERE " & PKField & " = " & Tf.GetValue
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM ADJPRM_ITEM WHERE " & PKField & " = " & Tf.GetValue
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM GL_AGE_AMOUNT WHERE " & PKField & " = " & Tf.GetValue
   Call m_ADOConn.Execute(SQL)
   
   SQL = "DELETE FROM GL_BACK_AMOUNT WHERE " & PKField & " = " & Tf.GetValue
   Call m_ADOConn.Execute(SQL)
   
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

Dim PARAM_ID As Long
Dim FROM_DATE As Date
Dim TO_DATE As Date
Dim PARAM_NO As String
Dim PARAM_AREA As Long
Dim COMMIT_FLAG As String
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long

   PARAM_ID = Val(GetFieldValue("PARAM_ID"))
   PARAM_AREA = Val(GetFieldValue("PARAM_AREA"))
   FROM_DATE = GetFieldValue("FROM_DATE")
   TO_DATE = GetFieldValue("TO_DATE")
   PARAM_NO = GetFieldValue("PARAM_NO")
   COMMIT_FLAG = GetFieldValue("COMMIT_FLAG")
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   If Ind = 1 Then
      SelectStr1 = "SELECT EM.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EM "
      
      WhereStr = ""
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
      WhereStr = WhereStr & TempStr & " (EM.PARAM_AREA = " & PARAM_AREA & ") " & vbCrLf
   End If
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (PARAM_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (PARAM_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
      
   If Len(Trim(PARAM_NO)) > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      If WildCard(PARAM_NO, SubLen, NewStr) Then
         WhereStr = WhereStr & TempStr & " (SUBSTR(PARAM_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
      Else
         WhereStr = WhereStr & TempStr & " (PARAM_NO = '" & ChangeQuote(Trim(PARAM_NO)) & "')"
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
         OrderByStr = "ORDER BY EM.PARAM_NO "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY EM.PARAM_DATE "
      Else
         OrderByStr = "ORDER BY EM.PARAM_NO "
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
      Call SetFieldValue("PARAM_ID", NVLI(Rs("PARAM_ID"), -1))
      Call SetFieldValue("PARAM_NO", NVLS(Rs("PARAM_NO"), ""))
      Call SetFieldValue("PARAM_DATE", InternalDateToDate(NVLS(Rs("PARAM_DATE"), "")))
      Call SetFieldValue("PARAM_AREA", NVLI(Rs("PARAM_AREA"), -1))
      Call SetFieldValue("FROM_SALE_DATE", InternalDateToDate(NVLS(Rs("FROM_SALE_DATE"), "")))
      Call SetFieldValue("TO_SALE_DATE", InternalDateToDate(NVLS(Rs("TO_SALE_DATE"), "")))
      Call SetFieldValue("FROM_AGE", NVLD(Rs("FROM_AGE"), 0))
      Call SetFieldValue("TO_AGE", NVLD(Rs("TO_AGE"), 0))
      Call SetFieldValue("PARAM_DESC", NVLS(Rs("PARAM_DESC"), ""))
      Call SetFieldValue("COMMIT_FLAG", NVLS(Rs("COMMIT_FLAG"), ""))
      Call SetFieldValue("PIG_TYPE", NVLI(Rs("PIG_TYPE"), -1))
      Call SetFieldValue("PIG_STATUS", NVLI(Rs("PIG_STATUS"), -1))
      Call SetFieldValue("ADG", NVLD(Rs("ADG"), 0))
      
      Call SetFieldValue("KEY_ID", NVLS(Rs("PARAM_ID"), ""))
      Call SetFieldValue("KEY_LOOKUP", NVLS(Rs("PARAM_NO"), ""))
   End If
End Sub

Public Sub RegisterFields()
   TableName = "PARAMETERS"
   SequenceName = "PARAMETER_SEQ"
   PKField = "PARAM_ID"
   
   Call AddField("PARAM_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("PARAM_NO", STRING_TYPE, "", DATA_CAT)
   Call AddField("PARAM_DATE", DATE_TYPE, "", DATA_CAT)
   Call AddField("PARAM_AREA", INT_TYPE, -1, DATA_CAT)
   Call AddField("FROM_SALE_DATE", DATE_TYPE, "", DATA_CAT)
   Call AddField("TO_SALE_DATE", DATE_TYPE, "", DATA_CAT)
   Call AddField("FROM_AGE", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("TO_AGE", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("PARAM_DESC", STRING_TYPE, "", DATA_CAT)
   Call AddField("COMMIT_FLAG", STRING_TYPE, "", DATA_CAT)
   Call AddField("PIG_STATUS", INT_TYPE, -1, DATA_CAT)
   Call AddField("PIG_TYPE", INT_TYPE, -1, DATA_CAT)
   Call AddField("ADG", MONEY_TYPE, 0, DATA_CAT)
   
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
End Sub



