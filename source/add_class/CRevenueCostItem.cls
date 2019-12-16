VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRevenueCostItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public TableName As String
Public SequenceName As String
Public PKField As String

Public m_FieldList As Collection
Private m_ADOConn As ADODB.Connection

Public ShowMode As SHOW_MODE_TYPE
Public Flag As String
Public QueryFlag As Long

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
Dim Tf As CTableField

   Set Tf = m_FieldList(PKField)

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

Dim REVENUE_COST_ITEM_ID As Long
Dim REVENUE_COST_ID As Long
Dim REVENUE_TYPE_ID As Long
Dim FROM_DATE As Date
Dim TO_DATE As Date
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long
Dim TempDate As String

   REVENUE_COST_ITEM_ID = Val(GetFieldValue("REVENUE_COST_ITEM_ID"))
   REVENUE_COST_ID = Val(GetFieldValue("REVENUE_COST_ID"))
   REVENUE_TYPE_ID = Val(GetFieldValue("REVENUE_TYPE_ID"))
   FROM_DATE = GetFieldValue("FROM_DATE")
   TO_DATE = GetFieldValue("TO_DATE")
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.* "
      SelectStr1 = SelectStr1 & ",RT.REVENUE_NAME REVENUE_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",PT.PRODUCT_TYPE_NAME PIG_TYPE_NAME,PT.PRODUCT_TYPE_NO PIG_TYPE_NO "
      SelectStr1 = SelectStr1 & ",PS.PRODUCT_STATUS_NAME PIG_STATUS_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN REVENUE_TYPE RT ON (RT.REVENUE_TYPE_ID = NM.REVENUE_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (PS.PRODUCT_STATUS_ID = NM.PIG_STATUS) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_TYPE PT ON (PT.PRODUCT_TYPE_ID = NM.PIG_TYPE) "
      
      WhereStr = ""
      
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT NM.* "
      SelectStr1 = SelectStr1 & ",RT.REVENUE_NAME REVENUE_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",PT.PRODUCT_TYPE_NAME PIG_TYPE_NAME,PT.PRODUCT_TYPE_NO PIG_TYPE_NO "
      SelectStr1 = SelectStr1 & ",PS.PRODUCT_STATUS_NAME PIG_STATUS_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      FromStr = FromStr & "LEFT OUTER JOIN REVENUE_COST RC ON (RC.REVENUE_COST_ID = NM.REVENUE_COST_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN REVENUE_TYPE RT ON (RT.REVENUE_TYPE_ID = NM.REVENUE_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (PS.PRODUCT_STATUS_ID = NM.PIG_STATUS) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_TYPE PT ON (PT.PRODUCT_TYPE_ID = NM.PIG_TYPE) "
      WhereStr = ""
   End If
   
   If REVENUE_COST_ITEM_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.REVENUE_COST_ITEM_ID = " & REVENUE_COST_ITEM_ID & ") " & vbCrLf
   End If
   
   If REVENUE_COST_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.REVENUE_COST_ID = " & REVENUE_COST_ID & ") " & vbCrLf
   End If
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (RC.REVENUE_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If

   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (RC.REVENUE_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY NM.REVENUE_COST_ITEM_ID "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " DESC "
   End If
   
   If Ind = 2 Then
      OrderByStr = ""
   End If
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("REVENUE_COST_ITEM_ID", NVLI(Rs("REVENUE_COST_ITEM_ID"), -1))
      Call SetFieldValue("REVENUE_COST_ID", NVLI(Rs("REVENUE_COST_ID"), -1))
      Call SetFieldValue("REVENUE_TYPE_ID", NVLI(Rs("REVENUE_TYPE_ID"), -1))
      Call SetFieldValue("REVENUE_COST_ITEM_AMOUNT", NVLD(Rs("REVENUE_COST_ITEM_AMOUNT"), 0))
      Call SetFieldValue("REVENUE_COST_ITEM_SELL", NVLD(Rs("REVENUE_COST_ITEM_SELL"), 0))
      Call SetFieldValue("PIG_STATUS", NVLI(Rs("PIG_STATUS"), -1))
      Call SetFieldValue("PIG_TYPE", NVLI(Rs("PIG_TYPE"), -1))
      
      Call SetFieldValue("REVENUE_TYPE_NAME", NVLS(Rs("REVENUE_TYPE_NAME"), ""))
      Call SetFieldValue("PIG_STATUS_NAME", NVLS(Rs("PIG_STATUS_NAME"), ""))
      Call SetFieldValue("PIG_TYPE_NO", NVLS(Rs("PIG_TYPE_NO"), ""))
      Call SetFieldValue("PIG_TYPE_NAME", NVLS(Rs("PIG_TYPE_NAME"), ""))
   ElseIf Ind = 2 Then
      Call SetFieldValue("REVENUE_TYPE_ID", NVLI(Rs("REVENUE_TYPE_ID"), -1))
      Call SetFieldValue("REVENUE_COST_ITEM_AMOUNT", NVLD(Rs("REVENUE_COST_ITEM_AMOUNT"), 0))
      Call SetFieldValue("REVENUE_COST_ITEM_SELL", NVLD(Rs("REVENUE_COST_ITEM_SELL"), 0))
      Call SetFieldValue("PIG_STATUS", NVLI(Rs("PIG_STATUS"), -1))
      
      Call SetFieldValue("REVENUE_TYPE_NAME", NVLS(Rs("REVENUE_TYPE_NAME"), ""))
      Call SetFieldValue("PIG_STATUS_NAME", NVLS(Rs("PIG_STATUS_NAME"), ""))
      Call SetFieldValue("PIG_TYPE_NO", NVLS(Rs("PIG_TYPE_NO"), ""))
      Call SetFieldValue("PIG_TYPE_NAME", NVLS(Rs("PIG_TYPE_NAME"), ""))
   End If
End Sub
Public Sub RegisterFields()
   TableName = "REVENUE_COST_ITEM"
   SequenceName = "REVENUE_COST_ITEM_SEQ"
   PKField = "REVENUE_COST_ITEM_ID"
   
   Call AddField("REVENUE_COST_ITEM_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("REVENUE_COST_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("REVENUE_TYPE_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("REVENUE_COST_ITEM_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("PIG_STATUS", INT_TYPE, -1, DATA_CAT)
   Call AddField("REVENUE_COST_ITEM_SELL", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("PIG_TYPE", INT_TYPE, -1, DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_DATE", DATE_TYPE, -1, TEMP_CAT)
   Call AddField("TO_DATE", DATE_TYPE, -1, TEMP_CAT)
   
   Call AddField("REVENUE_TYPE_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PIG_STATUS_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PIG_TYPE_NO", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PIG_TYPE_NAME", STRING_TYPE, "", TEMP_CAT)
End Sub
