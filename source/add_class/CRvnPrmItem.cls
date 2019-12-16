VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRvnPrmItem"
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

Dim RVNPRM_ITEM_ID As Long
Dim PARAM_ID As Long
Dim FROM_DATE As Date
Dim TO_DATE As Date
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long

   RVNPRM_ITEM_ID = Val(GetFieldValue("RVNPRM_ITEM_ID"))
   PARAM_ID = Val(GetFieldValue("PARAM_ID"))
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   If Ind = 1 Then
      SelectStr1 = "SELECT EM.*, RT.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EM "
      FromStr = FromStr & "LEFT OUTER JOIN REVENUE_TYPE RT ON (EM.REVENUE_ID = RT.REVENUE_TYPE_ID) "
      
      WhereStr = ""
   End If
   
   If RVNPRM_ITEM_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (EM.RVNPRM_ITEM_ID = " & RVNPRM_ITEM_ID & ") " & vbCrLf
   End If
   
   If PARAM_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (EM.PARAM_ID = " & PARAM_ID & ") " & vbCrLf
   End If
         
   If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY EM.FROM_SALE "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY EM.FROM_SALE "
      Else
         OrderByStr = "ORDER BY EM.FROM_SALE "
      End If
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("RVNPRM_ITEM_ID", NVLI(Rs("RVNPRM_ITEM_ID"), -1))
      Call SetFieldValue("PARAM_ID", NVLS(Rs("PARAM_ID"), ""))
      Call SetFieldValue("SALE_AMOUNT", NVLD(Rs("SALE_AMOUNT"), 0))
      Call SetFieldValue("TOTAL_PRICE", NVLD(Rs("TOTAL_PRICE"), 0))
      Call SetFieldValue("UNIT_PRICE", NVLD(Rs("UNIT_PRICE"), 0))
      Call SetFieldValue("FROM_SALE", InternalDateToDate(NVLS(Rs("FROM_SALE"), "")))
      Call SetFieldValue("TO_SALE", InternalDateToDate(NVLS(Rs("TO_SALE"), "")))
      Call SetFieldValue("REVENUE_ID", NVLI(Rs("REVENUE_ID"), -1))
      
      Call SetFieldValue("REVENUE_NO", NVLS(Rs("REVENUE_NO"), ""))
      Call SetFieldValue("REVENUE_NAME", NVLS(Rs("REVENUE_NAME"), ""))
      Call SetFieldValue("KEY_ID", NVLS(Rs("RVNPRM_ITEM_ID"), ""))
      Call SetFieldValue("KEY_LOOKUP", NVLS(Rs("PARAM_ID"), ""))
   End If
End Sub

Public Sub RegisterFields()
   TableName = "RVNPRM_ITEM"
   SequenceName = "RVNPRM_ITEM_SEQ"
   PKField = "RVNPRM_ITEM_ID"
   
   Call AddField("RVNPRM_ITEM_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("PARAM_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("FROM_SALE", DATE_TYPE, -1, DATA_CAT)
   Call AddField("TO_SALE", DATE_TYPE, -1, DATA_CAT)
   Call AddField("SALE_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("TOTAL_PRICE", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("UNIT_PRICE", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("REVENUE_ID", INT_TYPE, -1, DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("KEY_ID", INT_TYPE, "", TEMP_CAT)
   Call AddField("KEY_LOOKUP", STRING_TYPE, "", TEMP_CAT)
   Call AddField("REVENUE_NAME", STRING_TYPE, "", TEMP_CAT)
   Call AddField("REVENUE_NO", STRING_TYPE, "", TEMP_CAT)
   
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_DATE", DATE_TYPE, -1, TEMP_CAT)
   Call AddField("TO_DATE", DATE_TYPE, -1, TEMP_CAT)
End Sub
Public Sub CopyItemCollection(TempJi As CRvnPrmItem)
   Call TempJi.SetFieldValue("RVNPRM_ITEM_ID", GetFieldValue("RVNPRM_ITEM_ID"))
   Call TempJi.SetFieldValue("PARAM_ID", GetFieldValue("PARAM_ID"))
   Call TempJi.SetFieldValue("SALE_AMOUNT", GetFieldValue("SALE_AMOUNT"))
   Call TempJi.SetFieldValue("TOTAL_PRICE", GetFieldValue("TOTAL_PRICE"))
   Call TempJi.SetFieldValue("UNIT_PRICE", GetFieldValue("UNIT_PRICE"))
   
   Call TempJi.SetFieldValue("FROM_SALE", GetFieldValue("FROM_SALE"))
   Call TempJi.SetFieldValue("TO_SALE", GetFieldValue("TO_SALE"))
   Call TempJi.SetFieldValue("REVENUE_ID", GetFieldValue("REVENUE_ID"))
   Call TempJi.SetFieldValue("REVENUE_NO", GetFieldValue("REVENUE_NO"))
   Call TempJi.SetFieldValue("REVENUE_NAME", GetFieldValue("REVENUE_NAME"))
End Sub
