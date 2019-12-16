VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBatchItem"
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

Public Parameters As CParameter

Private Sub Class_Initialize()
   Set m_FieldList = New Collection
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   
   Set Parameters = New CParameter
   
   Call RegisterFields
End Sub
Private Sub Class_Terminate()
   Set m_FieldList = Nothing
   Set Parameters = Nothing
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

Public Sub QueryData(Ind As Long, Rs As ADODB.Recordset, itemcount As Long)
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

Dim BATCH_ITEM_ID As Long
Dim BATCH_ID As Long
Dim FROM_DATE As Date
Dim TO_DATE As Date
Dim PARAM_AREA As Long
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long

   BATCH_ITEM_ID = Val(GetFieldValue("BATCH_ITEM_ID"))
   BATCH_ID = Val(GetFieldValue("BATCH_ID"))
   PARAM_AREA = Val(GetFieldValue("PARAM_AREA"))
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   If Ind = 1 Then
      SelectStr1 = "SELECT EM.*, PM.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EM "
      FromStr = FromStr & "LEFT OUTER JOIN PARAMETERS PM ON (EM.PARAM_ID = PM.PARAM_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT PM.FROM_AGE, PM.TO_AGE, PM.ADG, PM.PIG_TYPE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EM "
      FromStr = FromStr & "LEFT OUTER JOIN PARAMETERS PM ON (EM.PARAM_ID = PM.PARAM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BATCH BT ON (EM.BATCH_ID = BT.BATCH_ID) "
      
      WhereStr = ""
   End If
   
   If BATCH_ITEM_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (EM.BATCH_ITEM_ID = " & BATCH_ITEM_ID & ") " & vbCrLf
   End If
   
   If BATCH_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EM.BATCH_ID = " & BATCH_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EM.BATCH_ID = " & BATCH_ID & ") "
      End If
   Else
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE ((EM.BATCH_ID <= 0) OR (BATCH_ID IS NULL)) "
      Else
         WhereStr = WhereStr & "AND ((EM.BATCH_ID <= 0) OR (BATCH_ID IS NULL)) "
      End If
   End If

   If PARAM_AREA > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (PM.PARAM_AREA = " & PARAM_AREA & ") " & vbCrLf
   End If
   
   If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY PM.PARAM_NO "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY PM.PARAM_NO "
      Else
         OrderByStr = "ORDER BY PM.PARAM_NO "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY PM.PIG_TYPE ASC, PM.FROM_AGE ASC, PM.TO_AGE "
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Sub

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("BATCH_ITEM_ID", NVLI(Rs("BATCH_ITEM_ID"), -1))
      Call SetFieldValue("BATCH_ID", NVLI(Rs("BATCH_ID"), -1))
      Call SetFieldValue("PARAM_ID", NVLI(Rs("PARAM_ID"), -1))

      Call SetFieldValue("PARAM_NO", NVLS(Rs("PARAM_NO"), ""))
      Call SetFieldValue("PARAM_DESC", NVLS(Rs("PARAM_DESC"), ""))
      Call SetFieldValue("PARAM_DATE", InternalDateToDate(NVLS(Rs("PARAM_DATE"), "")))
      Call SetFieldValue("FROM_AGE", NVLD(Rs("FROM_AGE"), 0))
      Call SetFieldValue("TO_AGE", NVLD(Rs("TO_AGE"), 0))
      Call SetFieldValue("PARAM_AREA", NVLI(Rs("PARAM_AREA"), -1))
      Call SetFieldValue("PIG_TYPE", NVLI(Rs("PIG_TYPE"), -1))

      Call SetFieldValue("KEY_ID", NVLS(Rs("BATCH_ITEM_ID"), ""))
      Call SetFieldValue("KEY_LOOKUP", NVLS(Rs("BATCH_ID"), ""))
   ElseIf Ind = 2 Then
      Call SetFieldValue("PIG_TYPE", NVLI(Rs("PIG_TYPE"), -1))
      Call SetFieldValue("FROM_AGE", NVLD(Rs("FROM_AGE"), 0))
      Call SetFieldValue("TO_AGE", NVLD(Rs("TO_AGE"), 0))
      Call SetFieldValue("ADG", NVLD(Rs("ADG"), 0))
   End If
End Sub

Public Sub RegisterFields()
   TableName = "BATCH_ITEM"
   SequenceName = "BATCH_ITEM_SEQ"
   PKField = "BATCH_ITEM_ID"
   
   Call AddField("BATCH_ITEM_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("BATCH_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("PARAM_ID", INT_TYPE, 0, DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("KEY_ID", INT_TYPE, "", TEMP_CAT)
   Call AddField("KEY_LOOKUP", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PARAM_NO", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PARAM_DESC", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PARAM_DATE", DATE_TYPE, -1, TEMP_CAT)
   Call AddField("PARAM_AREA", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_AGE", MONEY_TYPE, 0, TEMP_CAT)
   Call AddField("TO_AGE", MONEY_TYPE, 0, TEMP_CAT)
   Call AddField("PIG_TYPE", MONEY_TYPE, "", TEMP_CAT)
   Call AddField("ADG", MONEY_TYPE, 0, TEMP_CAT)
   
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_DATE", DATE_TYPE, -1, TEMP_CAT)
   Call AddField("TO_DATE", DATE_TYPE, -1, TEMP_CAT)
End Sub

