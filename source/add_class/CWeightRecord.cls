VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWeightRecord"
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

Public PIG_AGE As Double

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

Public Sub ClearData()
Dim SQL As String
Dim BATCH_ID As Long
   
   BATCH_ID = GetFieldValue("BATCH_ID")
   
   SQL = "DELETE FROM " & TableName & " WHERE BATCH_ID  = " & BATCH_ID
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

Dim WEIGHT_RECORD_ID As Long
Dim BATCH_ID As Long
Dim FROM_DATE As Date
Dim TO_DATE As Date
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long
Dim PIG_NO As String
Dim PIG_TYPE As String
   WEIGHT_RECORD_ID = Val(GetFieldValue("WEIGHT_RECORD_ID"))
   BATCH_ID = Val(GetFieldValue("BATCH_ID"))
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   FROM_DATE = GetFieldValue("FROM_DATE")
   TO_DATE = GetFieldValue("TO_DATE")
   PIG_NO = GetFieldValue("PIG_NO")
   PIG_TYPE = GetFieldValue("PIG_TYPE")
   
   If Ind = 1 Then
      SelectStr1 = "SELECT EM.*, PI.PART_NO PIG_NO, PI.PART_DESC PIG_DESC,PI.PIG_TYPE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EM "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EM.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT EM.*, PI.PART_NO PIG_NO, PI.PART_DESC PIG_DESC "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EM "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EM.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = ""
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT MIN(EM.AVG_WEIGHT) AVG_WEIGHT,MAX(EM.AVG_WEIGHT) WEIGHT_AMOUNT,EM.PIG_AGE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EM "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EM.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      'WhereStr = " WHERE PI.PART_NO = '255001' "
   
   End If
   
   If WEIGHT_RECORD_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EM.WEIGHT_RECORD_ID = " & WEIGHT_RECORD_ID & ") " & vbCrLf
   End If
      
   If BATCH_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EM.BATCH_ID = " & BATCH_ID & ") " & vbCrLf
   End If
      
   If PIG_AGE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EM.PIG_AGE = " & PIG_AGE & ") " & vbCrLf
   End If
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(FROM_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EM.RECORD_DATE >= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(TO_DATE)
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (EM.RECORD_DATE <= '" & ChangeQuote(Trim(TempDate)) & "')"
   End If
   
   If Len(Trim(PIG_NO)) > 0 Then
      If WildCard(PIG_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PI.PART_NO = '" & ChangeQuote(Trim(PIG_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (PI.PART_NO = '" & ChangeQuote(Trim(PIG_NO)) & "')"
         End If
      End If
   End If
   
   If Len(Trim(PIG_TYPE)) > 0 Then
      If WildCard(PIG_TYPE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(PI.PIG_TYPE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(PI.PIG_TYPE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PI.PIG_TYPE = '" & ChangeQuote(Trim(PIG_TYPE)) & "')"
         Else
            WhereStr = WhereStr & "AND (PI.PIG_TYPE = '" & ChangeQuote(Trim(PIG_TYPE)) & "')"
         End If
      End If
   End If
   
   GroupByStr = ""
   If Ind = 3 Then
      GroupByStr = "GROUP BY EM.PIG_AGE "
   End If
   
   If Ind = 1 Then
      If ORDER_BY = 1 Then
         OrderByStr = "ORDER BY EM.PIG_AGE "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY EM.PIG_AGE_INT "
      ElseIf ORDER_BY = 10 Then
         OrderByStr = "ORDER BY PI.PIG_TYPE ASC,PI.PART_NO ASC,EM.RECORD_DATE ASC,EM.AVG_WEIGHT "
      Else
         OrderByStr = "ORDER BY EM.RECORD_DATE "
      End If
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY EM.PIG_AGE "
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
      Call SetFieldValue("WEIGHT_RECORD_ID", NVLI(Rs("WEIGHT_RECORD_ID"), -1))
      Call SetFieldValue("RECORD_DATE", InternalDateToDate(NVLS(Rs("RECORD_DATE"), "")))
      Call SetFieldValue("PART_ITEM_ID", NVLD(Rs("PART_ITEM_ID"), 0))
      Call SetFieldValue("ITEM_AMOUNT", NVLD(Rs("ITEM_AMOUNT"), 0))
      Call SetFieldValue("PIG_AGE_INT", NVLD(Rs("PIG_AGE_INT"), 0))
      Call SetFieldValue("PIG_AGE", NVLD(Rs("PIG_AGE"), 0))
      Call SetFieldValue("AVG_WEIGHT", NVLD(Rs("AVG_WEIGHT"), 0))
      Call SetFieldValue("WEIGHT_AMOUNT", NVLD(Rs("WEIGHT_AMOUNT"), 0))
      Call SetFieldValue("BATCH_ID", NVLD(Rs("BATCH_ID"), -1))
      Call SetFieldValue("ADG", NVLD(Rs("ADG"), -1))
      
      Call SetFieldValue("PIG_NO", NVLS(Rs("PIG_NO"), ""))
      Call SetFieldValue("PIG_DESC", NVLS(Rs("PIG_DESC"), ""))
      Call SetFieldValue("PIG_TYPE", NVLS(Rs("PIG_TYPE"), ""))
      
      Call SetFieldValue("KEY_ID", NVLS(Rs("WEIGHT_RECORD_ID"), ""))
      Call SetFieldValue("KEY_LOOKUP", NVLS(Rs("RECORD_DATE"), ""))
   ElseIf Ind = 3 Then
      Call SetFieldValue("PIG_AGE", NVLD(Rs("PIG_AGE"), 0))
      Call SetFieldValue("AVG_WEIGHT", NVLD(Rs("AVG_WEIGHT"), 0))
      Call SetFieldValue("WEIGHT_AMOUNT", NVLD(Rs("WEIGHT_AMOUNT"), 0))
   End If
End Sub

Public Sub RegisterFields()
   TableName = "WEIGHT_RECORD"
   SequenceName = "WEIGHT_RECORD_SEQ"
   PKField = "WEIGHT_RECORD_ID"
   
   Call AddField("WEIGHT_RECORD_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("RECORD_DATE", DATE_TYPE, -1, DATA_CAT)
   Call AddField("PART_ITEM_ID", INT_TYPE, 0, DATA_CAT)
   Call AddField("ITEM_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("PIG_AGE_INT", INT_TYPE, 0, DATA_CAT)
   Call AddField("PIG_AGE", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("WEIGHT_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("AVG_WEIGHT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("BATCH_ID", INT_TYPE, 0, DATA_CAT)
   Call AddField("ADG", MONEY_TYPE, 0, DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("KEY_ID", INT_TYPE, "", TEMP_CAT)
   Call AddField("KEY_LOOKUP", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PIG_NO", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PIG_DESC", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PIG_TYPE", STRING_TYPE, "", TEMP_CAT)
   
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_DATE", DATE_TYPE, -1, TEMP_CAT)
   Call AddField("TO_DATE", DATE_TYPE, -1, TEMP_CAT)
End Sub


