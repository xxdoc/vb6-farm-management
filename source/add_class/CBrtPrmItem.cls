VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBrtPrmItem"
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

Dim BRTPRM_ITEM_ID As Long
Dim PARAM_ID As Long
Dim FROM_DATE As Date
Dim TO_DATE As Date
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long

   BRTPRM_ITEM_ID = Val(GetFieldValue("BRTPRM_ITEM_ID"))
   PARAM_ID = Val(GetFieldValue("PARAM_ID"))
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   If Ind = 1 Then
      SelectStr1 = "SELECT EM.*, PI1.PART_NO PIGG_NO, PI1.PART_DESC PIGG_DESC, PI1.PIG_TYPE PIGG_TYPE, PI2.PART_NO PIGL_NO, PI2.PART_DESC PIGL_DESC, PI2.PIG_TYPE PIGL_TYPE  "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EM "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI1 ON (EM.PIGG_ID = PI1.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI2 ON (EM.PIGL_ID = PI2.PART_ITEM_ID) "
      
      WhereStr = ""
   End If
   
   If BRTPRM_ITEM_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (EM.BRTPRM_ITEM_ID = " & BRTPRM_ITEM_ID & ") " & vbCrLf
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
         OrderByStr = "ORDER BY EM.FROM_BREED "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY EM.FROM_BREED "
      Else
         OrderByStr = "ORDER BY EM.FROM_BREED "
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
      Call SetFieldValue("BRTPRM_ITEM_ID", NVLI(Rs("BRTPRM_ITEM_ID"), -1))
      Call SetFieldValue("PARAM_ID", NVLS(Rs("PARAM_ID"), ""))
      Call SetFieldValue("BREED_AMOUNT", NVLD(Rs("BREED_AMOUNT"), 0))
      Call SetFieldValue("BIRTH_AMOUNT", NVLD(Rs("BIRTH_AMOUNT"), 0))
      Call SetFieldValue("BREED_RATE", NVLD(Rs("BREED_RATE"), 0))
      
      Call SetFieldValue("PREGNANT_RATE", NVLD(Rs("PREGNANT_RATE"), 0))
      Call SetFieldValue("BIRTH_RATE", NVLD(Rs("BIRTH_RATE"), 0))
      Call SetFieldValue("CHILD_RATE", NVLD(Rs("CHILD_RATE"), 0))
      Call SetFieldValue("FROM_BIRTH", InternalDateToDate(NVLS(Rs("FROM_BIRTH"), "")))
      Call SetFieldValue("TO_BIRTH", InternalDateToDate(NVLS(Rs("TO_BIRTH"), "")))
      
      Call SetFieldValue("FROM_BREED", InternalDateToDate(NVLS(Rs("FROM_BREED"), "")))
      Call SetFieldValue("TO_BREED", InternalDateToDate(NVLS(Rs("TO_BREED"), "")))
      Call SetFieldValue("DAY_COUNT", NVLD(Rs("DAY_COUNT"), 0))
      Call SetFieldValue("AVG_WEIGHT", NVLD(Rs("AVG_WEIGHT"), 0))
      Call SetFieldValue("PIGG_ID", NVLI(Rs("PIGG_ID"), -1))
      
      Call SetFieldValue("PIGL_ID", NVLI(Rs("PIGL_ID"), -1))
      Call SetFieldValue("PIGG_AMOUNT", NVLD(Rs("PIGG_AMOUNT"), 0))
      Call SetFieldValue("PIGL_AMOUNT", NVLD(Rs("PIGL_AMOUNT"), 0))
      Call SetFieldValue("PIGGL_AMOUNT", NVLD(Rs("PIGGL_AMOUNT"), 0))
      Call SetFieldValue("BREED_PERCENT", NVLD(Rs("BREED_PERCENT"), 0))
      
      Call SetFieldValue("BIRTH_COST", NVLD(Rs("BIRTH_COST"), 0))
      Call SetFieldValue("PIGG_NO", NVLS(Rs("PIGG_NO"), ""))
      Call SetFieldValue("PIGG_DESC", NVLS(Rs("PIGG_DESC"), ""))
      Call SetFieldValue("PIGG_TYPE", NVLS(Rs("PIGG_TYPE"), ""))
      Call SetFieldValue("PIGL_NO", NVLS(Rs("PIGL_NO"), ""))
      
      Call SetFieldValue("PIGL_DESC", NVLS(Rs("PIGL_DESC"), ""))
      Call SetFieldValue("PIGL_TYPE", NVLS(Rs("PIGL_TYPE"), ""))
      
      Call SetFieldValue("KEY_ID", NVLS(Rs("BRTPRM_ITEM_ID"), ""))
      Call SetFieldValue("KEY_LOOKUP", NVLS(Rs("PARAM_ID"), ""))
   End If
End Sub

Public Sub RegisterFields()
   TableName = "BRTPRM_ITEM"
   SequenceName = "BRTPRM_ITEM_SEQ"
   PKField = "BRTPRM_ITEM_ID"
   
   Call AddField("BRTPRM_ITEM_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("PARAM_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("BREED_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("BIRTH_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("FROM_BIRTH", DATE_TYPE, -1, DATA_CAT)
   Call AddField("TO_BIRTH", DATE_TYPE, -1, DATA_CAT)
   Call AddField("FROM_BREED", DATE_TYPE, -1, DATA_CAT)
   Call AddField("TO_BREED", DATE_TYPE, -1, DATA_CAT)
   Call AddField("BREED_RATE", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("PREGNANT_RATE", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("BIRTH_RATE", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("CHILD_RATE", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("DAY_COUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("AVG_WEIGHT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("PIGG_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("PIGL_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("PIGG_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("PIGL_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("PIGGL_AMOUNT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("BREED_PERCENT", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("BIRTH_COST", MONEY_TYPE, 0, DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("KEY_ID", INT_TYPE, "", TEMP_CAT)
   Call AddField("KEY_LOOKUP", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PIGG_NO", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PIGG_TYPE", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PIGG_DESC", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PIGL_NO", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PIGL_TYPE", STRING_TYPE, "", TEMP_CAT)
   Call AddField("PIGL_DESC", STRING_TYPE, "", TEMP_CAT)
   
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_DATE", DATE_TYPE, -1, TEMP_CAT)
   Call AddField("TO_DATE", DATE_TYPE, -1, TEMP_CAT)
End Sub
Public Sub CopyItemCollection(TempBi As CBrtPrmItem)
      Call TempBi.SetFieldValue("BRTPRM_ITEM_ID", GetFieldValue("BRTPRM_ITEM_ID"))
      Call TempBi.SetFieldValue("PARAM_ID", GetFieldValue("PARAM_ID"))
      Call TempBi.SetFieldValue("BREED_AMOUNT", GetFieldValue("BREED_AMOUNT"))
      Call TempBi.SetFieldValue("BIRTH_AMOUNT", GetFieldValue("BIRTH_AMOUNT"))
      Call TempBi.SetFieldValue("BREED_RATE", GetFieldValue("BREED_RATE"))
      
      Call TempBi.SetFieldValue("PREGNANT_RATE", GetFieldValue("PREGNANT_RATE"))
      Call TempBi.SetFieldValue("BIRTH_RATE", GetFieldValue("BIRTH_RATE"))
      Call TempBi.SetFieldValue("CHILD_RATE", GetFieldValue("CHILD_RATE"))
      Call TempBi.SetFieldValue("FROM_BIRTH", GetFieldValue("FROM_BIRTH"))
      Call TempBi.SetFieldValue("TO_BIRTH", GetFieldValue("TO_BIRTH"))
      
      Call TempBi.SetFieldValue("FROM_BREED", GetFieldValue("FROM_BREED"))
      Call TempBi.SetFieldValue("TO_BREED", GetFieldValue("TO_BREED"))
      Call TempBi.SetFieldValue("DAY_COUNT", GetFieldValue("DAY_COUNT"))
      Call TempBi.SetFieldValue("AVG_WEIGHT", GetFieldValue("AVG_WEIGHT"))
      Call TempBi.SetFieldValue("PIGG_ID", GetFieldValue("PIGG_ID"))
      
      Call TempBi.SetFieldValue("PIGL_ID", GetFieldValue("PIGL_ID"))
      Call TempBi.SetFieldValue("PIGG_AMOUNT", GetFieldValue("PIGG_AMOUNT"))
      Call TempBi.SetFieldValue("PIGL_AMOUNT", GetFieldValue("PIGL_AMOUNT"))
      Call TempBi.SetFieldValue("PIGGL_AMOUNT", GetFieldValue("PIGGL_AMOUNT"))
      Call TempBi.SetFieldValue("BREED_PERCENT", GetFieldValue("BREED_PERCENT"))
      
      Call TempBi.SetFieldValue("BIRTH_COST", GetFieldValue("BIRTH_COST"))
      Call TempBi.SetFieldValue("PIGG_NO", GetFieldValue("PIGG_NO"))
      Call TempBi.SetFieldValue("PIGG_DESC", GetFieldValue("PIGG_DESC"))
      Call TempBi.SetFieldValue("PIGG_TYPE", GetFieldValue("PIGG_TYPE"))
      Call TempBi.SetFieldValue("PIGL_NO", GetFieldValue("PIGL_NO"))
      
      Call TempBi.SetFieldValue("PIGL_DESC", GetFieldValue("PIGL_DESC"))
      Call TempBi.SetFieldValue("PIGL_TYPE", GetFieldValue("PIGL_TYPE"))
End Sub
Public Sub UpdatePigBirthCost()
Dim ServerDateTime As String
Dim SQL1 As String
Dim Tf As CTableField

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   Set Tf = m_FieldList(PKField)
   
   SQL1 = "UPDATE BRTPRM_ITEM SET " & _
                  "   BIRTH_COST =  " & GetFieldValue("BIRTH_COST") & ", " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE BRTPRM_ITEM_ID = " & Tf.GetValue
   m_ADOConn.Execute (SQL1)
End Sub



