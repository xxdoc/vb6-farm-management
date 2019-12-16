VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTrnPrmItem"
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

Dim TRNPRM_ITEM_ID As Long
Dim PARAM_ID As Long
Dim FROM_DATE As Date
Dim TO_DATE As Date
Dim ORDER_BY As Long
Dim ORDER_TYPE As Long

   TRNPRM_ITEM_ID = Val(GetFieldValue("TRNPRM_ITEM_ID"))
   PARAM_ID = Val(GetFieldValue("PARAM_ID"))
   ORDER_BY = Val(GetFieldValue("ORDER_BY"))
   ORDER_TYPE = Val(GetFieldValue("ORDER_TYPE"))
   
   If Ind = 1 Then
      SelectStr1 = "SELECT EM.*, PS.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " EM "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (EM.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
      
      WhereStr = ""
   End If
   
   If TRNPRM_ITEM_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = TempStr & " (EM.TRNPRM_ITEM_ID = " & TRNPRM_ITEM_ID & ") " & vbCrLf
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
         OrderByStr = "ORDER BY EM.TRNPRM_ITEM_ID "
      ElseIf ORDER_BY = 2 Then
         OrderByStr = "ORDER BY EM.TRNPRM_ITEM_ID "
      Else
         OrderByStr = "ORDER BY EM.TRNPRM_ITEM_ID "
      End If
   End If
   
   If ORDER_TYPE = 1 Then
      OrderByStr = OrderByStr & " ASC "
   Else
      OrderByStr = OrderByStr & " ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, False, itemcount, Rs, glbErrorLog)
End Sub

Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("TRNPRM_ITEM_ID", NVLI(Rs("TRNPRM_ITEM_ID"), -1))
      Call SetFieldValue("PARAM_ID", NVLS(Rs("PARAM_ID"), ""))
      Call SetFieldValue("PIG_STATUS", NVLD(Rs("PIG_STATUS"), 0))
      Call SetFieldValue("LOSS_RATE", NVLD(Rs("LOSS_RATE"), 0))
      Call SetFieldValue("LOSS_TYPE", NVLI(Rs("LOSS_TYPE"), -1))
      
      Call SetFieldValue("STATUS_NO", NVLS(Rs("PRODUCT_STATUS_NO"), ""))
      Call SetFieldValue("STATUS_NAME", NVLS(Rs("PRODUCT_STATUS_NAME"), ""))
      
      Call SetFieldValue("KEY_ID", NVLS(Rs("TRNPRM_ITEM_ID"), ""))
      Call SetFieldValue("KEY_LOOKUP", NVLS(Rs("PARAM_ID"), ""))
   End If
End Sub

Public Sub RegisterFields()
   TableName = "TRNPRM_ITEM"
   SequenceName = "TRNPRM_ITEM_SEQ"
   PKField = "TRNPRM_ITEM_ID"
   
   Call AddField("TRNPRM_ITEM_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("PARAM_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("PIG_STATUS", INT_TYPE, 0, DATA_CAT)
   Call AddField("LOSS_RATE", MONEY_TYPE, 0, DATA_CAT)
   Call AddField("LOSS_TYPE", INT_TYPE, -1, DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)
   
   Call AddField("KEY_ID", INT_TYPE, "", TEMP_CAT)
   Call AddField("KEY_LOOKUP", STRING_TYPE, "", TEMP_CAT)
   Call AddField("STATUS_NO", STRING_TYPE, "", TEMP_CAT)
   Call AddField("STATUS_NAME", STRING_TYPE, "", TEMP_CAT)
   
   Call AddField("ORDER_BY", INT_TYPE, "", TEMP_CAT)
   Call AddField("ORDER_TYPE", INT_TYPE, "", TEMP_CAT)
   Call AddField("FROM_DATE", DATE_TYPE, -1, TEMP_CAT)
   Call AddField("TO_DATE", DATE_TYPE, -1, TEMP_CAT)
End Sub

