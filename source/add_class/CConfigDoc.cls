VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConfigDoc"
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
   Call SetFieldValue("ENTERPRISE_ID", 99999)
   
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
Dim TempRs As ADODB.Recordset
Dim iCount As Long

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

Dim CONFIG_DOC_TYPE As Long
Dim CONFIG_DOC_ID As Long
Dim ENTERPRISE_ID As Long

   CONFIG_DOC_TYPE = Val(GetFieldValue("CONFIG_DOC_TYPE"))
   CONFIG_DOC_ID = Val(GetFieldValue("CONFIG_DOC_ID"))
   ENTERPRISE_ID = 99999
   
   If Ind = 1 Then
      SelectStr1 = "SELECT NM.* "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM " & TableName & " NM "
      
      WhereStr = ""
   End If
   
   If CONFIG_DOC_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CONFIG_DOC_ID = " & CONFIG_DOC_ID & ") " & vbCrLf
   End If
      
   If CONFIG_DOC_TYPE > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.CONFIG_DOC_TYPE = " & CONFIG_DOC_TYPE & ") " & vbCrLf
   End If
      
   If ENTERPRISE_ID > 0 Then
      If Len(WhereStr) > 0 Then
         TempStr = "AND "
      Else
         TempStr = "WHERE "
      End If
      WhereStr = WhereStr & TempStr & " (NM.ENTERPRISE_ID = " & ENTERPRISE_ID & ") " & vbCrLf
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Sub
Public Sub PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      Call SetFieldValue("CONFIG_DOC_ID", NVLI(Rs("CONFIG_DOC_ID"), -1))
      Call SetFieldValue("CONFIG_DOC_TYPE", NVLI(Rs("CONFIG_DOC_TYPE"), -1))
      Call SetFieldValue("ENTERPRISE_ID", NVLI(Rs("ENTERPRISE_ID"), -1))
      Call SetFieldValue("PREFIX", NVLS(Rs("PREFIX"), ""))
      Call SetFieldValue("CODE1", NVLS(Rs("CODE1"), ""))
      Call SetFieldValue("YEAR_TYPE", NVLI(Rs("YEAR_TYPE"), -1))
      Call SetFieldValue("CODE2", NVLS(Rs("CODE2"), ""))
      Call SetFieldValue("MONTH_TYPE", NVLI(Rs("MONTH_TYPE"), -1))
      Call SetFieldValue("CODE3", NVLS(Rs("CODE3"), ""))
      Call SetFieldValue("DIGIT_AMOUNT", NVLI(Rs("DIGIT_AMOUNT"), -1))
      Call SetFieldValue("RUNNING_NO", NVLI(Rs("RUNNING_NO"), 0))
      Call SetFieldValue("LAST_NO", NVLS(Rs("LAST_NO"), ""))
      
      Call SetFieldValue("MM", NVLS(Rs("MM"), ""))
      Call SetFieldValue("UPDATE_MONTH_FLAG", NVLS(Rs("UPDATE_MONTH_FLAG"), ""))
      Call SetFieldValue("UPDATE_YEAR_FLAG", NVLS(Rs("UPDATE_YEAR_FLAG"), ""))
      
   End If
End Sub
Public Sub RegisterFields()
   TableName = "CONFIG_DOC"
   SequenceName = "CONFIG_DOC_SEQ"
   PKField = "CONFIG_DOC_ID"
   
   Call AddField("CONFIG_DOC_ID", INT_TYPE, -1, ID_CAT)
   Call AddField("CONFIG_DOC_TYPE", INT_TYPE, "", DATA_CAT)
   Call AddField("ENTERPRISE_ID", INT_TYPE, -1, DATA_CAT)
   Call AddField("PREFIX", STRING_TYPE, "", DATA_CAT)
   Call AddField("CODE1", STRING_TYPE, "", DATA_CAT)
   Call AddField("YEAR_TYPE", INT_TYPE, "", DATA_CAT)
   Call AddField("CODE2", STRING_TYPE, "", DATA_CAT)
   Call AddField("MONTH_TYPE", INT_TYPE, "", DATA_CAT)
   Call AddField("CODE3", STRING_TYPE, "", DATA_CAT)
   Call AddField("DIGIT_AMOUNT", INT_TYPE, 0, DATA_CAT)
   Call AddField("RUNNING_NO", INT_TYPE, 0, DATA_CAT)
   Call AddField("LAST_NO", STRING_TYPE, "", DATA_CAT)
   Call AddField("MM", STRING_TYPE, 0, DATA_CAT)
   Call AddField("UPDATE_MONTH_FLAG", STRING_TYPE, "", DATA_CAT)
   Call AddField("UPDATE_YEAR_FLAG", STRING_TYPE, "", DATA_CAT)
   
   Call AddField("CREATE_DATE", DATE_TYPE, -1, CREATE_DATE_CAT)
   Call AddField("MODIFY_DATE", DATE_TYPE, -1, MODIFY_DATE_CAT)
   Call AddField("CREATE_BY", INT_TYPE, -1, CREATE_BY_CAT)
   Call AddField("MODIFY_BY", INT_TYPE, -1, MODIFY_BY_CAT)

End Sub
Public Sub UpdateRunningNo()
Dim ServerDateTime As String
Dim SQL1 As String
Dim TransformToSQLString As String
            
      If Val(99999) <= 0 Then
         TransformToSQLString = ""
      Else
         TransformToSQLString = " AND ENTERPRISE_ID = '" & Val(99999) & "'"
      End If
      
      SQL1 = "UPDATE CONFIG_DOC SET " & _
                  "    RUNNING_NO = " & Val(GetFieldValue("RUNNING_NO")) & ", " & _
                  "    LAST_NO = '" & GetFieldValue("LAST_NO") & "', " & _
                  "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE CONFIG_DOC_TYPE = " & Val(GetFieldValue("CONFIG_DOC_TYPE")) & TransformToSQLString

   Call m_ADOConn.Execute(SQL1)
End Sub
Public Sub UpdateYearMonthRunningNo()
Dim ServerDateTime As String
Dim SQL1 As String
Dim TransformToSQLString As String
                  
      Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
      If Val(99999) <= 0 Then
         TransformToSQLString = ""
      Else
         TransformToSQLString = " AND ENTERPRISE_ID = '" & Val(99999) & "'"
      End If
      
      SQL1 = "UPDATE CONFIG_DOC SET " & _
                  "    RUNNING_NO = " & Val(GetFieldValue("RUNNING_NO")) & ", " & _
                  "    MM = '" & GetFieldValue("MM") & "', " & _
                  "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE CONFIG_DOC_TYPE = " & Val(GetFieldValue("CONFIG_DOC_TYPE")) & TransformToSQLString

   Call m_ADOConn.Execute(SQL1)
End Sub


