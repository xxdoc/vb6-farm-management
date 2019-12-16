VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmployeeName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long
Public Flag As String
' End of generic data member

Public EMPLOYEE_NAME_ID As Long
Public EMP_ID As String
Public NAME_ID As Long

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String

   SelectStr1 = "SELECT UG.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM EMPLOYEE_NAME UG "
   
   WhereStr = ""
   If EMPLOYEE_NAME_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EMPLOYEE_NAME_ID = " & EMPLOYEE_NAME_ID & ") "
      Else
         WhereStr = "AND (EMPLOYEE_NAME_ID = " & EMPLOYEE_NAME_ID & ") "
      End If
   End If
   
   If EMP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EMP_ID = " & EMP_ID & ") "
      Else
         WhereStr = "AND (EMP_ID = " & EMP_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
      
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY EMPLOYEE_NAME_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY EMP_ID "
   Else
      OrderByStr = "ORDER BY EMPLOYEE_NAME_ID "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Rs As ADODB.Recordset)
   EMPLOYEE_NAME_ID = NVLI(Rs("EMPLOYEE_NAME_ID"), -1)
   EMP_ID = NVLI(Rs("EMP_ID"), -1)
   NAME_ID = NVLI(Rs("NAME_ID"), -1)
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("EMPLOYEE_NAME_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO EMPLOYEE_NAME " & _
                     "( " & _
                     "   EMPLOYEE_NAME_ID, " & _
                     "   EMP_ID, " & _
                     "   NAME_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(EMP_ID)) & "'," & _
                         NAME_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         EMPLOYEE_NAME_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE EMPLOYEE_NAME SET " & _
                     " NAME_ID = " & NAME_ID & ", " & _
                     "   EMP_ID = '" & ChangeQuote(Trim(EMP_ID)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE EMPLOYEE_NAME_ID = " & EMPLOYEE_NAME_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM UNIT WHERE EMPLOYEE_NAME_ID = " & EMPLOYEE_NAME_ID
      m_Conn.Execute (SQL1)
End Function


