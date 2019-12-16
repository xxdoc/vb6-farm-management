VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEnterpriseName"
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

Public ENTERPRISE_NAME_ID As Long
Public ENTERPRISE_ID As Long
Public NAME_ID As Long

Public Names As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set Names = New Collection
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
   
   FromStr = "FROM ENTERPRISE_NAME UG "

   WhereStr = ""
   If ENTERPRISE_NAME_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ENTERPRISE_NAME_ID = " & ENTERPRISE_NAME_ID & ") "
      Else
         WhereStr = "AND (ENTERPRISE_NAME_ID = " & ENTERPRISE_NAME_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   If OrderBy > 0 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY ENTERPRISE_NAME_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY ENTERPRISE_NAME_ID "
      Else
         OrderByStr = "ORDER BY ENTERPRISE_NAME_ID "
      End If
      
      If OrderType = 1 Then
         OrderTypeStr = "ASC "
      ElseIf OrderType = 2 Then
         OrderTypeStr = "DESC "
      Else
         OrderTypeStr = "DESC "
      End If
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Rs As ADODB.Recordset)
   ENTERPRISE_NAME_ID = NVLI(Rs("ENTERPRISE_NAME_ID"), -1)
   ENTERPRISE_ID = NVLI(Rs("ENTERPRISE_ID"), -1)
   NAME_ID = NVLI(Rs("NAME_ID"), -1)
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("ENTERPRISE_NAME_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO ENTERPRISE_NAME " & _
                     "( " & _
                     "   ENTERPRISE_NAME_ID, " & _
                     "   ENTERPRISE_ID, " & _
                     "   NAME_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "" & ENTERPRISE_ID & "," & _
                         "" & NAME_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         ENTERPRISE_NAME_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE ENTERPRISE_NAME SET " & _
                     "   ENTERPRISE_ID = " & ENTERPRISE_ID & ", " & _
                     "   NAME_ID = " & NAME_ID & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE ENTERPRISE_NAME_ID = " & ENTERPRISE_NAME_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM ENTERPRISE_NAME WHERE ENTERPRISE_NAME_ID = " & ENTERPRISE_NAME_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set Names = Nothing
End Sub
