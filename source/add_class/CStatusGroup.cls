VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStatusGroup"
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

Public STATUS_GROUP_ID As Long
Public STATUS_GROUP_NAME As String
Public STATUS_GROUP_NO As String

Public KEY_ID As Long
Public KEY_LOOKUP As String

Public HGroupItems As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set HGroupItems = New Collection
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
   
   FromStr = "FROM STATUS_GROUP UG "
   
   WhereStr = ""
   If STATUS_GROUP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (STATUS_GROUP_ID = " & STATUS_GROUP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (STATUS_GROUP_ID = " & STATUS_GROUP_ID & ") "
      End If
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY STATUS_GROUP_NO "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY STATUS_GROUP_NAME "
   Else
      OrderByStr = "ORDER BY STATUS_GROUP_ID "
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

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   STATUS_GROUP_ID = NVLI(Rs("STATUS_GROUP_ID"), -1)
   STATUS_GROUP_NAME = NVLS(Rs("STATUS_GROUP_NAME"), "")
   STATUS_GROUP_NO = NVLS(Rs("STATUS_GROUP_NO"), "")
   
   KEY_ID = STATUS_GROUP_ID
   KEY_LOOKUP = STATUS_GROUP_NO
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String

   TempID1 = STATUS_GROUP_ID
   If STATUS_GROUP_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("STATUS_GROUP_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO STATUS_GROUP " & _
                     "( " & _
                     "   STATUS_GROUP_ID, " & _
                     "   STATUS_GROUP_NAME, " & _
                     "   STATUS_GROUP_NO, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(STATUS_GROUP_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(STATUS_GROUP_NO)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         STATUS_GROUP_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE STATUS_GROUP SET " & _
                     "   STATUS_GROUP_NAME = '" & ChangeQuote(Trim(STATUS_GROUP_NAME)) & "', " & _
                     "   STATUS_GROUP_NO = '" & ChangeQuote(Trim(STATUS_GROUP_NO)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE STATUS_GROUP_ID = " & STATUS_GROUP_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM SGROUP_ITEM WHERE STATUS_GROUP_ID = " & STATUS_GROUP_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM STATUS_GROUP WHERE STATUS_GROUP_ID = " & STATUS_GROUP_ID
      m_Conn.Execute (SQL1)
End Function


