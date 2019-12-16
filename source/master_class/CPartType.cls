VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPartType"
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

Public PART_TYPE_ID As Long
Public PART_TYPE_NAME As String
Public RAW_FLAG As String
Public PART_TYPE_NO As String
Public PART_GROUP_ID As Long

Public PART_NO As String
Public PART_GROUP_NO As String
Public PART_GROUP_NAME As String

Public KEY_ID As Long
Public KEY_LOOKUP As String

Public RightItems As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set RightItems = New Collection
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

   SelectStr1 = "SELECT UG.*, PG.PART_GROUP_NO, PG.PART_GROUP_NAME  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PART_TYPE UG "
   FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (UG.PART_GROUP_ID = PG.PART_GROUP_ID) "
   
   WhereStr = ""
   If PART_TYPE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_TYPE_ID = " & PART_TYPE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PART_TYPE_ID = " & PART_TYPE_ID & ") "
      End If
   End If
   
   If PART_GROUP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UG.PART_GROUP_ID = " & PART_GROUP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (UG.PART_GROUP_ID = " & PART_GROUP_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PART_TYPE_NO "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PART_TYPE_NAME "
   Else
      OrderByStr = "ORDER BY PART_GROUP_NO "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   PART_TYPE_ID = NVLI(Rs("PART_TYPE_ID"), -1)
   PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
   PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
   RAW_FLAG = NVLS(Rs("RAW_FLAG"), "")
   PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), -1)
   
   PART_GROUP_NO = NVLS(Rs("PART_GROUP_NO"), "")
   PART_GROUP_NAME = NVLS(Rs("PART_GROUP_NAME"), "")
   
   KEY_ID = PART_TYPE_ID
   KEY_LOOKUP = PART_TYPE_NO
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String

   TempID1 = PART_GROUP_ID
   If PART_GROUP_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("PART_TYPE_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO PART_TYPE " & _
                     "( " & _
                     "   PART_TYPE_ID, " & _
                     "   PART_TYPE_NAME, " & _
                     "   RAW_FLAG, " & _
                     "   PART_TYPE_NO, " & _
                     "   PART_GROUP_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(PART_TYPE_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(RAW_FLAG)) & "'," & _
                         "'" & ChangeQuote(Trim(PART_TYPE_NO)) & "'," & _
                         TempID1 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         PART_TYPE_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE PART_TYPE SET " & _
                     "   PART_TYPE_NAME = '" & ChangeQuote(Trim(PART_TYPE_NAME)) & "', " & _
                     "   RAW_FLAG = '" & ChangeQuote(Trim(RAW_FLAG)) & "', " & _
                     "   PART_TYPE_NO = '" & ChangeQuote(Trim(PART_TYPE_NO)) & "', " & _
                     "   PART_GROUP_ID = " & TempID1 & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PART_TYPE_ID = " & PART_TYPE_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM PART_TYPE WHERE PART_TYPE_ID = " & PART_TYPE_ID
      m_Conn.Execute (SQL1)
End Function




