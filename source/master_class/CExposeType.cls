VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExposeType"
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

Public EXPOSE_TYPE_ID As Long
Public EXPOSE_TYPE_NAME As String
Public EXPOSE_TYPE_NO As String

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

   SelectStr1 = "SELECT UG.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM EXPOSE_TYPE UG "
   
   WhereStr = ""
   If EXPOSE_TYPE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EXPOSE_TYPE_ID = " & EXPOSE_TYPE_ID & ") "
      Else
         WhereStr = "AND (EXPOSE_TYPE_ID = " & EXPOSE_TYPE_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY EXPOSE_TYPE_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY EXPOSE_TYPE_NAME "
   Else
      OrderByStr = "ORDER BY EXPOSE_TYPE_NO "
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
   EXPOSE_TYPE_ID = NVLI(Rs("EXPOSE_TYPE_ID"), -1)
   EXPOSE_TYPE_NAME = NVLS(Rs("EXPOSE_TYPE_NAME"), "")
   EXPOSE_TYPE_NO = NVLS(Rs("EXPOSE_TYPE_NO"), "")
   
   KEY_ID = EXPOSE_TYPE_ID
   KEY_LOOKUP = EXPOSE_TYPE_NO
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String

'   TempID1 = EXPOSE_TYPE_ID
'   If EXPOSE_TYPE_ID <= 0 Then
'      TempID1 = "NULL"
'   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("EXPOSE_TYPE_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO EXPOSE_TYPE " & _
                     "( " & _
                     "   EXPOSE_TYPE_ID, " & _
                     "   EXPOSE_TYPE_NAME, " & _
                     "   EXPOSE_TYPE_NO, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(EXPOSE_TYPE_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(EXPOSE_TYPE_NO)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         EXPOSE_TYPE_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE EXPOSE_TYPE SET " & _
                     "   EXPOSE_TYPE_NAME = '" & ChangeQuote(Trim(EXPOSE_TYPE_NAME)) & "', " & _
                     "   EXPOSE_TYPE_NO = '" & ChangeQuote(Trim(EXPOSE_TYPE_NO)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE EXPOSE_TYPE_ID = " & EXPOSE_TYPE_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM EXPOSE_TYPE WHERE EXPOSE_TYPE_ID = " & EXPOSE_TYPE_ID
      m_Conn.Execute (SQL1)
End Function




