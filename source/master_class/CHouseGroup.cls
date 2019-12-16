VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHouseGroup"
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

Public HOUSE_GROUP_ID As Long
Public HOUSE_GROUP_NAME As String
Public HOUSE_GROUP_NO As String
Public EXTRA_FLAG As String

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
   
   FromStr = "FROM HOUSE_GROUP UG "
   
   WhereStr = ""
   If HOUSE_GROUP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UG.HOUSE_GROUP_ID = " & HOUSE_GROUP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (UG.HOUSE_GROUP_ID = " & HOUSE_GROUP_ID & ") "
      End If
   End If
   
   If Len(EXTRA_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EXTRA_FLAG = '" & EXTRA_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (EXTRA_FLAG = '" & EXTRA_FLAG & "') "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY HOUSE_GROUP_NO "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY HOUSE_GROUP_NAME "
   Else
      OrderByStr = "ORDER BY UG.HOUSE_GROUP_ID "
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
   HOUSE_GROUP_ID = NVLI(Rs("HOUSE_GROUP_ID"), -1)
   HOUSE_GROUP_NAME = NVLS(Rs("HOUSE_GROUP_NAME"), "")
   HOUSE_GROUP_NO = NVLS(Rs("HOUSE_GROUP_NO"), "")
   EXTRA_FLAG = NVLS(Rs("EXTRA_FLAG"), "")
   
   KEY_ID = HOUSE_GROUP_ID
   KEY_LOOKUP = HOUSE_GROUP_NO
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String

   TempID1 = HOUSE_GROUP_ID
   If HOUSE_GROUP_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("HOUSE_GROUP_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO HOUSE_GROUP " & _
                     "( " & _
                     "   HOUSE_GROUP_ID, " & _
                     "   HOUSE_GROUP_NAME, " & _
                     "   HOUSE_GROUP_NO, " & _
                     "   EXTRA_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(HOUSE_GROUP_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(HOUSE_GROUP_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(EXTRA_FLAG)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         HOUSE_GROUP_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE HOUSE_GROUP SET " & _
                     "   HOUSE_GROUP_NAME = '" & ChangeQuote(Trim(HOUSE_GROUP_NAME)) & "', " & _
                     "   HOUSE_GROUP_NO = '" & ChangeQuote(Trim(HOUSE_GROUP_NO)) & "', " & _
                     "   EXTRA_FLAG = '" & ChangeQuote(Trim(EXTRA_FLAG)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE HOUSE_GROUP_ID = " & HOUSE_GROUP_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM HGROUP_ITEM WHERE HOUSE_GROUP_ID = " & HOUSE_GROUP_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM HOUSE_GROUP WHERE HOUSE_GROUP_ID = " & HOUSE_GROUP_ID
      m_Conn.Execute (SQL1)
End Function




