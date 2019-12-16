VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAgeRange"
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

Public AGE_RANGE_ID As Long
Public AGE_RANGE_NO As String
Public AGE_RANGE_NAME As String
Public FROM_WEEK As Long
Public TO_WEEK As Long

Public KEY_ID As Long
Public KEY_LOOKUP As String

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
   
   FromStr = "FROM AGE_RANGE UG "
   
   WhereStr = ""
   If AGE_RANGE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (AGE_RANGE_ID = " & AGE_RANGE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (AGE_RANGE_ID = " & AGE_RANGE_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY AGE_RANGE_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY AGE_RANGE_NAME "
   Else
      OrderByStr = "ORDER BY AGE_RANGE_ID "
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
   AGE_RANGE_ID = NVLI(Rs("AGE_RANGE_ID"), -1)
   AGE_RANGE_NAME = NVLS(Rs("AGE_RANGE_NAME"), "")
   AGE_RANGE_NO = NVLS(Rs("AGE_RANGE_NO"), "")
   FROM_WEEK = NVLI(Rs("FROM_WEEK"), 0)
   TO_WEEK = NVLI(Rs("TO_WEEK"), 0)
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("AGE_RANGE_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO AGE_RANGE " & _
                     "( " & _
                     "   AGE_RANGE_ID, " & _
                     "   AGE_RANGE_NAME, " & _
                     "   AGE_RANGE_NO, " & _
                     "   FROM_WEEK, " & _
                     "   TO_WEEK, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(AGE_RANGE_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(AGE_RANGE_NO)) & "'," & _
                         FROM_WEEK & "," & _
                         TO_WEEK & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         AGE_RANGE_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE AGE_RANGE SET " & _
                     "   AGE_RANGE_NO = '" & ChangeQuote(Trim(AGE_RANGE_NO)) & "', " & _
                     "   AGE_RANGE_NAME = '" & ChangeQuote(Trim(AGE_RANGE_NAME)) & "', " & _
                     "   FROM_WEEK = " & FROM_WEEK & ", " & _
                     "   TO_WEEK = " & TO_WEEK & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE AGE_RANGE_ID = " & AGE_RANGE_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM AGE_RANGE WHERE AGE_RANGE_ID = " & AGE_RANGE_ID
      m_Conn.Execute (SQL1)
End Function


