VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmpPosition"
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

Public POSITION_ID As Long
Public PPOSITION_ID As Long
Public POSITION_NAME As String
Public POSITION_DESC As String

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
   
   FromStr = "FROM EMP_POSITION UG "
   
   WhereStr = ""
   If POSITION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (POSITION_ID = " & POSITION_ID & ") "
      Else
         WhereStr = "AND (POSITION_ID = " & POSITION_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY POSITION_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY POSITION_NAME "
   Else
      OrderByStr = "ORDER BY POSITION_ID "
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
   POSITION_ID = NVLI(Rs("POSITION_ID"), -1)
   PPOSITION_ID = NVLI(Rs("PPOSITION_ID"), -1)
   POSITION_NAME = NVLS(Rs("POSITION_NAME"), "")
   POSITION_DESC = NVLS(Rs("POSITION_DESC"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   TempID1 = PPOSITION_ID
   If PPOSITION_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("EMP_POSITION_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO EMP_POSITION " & _
                     "( " & _
                     "   POSITION_ID, " & _
                     "   PPOSITION_ID, " & _
                     "   POSITION_NAME, " & _
                     "   POSITION_DESC, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         TempID1 & "," & _
                         "'" & ChangeQuote(Trim(POSITION_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(POSITION_DESC)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         POSITION_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE EMP_POSITION SET " & _
                     "   PPOSITION_ID = " & TempID1 & ", " & _
                     "   POSITION_NAME = '" & ChangeQuote(Trim(POSITION_NAME)) & "', " & _
                     "   POSITION_DESC = '" & ChangeQuote(Trim(POSITION_DESC)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE POSITION_ID = " & POSITION_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM EMP_POSITION WHERE POSITION_ID = " & POSITION_ID
      m_Conn.Execute (SQL1)
End Function


