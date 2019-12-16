VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBankBranch"
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

Public BBRANCH_ID As Long
Public BBRANCH_NO As String
Public BBRANCH_NAME As String
Public BANK_ID As Long

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
   
   FromStr = "FROM BANK_BRANCH UG "
   
   WhereStr = ""
   If BBRANCH_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BBRANCH_ID = " & BBRANCH_ID & ") "
      Else
         WhereStr = "AND (BBRANCH_ID = " & BBRANCH_ID & ") "
      End If
   End If
   
   If BANK_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BANK_ID = " & BANK_ID & ") "
      Else
         WhereStr = "AND (BANK_ID = " & BANK_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY BBRANCH_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY BBRANCH_NAME "
   Else
      OrderByStr = "ORDER BY BBRANCH_ID "
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
   BBRANCH_ID = NVLI(Rs("BBRANCH_ID"), -1)
   BBRANCH_NAME = NVLS(Rs("BBRANCH_NAME"), "")
   BBRANCH_NO = NVLS(Rs("BBRANCH_NO"), "")
   BANK_ID = NVLI(Rs("BANK_ID"), -1)
   
   KEY_ID = BBRANCH_ID
   KEY_LOOKUP = BBRANCH_NO
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String

   TempID1 = BANK_ID
   If BANK_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("BBRANCH_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO BANK_BRANCH " & _
                     "( " & _
                     "   BBRANCH_ID, " & _
                     "   BBRANCH_NAME, " & _
                     "   BBRANCH_NO, " & _
                     "   BANK_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(BBRANCH_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(BBRANCH_NO)) & "'," & _
                         TempID1 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         BBRANCH_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE BANK_BRANCH SET " & _
                     "   BANK_ID = " & BANK_ID & ", " & _
                     "   BBRANCH_NAME = '" & ChangeQuote(Trim(BBRANCH_NAME)) & "', " & _
                     "   BBRANCH_NO = '" & ChangeQuote(Trim(BBRANCH_NO)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE BBRANCH_ID = " & BBRANCH_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM BANK_BRANCH WHERE BBRANCH_ID = " & BBRANCH_ID
      m_Conn.Execute (SQL1)
End Function
