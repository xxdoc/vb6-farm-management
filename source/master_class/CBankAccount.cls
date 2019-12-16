VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBankAccount"
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

Public BANK_ACCOUNT_ID As Long
Public ACCOUNT_NO As String
Public ACCOUNT_NAME As String
Public BANK_ID As Long
Public BBRANCH_ID As Long

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
   
   FromStr = "FROM BANK_ACCOUNT UG "
   
   WhereStr = ""
   If BANK_ACCOUNT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BANK_ACCOUNT_ID = " & BANK_ACCOUNT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BANK_ACCOUNT_ID = " & BANK_ACCOUNT_ID & ") "
      End If
   End If
   
   If BANK_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BANK_ID = " & BANK_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BANK_ID = " & BANK_ID & ") "
      End If
   End If
         
   If BBRANCH_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BBRANCH_ID = " & BBRANCH_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BBRANCH_ID = " & BBRANCH_ID & ") "
      End If
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY BANK_ACCOUNT_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY ACCOUNT_NAME "
   Else
      OrderByStr = "ORDER BY BANK_ACCOUNT_ID "
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
   BANK_ACCOUNT_ID = NVLI(Rs("BANK_ACCOUNT_ID"), -1)
   ACCOUNT_NAME = NVLS(Rs("ACCOUNT_NAME"), "")
   ACCOUNT_NO = NVLS(Rs("ACCOUNT_NO"), "")
   BANK_ID = NVLI(Rs("BANK_ID"), -1)
   BBRANCH_ID = NVLI(Rs("BBRANCH_ID"), -1)

   KEY_ID = BANK_ACCOUNT_ID
   KEY_LOOKUP = ACCOUNT_NO
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("BANK_ACCOUNT_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO BANK_ACCOUNT " & _
                     "( " & _
                     "   BANK_ACCOUNT_ID, " & _
                     "   ACCOUNT_NAME, " & _
                     "   ACCOUNT_NO, " & _
                     "   BANK_ID, " & _
                     "   BBRANCH_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(ACCOUNT_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(ACCOUNT_NO)) & "'," & _
                         BANK_ID & "," & _
                         BBRANCH_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         BANK_ACCOUNT_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE BANK_ACCOUNT SET " & _
                     "   ACCOUNT_NAME = '" & ChangeQuote(Trim(ACCOUNT_NAME)) & "', " & _
                     "   ACCOUNT_NO = '" & ChangeQuote(Trim(ACCOUNT_NO)) & "', " & _
                     "   BANK_ID = " & BANK_ID & ", " & _
                     "   BBRANCH_ID = " & BBRANCH_ID & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE BANK_ACCOUNT_ID = " & BANK_ACCOUNT_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM BANK_ACCOUNT WHERE BANK_ACCOUNT_ID = " & BANK_ACCOUNT_ID
      m_Conn.Execute (SQL1)
End Function
