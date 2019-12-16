VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExpenseType"
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

Public EXPENSE_TYPE_ID As Long
Public EXPENSE_TYPE_NO As String
Public EXPENSE_TYPE_NAME As String
Public BUY_FLAG As String
Public DEPLICATE_FLAG As String
Public DEPRECIATION_GOOD_FLAG As String
Public SELL_PARENT_FLAG As String

Public KEY_ID As Long
Public KEY_LOOKUP As String

Public ExpenseRatios As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set ExpenseRatios = New Collection
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
   
   FromStr = "FROM EXPENSE_TYPE UG "
   
   WhereStr = ""
   If EXPENSE_TYPE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EXPENSE_TYPE_ID = " & EXPENSE_TYPE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EXPENSE_TYPE_ID = " & EXPENSE_TYPE_ID & ") "
      End If
   End If
   
   If Len(BUY_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BUY_FLAG = '" & BUY_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (BUY_FLAG = '" & BUY_FLAG & "') "
      End If
   End If
   
   If Len(DEPLICATE_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DEPLICATE_FLAG = '" & DEPLICATE_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (DEPLICATE_FLAG = '" & DEPLICATE_FLAG & "') "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY EXPENSE_TYPE_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY EXPENSE_TYPE_NAME "
   Else
      OrderByStr = "ORDER BY EXPENSE_TYPE_NO "
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
   EXPENSE_TYPE_ID = NVLI(Rs("EXPENSE_TYPE_ID"), -1)
   EXPENSE_TYPE_NAME = NVLS(Rs("EXPENSE_TYPE_NAME"), "")
   EXPENSE_TYPE_NO = NVLS(Rs("EXPENSE_TYPE_NO"), "")
   BUY_FLAG = NVLS(Rs("BUY_FLAG"), "")
   DEPLICATE_FLAG = NVLS(Rs("DEPLICATE_FLAG"), "")
   SELL_PARENT_FLAG = NVLS(Rs("SELL_PARENT_FLAG"), "")
   DEPRECIATION_GOOD_FLAG = NVLS(Rs("DEPRECIATION_GOOD_FLAG"), "")
   
   KEY_LOOKUP = EXPENSE_TYPE_NO
   KEY_ID = EXPENSE_TYPE_ID
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("EXPENSE_TYPE_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO EXPENSE_TYPE " & _
                     "( " & _
                     "   EXPENSE_TYPE_ID, " & _
                     "   EXPENSE_TYPE_NAME, " & _
                     "   EXPENSE_TYPE_NO, " & _
                     "   BUY_FLAG, " & _
                     "   DEPLICATE_FLAG, " & _
                     "   SELL_PARENT_FLAG, " & _
                     "   DEPRECIATION_GOOD_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(EXPENSE_TYPE_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(EXPENSE_TYPE_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(BUY_FLAG)) & "'," & _
                         "'" & ChangeQuote(Trim(DEPLICATE_FLAG)) & "',"
            SQL1 = SQL1 & _
                        "'" & ChangeQuote(Trim(SELL_PARENT_FLAG)) & "'," & _
                        "'" & ChangeQuote(Trim(DEPRECIATION_GOOD_FLAG)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         EXPENSE_TYPE_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE EXPENSE_TYPE SET " & _
                     "   EXPENSE_TYPE_NO = '" & ChangeQuote(Trim(EXPENSE_TYPE_NO)) & "', " & _
                     "   EXPENSE_TYPE_NAME = '" & ChangeQuote(Trim(EXPENSE_TYPE_NAME)) & "', " & _
                     "   BUY_FLAG = '" & ChangeQuote(Trim(BUY_FLAG)) & "', " & _
                     "   SELL_PARENT_FLAG = '" & ChangeQuote(Trim(SELL_PARENT_FLAG)) & "', " & _
                     "   DEPRECIATION_GOOD_FLAG = '" & ChangeQuote(Trim(DEPRECIATION_GOOD_FLAG)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE EXPENSE_TYPE_ID = " & EXPENSE_TYPE_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM EXPENSE_RATIO WHERE EXPENSE_TYPE_ID = " & EXPENSE_TYPE_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM EXPENSE_TYPE WHERE EXPENSE_TYPE_ID = " & EXPENSE_TYPE_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set ExpenseRatios = Nothing
End Sub
