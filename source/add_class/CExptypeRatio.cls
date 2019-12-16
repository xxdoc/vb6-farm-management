VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExptypeRatio"
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

Public EXPTYPE_RATIO_ID As Long
Public EXPENSE_TYPE_ID As Long
Public LOCATION_ID As Long
Public RATIO As Double
Public RATIO_AMOUNT As Double

Public LOCATION_NAME As String
Public LOCATION_NO As String

Public KEY_ID As Long
Public KEY_LOOKUP As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String

   SelectStr1 = "SELECT UG.*, LC.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM EXPTYPE_RATIO UG, LOCATION LC "
   
   WhereStr = "WHERE (UG.LOCATION_ID = LC.LOCATION_ID) "
   If EXPTYPE_RATIO_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EXPTYPE_RATIO_ID = " & EXPTYPE_RATIO_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EXPTYPE_RATIO_ID = " & EXPTYPE_RATIO_ID & ") "
      End If
   End If
   
   If EXPENSE_TYPE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EXPENSE_TYPE_ID = " & EXPENSE_TYPE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EXPENSE_TYPE_ID = " & EXPENSE_TYPE_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY UG.LOCATION_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY EXPENSE_TYPE_NAME "
   Else
      OrderByStr = "ORDER BY LOCATION_NAME "
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
   EXPTYPE_RATIO_ID = NVLI(Rs("EXPTYPE_RATIO_ID"), -1)
   EXPENSE_TYPE_ID = NVLI(Rs("EXPENSE_TYPE_ID"), -1)
   LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
   RATIO = NVLD(Rs("RATIO"), 0)
   RATIO_AMOUNT = NVLD(Rs("RATIO_AMOUNT"), 0)
   
'''debug.print SELECT_FLAG
   LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
   LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("EXPTYPE_RATIO_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO EXPTYPE_RATIO " & _
                     "( " & _
                     "   EXPTYPE_RATIO_ID, " & _
                     "   EXPENSE_TYPE_ID, " & _
                     "   LOCATION_ID, " & _
                     "   RATIO, " & _
                     "   RATIO_AMOUNT, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(EXPENSE_TYPE_ID)) & "'," & _
                         "'" & ChangeQuote(Trim(LOCATION_ID)) & "'," & _
                         "'" & ChangeQuote(Trim(RATIO)) & "',"
               SQL1 = SQL1 & _
                         "'" & ChangeQuote(Trim(RATIO_AMOUNT)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         EXPENSE_TYPE_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE EXPTYPE_RATIO SET " & _
                     "   EXPENSE_TYPE_ID = '" & ChangeQuote(Trim(EXPENSE_TYPE_ID)) & "', " & _
                     "   LOCATION_ID = '" & ChangeQuote(Trim(LOCATION_ID)) & "', " & _
                     "   RATIO = '" & ChangeQuote(Trim(RATIO)) & "', " & _
                     "   RATIO_AMOUNT = '" & ChangeQuote(Trim(RATIO_AMOUNT)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE EXPTYPE_RATIO_ID = " & EXPTYPE_RATIO_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM EXPTYPE_RATIO WHERE EXPTYPE_RATIO_ID = " & EXPTYPE_RATIO_ID
      m_Conn.Execute (SQL1)
End Function


