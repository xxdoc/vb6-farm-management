VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExpenseRatio"
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

Public EXPENSE_RATIO_ID As Long
Public RO_ITEM_ID As Long
Public LOCATION_ID As Long
Public SELECT_FLAG As String
Public RATIO As Double
Public RATIO_AMOUNT As Double
Public PIG_COUNT As Double
Public BATCH_ID As Long

Public LOCATION_NAME As String
Public LOCATION_NO As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public EXPENSE_TYPE As Long
Public DEPLICATE_FLAG As String
Public SELL_PARENT_FLAG  As String

Public KEY_ID As Long
Public KEY_LOOKUP As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_DATE = -1
   TO_DATE = -1
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
Dim GroupByStr As String
Dim TempDate As String
   
   If Ind = 1 Then
      SelectStr1 = "SELECT ER.*, LC.*,ET.SELL_PARENT_FLAG  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM EXPENSE_RATIO ER "
      
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (ER.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN RO_ITEM RI ON (ER.RO_ITEM_ID = RI.RO_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EXPENSE_TYPE ET ON (RI.EXPENSE_TYPE = ET.EXPENSE_TYPE_ID) "
      
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT ER.LOCATION_ID, SUM(ER.RATIO_AMOUNT) RATIO_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM EXPENSE_RATIO ER "
      FromStr = FromStr & "LEFT OUTER JOIN RO_ITEM RI ON (ER.RO_ITEM_ID = RI.RO_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (RI.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EXPENSE_TYPE ET ON (RI.EXPENSE_TYPE = ET.EXPENSE_TYPE_ID) "
      
      WhereStr = ""
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT ER.LOCATION_ID, RI.EXPENSE_TYPE, SUM(ER.RATIO_AMOUNT) RATIO_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM EXPENSE_RATIO ER "
      FromStr = FromStr & "LEFT OUTER JOIN RO_ITEM RI ON (ER.RO_ITEM_ID = RI.RO_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (RI.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      
      WhereStr = ""
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT DISTINCT(ER.LOCATION_ID) "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM EXPENSE_RATIO ER "
      FromStr = FromStr & "LEFT OUTER JOIN RO_ITEM RI ON (ER.RO_ITEM_ID = RI.RO_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (RI.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (ER.LOCATION_ID = LC.LOCATION_ID) "
      
      WhereStr = ""
   ElseIf Ind = 5 Then
      SelectStr1 = "SELECT RI.EXPENSE_TYPE, SUM(ER.RATIO_AMOUNT) RATIO_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM EXPENSE_RATIO ER "
      FromStr = FromStr & "LEFT OUTER JOIN RO_ITEM RI ON (ER.RO_ITEM_ID = RI.RO_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (RI.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      
      WhereStr = ""
   End If
   
   If EXPENSE_RATIO_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ER.EXPENSE_RATIO_ID = " & EXPENSE_RATIO_ID & ") "
      Else
         WhereStr = WhereStr & "AND (ER.EXPENSE_RATIO_ID = " & EXPENSE_RATIO_ID & ") "
      End If
   End If
   
   If RO_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ER.RO_ITEM_ID = " & RO_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (ER.RO_ITEM_ID = " & RO_ITEM_ID & ") "
      End If
   End If
   
   If LOCATION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LC.LOCATION_ID = " & LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LC.LOCATION_ID = " & LOCATION_ID & ") "
      End If
   End If
   
    If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If Len(DEPLICATE_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DEPLICATE_FLAG = '" & DEPLICATE_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (DEPLICATE_FLAG = '" & DEPLICATE_FLAG & "') "
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
      GroupByStr = "GROUP BY ER.LOCATION_ID "
   ElseIf Ind = 3 Then
      GroupByStr = "GROUP BY ER.LOCATION_ID, RI.EXPENSE_TYPE "
   ElseIf Ind = 5 Then
      GroupByStr = "GROUP BY RI.EXPENSE_TYPE "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY LOCATION_NAME "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY EXPENSE_TYPE_NAME "
      Else
         OrderByStr = "ORDER BY LOCATION_NAME "
      End If
   ElseIf Ind = 4 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY LOCATION_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY LOCATION_NAME "
      Else
         OrderByStr = "ORDER BY LOCATION_NO "
      End If
   End If
    
    If Len(OrderByStr) > 0 Then
        If OrderType = 1 Then
           OrderTypeStr = "ASC "
        ElseIf OrderType = 2 Then
           OrderTypeStr = "DESC "
        Else
           OrderTypeStr = "ASC "
        End If
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      EXPENSE_RATIO_ID = NVLI(Rs("EXPENSE_RATIO_ID"), -1)
      RO_ITEM_ID = NVLI(Rs("RO_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      SELECT_FLAG = NVLS(Rs("SELECT_FLAG"), "")
      RATIO = NVLD(Rs("RATIO"), 0)
      RATIO_AMOUNT = NVLD(Rs("RATIO_AMOUNT"), 0)
      BATCH_ID = NVLI(Rs("BATCH_ID"), -1)
      
   '''debug.print SELECT_FLAG
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
      SELL_PARENT_FLAG = NVLS(Rs("SELL_PARENT_FLAG"), "")
      
   ElseIf Ind = 2 Then
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      RATIO_AMOUNT = NVLD(Rs("RATIO_AMOUNT"), 0)
   ElseIf Ind = 3 Then
      EXPENSE_TYPE = NVLI(Rs("EXPENSE_TYPE"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      RATIO_AMOUNT = NVLD(Rs("RATIO_AMOUNT"), 0)
   ElseIf Ind = 4 Then
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
   ElseIf Ind = 5 Then
      EXPENSE_TYPE = NVLI(Rs("EXPENSE_TYPE"), -1)
      RATIO_AMOUNT = NVLD(Rs("RATIO_AMOUNT"), 0)
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("EXPENSE_RATIO_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO EXPENSE_RATIO " & _
                     "( " & _
                     "   EXPENSE_RATIO_ID, " & _
                     "   RO_ITEM_ID, " & _
                     "   LOCATION_ID, " & _
                     "   SELECT_FLAG, " & _
                     "   RATIO, " & _
                     "   RATIO_AMOUNT, " & _
                     "   BATCH_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(RO_ITEM_ID)) & "'," & _
                         "'" & ChangeQuote(Trim(LOCATION_ID)) & "'," & _
                         "'" & ChangeQuote(Trim(SELECT_FLAG)) & "'," & _
                         "'" & ChangeQuote(Trim(RATIO)) & "',"
               SQL1 = SQL1 & _
                         "'" & ChangeQuote(Trim(RATIO_AMOUNT)) & "'," & _
                         BATCH_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         RO_ITEM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE EXPENSE_RATIO SET " & _
                     "   RO_ITEM_ID = '" & ChangeQuote(Trim(RO_ITEM_ID)) & "', " & _
                     "   LOCATION_ID = '" & ChangeQuote(Trim(LOCATION_ID)) & "', " & _
                     "   SELECT_FLAG = '" & ChangeQuote(Trim(SELECT_FLAG)) & "', " & _
                     "   RATIO = '" & ChangeQuote(Trim(RATIO)) & "', " & _
                     "   RATIO_AMOUNT = '" & ChangeQuote(Trim(RATIO_AMOUNT)) & "', " & _
                     "   BATCH_ID = " & BATCH_ID & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE EXPENSE_RATIO_ID = " & EXPENSE_RATIO_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM EXPENSE_RATIO WHERE EXPENSE_RATIO_ID = " & EXPENSE_RATIO_ID
      m_Conn.Execute (SQL1)
End Function


