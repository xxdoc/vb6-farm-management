VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPaymentItem"
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

Public PAYMENT_ITEM_ID As Long
Public PAYMENT_ID As Long
Public PAYMENT_TYPE As PAYMENT_TYPE
Public BANK_ID As Long
Public BANK_BRANCH As Long
Public CHECK_NO As String
Public BANK_NOTE As String
Public PAY_AMOUNT As Double

Public OLD_PAYMENT_TYPE As PAYMENT_TYPE
Public BANK_NAME As String
Public BANK_BRANCH_NAME As String
Public TX_TYPE As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public DOCUMENT_TYPE As Long
Public RECEIPT_TYPE As Long

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
Dim SubLen As Long
Dim NewStr As String
Dim GroupByStr As String
Dim TempDate As String

   If Ind = 1 Then
      SelectStr1 = "SELECT AC.*, BB.*, BK.* "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PAYMENT_ITEM AC "
      FromStr = FromStr & "LEFT OUTER JOIN BANK_BRANCH BB ON (BB.BBRANCH_ID = AC.BANK_BRANCH) "
      FromStr = FromStr & "LEFT OUTER JOIN BANK BK ON (BB.BANK_ID = BK.BANK_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT PM.TX_TYPE, AC.PAYMENT_TYPE, SUM(AC.PAY_AMOUNT) PAY_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM PAYMENT_ITEM AC "
      FromStr = FromStr & "LEFT OUTER JOIN PAYMENT PM ON (AC.PAYMENT_ID = PM.PAYMENT_ID) "
      
      WhereStr = ""
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT BD.DOCUMENT_TYPE, BD.RECEIPT_TYPE, PM.TX_TYPE, AC.PAYMENT_TYPE, SUM(AC.PAY_AMOUNT) PAY_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM PAYMENT_ITEM AC "
      FromStr = FromStr & "LEFT OUTER JOIN PAYMENT PM ON (AC.PAYMENT_ID = PM.PAYMENT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (AC.PAYMENT_ID = BD.PAYMENT_ID) "
      
      WhereStr = ""
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT AC.PAYMENT_TYPE,PM.TX_TYPE, SUM(AC.PAY_AMOUNT) PAY_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PAYMENT_ITEM AC "
      FromStr = FromStr & "LEFT OUTER JOIN PAYMENT PM ON (AC.PAYMENT_ID = PM.PAYMENT_ID) "
      
   End If
      
   If PAYMENT_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PAYMENT_ITEM_ID = " & PAYMENT_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PAYMENT_ITEM_ID = " & PAYMENT_ITEM_ID & ") "
      End If
   End If
               
   If PAYMENT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PAYMENT_ID = " & PAYMENT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PAYMENT_ID = " & PAYMENT_ID & ") "
      End If
   End If
   
   If PAYMENT_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PAYMENT_TYPE = " & PAYMENT_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (PAYMENT_TYPE = " & PAYMENT_TYPE & ") "
      End If
   End If
   
    If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (PAYMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (PAYMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (PAYMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (PAYMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   GroupByStr = ""
   
   If Ind = 2 Then
      GroupByStr = "GROUP BY PM.TX_TYPE, AC.PAYMENT_TYPE "
   ElseIf Ind = 3 Then
      GroupByStr = "GROUP BY BD.DOCUMENT_TYPE, BD.RECEIPT_TYPE, PM.TX_TYPE, AC.PAYMENT_TYPE "
   ElseIf Ind = 4 Then
      GroupByStr = "GROUP BY AC.PAYMENT_TYPE,PM.TX_TYPE "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PAYMENT_ITEM_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PAYMENT_ITEM_ID "
      Else
         OrderByStr = "ORDER BY PAYMENT_ITEM_ID "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY PAYMENT_TYPE "
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY PAYMENT_TYPE "
   End If
   
   If Ind <> 4 Then
      If OrderType = 1 Then
         OrderTypeStr = "ASC "
      ElseIf OrderType = 2 Then
         OrderTypeStr = "DESC "
      Else
         OrderTypeStr = "DESC "
      End If
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      PAYMENT_ITEM_ID = NVLI(Rs("PAYMENT_ITEM_ID"), -1)
      PAYMENT_ID = NVLS(Rs("PAYMENT_ID"), "")
      PAYMENT_TYPE = NVLI(Rs("PAYMENT_TYPE"), -1)
      CHECK_NO = NVLS(Rs("CHECK_NO"), "")
      BANK_ID = NVLI(Rs("BANK_ID"), -1)
      BANK_BRANCH = NVLI(Rs("BANK_BRANCH"), -1)
      BANK_NOTE = NVLS(Rs("BANK_NOTE"), "")
      PAY_AMOUNT = NVLD(Rs("PAY_AMOUNT"), 0)
      
      OLD_PAYMENT_TYPE = PAYMENT_TYPE
      BANK_NAME = NVLS(Rs("BANK_NAME"), "")
      BANK_BRANCH_NAME = NVLS(Rs("BBRANCH_NAME"), "")
   ElseIf Ind = 2 Then
      PAYMENT_TYPE = NVLI(Rs("PAYMENT_TYPE"), -1)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      PAY_AMOUNT = NVLD(Rs("PAY_AMOUNT"), 0)
   ElseIf Ind = 3 Then
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      RECEIPT_TYPE = NVLI(Rs("RECEIPT_TYPE"), -1)
      PAYMENT_TYPE = NVLI(Rs("PAYMENT_TYPE"), -1)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      PAY_AMOUNT = NVLD(Rs("PAY_AMOUNT"), 0)
   ElseIf Ind = 4 Then
      PAYMENT_TYPE = NVLI(Rs("PAYMENT_TYPE"), -1)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      PAY_AMOUNT = NVLD(Rs("PAY_AMOUNT"), 0)
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String

'   TempID1 = BANK_ID
'   If BANK_ID <= 0 Then
'      TempID1 = "NULL"
'   End If
'
'   TempID2 = BANK_BRANCH
'   If BANK_BRANCH <= 0 Then
'      TempID2 = "NULL"
'   End If
'
'   TempID3 = PAYMENT_TYPE
'   If PAYMENT_TYPE <= 0 Then
'      TempID3 = "NULL"
'   End If
'
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   If AddEditMode = SHOW_ADD Then
'      Call glbDatabaseMngr.GetSeqID("PAYMENT_ITEM_SEQ", ID, glbErrorLog)
'
'      SQL1 = "INSERT INTO PAYMENT_ITEM " & _
'                     "( " & _
'                     "   PAYMENT_ITEM_ID, " & _
'                     "   PAYMENT_ID, " & _
'                     "   CHECK_NO, " & _
'                     "   PAYMENT_TYPE, " & _
'                     "   BANK_BRANCH, " & _
'                     "   BANK_ID, " & _
'                     "   BANK_NOTE, " & _
'                     "   PAY_AMOUNT, " & _
'                     "   CREATE_DATE, " & _
'                    "    CREATE_BY, " & _
'                    "    MODIFY_DATE, " & _
'                    "    MODIFY_BY " & _
'                     ") " & _
'                     "VALUES "
'         SQL1 = SQL1 & "( " & _
'                         ID & "," & _
'                         "" & PAYMENT_ID & "," & _
'                         "'" & ChangeQuote(Trim(CHECK_NO)) & "'," & _
'                         "" & TempID3 & "," & _
'                         "" & TempID2 & "," & _
'                         "" & TempID1 & "," & _
'                         "'" & ChangeQuote(Trim(BANK_NOTE)) & "'," & _
'                         "" & PAY_AMOUNT & "," & _
'                         "'" & ServerDateTime & "'," & _
'                         glbUser.USER_ID & "," & _
'                         "'" & ServerDateTime & "'," & _
'                         glbUser.USER_ID & _
'                     ") "
'         PAYMENT_ITEM_ID = ID
'   Else 'Show edit
'      SQL1 = "UPDATE PAYMENT_ITEM SET " & _
'                     "   PAYMENT_ID = '" & ChangeQuote(Trim(PAYMENT_ID)) & "', " & _
'                     "   PAYMENT_TYPE = " & PAYMENT_TYPE & ", " & _
'                     "   CHECK_NO = '" & ChangeQuote(Trim(CHECK_NO)) & "', " & _
'                     "   BANK_BRANCH = " & TempID2 & ", " & _
'                     "   BANK_ID = " & TempID1 & ", " & _
'                     "   BANK_NOTE = '" & ChangeQuote(Trim(BANK_NOTE)) & "', " & _
'                     "   PAY_AMOUNT = " & PAY_AMOUNT & ", " & _
'                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
'                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
'                     "WHERE PAYMENT_ITEM_ID = " & PAYMENT_ITEM_ID
'   End If
'
'   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
'      m_Conn.Execute (SQL1)
'   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
            
      SQL1 = "DELETE FROM PAYMENT_ITEM WHERE PAYMENT_ITEM_ID = " & PAYMENT_ITEM_ID
      m_Conn.Execute (SQL1)
End Function


