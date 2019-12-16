VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPayment"
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

Public PAYMENT_ID As Long
Public PAYMENT_NO As String
Public PAYMENT_DATE As Date
Public TO_BANK_ID As Long
Public TO_BANK_BRANCH As Long
Public ACCOUNT_NO As String
Public COMMIT_FLAG As String
Public TX_TYPE As String
Public INTERNAL_FLAG As String
Public ACCEPT_BY As Long
Public TOTAL_AMOUNT As Double
Public BANK_ACCOUNT_ID As Long

Public YYYYMM As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public TO_BANK_NAME As String
Public TO_BRANCH_NAME As String
Public ACCOUNT_NAME As String

Public PaymentItems As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set PaymentItems = New Collection
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
Dim TempDate As String
   
   If Ind = 1 Then
      SelectStr1 = "SELECT AC.*, BB.*, BK.*, BA.ACCOUNT_NAME ACCOUNT_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PAYMENT AC "
      FromStr = FromStr & "LEFT OUTER JOIN BANK_BRANCH BB ON (BB.BBRANCH_ID = AC.TO_BANK_BRANCH) "
      FromStr = FromStr & "LEFT OUTER JOIN BANK BK ON (BB.BANK_ID = BK.BANK_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BANK_ACCOUNT BA ON (AC.BANK_ACCOUNT_ID = BA.BANK_ACCOUNT_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT DISTINCT AC.ACCOUNT_NO "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PAYMENT AC "
      
      WhereStr = ""
   End If
   
   If PAYMENT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PAYMENT_ID = " & PAYMENT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PAYMENT_ID = " & PAYMENT_ID & ") "
      End If
   End If
                  
   If TO_BANK_BRANCH > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (TO_BANK_BRANCH = " & TO_BANK_BRANCH & ") "
      Else
         WhereStr = WhereStr & "AND (TO_BANK_BRANCH = " & TO_BANK_BRANCH & ") "
      End If
   End If
               
   If TO_BANK_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BANK_ID = " & TO_BANK_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BANK_ID = " & TO_BANK_ID & ") "
      End If
   End If
   
   If Len(COMMIT_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
   End If
   
   If Len(INTERNAL_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (INTERNAL_FLAG = '" & INTERNAL_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (INTERNAL_FLAG = '" & INTERNAL_FLAG & "') "
      End If
   End If
   
   If Len(TX_TYPE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (TX_TYPE = '" & TX_TYPE & "') "
      Else
         WhereStr = WhereStr & "AND (TX_TYPE = '" & TX_TYPE & "') "
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
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY ACCOUNT_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PAYMENT_DATE "
      Else
         OrderByStr = "ORDER BY PAYMENT_ID "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY ACCOUNT_NO "
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
   If Ind = 1 Then
      PAYMENT_ID = NVLI(Rs("PAYMENT_ID"), -1)
      PAYMENT_NO = NVLS(Rs("PAYMENT_NO"), "")
      PAYMENT_DATE = InternalDateToDate(NVLS(Rs("PAYMENT_DATE"), ""))
      ACCOUNT_NO = NVLS(Rs("ACCOUNT_NO"), "")
      TO_BANK_ID = NVLI(Rs("TO_BANK_ID"), -1)
      TO_BANK_BRANCH = NVLI(Rs("TO_BANK_BRANCH"), -1)
      COMMIT_FLAG = NVLS(Rs("COMMIT_FLAG"), "")
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      INTERNAL_FLAG = NVLS(Rs("INTERNAL_FLAG"), "")
      ACCEPT_BY = NVLI(Rs("ACCEPT_BY"), -1)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      BANK_ACCOUNT_ID = NVLI(Rs("BANK_ACCOUNT_ID"), -1)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      
      TO_BANK_NAME = NVLS(Rs("BANK_NAME"), "")
      TO_BRANCH_NAME = NVLS(Rs("BBRANCH_NAME"), "")
      ACCOUNT_NAME = NVLS(Rs("ACCOUNT_NAME"), "")
   ElseIf Ind = 2 Then
      ACCOUNT_NO = NVLS(Rs("ACCOUNT_NO"), "")
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String

   TempID1 = TO_BANK_ID
   If TO_BANK_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = TO_BANK_BRANCH
   If TO_BANK_BRANCH <= 0 Then
      TempID2 = "NULL"
   End If

   TempID3 = ACCEPT_BY
   If ACCEPT_BY <= 0 Then
      TempID3 = "NULL"
   End If
   
   TempID4 = BANK_ACCOUNT_ID
   If BANK_ACCOUNT_ID <= 0 Then
      TempID4 = "NULL"
   End If
   
   YYYYMM = Mid(DateToStringInt(PAYMENT_DATE), 1, 7)
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("PAYMENT_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO PAYMENT " & _
                     "( " & _
                     "   PAYMENT_ID, " & _
                     "   PAYMENT_NO, " & _
                     "   ACCOUNT_NO, " & _
                     "   PAYMENT_DATE, " & _
                     "   TO_BANK_BRANCH, " & _
                     "   TO_BANK_ID, " & _
                     "   COMMIT_FLAG, " & _
                     "   TX_TYPE, " & _
                     "   INTERNAL_FLAG, " & _
                     "   ACCEPT_BY, " & _
                     "   TOTAL_AMOUNT, " & _
                     "   BANK_ACCOUNT_ID, " & _
                     "   YYYYMM, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(PAYMENT_NO)) & "'," & _
                         " '" & ChangeQuote(Trim(ACCOUNT_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(PAYMENT_DATE))) & "'," & _
                         "" & TempID2 & "," & _
                         "" & TempID1 & "," & _
                         " '" & ChangeQuote(Trim(COMMIT_FLAG)) & "'," & _
                         " '" & ChangeQuote(Trim(TX_TYPE)) & "'," & _
                         " '" & ChangeQuote(Trim(INTERNAL_FLAG)) & "'," & _
                         "" & TempID3 & "," & _
                         "" & TOTAL_AMOUNT & "," & _
                         "" & TempID4 & "," & _
                         " '" & ChangeQuote(Trim(YYYYMM)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         PAYMENT_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE PAYMENT SET " & _
                     "   PAYMENT_NO = '" & ChangeQuote(Trim(PAYMENT_NO)) & "', " & _
                     "   PAYMENT_DATE = '" & ChangeQuote(Trim(DateToStringInt(PAYMENT_DATE))) & "', " & _
                     "   ACCOUNT_NO = '" & ChangeQuote(Trim(ACCOUNT_NO)) & "', " & _
                     "   TO_BANK_BRANCH = " & TempID2 & ", " & _
                     "   TO_BANK_ID = " & TempID1 & ", " & _
                     "   COMMIT_FLAG = '" & ChangeQuote(Trim(COMMIT_FLAG)) & "', " & _
                     "   TX_TYPE = '" & ChangeQuote(Trim(TX_TYPE)) & "', " & _
                     "   INTERNAL_FLAG = '" & ChangeQuote(Trim(INTERNAL_FLAG)) & "', " & _
                     "   ACCEPT_BY = " & TempID3 & ", " & _
                     "   TOTAL_AMOUNT = " & TOTAL_AMOUNT & ", " & _
                     "   BANK_ACCOUNT_ID = " & TempID4 & ", " & _
                     "   YYYYMM = '" & ChangeQuote(Trim(YYYYMM)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PAYMENT_ID = " & PAYMENT_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM PAYMENT_ITEM WHERE PAYMENT_ID = " & PAYMENT_ID
      m_Conn.Execute (SQL1)
            
      SQL1 = "DELETE FROM PAYMENT WHERE PAYMENT_ID = " & PAYMENT_ID
      m_Conn.Execute (SQL1)
End Function

Public Function MigrateBankAccount()
Dim SQL1 As String

      SQL1 = "UPDATE PAYMENT PM SET PM.BANK_ACCOUNT_ID = (SELECT BA.BANK_ACCOUNT_ID FROM BANK_ACCOUNT BA WHERE BA.ACCOUNT_NAME = PM.ACCOUNT_NO) "
      m_Conn.Execute (SQL1)
End Function

Public Function ClearPayment()
Dim SQL1 As String

      SQL1 = "UPDATE PAYMENT SET TOTAL_AMOUNT = 0 WHERE INTERNAL_FLAG = 'Y' "
      m_Conn.Execute (SQL1)

      SQL1 = "UPDATE PAYMENT_ITEM SET PAY_AMOUNT = 0 WHERE PAYMENT_ID IN (SELECT PAYMENT_ID FROM PAYMENT WHERE INTERNAL_FLAG = 'Y') "
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set PaymentItems = Nothing
End Sub
