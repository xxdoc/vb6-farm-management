VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReceiptItem"
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

Public RECEIPT_ITEM_ID As Long
Public BILLING_DOC_ID As Long
Public DO_ID As Long
Public RECEIPT_ITEM_AMOUNT As Double
Public PAID_AMOUNT As Double
Public DEBIT_CREDIT_AMOUNT As Double
Public STATUS_ID As Long
Public REASON_ID As Long
Public DESCRIPTION_DETAIL As String

Public PRODUCT_STATUS_NAME As String
Public REASON_NAME As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public DOCUMENT_DATE As Date
Public DOCUMENT_NO As String
Public DOCUMENT_TYPE As Long
Public DOCUMENT_SUBTYPE As Long
Public RECEIPT_TYPE As Long
Public ACCOUNT_ID As Long
Public CUSTOMER_ID As Long
Public FROM_ITEM_DATE As Date
Public TO_ITEM_DATE As Date

Public FROM_DOC_DATE As Date
Public TO_DOC_DATE As Date

Public CNDN_ID As Long
Public CNDN_NO As String
Public CNDN_DATE As Date
Public DO_NO As String
Public DO_DATE As Date
Public CNDN_TYPE As Long
Public PART_NO As String

Public RECEIPT_NO As String
Public RECEIPT_DATE As Date
Public BILL_NO As String
Public BILL_DATE As Date
Public DISCOUNT_AMOUNT As Double
Public DEPOSIT_AMOUNT As Double
Public CASH_DISCOUNT As Double
Public CUSTOMER_CODE As String
Public CUSTOMER_NAME As String
Public SALE_NAME As String
Public SALE_LNAME As String

Public FROM_ITEM_DUE_DATE As Date
Public TO_ITEM_DUE_DATE As Date
Public MAX_DOC_DATE As Date
Public MIN_DOC_DATE As Date

Public DocTypeSet As String

Public VALID_DATE As Date

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_ITEM_DATE = -1
   TO_ITEM_DATE = -1
   FROM_DATE = -1
   TO_DATE = -1
   FROM_DOC_DATE = -1
   TO_DOC_DATE = -1
   VALID_DATE = -1
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
Dim TempDate As String
Dim SubLen As Long
Dim NewStr As String
Dim GroupByStr As String
    
   If Ind = 1 Then
      SelectStr1 = "SELECT UG.*, DI.DOCUMENT_DATE, DI.DOCUMENT_NO, DI.COMMIT_FLAG, DI.DOCUMENT_TYPE, PS.PRODUCT_STATUS_NAME, RS.REASON_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC DI ON (UG.DO_ID = DI.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (UG.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (UG.STATUS_ID = PS.PRODUCT_STATUS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CNDN_REASON RS ON (UG.REASON_ID = RS.REASON_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT BD.DOCUMENT_TYPE, BD.DOCUMENT_SUBTYPE, BD.RECEIPT_TYPE, SUM(UG.PAID_AMOUNT) PAID_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (UG.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      
      WhereStr = ""
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT BD.ACCOUNT_ID, BD.DOCUMENT_TYPE, BD.DOCUMENT_SUBTYPE, BD.RECEIPT_TYPE, SUM(UG.PAID_AMOUNT) PAID_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (UG.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      
      WhereStr = ""
   ElseIf Ind = 4 Then              'LoadPaidAmountByBill
      SelectStr1 = "SELECT UG.DO_ID, SUM(UG.PAID_AMOUNT) PAID_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (UG.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      
      WhereStr = ""
   ElseIf Ind = 5 Then
      SelectStr1 = "SELECT CM2.CUSTOMER_ID, SUM(UG.PAID_AMOUNT) PAID_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (UG.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (BD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC.CUSTOMER_ID = CM2.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM2.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID) "
      
      WhereStr = ""
   ElseIf Ind = 6 Then
      SelectStr1 = "SELECT DI.ACCOUNT_ID, SUM(UG.DEBIT_CREDIT_AMOUNT) DEBIT_CREDIT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (UG.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC DI ON (UG.DO_ID = DI.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC2 ON (DI.ACCOUNT_ID = AC2.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC2.CUSTOMER_ID = CM2.CUSTOMER_ID) "

      WhereStr = ""
   ElseIf Ind = 7 Then      'LoadDnCnAmountByBill
      SelectStr1 = "SELECT UG.DO_ID, SUM(UG.DEBIT_CREDIT_AMOUNT) DEBIT_CREDIT_AMOUNT "
      'SelectStr1 = "SELECT UG.DO_ID,BD.DOCUMENT_NO, SUM(UG.DEBIT_CREDIT_AMOUNT) DEBIT_CREDIT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (UG.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC DI ON (UG.DO_ID = DI.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC2 ON (DI.ACCOUNT_ID = AC2.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC2.CUSTOMER_ID = CM2.CUSTOMER_ID) "

      WhereStr = ""
   ElseIf Ind = 8 Then
      SelectStr1 = "SELECT BD.DOCUMENT_TYPE, SUM(UG.DEBIT_CREDIT_AMOUNT) DEBIT_CREDIT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (UG.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC DI ON (UG.DO_ID = DI.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC2 ON (DI.ACCOUNT_ID = AC2.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC2.CUSTOMER_ID = CM2.CUSTOMER_ID) "

      WhereStr = ""
   ElseIf Ind = 9 Then
      SelectStr1 = "SELECT AC2.CUSTOMER_ID, SUM(UG.DEBIT_CREDIT_AMOUNT) DEBIT_CREDIT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (UG.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC DI ON (UG.DO_ID = DI.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC2 ON (DI.ACCOUNT_ID = AC2.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC2.CUSTOMER_ID = CM2.CUSTOMER_ID) "

      WhereStr = ""
   ElseIf Ind = 10 Then
      SelectStr1 = "SELECT UG.DEBIT_CREDIT_AMOUNT , BD.DOCUMENT_TYPE CNDN_TYPE, BD.BILLING_DOC_ID CNDN_ID, BD.DOCUMENT_NO CNDN_NO, BD.DOCUMENT_DATE CNDN_DATE, DI.BILLING_DOC_ID DO_ID, DI.DOCUMENT_NO DO_NO, DI.DOCUMENT_DATE DO_DATE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (UG.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC DI ON (UG.DO_ID = DI.BILLING_DOC_ID) "

      WhereStr = "WHERE (BD.DOCUMENT_TYPE IN (3, 4)) "
   ElseIf Ind = 11 Then
      SelectStr1 = "SELECT BD.DOCUMENT_DATE, CM2.CUSTOMER_ID, SUM(UG.RECEIPT_ITEM_AMOUNT) RECEIPT_ITEM_AMOUNT, SUM(UG.PAID_AMOUNT) PAID_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (UG.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC2 ON (BD.ACCOUNT_ID = AC2.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC2.CUSTOMER_ID = CM2.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM2.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = CN.NAME_ID) "

      WhereStr = ""
   ElseIf Ind = 12 Then
      SelectStr1 = "SELECT BD.BILLING_DOC_ID, SUM(UG.RECEIPT_ITEM_AMOUNT) RECEIPT_ITEM_AMOUNT, SUM(UG.PAID_AMOUNT) PAID_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (UG.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC2 ON (BD.ACCOUNT_ID = AC2.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC2.CUSTOMER_ID = CM2.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM2.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = CN.NAME_ID) "
      
      WhereStr = ""
   ElseIf Ind = 13 Then
      SelectStr1 = "SELECT BD.DOCUMENT_NO RECEIPT_NO, BD.DOCUMENT_DATE RECEIPT_DATE, BD.DOCUMENT_TYPE, DI.DOCUMENT_NO BILL_NO, DI.DOCUMENT_DATE BILL_DATE, UG.*, CM2.CUSTOMER_CODE, NM.LONG_NAME CUSTOMER_NAME, UG.DEBIT_CREDIT_AMOUNT, ENM.LONG_NAME SALE_NAME, ENM.LAST_NAME SALE_LNAME "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (UG.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC DI ON (UG.DO_ID = DI.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC2 ON (DI.ACCOUNT_ID = AC2.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC2.CUSTOMER_ID = CM2.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM2.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = CN.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE ED ON (CM2.RESPONSE_BY = ED.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (ED.EMP_ID = EN.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME ENM ON (EN.NAME_ID = ENM.NAME_ID) "
      
      WhereStr = ""
   ElseIf Ind = 14 Then
      SelectStr1 = "SELECT UG.DO_ID, UG.BILLING_DOC_ID, UG.PAID_AMOUNT, UG.RECEIPT_ITEM_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC DI ON (UG.DO_ID = DI.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (UG.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC2 ON (BD.ACCOUNT_ID = AC2.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC2.CUSTOMER_ID = CM2.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM2.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = CN.NAME_ID) "

      WhereStr = ""
   ElseIf Ind = 15 Then
      SelectStr1 = "SELECT MAX(BD.DOCUMENT_DATE) MAX_DOC_DATE, MIN(BD.DOCUMENT_DATE) MIN_DOC_DATE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC DI ON (UG.DO_ID = DI.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (UG.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC2 ON (BD.ACCOUNT_ID = AC2.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC2.CUSTOMER_ID = CM2.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM2.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = CN.NAME_ID) "

      WhereStr = ""
   ElseIf Ind = 16 Then
      SelectStr1 = "SELECT UG.RECEIPT_ITEM_ID,UG.DO_ID "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (UG.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      
      WhereStr = " WHERE (BD.DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "') "
   ElseIf Ind = 17 Then
      SelectStr1 = "SELECT UG.DO_ID, BD.DOCUMENT_DATE, BD.DOCUMENT_NO, UG.PAID_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*)  "
      FromStr = "FROM RECEIPT_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (UG.BILLING_DOC_ID = BD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC2 ON (BD.ACCOUNT_ID = AC2.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM2 ON (AC2.CUSTOMER_ID = CM2.CUSTOMER_ID) "
   End If

   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BD.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (BD.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      End If
   End If
   
   If RECEIPT_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (RECEIPT_ITEM_ID = " & RECEIPT_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (RECEIPT_ITEM_ID = " & RECEIPT_ITEM_ID & ") "
      End If
   End If
   
   If DO_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DO_ID = " & DO_ID & ") "
      Else
         WhereStr = WhereStr & "AND (DO_ID = " & DO_ID & ") "
      End If
   End If
   
   If BILLING_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BD.BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BD.BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      End If
   End If
   
    If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BD.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BD.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BD.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BD.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

    If FROM_ITEM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_ITEM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DI.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DI.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_ITEM_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_ITEM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DI.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DI.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
    If FROM_DOC_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DOC_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BD.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BD.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DOC_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DOC_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BD.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BD.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If FROM_ITEM_DUE_DATE > 0 Then
      TempDate = DateToStringIntLow(Trim(FROM_ITEM_DUE_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DI.DUE_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DI.DUE_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_ITEM_DUE_DATE > 0 Then
      TempDate = DateToStringIntHi(Trim(TO_ITEM_DUE_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DI.DUE_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DI.DUE_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If VALID_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(VALID_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE ((UG.VALID_DATE >= '" & ChangeQuote(TempDate) & "') OR (UG.VALID_DATE IS NULL )) "
      Else
         WhereStr = WhereStr & "AND ((UG.VALID_DATE >= '" & ChangeQuote(TempDate) & "') OR (UG.VALID_DATE IS NULL)) "
      End If
   End If
   
   If Len(DocTypeSet) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BD.DOCUMENT_TYPE IN " & DocTypeSet & ") "
      Else
         WhereStr = WhereStr & "AND (BD.DOCUMENT_TYPE IN " & DocTypeSet & ") "
      End If
   End If
   
   If Len(Trim(CUSTOMER_CODE)) > 0 Then
      If WildCard(CUSTOMER_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(CUSTOMER_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(CUSTOMER_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CUSTOMER_CODE = '" & ChangeQuote(Trim(CUSTOMER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (CUSTOMER_CODE = '" & ChangeQuote(Trim(CUSTOMER_CODE)) & "')"
         End If
      End If
   End If

   GroupByStr = ""
   If Ind = 2 Then
      GroupByStr = "GROUP BY BD.DOCUMENT_TYPE, BD.DOCUMENT_SUBTYPE, BD.RECEIPT_TYPE "
   ElseIf Ind = 3 Then
      GroupByStr = "GROUP BY BD.ACCOUNT_ID, BD.DOCUMENT_TYPE, BD.DOCUMENT_SUBTYPE, BD.RECEIPT_TYPE "
   ElseIf Ind = 4 Then
      GroupByStr = "GROUP BY UG.DO_ID "
   ElseIf Ind = 5 Then
      GroupByStr = "GROUP BY CM2.CUSTOMER_ID "
   ElseIf Ind = 6 Then
      GroupByStr = "GROUP BY DI.ACCOUNT_ID "
   ElseIf Ind = 7 Then
      GroupByStr = "GROUP BY UG.DO_ID "
      'GroupByStr = "GROUP BY UG.DO_ID,BD.DOCUMENT_NO "
   ElseIf Ind = 8 Then
      GroupByStr = "GROUP BY BD.DOCUMENT_TYPE "
   ElseIf Ind = 9 Then
      GroupByStr = "GROUP BY AC2.CUSTOMER_ID "
   ElseIf Ind = 11 Then
      GroupByStr = "GROUP BY BD.DOCUMENT_DATE, CM2.CUSTOMER_ID "
   ElseIf Ind = 12 Then
      GroupByStr = "GROUP BY BD.BILLING_DOC_ID "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY RECEIPT_ITEM_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY RECEIPT_ITEM_ID "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY DOCUMENT_DATE DESC, TRANSACTION_SEQ "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY LOCATION_ID DESC, DOCUMENT_DATE DESC, TRANSACTION_SEQ "
      ElseIf OrderBy = 5 Then
         OrderByStr = "ORDER BY PT.PART_TYPE_NO "
      ElseIf OrderBy = 6 Then
         OrderByStr = "ORDER BY PI.PART_NO "
      ElseIf OrderBy = 7 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE "
      ElseIf OrderBy = 8 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_NO "
      ElseIf OrderBy = 9 Then
         OrderByStr = "ORDER BY LC.LOCATION_NO "
      Else
         OrderByStr = "ORDER BY RECEIPT_ITEM_ID "
      End If
   ElseIf Ind = 2 Then
      'OrderByStr = "ORDER BY RECEIPT_ITEM_ID "
   ElseIf Ind = 3 Then
      'OrderByStr = "ORDER BY RECEIPT_ITEM_ID "
   ElseIf Ind = 4 Then
      'OrderByStr = "ORDER BY UG.DO_ID "
   ElseIf Ind = 5 Then
      'OrderByStr = "ORDER BY CUSTOMER_ID "
   ElseIf Ind = 6 Then
      OrderByStr = "ORDER BY DI.ACCOUNT_ID "
   ElseIf Ind = 7 Then
      'OrderByStr = "ORDER BY UG.DO_ID "
   ElseIf Ind = 8 Then
      OrderByStr = "ORDER BY BD.DOCUMENT_TYPE "
   ElseIf Ind = 9 Then
      OrderByStr = "ORDER BY AC2.CUSTOMER_ID "
   ElseIf Ind = 10 Then
      'OrderByStr = "ORDER BY UG.DO_ID "
   ElseIf Ind = 11 Then
      OrderByStr = "ORDER BY BD.DOCUMENT_DATE "
   ElseIf Ind = 12 Then
      OrderByStr = "ORDER BY BD.BILLING_DOC_ID "
   ElseIf Ind = 13 Then
      OrderByStr = "ORDER BY CM2.CUSTOMER_CODE ASC, BD.DOCUMENT_DATE ASC, BD.DOCUMENT_NO, DI.DOCUMENT_DATE ASC, DI.DOCUMENT_NO "
   ElseIf Ind = 14 Then
      OrderByStr = "ORDER BY UG.DO_ID "         'อันนี้จำเป็นต้องเรียง
   ElseIf Ind = 15 Then
      'OrderByStr = "ORDER BY UG.RECEIPT_ITEM_ID "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   If OrderByStr = "" Then
      OrderTypeStr = ""
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      RECEIPT_ITEM_ID = NVLI(Rs("RECEIPT_ITEM_ID"), -1)
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      DO_ID = NVLI(Rs("DO_ID"), -1)
      RECEIPT_ITEM_AMOUNT = NVLD(Rs("RECEIPT_ITEM_AMOUNT"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      DEBIT_CREDIT_AMOUNT = NVLD(Rs("DEBIT_CREDIT_AMOUNT"), 0)
      STATUS_ID = NVLI(Rs("STATUS_ID"), -1)
      REASON_ID = NVLI(Rs("REASON_ID"), -1)
      
      REASON_NAME = NVLS(Rs("REASON_NAME"), "")
      PRODUCT_STATUS_NAME = NVLS(Rs("PRODUCT_STATUS_NAME"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DESCRIPTION_DETAIL = NVLS(Rs("DESCRIPTION_DETAIL"), "")
'      PART_NO = NVLS(Rs("PART_NO"), "")
   ElseIf Ind = 2 Then
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      DOCUMENT_SUBTYPE = NVLI(Rs("DOCUMENT_SUBTYPE"), 0)
      RECEIPT_TYPE = NVLI(Rs("RECEIPT_TYPE"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
   ElseIf Ind = 3 Then
      ACCOUNT_ID = NVLI(Rs("ACCOUNT_ID"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      DOCUMENT_SUBTYPE = NVLI(Rs("DOCUMENT_SUBTYPE"), 0)
      RECEIPT_TYPE = NVLI(Rs("RECEIPT_TYPE"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
   ElseIf Ind = 4 Then
      DO_ID = NVLI(Rs("DO_ID"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
   ElseIf Ind = 5 Then
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
   ElseIf Ind = 6 Then
      ACCOUNT_ID = NVLI(Rs("ACCOUNT_ID"), 0)
      DEBIT_CREDIT_AMOUNT = NVLD(Rs("DEBIT_CREDIT_AMOUNT"), 0)
   ElseIf Ind = 7 Then
      DO_ID = NVLI(Rs("DO_ID"), 0)
                                                                                                                  'DOCUMENT_NO = NVLI(Rs("DOCUMENT_NO"), 0)
      DEBIT_CREDIT_AMOUNT = NVLD(Rs("DEBIT_CREDIT_AMOUNT"), 0)
   ElseIf Ind = 8 Then
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      DEBIT_CREDIT_AMOUNT = NVLD(Rs("DEBIT_CREDIT_AMOUNT"), 0)
   ElseIf Ind = 9 Then
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), 0)
      DEBIT_CREDIT_AMOUNT = NVLD(Rs("DEBIT_CREDIT_AMOUNT"), 0)
   ElseIf Ind = 10 Then
      CNDN_ID = NVLI(Rs("CNDN_ID"), 0)
      CNDN_NO = NVLS(Rs("CNDN_NO"), "")
      CNDN_DATE = InternalDateToDate(NVLS(Rs("CNDN_DATE"), ""))
      DO_ID = NVLI(Rs("DO_ID"), 0)
      DO_NO = NVLS(Rs("DO_NO"), "")
      DO_DATE = InternalDateToDate(NVLS(Rs("DO_DATE"), ""))
      DEBIT_CREDIT_AMOUNT = NVLD(Rs("DEBIT_CREDIT_AMOUNT"), 0)
      CNDN_TYPE = NVLI(Rs("CNDN_TYPE"), 0)
   ElseIf Ind = 11 Then
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      RECEIPT_ITEM_AMOUNT = NVLD(Rs("RECEIPT_ITEM_AMOUNT"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
   ElseIf Ind = 12 Then
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      RECEIPT_ITEM_AMOUNT = NVLD(Rs("RECEIPT_ITEM_AMOUNT"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
   ElseIf Ind = 13 Then
      RECEIPT_NO = NVLS(Rs("RECEIPT_NO"), "")
      RECEIPT_DATE = InternalDateToDate(NVLS(Rs("RECEIPT_DATE"), ""))
      BILL_NO = NVLS(Rs("BILL_NO"), "")
      BILL_DATE = InternalDateToDate(NVLS(Rs("BILL_DATE"), ""))
      
      RECEIPT_ITEM_AMOUNT = NVLD(Rs("RECEIPT_ITEM_AMOUNT"), 0)
      'DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      'DEPOSIT_AMOUNT = NVLD(Rs("DEPOSIT_AMOUNT"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      'CASH_DISCOUNT = NVLD(Rs("CASH_DISCOUNT"), 0)
      DEBIT_CREDIT_AMOUNT = NVLD(Rs("DEBIT_CREDIT_AMOUNT"), 0)
      
      CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      SALE_NAME = NVLS(Rs("SALE_NAME"), "")
      SALE_LNAME = NVLS(Rs("SALE_LNAME"), "")
      SALE_NAME = SALE_NAME & " " & SALE_LNAME
   ElseIf Ind = 14 Then
      DO_ID = NVLI(Rs("DO_ID"), -1)
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      RECEIPT_ITEM_AMOUNT = NVLD(Rs("RECEIPT_ITEM_AMOUNT"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
   ElseIf Ind = 15 Then
      MAX_DOC_DATE = InternalDateToDate(NVLS(Rs("MAX_DOC_DATE"), ""))
      MIN_DOC_DATE = InternalDateToDate(NVLS(Rs("MIN_DOC_DATE"), ""))
   ElseIf Ind = 16 Then
      RECEIPT_ITEM_ID = NVLI(Rs("RECEIPT_ITEM_ID"), -1)
      DO_ID = NVLI(Rs("DO_ID"), -1)
   ElseIf Ind = 17 Then
      DO_ID = NVLI(Rs("DO_ID"), -1)
      
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim ID2 As Long
Dim Bd As CBillingDoc
Dim TempID1 As String
Dim TempID2 As String

   TempID1 = STATUS_ID
   If STATUS_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = REASON_ID
   If REASON_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("RECEIPT_ITEM_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO RECEIPT_ITEM " & _
                     "( " & _
                     "   RECEIPT_ITEM_ID, " & _
                     "   BILLING_DOC_ID, " & _
                     "   DO_ID, " & _
                     "   RECEIPT_ITEM_AMOUNT, "
            SQL1 = SQL1 & _
                     "   PAID_AMOUNT, " & _
                     "   DEBIT_CREDIT_AMOUNT, " & _
                     "   STATUS_ID, " & _
                     "   REASON_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY, " & _
                    "    DESCRIPTION_DETAIL " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "" & BILLING_DOC_ID & "," & _
                         "" & DO_ID & "," & _
                         "" & RECEIPT_ITEM_AMOUNT & ","
            SQL1 = SQL1 & _
                        "" & PAID_AMOUNT & "," & _
                        "" & DEBIT_CREDIT_AMOUNT & "," & _
                        TempID1 & "," & _
                        TempID2 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ChangeQuote(Trim(DESCRIPTION_DETAIL)) & "'" & _
                     ") "
         RECEIPT_ITEM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE RECEIPT_ITEM SET " & _
                     "   BILLING_DOC_ID = " & BILLING_DOC_ID & ", " & _
                     "   DO_ID = " & DO_ID & ", " & _
                     "   RECEIPT_ITEM_AMOUNT = " & RECEIPT_ITEM_AMOUNT & ", " & _
                     "   PAID_AMOUNT = " & PAID_AMOUNT & ", " & _
                     "   DEBIT_CREDIT_AMOUNT = " & DEBIT_CREDIT_AMOUNT & ", " & _
                     "   STATUS_ID = " & TempID1 & ", " & _
                     "   REASON_ID = " & TempID2 & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & ", " & _
                    "    DESCRIPTION_DETAIL = '" & ChangeQuote(Trim(DESCRIPTION_DETAIL)) & "'" & _
                     "WHERE RECEIPT_ITEM_ID = " & RECEIPT_ITEM_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
      
      Set Bd = New CBillingDoc
      Bd.BILLING_DOC_ID = DO_ID
      Call Bd.UpdatePaidAmount
      Call Bd.UpdateCnDnAmount
      Set Bd = Nothing
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
Dim Bd As CBillingDoc

   SQL1 = "DELETE FROM RECEIPT_ITEM WHERE RECEIPT_ITEM_ID = " & RECEIPT_ITEM_ID
   m_Conn.Execute (SQL1)
   
   Set Bd = New CBillingDoc
   Bd.BILLING_DOC_ID = DO_ID
   Call Bd.UpdatePaidAmount
   Call Bd.UpdateCnDnAmount
   Set Bd = Nothing
End Function
Public Function DeleteDataFromDoID()
Dim SQL1 As String
   SQL1 = "DELETE FROM RECEIPT_ITEM WHERE DO_ID = " & DO_ID
   m_Conn.Execute (SQL1)
End Function
Public Function UpdateValidDate() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE RECEIPT_ITEM SET " & _
                  "   VALID_DATE = '" & ChangeQuote(Trim(DateToStringInt(VALID_DATE))) & "', " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE DO_ID = " & DO_ID
                  
    m_Conn.Execute (SQL1)
End Function
Public Function DeleteDataFromReceiptNo()
Dim SQL1 As String
   SQL1 = "DELETE FROM RECEIPT_ITEM WHERE BILLING_DOC_ID = (SELECT UG.BILLING_DOC_ID FROM BILLING_DOC UG WHERE DOCUMENT_NO = '" & DOCUMENT_NO & "')"
   m_Conn.Execute (SQL1)
   
End Function

