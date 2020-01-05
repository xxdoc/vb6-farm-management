VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBillingDoc"
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
Public ExtraFlag As Long
' End of generic data member

Public BILLING_DOC_ID As Long
Public ACCOUNT_ID As Long
Public DOCUMENT_NO As String
Public DOCUMENT_DATE As Date
Public DOCUMENT_TYPE As Long
Public DOCUMENT_SUBTYPE As Long
Public DUE_DATE As Date
Public NOTE As String
Public ENTERPRISE_ID As Long
Public BILLING_ADDRESS_ID As Long
Public ENTERPRISE_ADDRESS_ID As Long
Public TOTAL_AMOUNT As Double
Public BILL_TOTAL_AMOUNT As Double
Public VAT_PERCENT As Double
Public VAT_AMOUNT As Double
Public WH_PERCENT As Double
Public WH_AMOUNT As Double
Public DISCOUNT_AMOUNT As Double
Public PAYMENT_TERM As Long
Public ACCEPT_BY As Long
Public RECEIVE_BY As Long
Public EXCEPTION_FLAG As String
Public PAYEE_NAME As String
Public INVENTORY_DOC_ID As Long
Public COMMIT_FLAG As String
Public SUPPLIER_ID As Long
Public RECEIPT_TYPE As Long
Public BANK_BRANCH_ID As Long
Public BANK_NOTE As String
Public PAYMENT_TYPE As PAYMENT_TYPE
Public CHECK_NO As String
Public PAYMENT_ID As Long
Public REGION_ID As Long
Public PAID_AMOUNT As Double
Public DEBIT_AMOUNT As Double
Public CREDIT_AMOUNT As Double
Public BATCH_ID As Long
Public SIMULATE_FLAG As String
Public YYYYMM2 As String
Public YYYYMM As String
Public OLD_COMMIT_FLAG As String
Public ACCOUNT_NO As String
Public CUSTOMER_CODE As String
Public CUSTOMER_NAME As String
Public SUPPLIER_CODE As String
Public SUPPLIER_NAME As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public CUSTOMER_ID As Long
Public RECEIPT_REF_COUNT As Long
Public DO_TOTAL_PRICE As Double

Public DO_TOTAL_AMOUNT As Double
Public RECEIPT_ITEM_ID As Long
Public BANK_ID As Long
Public PAID_TYPE As Long
Public TEMP_PAID_AMOUNT As Double
Public RECEIPT_PAID_AMOUNT As Double
Public REVENUE_TOTAL_PRICE As Double
Public Credit As Long
Public DEBIT_CREDIT_AMOUNT As Double
Public STATUS_ID As Long
Public REASON_ID As Long
Public PRODUCT_STATUS_NAME As String
Public DESCRIPTION_DETAIL As String
Public REASON_NAME As String
Public BANK_NAME As String
Public BRANCH_NAME As String
Public TOTAL_SUM_WEIGHT As Double
Public CNDN_TOTAL_PRICE As Double

Public RUNNING_NO As Long
Public CONFIG_DOC_TYPE As Long
Public EMP_CODE As String

Public PKG_TYPE  As Long

Public RECEIPT_DATE  As Date
Public RECEIPT_NO As String

Public FROM_EMP_CODE As String
Public TO_EMP_CODE As String

Public DoItems As Collection
Public RoItems As Collection
Public InventoryDocs As Collection
Public ReceiptItems As Collection
Public CapitalMovements As Collection
Public Revenues As Collection
Public ReceiptCnDns As Collection
Public DocTypeSet As String
Public DocSubTypeSet As String
'----------------------------------------------------------------------------------------------------------------------------------------------------------------
Public ItemSumFlag As Boolean

Public WeightSumFlag As Boolean
Public PriceSumFlag As Boolean
Public AmountSumFlag  As Boolean
Public CountBillSumFlag  As Boolean
Public PaidSumFlag As Boolean
Public DocType34 As Boolean
Public DCSumFlag  As Boolean
'----------------------------------------------------------------------------------------------------------------------------------------------------------------

Public Payments  As Collection

Public FROM_DUE_DATE As Date
Public TO_DUE_DATE As Date
Public TEMP_CHECK_DATE As String
Public TEMP_CHECK_NO As String

Public VALID_DATE As Date
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set DoItems = New Collection
   Set RoItems = New Collection
   Set InventoryDocs = New Collection
   FROM_DATE = -1
   TO_DATE = -1
   FROM_DUE_DATE = -1
   TO_DUE_DATE = -1
   VALID_DATE = -1
   Set ReceiptItems = New Collection
   Set Revenues = New Collection
   Set CapitalMovements = New Collection
   Set ReceiptCnDns = New Collection
   Set Payments = New Collection
   
   '       Set                ItemSumFlag            เป็น False ไว้ นอกนั้นเป็น True แบ้วค่อยไปแก้ในตอนเรียกใช้
   ItemSumFlag = False
   
   WeightSumFlag = True
   PriceSumFlag = True
   AmountSumFlag = True
   CountBillSumFlag = True
   PaidSumFlag = True
   DCSumFlag = True
    '---------------------------------------------------------------------------------------------------------------------------------------
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, itemcount As Long)
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

   SIMULATE_FLAG = glbUser.SIMULATE_FLAG
   
   If Ind = 1 Then
      SelectStr1 = "SELECT BD.*, SNM.LONG_NAME SUPPLIER_NAME, SP.SUPPLIER_CODE, CM.CUSTOMER_CODE, CM.CUSTOMER_ID, CM.CREDIT, NM.LONG_NAME CUSTOMER_NAME, AC.ACCOUNT_NO, NULL RECEIPT_ITEM_ID, BB.BANK_ID, BK.BANK_NAME, BB.BBRANCH_NAME "
      SelectStr1 = SelectStr1 & ",EM.EMP_CODE "
      If ItemSumFlag Then
         If PriceSumFlag Then
            SelectStr1 = SelectStr1 & ", (SELECT SUM(DI.TOTAL_PRICE) FROM DO_ITEM DI WHERE (DI.DO_ID = BD.BILLING_DOC_ID) AND (DI.REVENUE_ID IS NOT NULL)) REVENUE_TOTAL_PRICE "
            SelectStr1 = SelectStr1 & ", (SELECT SUM(DI.TOTAL_PRICE) FROM DO_ITEM DI WHERE (DI.DO_ID = BD.BILLING_DOC_ID) AND (DI.PART_ITEM_ID IS NOT NULL)) DO_TOTAL_PRICE "
            SelectStr1 = SelectStr1 & ", (SELECT SUM(DI.DISCOUNT_AMOUNT) FROM DO_ITEM DI WHERE (DI.DO_ID = BD.BILLING_DOC_ID) AND (DI.PART_ITEM_ID IS NOT NULL)) DISCOUNT_AMOUNT "
        Else
            SelectStr1 = SelectStr1 & ", 0 REVENUE_TOTAL_PRICE "
            SelectStr1 = SelectStr1 & ", 0 DO_TOTAL_PRICE "
            SelectStr1 = SelectStr1 & ", 0 DISCOUNT_AMOUNT "
         End If
         If WeightSumFlag Then
            SelectStr1 = SelectStr1 & ", (SELECT SUM(DI.TOTAL_WEIGHT) FROM DO_ITEM DI WHERE (DI.DO_ID = BD.BILLING_DOC_ID) AND (DI.PART_ITEM_ID IS NOT NULL)) TOTAL_SUM_WEIGHT "
        Else
            SelectStr1 = SelectStr1 & ", 0 TOTAL_SUM_WEIGHT "
         End If
         If AmountSumFlag Then
            SelectStr1 = SelectStr1 & ", (SELECT SUM(DI.ITEM_AMOUNT) FROM DO_ITEM DI WHERE (DI.DO_ID = BD.BILLING_DOC_ID) AND (DI.PART_ITEM_ID IS NOT NULL)) DO_TOTAL_AMOUNT "
        Else
            SelectStr1 = SelectStr1 & ", 0 DO_TOTAL_AMOUNT "
         End If
         If CountBillSumFlag Then
            SelectStr1 = SelectStr1 & ", (SELECT COUNT(*) FROM RECEIPT_ITEM RI1, BILLING_DOC BD1 WHERE (RI1.BILLING_DOC_ID = BD1.BILLING_DOC_ID) AND (BD1.DOCUMENT_TYPE = 2) AND (RI1.DO_ID = BD.BILLING_DOC_ID)) RECEIPT_REF_COUNT  "
         Else
            SelectStr1 = SelectStr1 & ", 0 RECEIPT_REF_COUNT "
         End If
         If PaidSumFlag Then        'รับชำระทั้งใบเสร็จรับเงินเท่าไหร่
            SelectStr1 = SelectStr1 & ", (SELECT SUM(PAID_AMOUNT) FROM RECEIPT_ITEM RI2 WHERE RI2.BILLING_DOC_ID = BD.BILLING_DOC_ID) RECEIPT_PAID_AMOUNT "
         Else
            SelectStr1 = SelectStr1 & ", 0 RECEIPT_PAID_AMOUNT "
         End If
         If DCSumFlag Then              'ลดหนี้เพิ่มหนี้ทั้งใบเท่าไหร่
            SelectStr1 = SelectStr1 & ", (SELECT SUM(DEBIT_CREDIT_AMOUNT) FROM RECEIPT_ITEM RI3 WHERE RI3.BILLING_DOC_ID = BD.BILLING_DOC_ID) CNDN_TOTAL_PRICE "
         Else
            SelectStr1 = SelectStr1 & ", 0 CNDN_TOTAL_PRICE "
         End If
      Else
         SelectStr1 = SelectStr1 & ", 0 REVENUE_TOTAL_PRICE "
         SelectStr1 = SelectStr1 & ", 0 DO_TOTAL_PRICE "
         SelectStr1 = SelectStr1 & ", 0 DO_TOTAL_AMOUNT "
         SelectStr1 = SelectStr1 & ", 0 RECEIPT_REF_COUNT  "
         SelectStr1 = SelectStr1 & ", 0 RECEIPT_PAID_AMOUNT  "
         SelectStr1 = SelectStr1 & ", 0 CNDN_TOTAL_PRICE  "
         SelectStr1 = SelectStr1 & ", 0 TOTAL_SUM_WEIGHT "
      End If
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BILLING_DOC BD "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (BD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (BD.SUPPLIER_ID = SP.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SP.SUPPLIER_ID = SN.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME SNM ON (SN.NAME_ID = SNM.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BANK_BRANCH BB ON (BD.BANK_BRANCH_ID = BB.BBRANCH_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BANK BK ON (BB.BANK_ID = BK.BANK_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EM ON (BD.ACCEPT_BY = EM.EMP_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (EN.EMP_ID = EM.EMP_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN NAME ENM ON (ENM.NAME_ID = EN.NAME_ID) "
      
      If DocType34 = True Then
          WhereStr = "WHERE (DOCUMENT_TYPE = 3 OR DOCUMENT_TYPE = 4) "
      End If
      
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT DISTINCT (BD.ACCOUNT_ID) "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BILLING_DOC BD "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (BD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      
      WhereStr = "WHERE (BD.ACCOUNT_ID IS NOT NULL) AND (BD.ACCOUNT_ID > 0) "
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT BD.BILLING_DOC_ID, CM.CUSTOMER_CODE, CM.CUSTOMER_ID, NM.LONG_NAME CUSTOMER_NAME  "
      SelectStr1 = SelectStr1 & ",BD.DOCUMENT_DATE,BD.DOCUMENT_NO,BD.RECEIPT_TYPE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BILLING_DOC BD "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (BD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID) "
      
      WhereStr = ""
   ElseIf Ind = 4 Then
      SelectStr1 = "BD.BILLING_DOC_ID,BD.PAYMENT_ID,BD.DOCUMENT_NO,BD.DOCUMENT_DATE  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BILLING_DOC BD "
      
      WhereStr = ""
    ElseIf Ind = 5 Then
       SelectStr1 = "SELECT DISTINCT CM.CUSTOMER_CODE, CM.CUSTOMER_ID, NM.LONG_NAME CUSTOMER_NAME  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BILLING_DOC BD "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (BD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CM.CUSTOMER_ID = CN.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID) "
      
      WhereStr = ""
   End If
   
   If BILLING_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BD.BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BD.BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      End If
   End If
   
   If Len(EXCEPTION_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EXCEPTION_FLAG = '" & EXCEPTION_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (EXCEPTION_FLAG = '" & EXCEPTION_FLAG & "') "
      End If
   End If
   
   If SIMULATE_FLAG = "Y" Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SIMULATE_FLAG = '" & SIMULATE_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (SIMULATE_FLAG = '" & SIMULATE_FLAG & "') "
      End If
   Else
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE ((SIMULATE_FLAG <> 'Y') OR (SIMULATE_FLAG IS NULL)) "
      Else
         WhereStr = WhereStr & "AND ((SIMULATE_FLAG <> 'Y') OR (SIMULATE_FLAG IS NULL)) "
      End If
   End If
   
   If Len(COMMIT_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
   End If
   
   If Val(ACCOUNT_ID) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (AC.ACCOUNT_ID = " & ACCOUNT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (AC.ACCOUNT_ID = " & ACCOUNT_ID & ") "
      End If
   End If
   
   If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CM.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If
   
   If Len(Trim(DOCUMENT_NO)) > 0 Then
      If WildCard(DOCUMENT_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "')"
         End If
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

   If Len(Trim(CUSTOMER_NAME)) > 0 Then
      If WildCard(CUSTOMER_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(NM.LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(NM.LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (NM.LONG_NAME = '" & ChangeQuote(Trim(CUSTOMER_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (NM.LONG_NAME = '" & ChangeQuote(Trim(CUSTOMER_NAME)) & "')"
         End If
      End If
   End If
   
   If Len(FROM_EMP_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EM.EMP_CODE >= '" & Trim(ChangeQuote(FROM_EMP_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (EM.EMP_CODE >= '" & Trim(ChangeQuote(FROM_EMP_CODE)) & "') "
      End If
   End If
      
   If Len(TO_EMP_CODE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EM.EMP_CODE <= '" & Trim(ChangeQuote(TO_EMP_CODE)) & "') "
      Else
         WhereStr = WhereStr & "AND (EM.EMP_CODE  <= '" & Trim(ChangeQuote(TO_EMP_CODE)) & "') "
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
    
    If VALID_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(VALID_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE ((VALID_DATE >= '" & ChangeQuote(TempDate) & "') OR (VALID_DATE IS NULL)) "
      Else
         WhereStr = WhereStr & "AND ((VALID_DATE >= '" & ChangeQuote(TempDate) & "') OR (VALID_DATE IS NULL)) "
      End If
   End If

   
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      End If
   End If

   If PAYMENT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PAYMENT_ID = " & PAYMENT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PAYMENT_ID = " & PAYMENT_ID & ") "
      End If
   End If
   
   If Len(DocTypeSet) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DOCUMENT_TYPE IN " & DocTypeSet & ") "
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_TYPE IN " & DocTypeSet & ") "
      End If
   End If
   
   If Len(DocSubTypeSet) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DOCUMENT_SUBTYPE IN " & DocSubTypeSet & ") "
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_SUBTYPE IN " & DocSubTypeSet & ") "
      End If
   End If
   
   If BATCH_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BATCH_ID = " & BATCH_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BATCH_ID = " & BATCH_ID & ") "
      End If
   Else
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE ((BATCH_ID <= 0) OR (BATCH_ID IS NULL)) "
      Else
         WhereStr = WhereStr & "AND ((BATCH_ID <= 0) OR (BATCH_ID IS NULL)) "
      End If
   End If
   
   If FROM_DUE_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DUE_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BD.DUE_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BD.DUE_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DUE_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DUE_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BD.DUE_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BD.DUE_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If PKG_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE ((SELECT COUNT(*) FROM DO_ITEM DI WHERE (DI.DO_ID = BD.BILLING_DOC_ID) AND DI.PKG_TYPE = " & PKG_TYPE & ") > 0)"
      Else
         WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM DO_ITEM DI WHERE (DI.DO_ID = BD.BILLING_DOC_ID) AND DI.PKG_TYPE = " & PKG_TYPE & ") > 0)"
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         If OrderType = 1 Then
            OrderByStr = "ORDER BY DOCUMENT_NO ASC, BILLING_DOC_ID "
         Else
            OrderByStr = "ORDER BY DOCUMENT_NO DESC, BILLING_DOC_ID "
         End If
      ElseIf OrderBy = 2 Then
         If OrderType = 1 Then
            OrderByStr = "ORDER BY DOCUMENT_DATE ASC, BILLING_DOC_ID "
         Else
            OrderByStr = "ORDER BY DOCUMENT_DATE DESC, BILLING_DOC_ID "
         End If
      ElseIf OrderBy = 3 Then
         If OrderType = 1 Then
            OrderByStr = "ORDER BY CUSTOMER_CODE ASC, BILLING_DOC_ID "
         Else
            OrderByStr = "ORDER BY CUSTOMER_CODE DESC, BILLING_DOC_ID "
         End If
      ElseIf OrderBy = 4 Then
         If OrderType = 1 Then
            OrderByStr = "ORDER BY CUSTOMER_CODE ASC, DOCUMENT_DATE "
         Else
            OrderByStr = "ORDER BY CUSTOMER_CODE DESC, DOCUMENT_DATE "
         End If
      ElseIf OrderBy = 5 Then
         If OrderType = 1 Then
            OrderByStr = "ORDER BY NM.LONG_NAME ASC, DOCUMENT_DATE "
         Else
            OrderByStr = "ORDER BY NM.LONG_NAME DESC, DOCUMENT_DATE "
         End If
      ElseIf OrderBy = 6 Then
         If OrderType = 1 Then
            OrderByStr = "ORDER BY DOCUMENT_DATE ASC, DOCUMENT_NO "
         Else
            OrderByStr = "ORDER BY DOCUMENT_DATE ASC, DOCUMENT_NO "
         End If
      ElseIf OrderBy = 7 Then
         If OrderType = 1 Then
            OrderByStr = "ORDER BY BD.DOCUMENT_DATE ASC, BD.DOCUMENT_NO "
         Else
            OrderByStr = "ORDER BY BD.DOCUMENT_DATE DESC, BD.DOCUMENT_NO "
         End If
      ElseIf OrderBy = 8 Then
            OrderByStr = "ORDER BY CUSTOMER_CODE ASC, BD.DOCUMENT_DATE "
      ElseIf OrderBy = 9 Then
            OrderByStr = "ORDER BY NM.LONG_NAME ASC, BD.DOCUMENT_DATE "
      ElseIf OrderBy = 10 Then
         OrderByStr = "ORDER BY CM.CUSTOMER_CODE ASC, DOCUMENT_NO ASC, DUE_DATE "
      Else
         OrderByStr = "ORDER BY BD.BILLING_DOC_ID "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY CUSTOMER_CODE "
   ElseIf Ind = 3 Then
      If OrderType = 1 Then
         OrderByStr = "ORDER BY CM.CUSTOMER_CODE ASC, BD.DOCUMENT_DATE  "
      Else
         OrderByStr = "ORDER BY CM.CUSTOMER_CODE ASC, BD.DOCUMENT_NO  "
      End If
    ElseIf Ind = 5 Then
        OrderByStr = "ORDER BY CM.CUSTOMER_CODE "
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Function
Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      NOTE = NVLS(Rs("NOTE"), "")
      BILLING_ADDRESS_ID = NVLI(Rs("BILLING_ADDRESS_ID"), -1)
      ENTERPRISE_ADDRESS_ID = NVLI(Rs("ENTERPRISE_ADDRESS_ID"), -1)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DUE_DATE = InternalDateToDate(NVLS(Rs("DUE_DATE"), ""))
      ACCOUNT_ID = NVLI(Rs("ACCOUNT_ID"), -1)
      WH_PERCENT = NVLD(Rs("WH_PERCENT"), 0)
      WH_AMOUNT = NVLD(Rs("WH_AMOUNT"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      VAT_PERCENT = NVLI(Rs("VAT_PERCENT"), 0)
      TOTAL_AMOUNT = NVLD(Rs("DO_TOTAL_AMOUNT"), 0)
      BILL_TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      'BILL_TOTAL_AMOUNT
      ENTERPRISE_ID = NVLI(Rs("ENTERPRISE_ID"), -1)
      ACCEPT_BY = NVLI(Rs("ACCEPT_BY"), -1)
      RECEIVE_BY = NVLI(Rs("RECEIVE_BY"), -1)
      PAYMENT_TERM = NVLI(Rs("PAYMENT_TERM"), -1)
      PAYEE_NAME = NVLS(Rs("PAYEE_NAME"), "")
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      EXCEPTION_FLAG = NVLS(Rs("EXCEPTION_FLAG"), "")
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      COMMIT_FLAG = NVLS(Rs("COMMIT_FLAG"), "N")
      OLD_COMMIT_FLAG = COMMIT_FLAG
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), -1)
      RECEIPT_TYPE = NVLI(Rs("RECEIPT_TYPE"), -1)
      DOCUMENT_SUBTYPE = NVLI(Rs("DOCUMENT_SUBTYPE"), -1)
      BANK_BRANCH_ID = NVLI(Rs("BANK_BRANCH_ID"), -1)
      BANK_NOTE = NVLS(Rs("BANK_NOTE"), "")
      CHECK_NO = NVLS(Rs("CHECK_NO"), "")
      PAYMENT_TYPE = NVLI(Rs("PAYMENT_TYPE"), -1)
      PAYMENT_ID = NVLI(Rs("PAYMENT_ID"), -1)
      REGION_ID = NVLI(Rs("REGION_ID"), -1)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      DEBIT_AMOUNT = NVLD(Rs("DEBIT_AMOUNT"), 0)
      CREDIT_AMOUNT = NVLD(Rs("CREDIT_AMOUNT"), 0)
      BATCH_ID = NVLI(Rs("BATCH_ID"), -1)
      SIMULATE_FLAG = NVLS(Rs("SIMULATE_FLAG"), "")
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      YYYYMM2 = NVLS(Rs("YYYYMM2"), "")
      
      EMP_CODE = NVLS(Rs("EMP_CODE"), "")
      
      CNDN_TOTAL_PRICE = NVLD(Rs("CNDN_TOTAL_PRICE"), 0)
      ACCOUNT_NO = NVLS(Rs("ACCOUNT_NO"), "")
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
      SUPPLIER_NAME = NVLS(Rs("SUPPLIER_NAME"), "")
      RECEIPT_REF_COUNT = NVLI(Rs("RECEIPT_REF_COUNT"), -1)
      REVENUE_TOTAL_PRICE = NVLD(Rs("REVENUE_TOTAL_PRICE"), 0)
      DO_TOTAL_PRICE = NVLD(Rs("DO_TOTAL_PRICE"), 0)
      DO_TOTAL_AMOUNT = NVLD(Rs("DO_TOTAL_AMOUNT"), 0)
      BILL_TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      RECEIPT_PAID_AMOUNT = NVLD(Rs("RECEIPT_PAID_AMOUNT"), 0)
      BANK_ID = NVLI(Rs("BANK_ID"), -1)
      
      RECEIPT_ITEM_ID = NVLI(Rs("RECEIPT_ITEM_ID"), -1)
      Credit = NVLI(Rs("CREDIT"), 0)
      
      BANK_NAME = NVLS(Rs("BANK_NAME"), "")
      BRANCH_NAME = NVLS(Rs("BBRANCH_NAME"), "")
      BILLING_DOC_ID = NVLD(Rs("BILLING_DOC_ID"), 0)
      
      TOTAL_SUM_WEIGHT = NVLD(Rs("TOTAL_SUM_WEIGHT"), 0)
      
   ElseIf Ind = 2 Then
      ACCOUNT_ID = NVLI(Rs("ACCOUNT_ID"), -1)
   ElseIf Ind = 3 Then
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      RECEIPT_TYPE = NVLI(Rs("RECEIPT_TYPE"), -1)
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
      
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
   ElseIf Ind = 4 Then
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      PAYMENT_ID = NVLI(Rs("PAYMENT_ID"), -1)
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
   ElseIf Ind = 5 Then
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
      CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
   End If
End Function

Public Function PatchDocDate() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE INVENTORY_DOC SET " & _
                  "   DOCUMENT_DATE = '" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "', " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE INVENTORY_DOC_ID = " & INVENTORY_DOC_ID
                  
   m_Conn.Execute (SQL1)
End Function

Public Function UpdateCommitFlag() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE BILLING_DOC SET " & _
                  "   COMMIT_FLAG = '" & COMMIT_FLAG & "', " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE BILLING_DOC_ID = " & BILLING_DOC_ID
                  
   m_Conn.Execute (SQL1)
End Function

Public Function PatchDueDate() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL1 = "UPDATE BILLING_DOC SET " & _
                  "   DUE_DATE = '" & ChangeQuote(Trim(DateToStringInt(DUE_DATE))) & "', " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE BILLING_DOC_ID = " & BILLING_DOC_ID

   m_Conn.Execute (SQL1)
End Function

Public Function UpdatePaidAmount() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL1 = "UPDATE BILLING_DOC SET " & _
                  "   PAID_AMOUNT =  (SELECT SUM(RI.PAID_AMOUNT) FROM RECEIPT_ITEM RI WHERE (RI.DO_ID = " & BILLING_DOC_ID & ")), " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE BILLING_DOC_ID = " & BILLING_DOC_ID

   m_Conn.Execute (SQL1)
End Function

Public Function UpdateCnDnAmount() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL1 = "UPDATE BILLING_DOC SET " & _
                  "   DEBIT_AMOUNT =  (SELECT SUM(RI.DEBIT_CREDIT_AMOUNT) FROM RECEIPT_ITEM RI, BILLING_DOC BD1 WHERE (RI.BILLING_DOC_ID = BD1.BILLING_DOC_ID) AND (BD1.DOCUMENT_TYPE = 3) AND (RI.DO_ID = " & BILLING_DOC_ID & ")), " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE BILLING_DOC_ID = " & BILLING_DOC_ID
   m_Conn.Execute (SQL1)

   SQL1 = "UPDATE BILLING_DOC SET " & _
                  "   CREDIT_AMOUNT =  (SELECT SUM(RI.DEBIT_CREDIT_AMOUNT) FROM RECEIPT_ITEM RI, BILLING_DOC BD1 WHERE (RI.BILLING_DOC_ID = BD1.BILLING_DOC_ID) AND (BD1.DOCUMENT_TYPE = 4) AND (RI.DO_ID = " & BILLING_DOC_ID & ")), " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE BILLING_DOC_ID = " & BILLING_DOC_ID
   m_Conn.Execute (SQL1)
End Function
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String
Dim TempID5 As String
Dim TempID6 As String
Dim TempID7 As String
Dim TempID8 As String
Dim TempID9 As String
Dim TempID10 As String
Dim TempID11 As String
Dim TempID12 As String
Dim TempID13 As String
Dim TempID14 As String
Dim TempID15 As String
Dim TempID16 As String
Dim TempID17 As String
Dim TempID18 As String

   TempID1 = BILLING_ADDRESS_ID
   If BILLING_ADDRESS_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = ENTERPRISE_ADDRESS_ID
   If ENTERPRISE_ADDRESS_ID <= 0 Then
      TempID2 = "NULL"
   End If
      
   TempID3 = DOCUMENT_TYPE
   If DOCUMENT_TYPE <= 0 Then
      TempID3 = "NULL"
   End If
      
   TempID4 = VAT_PERCENT
   If VAT_PERCENT <= 0 Then
      TempID4 = "NULL"
   End If
      
   TempID5 = ENTERPRISE_ID
   If ENTERPRISE_ID <= 0 Then
      TempID5 = "NULL"
   End If

   TempID6 = WH_PERCENT
   If WH_PERCENT <= 0 Then
      TempID6 = "NULL"
   End If

   TempID8 = WH_AMOUNT
   If WH_AMOUNT <= 0 Then
      TempID8 = "NULL"
   End If

   TempID9 = ACCEPT_BY
   If ACCEPT_BY <= 0 Then
      TempID9 = "NULL"
   End If
   
   TempID10 = RECEIVE_BY
   If RECEIVE_BY <= 0 Then
      TempID10 = "NULL"
   End If
   
   TempID11 = PAYMENT_TERM
   If PAYMENT_TERM <= 0 Then
      TempID11 = "NULL"
   End If
   
   TempID12 = INVENTORY_DOC_ID
   If INVENTORY_DOC_ID <= 0 Then
      TempID12 = "NULL"
   End If
   
   TempID13 = SUPPLIER_ID
   If SUPPLIER_ID <= 0 Then
      TempID13 = "NULL"
   End If
   
   TempID14 = RECEIPT_TYPE
   If RECEIPT_TYPE <= 0 Then
      TempID14 = "NULL"
   End If
   
   TempID15 = BANK_BRANCH_ID
   If BANK_BRANCH_ID <= 0 Then
      TempID15 = "NULL"
   End If
   
   TempID16 = PAYMENT_ID
   If PAYMENT_ID <= 0 Then
      TempID16 = "NULL"
   End If
   
   TempID17 = REGION_ID
   If REGION_ID <= 0 Then
      TempID17 = "NULL"
   End If
   
   TempID18 = BATCH_ID
   If BATCH_ID <= 0 Then
      TempID18 = "NULL"
   End If
   
   If DUE_DATE <= 0 Then
      DUE_DATE = DOCUMENT_DATE
   End If
   
   SIMULATE_FLAG = glbUser.SIMULATE_FLAG
   YYYYMM = Mid(DateToStringInt(DOCUMENT_DATE), 1, 7)
   YYYYMM2 = Mid(DateToStringInt(DUE_DATE), 1, 7)
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("BILLING_DOC_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO BILLING_DOC " & _
                     "( " & _
                     "   BILLING_DOC_ID, " & _
                     "   DOCUMENT_NO, " & _
                     "   NOTE, " & _
                     "   BILLING_ADDRESS_ID, " & _
                     "   ENTERPRISE_ADDRESS_ID, " & _
                     "   DOCUMENT_DATE, " & _
                     "   DUE_DATE, " & _
                     "   ACCOUNT_ID, " & _
                     "   DOCUMENT_TYPE, " & _
                     "   TOTAL_AMOUNT, " & _
                     "   VAT_PERCENT, "
               SQL1 = SQL1 & _
                     "   VAT_AMOUNT, " & _
                     "   WH_PERCENT, " & _
                     "   WH_AMOUNT, " & _
                      "   ENTERPRISE_ID, " & _
                      "   ACCEPT_BY, " & _
                      "   RECEIVE_BY, " & _
                      "   PAYMENT_TERM, " & _
                      "   PAYEE_NAME, " & _
                      "   DISCOUNT_AMOUNT, " & _
                      "   EXCEPTION_FLAG, " & _
                      "   INVENTORY_DOC_ID, " & _
                      "   COMMIT_FLAG, " & _
                      "   SUPPLIER_ID, " & _
                      "   RECEIPT_TYPE, " & _
                      "   DOCUMENT_SUBTYPE, " & _
                      "   BANK_BRANCH_ID, " & _
                      "   BANK_NOTE, " & _
                      "   PAYMENT_TYPE, " & _
                      "   CHECK_NO, "
         SQL1 = SQL1 & _
                     "   PAYMENT_ID, " & _
                     "   REGION_ID, " & _
                     "   PAID_AMOUNT, " & _
                     "   CREDIT_AMOUNT, " & _
                     "   DEBIT_AMOUNT, " & _
                     "   BATCH_ID, " & _
                     "   SIMULATE_FLAG, " & _
                     "   YYYYMM, " & _
                     "   YYYYMM2, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
                  SQL1 = SQL1 & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(DOCUMENT_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(NOTE)) & "'," & _
                         TempID1 & "," & _
                         TempID2 & "," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(DUE_DATE))) & "'," & _
                         ACCOUNT_ID & "," & _
                         TempID3 & "," & _
                         TOTAL_AMOUNT & "," & _
                         TempID4 & ","
                  SQL1 = SQL1 & _
                         VAT_AMOUNT & "," & _
                         TempID6 & "," & _
                         TempID8 & "," & _
                         TempID5 & "," & _
                         TempID9 & "," & _
                         TempID10 & "," & _
                         TempID11 & "," & _
                         "'" & ChangeQuote(Trim(PAYEE_NAME)) & "'," & _
                         DISCOUNT_AMOUNT & "," & _
                         "'" & ChangeQuote(Trim(EXCEPTION_FLAG)) & "'," & _
                         TempID12 & "," & _
                         "'" & ChangeQuote(Trim(COMMIT_FLAG)) & "'," & _
                         TempID13 & "," & _
                         TempID14 & "," & _
                         DOCUMENT_SUBTYPE & "," & _
                         TempID15 & "," & _
                         "'" & ChangeQuote(Trim(BANK_NOTE)) & "'," & _
                         PAYMENT_TYPE & "," & _
                         "'" & ChangeQuote(Trim(CHECK_NO)) & "',"
            SQL1 = SQL1 & _
                         TempID16 & "," & _
                         TempID17 & "," & _
                         PAID_AMOUNT & "," & _
                         CREDIT_AMOUNT & "," & _
                         DEBIT_AMOUNT & "," & _
                         TempID18 & "," & _
                         "'" & ChangeQuote(Trim(SIMULATE_FLAG)) & "'," & _
                         "'" & ChangeQuote(Trim(YYYYMM)) & "'," & _
                         "'" & ChangeQuote(Trim(YYYYMM2)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         BILLING_DOC_ID = ID
         
   Else 'Show edit
      SQL1 = "UPDATE BILLING_DOC SET " & _
                     "   DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "', " & _
                     "   NOTE = '" & ChangeQuote(Trim(NOTE)) & "', " & _
                     "    BILLING_ADDRESS_ID = " & TempID1 & ", " & _
                     "    ENTERPRISE_ADDRESS_ID = " & TempID2 & ", " & _
                     "   DOCUMENT_DATE = '" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "', " & _
                     "   DUE_DATE = '" & ChangeQuote(Trim(DateToStringInt(DUE_DATE))) & "', " & _
                     "    ACCOUNT_ID = " & ACCOUNT_ID & ", " & _
                     "    DOCUMENT_TYPE = " & TempID3 & ", " & _
                     "    TOTAL_AMOUNT = " & TOTAL_AMOUNT & ", " & _
                     "    VAT_PERCENT = " & TempID4 & ", " & _
                     "    VAT_AMOUNT = " & VAT_AMOUNT & ", "
            SQL1 = SQL1 & _
                     "    WH_PERCENT = " & TempID6 & ", " & _
                     "    WH_AMOUNT = " & TempID8 & ", " & _
                     "    ENTERPRISE_ID = " & TempID5 & ", " & _
                     "    ACCEPT_BY = " & TempID9 & ", " & _
                     "    RECEIVE_BY = " & TempID10 & ", " & _
                     "    PAYMENT_TERM = " & TempID11 & ", " & _
                     "   PAYEE_NAME = '" & ChangeQuote(Trim(PAYEE_NAME)) & "', " & _
                     "    DISCOUNT_AMOUNT = " & DISCOUNT_AMOUNT & ", " & _
                     "   EXCEPTION_FLAG = '" & ChangeQuote(Trim(EXCEPTION_FLAG)) & "', " & _
                     "    INVENTORY_DOC_ID = " & TempID12 & ", " & _
                     "   COMMIT_FLAG = '" & ChangeQuote(Trim(COMMIT_FLAG)) & "', " & _
                     "    SUPPLIER_ID = " & TempID13 & ", " & _
                     "    RECEIPT_TYPE = " & TempID14 & ", " & _
                     "    DOCUMENT_SUBTYPE = " & DOCUMENT_SUBTYPE & ", " & _
                     "    BANK_BRANCH_ID = " & TempID15 & ", " & _
                     "   BANK_NOTE = '" & ChangeQuote(Trim(BANK_NOTE)) & "', " & _
                     "    PAYMENT_TYPE = " & PAYMENT_TYPE & ", " & _
                     "   CHECK_NO = '" & ChangeQuote(Trim(CHECK_NO)) & "', " & _
                     "    PAYMENT_ID = " & TempID16 & ", " & _
                     "    REGION_ID = " & TempID17 & ", "
            SQL1 = SQL1 & _
                     "    PAID_AMOUNT = " & PAID_AMOUNT & ", " & _
                     "    CREDIT_AMOUNT = " & CREDIT_AMOUNT & ", " & _
                     "    DEBIT_AMOUNT = " & DEBIT_AMOUNT & ", " & _
                     "    BATCH_ID = " & TempID18 & ", " & _
                     "   SIMULATE_FLAG = '" & ChangeQuote(Trim(SIMULATE_FLAG)) & "', " & _
                     "   YYYYMM = '" & ChangeQuote(Trim(YYYYMM)) & "', " & _
                     "   YYYYMM2 = '" & ChangeQuote(Trim(YYYYMM2)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE BILLING_DOC_ID = " & BILLING_DOC_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
Dim Cm As CCapitalMovement
Dim Cl As CCapitalLoss
Dim Ri As CReceiptItem
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim Ct As CCashTran
Dim Cq As CCheque

   Set Cm = New CCapitalMovement
   Cm.CAPITAL_MOVEMENT_ID = -1
   Cm.BILLING_DOC_ID = BILLING_DOC_ID
   Call Cm.DeleteDataFromBillingDoc
   Set Cm = Nothing
   
   Set Cl = New CCapitalLoss
   Cl.CAPITAL_LOSS_ID = -1
   Cl.BILLING_DOC_ID = BILLING_DOC_ID
   Call Cl.DeleteDataFromBillingDoc
   Set Cl = Nothing
   
   Set TempRs = New ADODB.Recordset
   Set Ri = New CReceiptItem
   Ri.BILLING_DOC_ID = BILLING_DOC_ID
   Call Ri.QueryData(1, TempRs, iCount)
   While Not TempRs.EOF
      Call Ri.PopulateFromRS(1, TempRs)
      Call Ri.DeleteData
      
      TempRs.MoveNext
   Wend
   Set Ri = Nothing
   Set TempRs = Nothing
      
   
   Set TempRs = New ADODB.Recordset
   Set Ct = New CCashTran
   Call Ct.SetFieldValue("BILLING_DOC_ID", BILLING_DOC_ID)
   Call Ct.QueryData(1, TempRs, iCount)
   Set Ct = Nothing
   While Not TempRs.EOF
      Set Ct = New CCashTran
      Call Ct.PopulateFromRS(1, TempRs)
      Call Ct.DeleteData
      If Ct.GetFieldValue("CHECK_ID") > 0 Then
         Set Cq = New CCheque
         Call Cq.SetFieldValue("CHEQUE_ID", Ct.GetFieldValue("CHECK_ID"))
         Call Cq.DeleteData
         Set Cq = Nothing
      End If
      Set Ct = Nothing
      Call TempRs.MoveNext
   Wend


   SQL1 = "DELETE FROM DO_ITEM WHERE DO_ID = " & BILLING_DOC_ID
   m_Conn.Execute (SQL1)
   
   SQL1 = "DELETE FROM RO_ITEM WHERE BILLING_DOC_ID = " & BILLING_DOC_ID
   m_Conn.Execute (SQL1)

   SQL1 = "DELETE FROM RECEIPT_CNDN WHERE BILLING_DOC_ID = " & BILLING_DOC_ID
   m_Conn.Execute (SQL1)
   
   SQL1 = "DELETE FROM BILLING_ADJUST WHERE BILLING_DOC_ID = " & BILLING_DOC_ID
   m_Conn.Execute (SQL1)

   SQL1 = "DELETE FROM BILLING_DOC WHERE BILLING_DOC_ID = " & BILLING_DOC_ID & " "
   m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set DoItems = Nothing
   Set RoItems = New Collection
   Set InventoryDocs = Nothing
   Set ReceiptItems = Nothing
   Set CapitalMovements = Nothing
   Set Revenues = Nothing
   Set ReceiptCnDns = Nothing
   Set Payments = Nothing
End Sub

Public Sub RemoveARBalance()
Dim SQL1 As String
Dim iCount As Long
Dim TempDate As String

   TempDate = DateToStringIntHi(TO_DATE)
   
   SQL1 = "UPDATE BILLING_DOC SET ACCOUNT_ID = (SELECT ACCOUNT_ID FROM ACCOUNT WHERE CUSTOMER_ID = " & CUSTOMER_ID & ") "
   m_Conn.Execute (SQL1)

   SQL1 = "DELETE FROM ACCOUNT WHERE CUSTOMER_ID <> " & CUSTOMER_ID & " "
   m_Conn.Execute (SQL1)

   SQL1 = "DELETE FROM CUSTOMER_ADDRESS WHERE (CUSTOMER_ID <> " & CUSTOMER_ID & ") "
   m_Conn.Execute (SQL1)
   
   SQL1 = "DELETE FROM CUSTOMER_NAME WHERE (CUSTOMER_ID <> " & CUSTOMER_ID & ") "
   m_Conn.Execute (SQL1)
   
   SQL1 = "DELETE FROM CUSTOMER WHERE (CUSTOMER_ID <> " & CUSTOMER_ID & ") "
   m_Conn.Execute (SQL1)
   
   SQL1 = "UPDATE BILLING_DOC SET DOCUMENT_TYPE = 2, RECEIPT_TYPE = 1 WHERE DOCUMENT_TYPE = 1 AND DOCUMENT_DATE <= '" & TempDate & "' "
   m_Conn.Execute (SQL1)
End Sub
Public Function UpdateValidDate() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE BILLING_DOC SET " & _
                  "   VALID_DATE = '" & ChangeQuote(Trim(DateToStringInt(VALID_DATE))) & "', " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE BILLING_DOC_ID = " & BILLING_DOC_ID
                  
    m_Conn.Execute (SQL1)
End Function
Public Sub SetFlag(WeightFlag As Boolean, PriceFlag As Boolean, AmountFlag As Boolean, CountBillFlag As Boolean, PaidFlag As Boolean, DCFlag As Boolean)
   WeightSumFlag = WeightFlag
   PriceSumFlag = PriceFlag
   AmountSumFlag = AmountFlag
   CountBillSumFlag = CountBillFlag
   PaidSumFlag = PaidFlag
   DCSumFlag = DCFlag
End Sub

