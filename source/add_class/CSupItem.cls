VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSupItem"
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

Public SUP_ITEM_ID As Long
Public DO_ID As Long
Public PART_ITEM_ID As Long
Public LOCATION_ID As Long
Public ACTUAL_UNIT_PRICE As Double
Public INCLUDE_UNIT_PRICE As Double
Public PREVIOUS_PRICE As Double
Public PREVIOUS_AMOUNT As Double
Public NEW_PRICE As Double
Public TX_AMOUNT As Double
Public NEW_AMOUNT As Double
Public TRANSACTION_SEQ As Long
Public GUI_ID As Long
Public CALCULATE_FLAG As String
Public FATHER_NO As String
Public MOTHER_NO As String
Public TOTAL_WEIGHT As Double
Public TOTAL_ACTUAL_PRICE As Double
Public TOTAL_INCLUDE_PRICE As Double
Public TX_TYPE As String
Public LEFT_AMOUNT As Double
Public LOT_MANUAL As String
Public MULTIPLIER As Long
Public LAYOUT_ID As Long
Public LINK_ID As Long
Public PACKAGING_AMT As Double
Public ENTRY_WEIGHT As Double
Public EXIT_WEIGHT As Double
Public EXPORT_AMOUNT As Double
Public WEIGHT_AMOUNT As Double
Public PACKAGE_WEIGHT As Double
Public OTHER_WEIGHT As Double
Public PERCENT_HUMID As Double
Public HUMID_WEIGHT As Double
Public PACKAGING_ID As Long
Public PACKAGING_NAME As String
Public SUPPLIER_WEIGHT As Double
Public CALCULATE_TYPE As Long
Public PACKAGE_AMOUNT As Double
Public ACTUAL_PKG_AMOUNT As Double
Public PUREXP_ID1 As Long
Public PUREXP_ID2 As Long
Public EXPENSE1 As Double
Public EXPENSE2 As Double
Public RAW_PRICE As Double
Public DISCOUNT_AMT As Double
Public RAW_TOT_PRICE As Double
Public ITEM_DESC As String
Public TO_DEPARTMENT As Long
Public EXTRA_NAME As String
Public SALE_TOT_PRICE As Double
Public CALCULATE_WEIGHT As Double
Public TOTAL_NEW_PRICE As Double
Public BAG_RETURN As String
Public AUTO_PRICE As String
Public ACTUAL_AMOUNT As Double
Public ACTUAL_PRICE As Double
Public CURRENT_AMOUNT As Double
Public CURRENT_PRICE As Double
Public NEED_TOTAL_AMOUNT As Double
Public NEED_TOTAL_PRICE As Double
Public NEED_AVG_PRICE As Double
Public MANUAL_PRICE As String
Public DOCUMENT_TYPE_SET As String
Public CUSTOMER_TYPE As Long


'Public CURRENT_AMOUNT As Double
Public BALANCE_AMOUNT As Double
Public ALL_IMPORT_AMT As Double
Public ALL_EXPORT_AMT As Double
Public EXPORT_AVG_PRICE As Double
Public EXPORT_TOTAL_PRICE As Double
Public IMPORT_AMOUNT As Double

Public PRODUCT_ID As Long
Public PRODUCT_NO As String
Public PRODUCT_NAME As String
Public PART_NO As String
Public PART_DESC As String
Public PART_TYPE As Long
Public PART_TYPE_NAME As String
Public PART_SUPPLIER_NAME As String
Public PART_NOTE As String
Public PART_TYPE_NO As String
Public LOCATION_NO As String
Public LOCATION_NAME As String
Public OLD_TX_AMOUNT As Double
Public OLD_LOCATION_ID As Long
Public PIG_TYPE As String
Public UNIT_NAME As String
Public DOCUMENT_DATE As Date
Public COMMIT_FLAG As String
Public DOCUMENT_TYPE As Long
Public PIG_FLAG As String
Public SUM_ALL_IMPORT_AMT As Double
Public SUM_ALL_IMPORT As Double
Public SALE_FLAG As String
Public DOCUMENT_NO As String
Public PO_DOCUMENT_NO As String
Public RO_DOCUMENT_NO As String
Public SUPPLIER_ID As Long
Public QUE_NO As String
Public RECEIVE_NAME As String
Public TRUCK_NO As String
Public SUPPLIER_CODE As String
Public SUPPLIER_NAME As String
Public SUPPLIER_TYPE As Long
Public RAW_COST As Double
Public EXPENSE_COST  As Double
Public PROCESS_ID As Long
Public EXPENSE_TYPE As Long
Public SENDER_NAME As String
Public PROJECT_NAME_ID As Long


Public FROM_TX_TYPE As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public TO_DATE1 As Date
Public USED_AMOUNT As Double
Public FROM_LOCATION As String
Public TO_LOCATION As String
Public FROM_LOCATION_ID As Long
Public TO_LOCATION_ID As Long
Public PRIORITY1 As Long
Public PACKAGING_CODE As String
Public PUREXP_NO1 As String
Public PUREXP_NO2 As String
Public PUREXP_NAME1 As String
Public PUREXP_NAME2 As String
Public EXPENSE_RATE1 As Double
Public EXPENSE_RATE2 As Double
Public PART_GROUP_ID As Long
Public EXPENSE_NAME As String
Public DEP_NAME As String
Public DEP_NO As String
Public DO_NO As String
Public ITEM_DESC_ID As Long
Public ITEM_DESC_NO As String
Public ITEM_DESC_NAME As String
Public DOCUMENT_SUBTYPE As Long

Public DEPARTMENT_NAME As String
Public CSTTYPE_NAME As String
Public PR_NO As String

Public PO_ID As Long
Public PO_NO As String

'-------------------------------------------------------------------------------------------------------------------------------->
'For Export  tabel BillingDoc
Public BILLING_DOC_ID As Long
Public ACCOUNT_ID As Long
Public DUE_DATE As Date
Public NOTE As String
Public ENTERPRISE_ID As Long
Public BILLING_ADDRESS_ID As Long
Public ENTERPRISE_ADDRESS_ID As Long
Public TOTAL_AMOUNT As Double
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
Public RECEIPT_TYPE As Long
Public DEPOSIT_AMOUNT As Double
Public TOTAL_PRICE As Double
Public ESTIMATE_DATE As Date
Public APPROVE_DATE As Date
Public RESOURCE_ID As Long
Public PAYMENT_TYPE As Long
Public BANK_ID As Long
Public BBRANCH_ID As Long
Public BANK_NAME As String
Public BBRANCH_NAME As String
Public CHECK_NO As String
Public CHECK_DATE As Date
Public AREA_CODE As String
Public CD_PERCENT As Double
Public CD_AMOUNT As Double
Public PACKAGE_TYPE As Long
Public TEMP_DO_NO As String
Public PAID_AMOUNT As Double
Public DEBIT_AMOUNT As Double
Public CREDIT_AMOUNT As Double
Public DO_TOTAL_PRICE As Double
Public REVENUE_TOTAL_PRICE As Double
Public PAYMENT_ID As Long
Public BANK_BRANCH_ID As Long
Public BANK_NOTE As String
Public BATCH_ID As Long
Public SIMULATE_FLAG As String
Public ENTRY_DATE As Date
Public EXIT_DATE As Date

Public RETURN_AMOUNT As Double

Public SUM_DISCOUNT_AMOUNT As Double
Public SUM_PAID_AMOUNT As Double
Public TOTAL_DEBIT_AMOUNT As Double
Public TOTAL_CREDIT_AMOUNT As Double
Public SUM_CASH_DISCOUNT As Double
Public OLD_COMMIT_FLAG As String
Public ACCOUNT_NO As String
Public CUSTOMER_CODE As String
Public CUSTOMER_NAME As String
Public FROM_DUE_DATE As Date
Public TO_DUE_DATE As Date
Public CUSTOMER_ID As Long
Public DEBIT_CREDIT_AMOUNT As Double
Public RECEIPT_ITEM_ID As Long
Public RECEIPT_REF_COUNT As Long
Public DBN_REF_COUNT As Long
Public CDN_REF_COUNT As Long
Public BLS_REF_COUNT As Long
Public RECEIPT_REF_COUNT2 As Long
Public DBN_REF_COUNT2 As Long
Public CDN_REF_COUNT2 As Long
Public TAX_FLAG As String
Public WH_FLAG As String
Public EMP_CODE As String
Public EMP_NAME As String
Public EMP_LNAME As String
Public RESOURCE_NO As String
Public FEATURE_CODE As String
Public VADILITY As Date
Public DELIVERY As Date
Public PAYMENT As Date
Public PAYMENT_DESC As String
Public SHIPMENT As Date
Public TEMP_PAID_AMOUNT As Double
Public TEMP_CHECK_DATE As String
Public TEMP_CHECK_NO As String
Public PAID_TYPE As Long
Public RECEIPT_PAID_AMOUNT As Double
Public REASON_ID As Long
Public CASH_DISCOUNT As Double
Public DO_RECEIPT_FLAG As String
Public CUSTOMER_GRADE As Long
Public SUPPLIER_GRADE As Long

Public REFER_INV As String
Public PACKING_OF As String
Public SHIP_NAME As String
Public SHIPPING_MARKS As String
Public SHIPON_DATE As Date
Public REF As String
Public Credit As Long
Public CREDIT_LIMIT As Double
Public CUSTOMER_CREDIT As Long
Public NEW_ACCOUNT_ID As Long

Public TOTAL_RCP As Double
Public AGREEMENT_DATA As String
Public AGREEMENT_FINANCE As String

Public RUNNING_NO As Long
Public CONFIG_DOC_TYPE As CONFIG_DOC_TYPE

Public OLD_CREDIT_AMOUNT As Double

Public FROM_CUSTOMER_CODE As String
Public TO_CUSTOMER_CODE As String

'------------------------------------------------------------------------------------------------------------------------------------------------>
Public DELIVERY_FEE As Double
Public DEPARTMENT_ID As Long
'------------------------------------------------------------------------------------------------------------------------------------------------>

'------------------------------------------------------------------------------------------------------------------------------------------------>
Public CONDITION_NAME As String
Public PAID_TYPE_NAME As String
Public CLOSE_FLAG As String
Public PO_APPROVED_FLAG As String
Public AUTO_GEN_FLAG As String

'------------------------------------------------------------------------------------------------------------------------------------------------>


'-------------------------------------------------------------------------------------------------------------------------------->
Public HiFlag As Boolean

Public SubLotItems As Collection
Public IncludeFlag As Boolean

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_DATE = -1
   TO_DATE = -1
    CLOSE_FLAG = ""
    PO_APPROVED_FLAG = ""
   Set SubLotItems = New Collection
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
      SelectStr1 = "SELECT UG.*, PT.*, LC.*, PI.*, UN.*, IVD.SUPPLIER_ID, IVD.NOTE,IVD.DOCUMENT_DATE, IVD.DOCUMENT_NO, IVD.COMMIT_FLAG, IVD.DOCUMENT_TYPE, IVD.QUE_NO, IVD.RECEIVE_NAME, IVD.TRUCK_NO, IVD.DO_NO, IVD.SENDER_NAME, SP.SUPPLIER_CODE, NM1.LONG_NAME SUPPLIER_NAME, DP.PRIORITY1, PG.PACKAGING_NO, PE1.PUREXP_NO PUREXP_NO1, PE1.PUREXP_NAME PUREXP_NAME1, PE1.EXPENSE_RATE EXPENSE_RATE1, PE2.PUREXP_NO PUREXP_NO2, PE2.PUREXP_NAME PUREXP_NAME2, PE2.EXPENSE_RATE EXPENSE_RATE2, PT.PART_GROUP_ID PART_GROUP_ID1 "
      SelectStr1 = SelectStr1 & ",MR1.KEY_CODE ITEM_DESC_NO,MR1.KEY_NAME ITEM_DESC_NAME,IVD.PR_NO "
      SelectStr1 = SelectStr1 & ",LO.LAY_OUT_NAME DEPARTMENT_NAME "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (SN.NAME_ID = NM1.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN DOCUMENT_PRIORITY DP ON ((IVD.DOCUMENT_TYPE = DP.DOCUMENT_TYPE) AND (UG.TX_TYPE = DP.TX_TYPE) AND (DP.AREA = 1)) "
      FromStr = FromStr & "LEFT OUTER JOIN PACKAGING PG ON (UG.PACKAGING_ID = PG.PACKAGING_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PURCHASE_EXPENSE PE1 ON (UG.PUREXP_ID1 = PE1.PUREXP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PURCHASE_EXPENSE PE2 ON (UG.PUREXP_ID2 = PE2.PUREXP_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (MR1.KEY_ID = UG.ITEM_DESC_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN LAY_OUT LO ON (IVD.DEPARTMENT_ID = LO.LAY_OUT_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT UG.*, PT.*, LC.*, PI.*, UN.*, IVD.DOCUMENT_DATE, IVD.DOCUMENT_NO, IVD.COMMIT_FLAG, IVD.DOCUMENT_TYPE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      
      WhereStr = "WHERE (UG.LEFT_AMOUNT > 0) "
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT LI.PART_ITEM_ID, SUM(LI.MULTIPLIER * LI.TX_AMOUNT) LEFT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SUP_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (LI.DO_ID = IVD.BILLING_DOC_ID) "
      
      WhereStr = ""
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT LI1.PART_ITEM_ID, LI1.LOCATION_ID, LI1.NEW_PRICE, LI1.PREVIOUS_PRICE, LI1.NEW_AMOUNT, LI1.PREVIOUS_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM SUP_ITEM LI1 "

      WhereStr = "WHERE LI1.TRANSACTION_SEQ IN "
      WhereStr = WhereStr & "( "
      WhereStr = WhereStr & "   SELECT MAX(LI2.TRANSACTION_SEQ) "
      WhereStr = WhereStr & "   FROM SUP_ITEM LI2 "
      WhereStr = WhereStr & "   LEFT OUTER JOIN BILLING_DOC IVD2 ON (LI2.BILLING_DOC_ID = IVD2.BILLING_DOC_ID) "
      If HiFlag Then
         WhereStr = WhereStr & "   WHERE (IVD2.DOCUMENT_DATE <= '" & DateToStringIntHi(TO_DATE1) & "') "
      Else
         WhereStr = WhereStr & "   WHERE (IVD2.DOCUMENT_DATE <= '" & DateToStringIntLow(TO_DATE1) & "') "
      End If
      WhereStr = WhereStr & "   AND (LI2.PART_ITEM_ID = LI1.PART_ITEM_ID) "
      WhereStr = WhereStr & "   AND (LI2.LOCATION_ID = LI1.LOCATION_ID) "
      WhereStr = WhereStr & ") "
   ElseIf Ind = 5 Then
      SelectStr1 = "SELECT DISTINCT PT.PART_TYPE_ID "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "

      WhereStr = ""
   ElseIf Ind = 6 Then
      SelectStr1 = "SELECT DISTINCT LI.PART_ITEM_ID "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SUP_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (LI.DO_ID = IVD.BILLING_DOC_ID) "

      WhereStr = ""
   ElseIf (Ind = 7) Then
      SelectStr1 = "SELECT IVD.SUPPLIER_ID, SUM(LI.TX_AMOUNT) TX_AMOUNT, SUM(TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM SUP_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (LI.DO_ID = IVD.BILLING_DOC_ID) "

      WhereStr = ""
   ElseIf Ind = 8 Then
      SelectStr1 = "SELECT LI.PART_ITEM_ID, LI.LOCATION_ID, LI.LAYOUT_ID, SUM(LI.LEFT_AMOUNT) LEFT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM SUP_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (LI.DO_ID = IVD.BILLING_DOC_ID) "

      WhereStr = "WHERE (LI.LAYOUT_ID IS NOT NULL) "
   ElseIf Ind = 9 Then
      SelectStr1 = "SELECT UG.*, PT.*, LC.*, PI.*, UN.*, IVD.DOCUMENT_DATE, IVD.DOCUMENT_NO, IVD.COMMIT_FLAG, IVD.DOCUMENT_TYPE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      
      WhereStr = "WHERE (UG.LEFT_AMOUNT > 0) "
   ElseIf Ind = 10 Then
      SelectStr1 = "SELECT DISTINCT UG.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      
      WhereStr = ""
   ElseIf Ind = 11 Then
      SelectStr1 = "SELECT UG.*, PT.*, LC1.LOCATION_NAME FROM_LOCATION, LC2.LOCATION_NAME TO_LOCATION, PI.*, UN.*, IVD.DOCUMENT_DATE, IVD.DOCUMENT_NO "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN SUP_ITEM LI2 ON ((UG.GUI_ID = LI2.GUI_ID) AND (LI2.TX_TYPE = 'I')) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC1 ON (UG.LOCATION_ID = LC1.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC2 ON (LI2.LOCATION_ID = LC2.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      
      WhereStr = ""
   ElseIf Ind = 12 Then
      SelectStr1 = "SELECT UG.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, UG.LOCATION_ID, SUM(UG.WEIGHT_AMOUNT) WEIGHT_AMOUNT, SUM(UG.PACKAGE_AMOUNT) PACKAGE_AMOUNT, SUM(UG.ACTUAL_PKG_AMOUNT) ACTUAL_PKG_AMOUNT "
      SelectStr1 = SelectStr1 & ", SUM(UG.PACKAGE_WEIGHT) PACKAGE_WEIGHT, SUM(UG.HUMID_WEIGHT) HUMID_WEIGHT, SUM(UG.OTHER_WEIGHT) OTHER_WEIGHT "
      SelectStr1 = SelectStr1 & ", SUM(UG.TOTAL_WEIGHT) TOTAL_WEIGHT,  SUM(UG.TX_AMOUNT) TX_AMOUNT,  SUM(UG.TOTAL_ACTUAL_PRICE) TOTAL_ACTUAL_PRICE,  SUM(UG.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE  "
      SelectStr1 = SelectStr1 & ", SUM(UG.CALCULATE_WEIGHT) CALCULATE_WEIGHT , SUM(UG.EXPENSE1) EXPENSE1, SUM(UG.EXPENSE2) EXPENSE2 "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      
      WhereStr = ""
   ElseIf Ind = 13 Then
      SelectStr1 = "SELECT IVD.SUPPLIER_ID, SUM(UG.WEIGHT_AMOUNT) WEIGHT_AMOUNT, SUM(UG.PACKAGE_AMOUNT) PACKAGE_AMOUNT "
      SelectStr1 = SelectStr1 & ", SUM(UG.PACKAGE_WEIGHT) PACKAGE_WEIGHT, SUM(UG.HUMID_WEIGHT) HUMID_WEIGHT, SUM(UG.OTHER_WEIGHT) OTHER_WEIGHT "
      SelectStr1 = SelectStr1 & ", SUM(UG.TOTAL_WEIGHT) TOTAL_WEIGHT,  SUM(UG.TX_AMOUNT) TX_AMOUNT,  SUM(UG.TOTAL_ACTUAL_PRICE) TOTAL_ACTUAL_PRICE,  SUM(UG.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE  "
      SelectStr1 = SelectStr1 & ", SUM(UG.CALCULATE_WEIGHT) CALCULATE_WEIGHT , SUM(UG.EXPENSE1) EXPENSE1, SUM(UG.EXPENSE2) EXPENSE2 "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (IVD.SUPPLIER_ID = SP.SUPPLIER_ID) "
      
      WhereStr = ""
   ElseIf Ind = 14 Then
      SelectStr1 = "SELECT DISTINCT UG.LOCATION_ID, LC.LOCATION_NO, LC.LOCATION_NAME, UG.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
      SelectStr1 = SelectStr1 & ",PT.PART_TYPE_NO, PT.PART_TYPE_NAME "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "

      WhereStr = ""
   ElseIf Ind = 15 Then
      SelectStr1 = "SELECT UG.LOCATION_ID, UG.PART_ITEM_ID, UG.TX_TYPE, SUM(UG.TX_AMOUNT) TX_AMOUNT, SUM(UG.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "

      WhereStr = ""
   ElseIf Ind = 16 Then
      SelectStr1 = "SELECT UG.PART_ITEM_ID, UG.TX_TYPE, SUM(UG.TX_AMOUNT) TX_AMOUNT, SUM(UG.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "

      WhereStr = ""
   ElseIf Ind = 17 Then
      SelectStr1 = "SELECT UG.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, SUM(UG.TX_AMOUNT) TX_AMOUNT, SUM(UG.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN JOB JB ON (IVD.BILLING_DOC_ID = JB.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "

      WhereStr = ""
   ElseIf Ind = 18 Then
      SelectStr1 = "SELECT UG.PART_ITEM_ID, UG.TX_TYPE, IVD.DOCUMENT_TYPE, SUM(UG.TX_AMOUNT) TX_AMOUNT, SUM(UG.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "

      WhereStr = ""
   ElseIf Ind = 19 Then
      SelectStr1 = "SELECT UG.PART_ITEM_ID PART_ITEM_ID, PI.PART_NO PART_NO, PI.PART_DESC PART_DESC,  JB.PART_ITEM_ID PRODUCT_ID, PI2.PART_NO PRODUCT_NO, PI2.PART_DESC PRODUCT_NAME, SUM(UG.TX_AMOUNT) TX_AMOUNT, SUM(UG.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN JOB JB ON (IVD.BILLING_DOC_ID = JB.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI2 ON (JB.PART_ITEM_ID = PI2.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "

      WhereStr = ""
   ElseIf Ind = 20 Then
      SelectStr1 = "SELECT UG.LOCATION_ID, UG.PART_ITEM_ID, UG.TX_TYPE, IVD.DOCUMENT_TYPE, SUM(UG.TX_AMOUNT) TX_AMOUNT, SUM(UG.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "

      WhereStr = ""
   ElseIf Ind = 21 Then
      SelectStr1 = "SELECT UG.*, PI.PART_NO, PI.PART_DESC, MR1.KEY_NAME EXPENSE_NAME, LT.LAY_OUT_NAME DEP_NAME, LT.LAY_OUT_NO DEP_NO, IVD.DOCUMENT_DATE, IVD.DOCUMENT_NO, UN.UNIT_NAME  "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME ITEM_DESC_NAME "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (UG.EXPENSE_TYPE = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LAY_OUT LT ON (UG.TO_DEPARTMENT = LT.LAY_OUT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
   
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (MR1.KEY_ID = UG.ITEM_DESC_ID) "
      
      WhereStr = ""
   ElseIf Ind = 22 Then
      SelectStr1 = "SELECT DISTINCT PI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, PI.PART_TYPE, PT.PART_TYPE_NO, PT.PART_TYPE_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      
      WhereStr = "WHERE (UG.LOCATION_ID = " & LOCATION_ID & ") "
   ElseIf Ind = 23 Then
      SelectStr1 = "SELECT UG.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, UT.UNIT_NAME, UG.EXPENSE_TYPE, UG.TO_DEPARTMENT, MR1.KEY_NAME EXPENSE_NAME, LT.LAY_OUT_NAME DEP_NAME, SUM(UG.TX_AMOUNT) TX_AMOUNT, SUM(UG.TOTAL_INCLUDE_PRICE)  TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (UG.EXPENSE_TYPE = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LAY_OUT LT ON (UG.TO_DEPARTMENT = LT.LAY_OUT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UT ON (PI.UNIT_COUNT = UT.UNIT_ID) "

      WhereStr = ""
   ElseIf Ind = 24 Then
      SelectStr1 = "SELECT UG.PART_ITEM_ID, UG.TX_TYPE, SUM(UG.TX_AMOUNT) TX_AMOUNT, SUM(UG.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr1 = SelectStr1 & ",CM.CUSTOMER_TYPE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (BD.BILLING_DOC_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (BD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      
      WhereStr = " WHERE (IVD.DOCUMENT_TYPE IN ( 10,21 ))  "
   ElseIf Ind = 25 Then
      SelectStr1 = "SELECT DISTINCT CM.CUSTOMER_TYPE, CT.CSTTYPE_NAME "
      SelectStr2 = "SELECT COUNT(DISTINCT CM.CUSTOMER_TYPE)  "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (BD.BILLING_DOC_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (BD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_TYPE CT ON (CM.CUSTOMER_TYPE = CT.CSTTYPE_ID) "
      
      WhereStr = " WHERE (IVD.DOCUMENT_TYPE IN ( 10,21 ) AND TX_AMOUNT > 0 )  "
   ElseIf Ind = 26 Then '  เช็ค ว่า PO หรือเอกสาร มี การ อ้างอิง หรือ ยังเพื่อไม่ให้ลบหรือแก้ไข
      SelectStr1 = "SELECT UG.*  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      
      WhereStr = " WHERE (PO_ID  =   " & PO_ID & " ) "
   ElseIf Ind = 100 Then
      SelectStr1 = "SELECT UG.DO_ID, SUM(UG.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = SP.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = SN.NAME_ID) "
      
      WhereStr = "WHERE (IVD.SUPPLIER_ID IS NOT NULL) "
   ElseIf Ind = 101 Then
      SelectStr1 = "SELECT IVD.SUPPLIER_ID, SUM(UG.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (IVD.BILLING_DOC_ID = UG.DO_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = SP.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (NM.NAME_ID = SN.NAME_ID) "
      
      WhereStr = "WHERE (IVD.SUPPLIER_ID IS NOT NULL) "
   ElseIf Ind = 102 Then        'frmExportSupItem
      SelectStr1 = "SELECT UG.*, IVD.* "
      SelectStr1 = SelectStr1 & ",SP.SUPPLIER_CODE "
      SelectStr1 = SelectStr1 & ",PI.PART_NO "
      SelectStr1 = SelectStr1 & ",LC.LOCATION_NO "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      
      WhereStr = "WHERE (IVD.SUPPLIER_ID IS NOT NULL) "
   ElseIf Ind = 103 Then
      SelectStr1 = "SELECT IVD.SUPPLIER_ID, IVD.DOCUMENT_TYPE, IVD.DOCUMENT_SUBTYPE, IVD.RECEIPT_TYPE, SUM(UG.TOTAL_INCLUDE_PRICE) TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      
      WhereStr = "WHERE (IVD.SUPPLIER_ID IS NOT NULL) "
   ElseIf Ind = 104 Then               'GetUpDateFromPo
      SelectStr1 = "SELECT SUM(UG.TX_AMOUNT) TX_AMOUNT,UG.PART_ITEM_ID  "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM SUP_ITEM UG "
      
      WhereStr = ""
      WhereStr = " WHERE (PO_ID  =   " & PO_ID & " ) "
   ElseIf Ind = 105 Then      'CReportAp013
      SelectStr1 = "SELECT IVD.SUPPLIER_ID,IVD.BILLING_DOC_ID,IVD.DOCUMENT_DATE,IVD.DOCUMENT_NO "
      SelectStr1 = SelectStr1 & ",SP.SUPPLIER_CODE, NM1.LONG_NAME SUPPLIER_NAME "
      SelectStr1 = SelectStr1 & ",PI.PART_ITEM_ID,PI.PART_NO, PI.PART_DESC "
      SelectStr1 = SelectStr1 & ",UG.TX_AMOUNT "
      SelectStr1 = SelectStr1 & ",UN.UNIT_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (SN.NAME_ID = NM1.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
   ElseIf Ind = 106 Then      'LoadSupItemByRo
      SelectStr1 = "SELECT IVD.SUPPLIER_ID,UG.PO_ID "
      SelectStr1 = SelectStr1 & ",UG.PART_ITEM_ID "
      SelectStr1 = SelectStr1 & ",SUM(UG.TX_AMOUNT) TX_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD.SUPPLIER_ID) "
     
   ElseIf Ind = 107 Then      'CReportAp016
      SelectStr1 = "SELECT IVD.SUPPLIER_ID,IVD.DOCUMENT_DATE,IVD.DOCUMENT_NO,IVD.DUE_DATE "
      SelectStr1 = SelectStr1 & ",SP.SUPPLIER_CODE "
      SelectStr1 = SelectStr1 & ",PI.PART_ITEM_ID,PI.PART_NO, PI.PART_DESC "
      SelectStr1 = SelectStr1 & ",UG.TX_AMOUNT,UG.DISCOUNT_AMT "
      SelectStr1 = SelectStr1 & ",UN.UNIT_NAME "
      SelectStr1 = SelectStr1 & ",UG.TOTAL_INCLUDE_PRICE,IVD.TOTAL_PRICE,IVD.DISCOUNT_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      
      
    ElseIf Ind = 108 Then      'CReportAp017 'ยกเลิกใช้งานเนื่องจากช้ามาก
      SelectStr1 = "SELECT IVD.CLOSE_FLAG,IVD.SUPPLIER_ID,IVD.BILLING_DOC_ID,IVD.DOCUMENT_DATE,IVD.DOCUMENT_NO,IVD.NOTE "
      SelectStr1 = SelectStr1 & ",SP.SUPPLIER_CODE, NM1.LONG_NAME SUPPLIER_NAME "
      SelectStr1 = SelectStr1 & ",PI.PART_ITEM_ID,PI.PART_NO, PI.PART_DESC "
      SelectStr1 = SelectStr1 & ",UG.TX_AMOUNT ,UG.ACTUAL_UNIT_PRICE   "
      SelectStr1 = SelectStr1 & ",UN.UNIT_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (SN.NAME_ID = NM1.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      
      
      WhereStr = "WHERE ((IVD.TOTAL_AMOUNT > (SELECT SUM(SUP.TX_AMOUNT) FROM SUP_ITEM SUP LEFT OUTER JOIN BILLING_DOC BDD ON (SUP.DO_ID = BDD.BILLING_DOC_ID) "
      WhereStr = WhereStr & " WHERE SUP.PO_ID = IVD.BILLING_DOC_ID AND BDD.DOCUMENT_DATE <= '" & DateToStringIntHi(DateAdd("D", -1, TO_DATE)) & "')) "    'ลดใบ RO 1 วันเพื่อแสดงให้เห็นว่าวันนี้รับเข้าแล้วค้างเป็น 0 นะ ครับ
      WhereStr = WhereStr & " OR ((SELECT COUNT(*) FROM SUP_ITEM SUP LEFT OUTER JOIN BILLING_DOC BDD ON (SUP.DO_ID = BDD.BILLING_DOC_ID) WHERE SUP.PO_ID = IVD.BILLING_DOC_ID  AND BDD.DOCUMENT_DATE <= '" & DateToStringIntHi(DateAdd("D", -1, TO_DATE)) & "') = 0 )) "
         
   ElseIf Ind = 109 Then      'CReportAp018
      SelectStr1 = "SELECT IVD.SUPPLIER_ID,IVD.BILLING_DOC_ID,IVD.DOCUMENT_DATE,IVD.DOCUMENT_NO ,IVD.DUE_DATE,IVD.NOTE "
      SelectStr1 = SelectStr1 & ",SP.SUPPLIER_CODE, NM1.LONG_NAME SUPPLIER_NAME,IVD.CREDIT "
      SelectStr1 = SelectStr1 & ",PI.PART_ITEM_ID,PI.PART_NO, PI.PART_DESC "
      SelectStr1 = SelectStr1 & ",UG.TX_AMOUNT ,UG.ACTUAL_UNIT_PRICE   "
      SelectStr1 = SelectStr1 & ",UN.UNIT_NAME "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME CONDITION_NAME,   MR3.KEY_NAME PAID_TYPE_NAME  "
      
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (SN.NAME_ID = NM1.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & " LEFT OUTER JOIN MASTER_REF MR2 ON (IVD.CONDITION = MR2.KEY_ID)"
      FromStr = FromStr & " LEFT OUTER JOIN MASTER_REF MR3 ON (IVD.PAID_TYPE = MR3.KEY_ID)"
      
    ElseIf Ind = 110 Then      'LoadSupItemComeIn
      SelectStr1 = "SELECT IVD.SUPPLIER_ID,UG.PO_ID "
      SelectStr1 = SelectStr1 & ",UG.PART_ITEM_ID "
      SelectStr1 = SelectStr1 & ",SUM(UG.TX_AMOUNT) TX_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(UG.TOTAL_WEIGHT) TOTAL_WEIGHT "
      SelectStr1 = SelectStr1 & ",SUM(UG.WEIGHT_AMOUNT) WEIGHT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(UG.SUPPLIER_WEIGHT) SUPPLIER_WEIGHT "
      SelectStr1 = SelectStr1 & ",SUM(UG.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      
      WhereStr = "WHERE (UG.PO_ID IS NOT NULL) "
   ElseIf Ind = 111 Then      'LoadSupItemByRo
      SelectStr1 = "SELECT UG.PO_ID "
      SelectStr1 = SelectStr1 & ",UG.PART_ITEM_ID "
      SelectStr1 = SelectStr1 & ",SUM(UG.TX_AMOUNT) TX_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
    ElseIf Ind = 112 Then      'CReportAp017,CReportAp019
      SelectStr1 = "SELECT IVD.CLOSE_FLAG,IVD.SUPPLIER_ID,IVD.BILLING_DOC_ID,IVD.DOCUMENT_DATE,IVD.DOCUMENT_NO,IVD.NOTE "
      SelectStr1 = SelectStr1 & ",SP.SUPPLIER_CODE, NM1.LONG_NAME SUPPLIER_NAME "
      SelectStr1 = SelectStr1 & ",PI.PART_ITEM_ID,PI.PART_NO, PI.PART_DESC "
      SelectStr1 = SelectStr1 & ",UG.TX_AMOUNT ,UG.ACTUAL_UNIT_PRICE,UG.CALCULATE_TYPE   "
      SelectStr1 = SelectStr1 & ",UN.UNIT_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (SN.NAME_ID = NM1.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
   ElseIf Ind = 113 Then
       SelectStr1 = "SELECT UG.DO_ID,UG.PO_ID ,UG.ACTUAL_UNIT_PRICE,UG.INCLUDE_UNIT_PRICE,UG.TX_AMOUNT, "
      SelectStr1 = SelectStr1 & " UG.CALCULATE_TYPE,UG.TOTAL_WEIGHT,UG.WEIGHT_AMOUNT,UG.SUPPLIER_WEIGHT,UG.TOTAL_ACTUAL_PRICE,UG.ACTUAL_PRICE,  " '
      SelectStr1 = SelectStr1 & " IVD.DOCUMENT_DATE,IVD.DOCUMENT_NO,IVD.BILLING_DOC_ID, "
      SelectStr1 = SelectStr1 & "PI.PART_DESC,PI.PART_NO "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM SUP_ITEM UG LEFT OUTER JOIN BILLING_DOC IVD ON(UG.DO_ID=IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON(UG.PART_ITEM_ID=PI.PART_ITEM_ID) "
   ElseIf Ind = 114 Then
      SelectStr1 = "SELECT IVD2.DOCUMENT_NO AS PO_DOCUMENT_NO,IVD.SUPPLIER_ID,IVD.DOCUMENT_DATE "
      SelectStr1 = SelectStr1 & ",IVD.DOCUMENT_NO AS RO_DOCUMENT_NO,IVD.DUE_DATE,SP.SUPPLIER_CODE,NM1.LONG_NAME "
      SelectStr1 = SelectStr1 & ",PI.PART_ITEM_ID,PI.PART_NO, PI.PART_DESC "
      SelectStr1 = SelectStr1 & ",UG.CALCULATE_TYPE,UG.SUPPLIER_WEIGHT,UG.WEIGHT_AMOUNT,UG.TOTAL_WEIGHT "
      SelectStr1 = SelectStr1 & ",UG.TX_AMOUNT,UG.DISCOUNT_AMT,UG.ACTUAL_UNIT_PRICE "
      SelectStr1 = SelectStr1 & ",UN.UNIT_NAME "
      SelectStr1 = SelectStr1 & ",UG.TOTAL_INCLUDE_PRICE,IVD.TOTAL_PRICE,IVD.DISCOUNT_AMOUNT "
      SelectStr1 = SelectStr1 & ",LC.LOCATION_NAME,PG.PACKAGING_NAME "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD2 ON (UG.PO_ID = IVD2.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PACKAGING PG ON (UG.PACKAGING_ID = PG.PACKAGING_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (SN.NAME_ID = NM1.NAME_ID) "
   ElseIf Ind = 115 Then
      SelectStr1 = "SELECT IVD.DOCUMENT_DATE "
      SelectStr1 = SelectStr1 & ",IVD.DOCUMENT_NO,IVD.SUPPLIER_ID,SP.SUPPLIER_CODE,NM1.LONG_NAME "
      SelectStr1 = SelectStr1 & ",PI.PART_NO, PI.PART_DESC "
      SelectStr1 = SelectStr1 & ",UG.TX_AMOUNT,UG.DISCOUNT_AMT,UG.ACTUAL_UNIT_PRICE "
      SelectStr1 = SelectStr1 & ",UG.TOTAL_INCLUDE_PRICE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (SN.NAME_ID = NM1.NAME_ID) "
   ElseIf Ind = 116 Then
       SelectStr1 = "SELECT IVD.DOCUMENT_DATE  " ',IVD.DOCUMENT_NO
      SelectStr1 = SelectStr1 & ",UG.PART_ITEM_ID,UG.CALCULATE_TYPE  "
      SelectStr1 = SelectStr1 & ",SUM(UG.TX_AMOUNT) TX_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(UG.TOTAL_WEIGHT) TOTAL_WEIGHT "
      SelectStr1 = SelectStr1 & ",SUM(UG.WEIGHT_AMOUNT) WEIGHT_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(UG.SUPPLIER_WEIGHT) SUPPLIER_WEIGHT "
      SelectStr1 = SelectStr1 & ",SUM(UG.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM SUP_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.DO_ID = IVD.BILLING_DOC_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      
      WhereStr = ""
   End If
   
   If PO_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PO_ID = '" & PO_ID & "') "
      Else
         WhereStr = WhereStr & "AND (PO_ID = '" & PO_ID & "') "
      End If
   End If
   
   If SUP_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUP_ITEM_ID = " & SUP_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SUP_ITEM_ID = " & SUP_ITEM_ID & ") "
      End If
   End If
   
   If DO_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UG.DO_ID = " & DO_ID & ") "
      Else
         WhereStr = WhereStr & "AND (UG.DO_ID = " & DO_ID & ") "
      End If
   End If
   
   If LOCATION_ID > 0 And (Ind <> 22) And (Ind <> 1) Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LOCATION_ID = " & LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LOCATION_ID = " & LOCATION_ID & ") "
      End If
   End If
   
   If LOCATION_ID > 0 And ((Ind = 1) Or (Ind = 22)) Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UG.LOCATION_ID = " & LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (UG.LOCATION_ID = " & LOCATION_ID & ") "
      End If
   End If
   
   If FROM_LOCATION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UG.LOCATION_ID = " & FROM_LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (UG.LOCATION_ID = " & FROM_LOCATION_ID & ") "
      End If
   End If
   
   If TO_LOCATION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LI2.LOCATION_ID = " & TO_LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LI2.LOCATION_ID = " & TO_LOCATION_ID & ") "
      End If
   End If
   
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   If Len(COMMIT_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
   End If
   
   If Len(CLOSE_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CLOSE_FLAG = '" & CLOSE_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (CLOSE_FLAG = '" & CLOSE_FLAG & "') "
      End If
   End If

  If Len(AUTO_GEN_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (AUTO_GEN_FLAG = '" & AUTO_GEN_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (AUTO_GEN_FLAG = '" & AUTO_GEN_FLAG & "') "
      End If
   End If
   
     If Len(PO_APPROVED_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PO_APPROVED_FLAG = '" & PO_APPROVED_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (PO_APPROVED_FLAG = '" & PO_APPROVED_FLAG & "') "
      End If
   End If
   
   
   If Ind = 7 Then
      If Len(TX_TYPE) > 0 Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (LI.TX_TYPE = '" & TX_TYPE & "') "
         Else
            WhereStr = WhereStr & "AND (LI.TX_TYPE = '" & TX_TYPE & "') "
         End If
      End If
   Else
      If Len(TX_TYPE) > 0 Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (UG.TX_TYPE = '" & TX_TYPE & "') "
         Else
            WhereStr = WhereStr & "AND (UG.TX_TYPE = '" & TX_TYPE & "') "
         End If
      End If
   End If
   If Len(FROM_TX_TYPE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UG.TX_TYPE = '" & FROM_TX_TYPE & "') "
      Else
         WhereStr = WhereStr & "AND (UG.TX_TYPE = '" & FROM_TX_TYPE & "') "
      End If
   End If
   
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (IVD.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (IVD.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      End If
   End If
   
   If Len(DOCUMENT_TYPE_SET) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (IVD.DOCUMENT_TYPE IN " & DOCUMENT_TYPE_SET & ") "
      Else
         WhereStr = WhereStr & "AND (IVD.DOCUMENT_TYPE IN " & DOCUMENT_TYPE_SET & ") "
      End If
   End If
   
   If PART_GROUP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PT.PART_GROUP_ID = " & PART_GROUP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PT.PART_GROUP_ID = " & PART_GROUP_ID & ") "
      End If
   End If
      
   If PART_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_TYPE = " & PART_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (PART_TYPE = " & PART_TYPE & ") "
      End If
   End If
   
   If PROCESS_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (JB.PROCESS_ID = " & PROCESS_ID & ") "
      Else
         WhereStr = WhereStr & "AND (JB.PROCESS_ID = " & PROCESS_ID & ") "
      End If
   End If
   
    If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (IVD.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (IVD.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (IVD.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (IVD.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If Trim(PART_NO) <> "" Then
      If WildCard(PART_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PI.PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (PI.PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')"
         End If
      End If
   End If
   
   If Trim(SUPPLIER_CODE) <> "" Then
      Call AddSearchSupplierCode(FromStr)
      If WildCard(SUPPLIER_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(SUPPLIER_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(SUPPLIER_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (SUPPLIER_CODE = '" & ChangeQuote(Trim(SUPPLIER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUPPLIER_CODE = '" & ChangeQuote(Trim(SUPPLIER_CODE)) & "')"
         End If
      End If
   End If
   
   If Trim(PART_DESC) <> "" Then
      If WildCard(PART_DESC, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = GenerateSearchLike("WHERE", "PART_DESC", SubLen, NewStr)
         Else
            WhereStr = WhereStr & GenerateSearchLike("AND", "PART_DESC", SubLen, NewStr)
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PART_DESC = '" & ChangeQuote(Trim(PART_DESC)) & "')"
         Else
            WhereStr = WhereStr & "AND (PART_DESC = '" & ChangeQuote(Trim(PART_DESC)) & "')"
         End If
      End If
   End If
   
   If SUPPLIER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUPPLIER_ID = " & SUPPLIER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SUPPLIER_ID = " & SUPPLIER_ID & ") "
      End If
   End If
   
   If SUPPLIER_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SP.SUPPLIER_TYPE = " & SUPPLIER_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (SP.SUPPLIER_TYPE = " & SUPPLIER_TYPE & ") "
      End If
   End If
   
   If GUI_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (GUI_ID = " & GUI_ID & ") "
      Else
         WhereStr = WhereStr & "AND (GUI_ID = " & GUI_ID & ") "
      End If
   End If
   
   If LAYOUT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UG.TO_DEPARTMENT = " & LAYOUT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (UG.TO_DEPARTMENT = " & LAYOUT_ID & ") "
      End If
   End If
   
   GroupByStr = ""
   If Ind = 3 Then
      GroupByStr = "GROUP BY LI.PART_ITEM_ID "
   ElseIf Ind = 7 Then
      GroupByStr = "GROUP BY IVD.SUPPLIER_ID "
   ElseIf Ind = 8 Then
      GroupByStr = "GROUP BY LI.PART_ITEM_ID, LI.LOCATION_ID, LI.LAYOUT_ID "
   ElseIf Ind = 12 Then
      GroupByStr = "GROUP BY UG.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, UG.LOCATION_ID "
   ElseIf Ind = 13 Then
      GroupByStr = "GROUP BY IVD.SUPPLIER_ID "
   ElseIf Ind = 15 Then
      GroupByStr = "GROUP BY UG.PART_ITEM_ID, UG.LOCATION_ID, UG.TX_TYPE "
   ElseIf Ind = 16 Then
      GroupByStr = "GROUP BY UG.PART_ITEM_ID, UG.TX_TYPE "
   ElseIf Ind = 17 Then
      GroupByStr = "GROUP BY UG.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
   ElseIf Ind = 18 Then
      GroupByStr = "GROUP BY UG.PART_ITEM_ID, UG.TX_TYPE, IVD.DOCUMENT_TYPE "
   ElseIf Ind = 19 Then
      GroupByStr = "GROUP BY UG.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, JB.PART_ITEM_ID, PI2.PART_NO, PI2.PART_DESC "
   ElseIf Ind = 20 Then
      GroupByStr = "GROUP BY UG.LOCATION_ID, UG.PART_ITEM_ID, UG.TX_TYPE, IVD.DOCUMENT_TYPE "
   ElseIf Ind = 23 Then
      GroupByStr = "GROUP BY UG.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, UT.UNIT_NAME, UG.EXPENSE_TYPE, UG.TO_DEPARTMENT, MR1.KEY_NAME, LT.LAY_OUT_NAME "
   ElseIf Ind = 24 Then
      GroupByStr = "GROUP BY UG.PART_ITEM_ID, UG.TX_TYPE, IVD.DOCUMENT_TYPE,CM.CUSTOMER_TYPE "
   ElseIf Ind = 100 Then
      GroupByStr = "GROUP BY UG.DO_ID "
   ElseIf Ind = 101 Then
      GroupByStr = "GROUP BY IVD.SUPPLIER_ID "
   ElseIf Ind = 103 Then
      GroupByStr = "GROUP BY IVD.SUPPLIER_ID, IVD.DOCUMENT_TYPE, IVD.DOCUMENT_SUBTYPE, IVD.RECEIPT_TYPE "
   ElseIf Ind = 104 Then
      GroupByStr = "GROUP BY UG.PART_ITEM_ID "
   ElseIf Ind = 106 Then
      GroupByStr = "GROUP BY IVD.SUPPLIER_ID,UG.PO_ID,UG.PART_ITEM_ID "
   ElseIf Ind = 110 Then
      GroupByStr = "GROUP BY IVD.SUPPLIER_ID,UG.PO_ID,UG.PART_ITEM_ID "
   ElseIf Ind = 111 Then
      GroupByStr = "GROUP BY UG.PO_ID,UG.PART_ITEM_ID "
   ElseIf Ind = 116 Then
      GroupByStr = "GROUP BY IVD.DOCUMENT_DATE,UG.PART_ITEM_ID,UG.CALCULATE_TYPE " ',IVD.DOCUMENT_NO
   End If

   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY SUP_ITEM_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY SUP_ITEM_ID "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY DOCUMENT_DATE DESC, TRANSACTION_SEQ "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY LOCATION_ID DESC, DOCUMENT_DATE DESC, TRANSACTION_SEQ "
      ElseIf OrderBy = 5 Then
         OrderByStr = "ORDER BY PT.PART_TYPE_NO "
      ElseIf OrderBy = 6 Then
         OrderByStr = "ORDER BY PI.PART_NO ASC, IVD.DOCUMENT_NO "
      ElseIf OrderBy = 7 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE "
      ElseIf OrderBy = 8 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_NO "
      ElseIf OrderBy = 9 Then
         OrderByStr = "ORDER BY LC.LOCATION_NO "
      ElseIf OrderBy = 10 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE ASC, UG.TRANSACTION_SEQ  "
      ElseIf OrderBy = 11 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE ASC, DP.PRIORITY1 ASC, IVD.DOCUMENT_NO ASC, UG.TRANSACTION_SEQ  "
      ElseIf OrderBy = 12 Then
         OrderByStr = "ORDER BY SP.SUPPLIER_CODE ASC, IVD.DOCUMENT_NO "
      ElseIf OrderBy = 13 Then
         OrderByStr = "ORDER BY SN.LONG_NAME "
      ElseIf OrderBy = 14 Then
         OrderByStr = "ORDER BY PT.PART_TYPE_NO ASC, PI.PART_NO ASC, PI.PART_ITEM_ID, IVD.DOCUMENT_DATE ASC, DP.PRIORITY1 ASC, IVD.DOCUMENT_NO ASC, UG.TRANSACTION_SEQ  "
      ElseIf OrderBy = 15 Then
         OrderByStr = "ORDER BY PI.PART_NO ASC, IVD.DOCUMENT_DATE ASC, IVD.DOCUMENT_NO "
      ElseIf OrderBy = 16 Then
         OrderByStr = "ORDER BY PI.PART_DESC ASC, PI.PART_NO ASC, IVD.DOCUMENT_DATE ASC, IVD.DOCUMENT_NO "
      Else
         OrderByStr = "ORDER BY SUP_ITEM_ID "
      End If
   ElseIf Ind = 2 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PART_TYPE_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PART_NO "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY DOCUMENT_NO "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY DOCUMENT_DATE "
      Else
         OrderByStr = "ORDER BY SUP_ITEM_ID "
      End If
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY LI.PART_ITEM_ID "
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY PART_ITEM_ID, LOCATION_ID "
   ElseIf Ind = 5 Then
      OrderByStr = "ORDER BY PI.PART_TYPE "
   ElseIf Ind = 6 Then
      OrderByStr = "ORDER BY PT.PART_TYPE_NAME ASC, PI.PART_NO "
   ElseIf Ind = 7 Then
      OrderByStr = "ORDER BY IVD.SUPPLIER_ID "
   ElseIf Ind = 8 Then
      OrderByStr = "ORDER BY LI.LAYOUT_ID "
   ElseIf Ind = 9 Then
      OrderByStr = "ORDER BY PART_ITEM_ID "
   ElseIf Ind = 10 Then
      OrderByStr = "ORDER BY PI.PART_NO "
   ElseIf Ind = 11 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY SUP_ITEM_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY SUP_ITEM_ID "
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
      ElseIf OrderBy = 10 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE ASC, UG.TRANSACTION_SEQ  "
      ElseIf OrderBy = 11 Then
         OrderByStr = "ORDER BY UG.TRANSACTION_SEQ  "
      Else
         OrderByStr = "ORDER BY SUP_ITEM_ID "
      End If
   ElseIf Ind = 12 Then
      OrderByStr = "ORDER BY PI.PART_NO "
   ElseIf Ind = 13 Then
      OrderByStr = "ORDER BY SP.SUPPLIER_CODE "
   ElseIf Ind = 14 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY LC.LOCATION_NAME ASC, PI.PART_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PI.PART_NO ASC, LC.LOCATION_NAME "
      End If
   ElseIf Ind = 15 Then
      OrderByStr = "ORDER BY UG.LOCATION_ID, UG.PART_ITEM_ID "
   ElseIf Ind = 16 Then
      OrderByStr = "ORDER BY UG.PART_ITEM_ID "
   ElseIf Ind = 17 Then
      OrderByStr = "ORDER BY PI.PART_NO "
   ElseIf Ind = 18 Then
      OrderByStr = "ORDER BY UG.PART_ITEM_ID "
   ElseIf Ind = 19 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PI2.PART_NO ASC, PI.PART_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PI.PART_NO ASC, PI2.PART_NO "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY PI2.PART_NO "
      End If
   ElseIf Ind = 20 Then
      OrderByStr = "ORDER BY UG.PART_ITEM_ID "
   ElseIf Ind = 21 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY LT.LAY_OUT_NAME ASC, MR1.KEY_NAME "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE ASC, IVD.DOCUMENT_NO ASC, LT.LAY_OUT_NAME ASC, MR1.KEY_NAME "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY PI.PART_NO ASC, IVD.DOCUMENT_DATE ASC, IVD.DOCUMENT_NO "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY LT.LAY_OUT_NAME ASC, MR1.KEY_NAME ASC, PI.PART_NO ASC, IVD.DOCUMENT_DATE ASC, IVD.DOCUMENT_NO "
      Else
         OrderByStr = "ORDER BY MR1.KEY_NAME ASC, LT.LAY_OUT_NAME "
      End If
   ElseIf Ind = 22 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PT.PART_TYPE_NO, PI.PART_NO ASC, PI.PART_ITEM_ID "
      Else
         OrderByStr = "ORDER BY PT.PART_TYPE_NO ASC, PI.PART_NO ASC, PI.PART_ITEM_ID  "
      End If
   ElseIf Ind = 23 Then
      OrderByStr = "ORDER BY LT.LAY_OUT_NAME ASC, MR1.KEY_NAME ASC, PI.PART_NO "
   ElseIf Ind = 100 Then
      OrderByStr = "ORDER BY UG.DO_ID "
   ElseIf Ind = 101 Then
      OrderByStr = "ORDER BY IVD.SUPPLIER_ID "
   ElseIf Ind = 102 Then
      OrderByStr = "ORDER BY IVD.DOCUMENT_DATE,IVD.DOCUMENT_NO "
   ElseIf Ind = 103 Then
      OrderByStr = "ORDER BY IVD.SUPPLIER_ID "
   ElseIf Ind = 105 Then
      OrderByStr = "ORDER BY SP.SUPPLIER_CODE,IVD.DOCUMENT_DATE,IVD.DOCUMENT_NO,PI.PART_NO "
   ElseIf Ind = 107 Then
      OrderByStr = "ORDER BY IVD.DOCUMENT_DATE ASC,IVD.DOCUMENT_NO ASC,PI.PART_NO "
   ElseIf Ind = 108 Then
      OrderByStr = "ORDER BY PI.PART_NO,IVD.DOCUMENT_DATE,IVD.DOCUMENT_NO ,SP.SUPPLIER_CODE    "
   ElseIf Ind = 109 Or Ind = 112 Then
      Select Case OrderBy
      Case 1
         OrderByStr = "ORDER BY SP.SUPPLIER_CODE   "
      Case 2
         OrderByStr = "ORDER BY SUPPLIER_NAME   "
      Case 3
         OrderByStr = "ORDER BY IVD.DOCUMENT_NO  "
      Case 4
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE  "
      Case 5
         OrderByStr = "ORDER BY PI.PART_NO  "
      Case 6
         OrderByStr = "ORDER BY PI.PART_DESC  "
      Case 7
         OrderByStr = "ORDER BY PI.PART_NO  ASC,IVD.DOCUMENT_NO "
      Case 8
         OrderByStr = "ORDER BY SP.SUPPLIER_CODE,IVD.DOCUMENT_NO "
      End Select
    ElseIf Ind = 113 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE  "
      End If
    ElseIf Ind = 114 Then
        OrderByStr = "ORDER BY SP.SUPPLIER_CODE ASC,IVD.DOCUMENT_DATE,IVD.DOCUMENT_NO "
   ElseIf Ind = 115 Then
        OrderByStr = "ORDER BY SP.SUPPLIER_CODE ASC,IVD.DOCUMENT_DATE,IVD.DOCUMENT_NO "
   End If
   
   If Not (Ind = 7 Or Ind = 24 Or Ind = 25 Or Ind = 26 Or Ind = 104 Or Ind = 106 Or Ind = 110) Then
      If OrderType = 1 Then
         OrderTypeStr = "ASC "
      ElseIf OrderType = 2 Then
         OrderTypeStr = "DESC "
      Else
         OrderTypeStr = "ASC "
      End If
   End If
   
   If Len(OrderByStr) <= 0 Then
      OrderTypeStr = ""
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      SUP_ITEM_ID = NVLI(Rs("SUP_ITEM_ID"), -1)
      DO_ID = NVLI(Rs("DO_ID"), -1)
      NOTE = NVLS(Rs("NOTE"), "")
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      ACTUAL_UNIT_PRICE = NVLD(Rs("ACTUAL_UNIT_PRICE"), 0)
      INCLUDE_UNIT_PRICE = NVLD(Rs("INCLUDE_UNIT_PRICE"), 0)
      PREVIOUS_PRICE = NVLD(Rs("PREVIOUS_PRICE"), 0)
      PREVIOUS_AMOUNT = NVLD(Rs("PREVIOUS_AMOUNT"), 0)
      NEW_PRICE = NVLD(Rs("NEW_PRICE"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      NEW_AMOUNT = NVLD(Rs("NEW_AMOUNT"), 0)
      TRANSACTION_SEQ = NVLI(Rs("TRANSACTION_SEQ"), 0)
      GUI_ID = NVLI(Rs("GUI_ID"), -1)
      CALCULATE_FLAG = NVLS(Rs("CALCULATE_FLAG"), "")
      TOTAL_ACTUAL_PRICE = NVLD(Rs("TOTAL_ACTUAL_PRICE"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      LEFT_AMOUNT = NVLD(Rs("LEFT_AMOUNT"), 0)
      LAYOUT_ID = NVLI(Rs("LAYOUT_ID"), -1)
      LINK_ID = NVLI(Rs("LINK_ID"), -1)
      PACKAGING_AMT = NVLD(Rs("PACKAGING_AMT"), 0)
      ENTRY_WEIGHT = NVLD(Rs("ENTRY_WEIGHT"), 0)
      EXIT_WEIGHT = NVLD(Rs("EXIT_WEIGHT"), 0)
      WEIGHT_AMOUNT = NVLD(Rs("WEIGHT_AMOUNT"), 0)
      PACKAGE_WEIGHT = NVLD(Rs("PACKAGE_WEIGHT"), 0)
      OTHER_WEIGHT = NVLD(Rs("OTHER_WEIGHT"), 0)
      PERCENT_HUMID = NVLD(Rs("PERCENT_HUMID"), 0)
      HUMID_WEIGHT = NVLD(Rs("HUMID_WEIGHT"), 0)
      PACKAGING_ID = NVLD(Rs("PACKAGING_ID"), 0)
      SUPPLIER_WEIGHT = NVLD(Rs("SUPPLIER_WEIGHT"), 0)
      CALCULATE_TYPE = NVLI(Rs("CALCULATE_TYPE"), 0)
      PACKAGE_AMOUNT = NVLD(Rs("PACKAGE_AMOUNT"), 0)
      ACTUAL_PKG_AMOUNT = NVLD(Rs("ACTUAL_PKG_AMOUNT"), 0)
      PUREXP_ID1 = NVLI(Rs("PUREXP_ID1"), 0)
      PUREXP_ID2 = NVLI(Rs("PUREXP_ID2"), 0)
      EXPENSE1 = NVLD(Rs("EXPENSE1"), 0)
      EXPENSE2 = NVLD(Rs("EXPENSE2"), 0)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      RAW_PRICE = NVLD(Rs("RAW_PRICE"), 0)
      DISCOUNT_AMT = NVLD(Rs("DISCOUNT_AMT"), 0)
      RAW_TOT_PRICE = NVLD(Rs("RAW_TOT_PRICE"), 0)
      TO_DEPARTMENT = NVLI(Rs("TO_DEPARTMENT"), -1)
      ITEM_DESC = NVLS(Rs("ITEM_DESC"), "")
      EXTRA_NAME = NVLS(Rs("EXTRA_NAME"), "")
      SALE_TOT_PRICE = NVLD(Rs("SALE_TOT_PRICE"), 0)
      CALCULATE_WEIGHT = NVLD(Rs("CALCULATE_WEIGHT"), 0)
      RAW_COST = NVLD(Rs("RAW_COST"), 0)
      EXPENSE_COST = NVLD(Rs("EXPENSE_COST"), 0)
      TOTAL_NEW_PRICE = NVLD(Rs("TOTAL_NEW_PRICE"), 0)
      BAG_RETURN = NVLS(Rs("BAG_RETURN"), "")
      ACTUAL_AMOUNT = NVLD(Rs("ACTUAL_AMOUNT"), 0)
      ACTUAL_PRICE = NVLD(Rs("ACTUAL_PRICE"), 0)
      CURRENT_AMOUNT = NVLD(Rs("CURRENT_AMOUNT"), 0)
      CURRENT_PRICE = NVLD(Rs("CURRENT_PRICE"), 0)
      NEED_TOTAL_PRICE = NVLD(Rs("NEED_TOTAL_PRICE"), 0)
      NEED_TOTAL_AMOUNT = NVLD(Rs("NEED_TOTAL_AMOUNT"), 0)
      NEED_AVG_PRICE = NVLD(Rs("NEED_AVG_PRICE"), 0)
      MANUAL_PRICE = NVLS(Rs("MANUAL_PRICE"), "")
      EXPENSE_TYPE = NVLI(Rs("EXPENSE_TYPE"), 0)
      AUTO_PRICE = NVLS(Rs("AUTO_PRICE"), "")
      ITEM_DESC_ID = NVLI(Rs("ITEM_DESC_ID"), -1)
      ITEM_DESC_NO = NVLS(Rs("ITEM_DESC_NO"), "")
      ITEM_DESC_NAME = NVLS(Rs("ITEM_DESC_NAME"), "")
      PUREXP_NO1 = NVLS(Rs("PUREXP_NO1"), "")
      PUREXP_NO2 = NVLS(Rs("PUREXP_NO2"), "")
      PUREXP_NAME1 = NVLS(Rs("PUREXP_NAME1"), "")
      PUREXP_NAME2 = NVLS(Rs("PUREXP_NAME2"), "")
      EXPENSE_RATE1 = NVLD(Rs("EXPENSE_RATE1"), 0)
      EXPENSE_RATE2 = NVLD(Rs("EXPENSE_RATE2"), 0)
      
      PROJECT_NAME_ID = NVLI(Rs("PROJECT_NAME_ID"), -1)
      
      PO_ID = NVLI(Rs("PO_ID"), -1)
      PO_NO = NVLS(Rs("PO_NO"), "")
      
      PACKAGING_CODE = NVLS(Rs("PACKAGING_NO"), "")
      PACKAGING_CODE = EmptyToString(PACKAGING_CODE, " ")
      
      PRIORITY1 = NVLI(Rs("PRIORITY1"), 0)
      EXPORT_AVG_PRICE = INCLUDE_UNIT_PRICE
      EXPORT_TOTAL_PRICE = TOTAL_INCLUDE_PRICE
      IMPORT_AMOUNT = TX_AMOUNT
      EXPORT_AMOUNT = TX_AMOUNT

      OLD_TX_AMOUNT = TX_AMOUNT
      OLD_LOCATION_ID = LOCATION_ID
      
      PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID1"), -1)
      PART_TYPE = NVLI(Rs("PART_TYPE"), -1)
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
      PART_SUPPLIER_NAME = NVLS(Rs("PART_SUPPLIER_NAME"), "") 'add 2016-03-09
      PART_NOTE = NVLS(Rs("PART_NOTE"), "")
      PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      PIG_FLAG = NVLS(Rs("PIG_FLAG"), "")
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DO_NO = NVLS(Rs("DO_NO"), "")
      SENDER_NAME = NVLS(Rs("SENDER_NAME"), "")
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      RECEIVE_NAME = NVLS(Rs("RECEIVE_NAME"), "")
      QUE_NO = NVLS(Rs("QUE_NO"), "")
      TRUCK_NO = NVLS(Rs("TRUCK_NO"), "")
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), -1)
      SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
      SUPPLIER_NAME = NVLS(Rs("SUPPLIER_NAME"), "")
      
      PR_NO = NVLS(Rs("PR_NO"), "")
      
      DEPARTMENT_NAME = NVLS(Rs("DEPARTMENT_NAME"), "")

   ElseIf Ind = 3 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LEFT_AMOUNT = NVLD(Rs("LEFT_AMOUNT"), 0)
   ElseIf Ind = 4 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      NEW_PRICE = NVLD(Rs("NEW_PRICE"), 0)
      PREVIOUS_PRICE = NVLD(Rs("PREVIOUS_PRICE"), 0)
      NEW_AMOUNT = NVLD(Rs("NEW_AMOUNT"), 0)
      PREVIOUS_AMOUNT = NVLD(Rs("PREVIOUS_AMOUNT"), 0)
   ElseIf Ind = 5 Then
      PART_TYPE = NVLI(Rs("PART_TYPE_ID"), 0)
   ElseIf Ind = 6 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
   ElseIf Ind = 7 Then
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)

'      LAYOUT_ID = NVLI(Rs("LAYOUT_ID"), 0)
'      LEFT_AMOUNT = NVLD(Rs("LEFT_AMOUNT"), 0)
   ElseIf Ind = 9 Then
      SUP_ITEM_ID = NVLI(Rs("SUP_ITEM_ID"), -1)
      DO_ID = NVLI(Rs("DO_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      ACTUAL_UNIT_PRICE = NVLD(Rs("ACTUAL_UNIT_PRICE"), 0)
      INCLUDE_UNIT_PRICE = NVLD(Rs("INCLUDE_UNIT_PRICE"), 0)
      PREVIOUS_PRICE = NVLD(Rs("PREVIOUS_PRICE"), 0)
      PREVIOUS_AMOUNT = NVLD(Rs("PREVIOUS_AMOUNT"), 0)
      NEW_PRICE = NVLD(Rs("NEW_PRICE"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      NEW_AMOUNT = NVLD(Rs("NEW_AMOUNT"), 0)
      TRANSACTION_SEQ = NVLI(Rs("TRANSACTION_SEQ"), 0)
      GUI_ID = NVLI(Rs("GUI_ID"), -1)
      CALCULATE_FLAG = NVLS(Rs("CALCULATE_FLAG"), "")
      TOTAL_ACTUAL_PRICE = NVLD(Rs("TOTAL_ACTUAL_PRICE"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      LEFT_AMOUNT = NVLD(Rs("LEFT_AMOUNT"), 0)
      LAYOUT_ID = NVLI(Rs("LAYOUT_ID"), -1)
      
      OLD_TX_AMOUNT = TX_AMOUNT
      OLD_LOCATION_ID = LOCATION_ID
      
      PART_TYPE = NVLI(Rs("PART_TYPE"), -1)
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
      PART_SUPPLIER_NAME = NVLS(Rs("PART_SUPPLIER_NAME"), "")
      PART_NOTE = NVLS(Rs("PART_NOTE"), "")
      PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      
      USED_AMOUNT = LEFT_AMOUNT
   ElseIf Ind = 10 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
   ElseIf Ind = 11 Then
      SUP_ITEM_ID = NVLI(Rs("SUP_ITEM_ID"), -1)
      DO_ID = NVLI(Rs("DO_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      PACKAGING_AMT = NVLD(Rs("PACKAGING_AMT"), 0)
     
      PART_TYPE = NVLI(Rs("PART_TYPE"), -1)
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
      PART_SUPPLIER_NAME = NVLS(Rs("PART_SUPPLIER_NAME"), "")
      PART_NOTE = NVLS(Rs("PART_NOTE"), "")
      PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
      FROM_LOCATION = NVLS(Rs("FROM_LOCATION"), "")
      TO_LOCATION = NVLS(Rs("TO_LOCATION"), "")
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
   ElseIf Ind = 12 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_ACTUAL_PRICE = NVLD(Rs("TOTAL_ACTUAL_PRICE"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      PACKAGE_AMOUNT = NVLD(Rs("PACKAGE_AMOUNT"), 0)
      ACTUAL_PKG_AMOUNT = NVLD(Rs("ACTUAL_PKG_AMOUNT"), 0)
      WEIGHT_AMOUNT = NVLD(Rs("WEIGHT_AMOUNT"), 0)
      PACKAGE_WEIGHT = NVLD(Rs("PACKAGE_WEIGHT"), 0)
      OTHER_WEIGHT = NVLD(Rs("OTHER_WEIGHT"), 0)
      HUMID_WEIGHT = NVLD(Rs("HUMID_WEIGHT"), 0)
      CALCULATE_WEIGHT = NVLD(Rs("CALCULATE_WEIGHT"), 0)
      PACKAGE_AMOUNT = NVLD(Rs("PACKAGE_AMOUNT"), 0)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      EXPENSE1 = NVLD(Rs("EXPENSE1"), 0)
      EXPENSE2 = NVLD(Rs("EXPENSE2"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
   ElseIf Ind = 13 Then
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), -1)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_ACTUAL_PRICE = NVLD(Rs("TOTAL_ACTUAL_PRICE"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      PACKAGE_AMOUNT = NVLD(Rs("PACKAGE_AMOUNT"), 0)
      WEIGHT_AMOUNT = NVLD(Rs("WEIGHT_AMOUNT"), 0)
      PACKAGE_WEIGHT = NVLD(Rs("PACKAGE_WEIGHT"), 0)
      OTHER_WEIGHT = NVLD(Rs("OTHER_WEIGHT"), 0)
      HUMID_WEIGHT = NVLD(Rs("HUMID_WEIGHT"), 0)
      CALCULATE_WEIGHT = NVLD(Rs("CALCULATE_WEIGHT"), 0)
      PACKAGE_AMOUNT = NVLD(Rs("PACKAGE_AMOUNT"), 0)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      EXPENSE1 = NVLD(Rs("EXPENSE1"), 0)
      EXPENSE2 = NVLD(Rs("EXPENSE2"), 0)
   ElseIf Ind = 14 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
      PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
   ElseIf Ind = 15 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
   ElseIf Ind = 16 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
   ElseIf Ind = 17 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
   ElseIf Ind = 18 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
   ElseIf Ind = 19 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      PRODUCT_ID = NVLI(Rs("PRODUCT_ID"), 0)
      PRODUCT_NO = NVLS(Rs("PRODUCT_NO"), "")
      PRODUCT_NAME = NVLS(Rs("PRODUCT_NAME"), "")
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
   ElseIf Ind = 20 Then
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
   ElseIf Ind = 21 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      DEP_NAME = NVLS(Rs("DEP_NAME"), "")
      DEP_NO = NVLS(Rs("DEP_NO"), "")
      EXPENSE_NAME = NVLS(Rs("EXPENSE_NAME"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      
      ITEM_DESC = NVLS(Rs("ITEM_DESC_NAME"), "") & " " & NVLS(Rs("ITEM_DESC"), "")
   ElseIf Ind = 22 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      PART_TYPE = NVLI(Rs("PART_TYPE"), -1)
      PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
   ElseIf Ind = 23 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
'      PART_TYPE = NVLI(Rs("PART_TYPE"), -1)
'      PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
'      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      DEP_NAME = NVLS(Rs("DEP_NAME"), "")
      EXPENSE_NAME = NVLS(Rs("EXPENSE_NAME"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
   ElseIf Ind = 24 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      CUSTOMER_TYPE = NVLI(Rs("CUSTOMER_TYPE"), 0)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
   ElseIf Ind = 25 Then
      CUSTOMER_TYPE = NVLI(Rs("CUSTOMER_TYPE"), 0)
      CSTTYPE_NAME = NVLS(Rs("CSTTYPE_NAME"), "")
   ElseIf Ind = 100 Then
      DO_ID = NVLI(Rs("DO_ID"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
   ElseIf Ind = 101 Then
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
'      ITEM_AMOUNT = NVLD(Rs("ITEM_AMOUNT"), 0)
   ElseIf Ind = 102 Then
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      
      DUE_DATE = InternalDateToDate(NVLS(Rs("DUE_DATE"), ""))
      NOTE = NVLS(Rs("NOTE"), "")
      BILLING_ADDRESS_ID = NVLI(Rs("BILLING_ADDRESS_ID"), 0)
      ENTERPRISE_ID = NVLI(Rs("ENTERPRISE_ID"), 0)
      ENTERPRISE_ADDRESS_ID = NVLI(Rs("ENTERPRISE_ADDRESS_ID"), 0)
      TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), 0)
      VAT_PERCENT = NVLD(Rs("VAT_PERCENT"), 0)
      VAT_AMOUNT = NVLD(Rs("VAT_AMOUNT"), 0)
      WH_PERCENT = NVLD(Rs("WH_PERCENT"), 0)
      WH_AMOUNT = NVLD(Rs("WH_AMOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      
      PAYMENT_TERM = NVLI(Rs("PAYMENT_TERM"), 0)
      ACCEPT_BY = NVLI(Rs("ACCEPT_BY"), 0)
      RECEIVE_BY = NVLI(Rs("RECEIVE_BY"), 0)
      
      EXCEPTION_FLAG = NVLS(Rs("EXCEPTION_FLAG"), "")
      PAYEE_NAME = NVLS(Rs("PAYEE_NAME"), "")
      
      '"INVENTORY_DOC_ID" "ID_TYPE",
      
      COMMIT_FLAG = NVLS(Rs("COMMIT_FLAG"), "")
      
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), 0)
      SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
      
      RECEIPT_TYPE = NVLI(Rs("RECEIPT_TYPE"), 0)
      ACCOUNT_ID = NVLI(Rs("ACCOUNT_ID"), 0)
      
      DEPOSIT_AMOUNT = NVLD(Rs("DEPOSIT_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      
      APPROVE_DATE = InternalDateToDate(NVLS(Rs("APPROVE_DATE"), ""))
      ESTIMATE_DATE = InternalDateToDate(NVLS(Rs("ESTIMATE_DATE"), ""))
      
      RESOURCE_ID = NVLI(Rs("RESOURCE_ID"), 0)
      PAYMENT_TYPE = NVLI(Rs("PAYMENT_TYPE"), 0)
      
      BANK_ID = NVLI(Rs("BANK_ID"), 0)
      BBRANCH_ID = NVLI(Rs("BBRANCH_ID"), 0)
      
      CHECK_NO = NVLS(Rs("CHECK_NO"), "")
      CHECK_DATE = InternalDateToDate(NVLS(Rs("CHECK_DATE"), ""))
      VADILITY = InternalDateToDate(NVLS(Rs("VADILITY"), ""))
      DELIVERY = InternalDateToDate(NVLS(Rs("DELIVERY"), ""))
      PAYMENT = InternalDateToDate(NVLS(Rs("PAYMENT"), ""))
      
      PAYMENT_DESC = NVLS(Rs("PAYMENT_DESC"), "")
      SHIPMENT = (NVLS(Rs("SHIPMENT"), ""))
      REFER_INV = NVLS(Rs("REFER_INV"), "")
      PACKING_OF = NVLS(Rs("PACKING_OF"), "")
      SHIPON_DATE = InternalDateToDate(NVLS(Rs("SHIPON_DATE"), ""))
      REF = NVLS(Rs("REF"), "")
      
      Credit = NVLD(Rs("CREDIT"), 0)
      'NUM_NO = NVLI(Rs("NUM_NO"), 0)
      AREA_CODE = NVLS(Rs("AREA_CODE"), "")
      'PO_DATE = InternalDateToDate(NVLS(Rs("PO_DATE"), ""))
      
      SHIP_NAME = NVLS(Rs("SHIP_NAME"), "")
      SHIPPING_MARKS = NVLS(Rs("SHIPPING_MARKS"), "")
      CD_PERCENT = NVLD(Rs("CD_PERCENT"), 0)
      CD_AMOUNT = NVLD(Rs("CD_AMOUNT"), 0)
      PACKAGE_TYPE = NVLI(Rs("PACKAGE_TYPE"), 0)
      TEMP_DO_NO = NVLS(Rs("TEMP_DO_NO"), "")
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      
      DEBIT_AMOUNT = NVLD(Rs("DEBIT_AMOUNT"), 0)
      CREDIT_AMOUNT = NVLD(Rs("CREDIT_AMOUNT"), 0)
      DO_TOTAL_PRICE = NVLD(Rs("DO_TOTAL_PRICE"), 0)
      REVENUE_TOTAL_PRICE = NVLD(Rs("REVENUE_TOTAL_PRICE"), 0)
      'DOCUMENT_SUBTYPE = NVLI(Rs("DOCUMENT_SUBTYPE"), 0)
      PAYMENT_ID = NVLI(Rs("PAYMENT_ID"), 0)
      BANK_BRANCH_ID = NVLI(Rs("BANK_BRANCH_ID"), 0)
      
      BANK_NOTE = NVLS(Rs("BANK_NOTE"), "")
      TOTAL_RCP = NVLD(Rs("TOTAL_RCP"), 0)
      RUNNING_NO = NVLI(Rs("RUNNING_NO"), 0)
      AGREEMENT_DATA = NVLS(Rs("AGREEMENT_DATA"), "")
      AGREEMENT_FINANCE = NVLS(Rs("AGREEMENT_FINANCE"), "")
      
      OLD_CREDIT_AMOUNT = NVLD(Rs("OLD_CREDIT_AMOUNT"), 0)
      ENTRY_DATE = InternalDateToDate(NVLS(Rs("ENTRY_DATE"), ""))
      EXIT_DATE = InternalDateToDate(NVLS(Rs("EXIT_DATE"), ""))
      DO_NO = NVLS(Rs("DO_NO"), "")
      TRUCK_NO = NVLS(Rs("TRUCK_NO"), "")
      
      DELIVERY_FEE = NVLD(Rs("DELIVERY_FEE"), 0)
      SENDER_NAME = NVLS(Rs("SENDER_NAME"), "")
      RECEIVE_NAME = NVLS(Rs("RECEIVE_NAME"), "")
      DEPARTMENT_ID = NVLI(Rs("DEPARTMENT_ID"), 0)
      QUE_NO = NVLS(Rs("QUE_NO"), "")
      PR_NO = NVLS(Rs("PR_NO"), "")
      
      DO_ID = NVLI(Rs("DO_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
      ACTUAL_UNIT_PRICE = NVLD(Rs("ACTUAL_UNIT_PRICE"), 0)
      INCLUDE_UNIT_PRICE = NVLD(Rs("INCLUDE_UNIT_PRICE"), 0)
      PREVIOUS_PRICE = NVLD(Rs("PREVIOUS_PRICE"), 0)
      PREVIOUS_AMOUNT = NVLD(Rs("PREVIOUS_AMOUNT"), 0)
      NEW_PRICE = NVLD(Rs("NEW_PRICE"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      NEW_AMOUNT = NVLD(Rs("NEW_AMOUNT"), 0)
      TRANSACTION_SEQ = NVLI(Rs("TRANSACTION_SEQ"), 0)
      GUI_ID = NVLI(Rs("GUI_ID"), -1)
      CALCULATE_FLAG = NVLS(Rs("CALCULATE_FLAG"), "")
      TOTAL_ACTUAL_PRICE = NVLD(Rs("TOTAL_ACTUAL_PRICE"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      LEFT_AMOUNT = NVLD(Rs("LEFT_AMOUNT"), 0)
      LAYOUT_ID = NVLI(Rs("LAYOUT_ID"), -1)
      LINK_ID = NVLI(Rs("LINK_ID"), -1)
      PACKAGING_AMT = NVLD(Rs("PACKAGING_AMT"), 0)
      ENTRY_WEIGHT = NVLD(Rs("ENTRY_WEIGHT"), 0)
      EXIT_WEIGHT = NVLD(Rs("EXIT_WEIGHT"), 0)
      WEIGHT_AMOUNT = NVLD(Rs("WEIGHT_AMOUNT"), 0)
      PACKAGE_WEIGHT = NVLD(Rs("PACKAGE_WEIGHT"), 0)
      OTHER_WEIGHT = NVLD(Rs("OTHER_WEIGHT"), 0)
      PERCENT_HUMID = NVLD(Rs("PERCENT_HUMID"), 0)
      HUMID_WEIGHT = NVLD(Rs("HUMID_WEIGHT"), 0)
      PACKAGING_ID = NVLD(Rs("PACKAGING_ID"), 0)
      SUPPLIER_WEIGHT = NVLD(Rs("SUPPLIER_WEIGHT"), 0)
      CALCULATE_TYPE = NVLI(Rs("CALCULATE_TYPE"), 0)
      PACKAGE_AMOUNT = NVLD(Rs("PACKAGE_AMOUNT"), 0)
      ACTUAL_PKG_AMOUNT = NVLD(Rs("ACTUAL_PKG_AMOUNT"), 0)
      PUREXP_ID1 = NVLI(Rs("PUREXP_ID1"), 0)
      PUREXP_ID2 = NVLI(Rs("PUREXP_ID2"), 0)
      EXPENSE1 = NVLD(Rs("EXPENSE1"), 0)
      EXPENSE2 = NVLD(Rs("EXPENSE2"), 0)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      RAW_PRICE = NVLD(Rs("RAW_PRICE"), 0)
      DISCOUNT_AMT = NVLD(Rs("DISCOUNT_AMT"), 0)
      RAW_TOT_PRICE = NVLD(Rs("RAW_TOT_PRICE"), 0)
      TO_DEPARTMENT = NVLI(Rs("TO_DEPARTMENT"), -1)
      ITEM_DESC = NVLS(Rs("ITEM_DESC"), "")
      EXTRA_NAME = NVLS(Rs("EXTRA_NAME"), "")
      SALE_TOT_PRICE = NVLD(Rs("SALE_TOT_PRICE"), 0)
      CALCULATE_WEIGHT = NVLD(Rs("CALCULATE_WEIGHT"), 0)
      RAW_COST = NVLD(Rs("RAW_COST"), 0)
      EXPENSE_COST = NVLD(Rs("EXPENSE_COST"), 0)
      TOTAL_NEW_PRICE = NVLD(Rs("TOTAL_NEW_PRICE"), 0)
      BAG_RETURN = NVLS(Rs("BAG_RETURN"), "")
      ACTUAL_AMOUNT = NVLD(Rs("ACTUAL_AMOUNT"), 0)
      ACTUAL_PRICE = NVLD(Rs("ACTUAL_PRICE"), 0)
      CURRENT_AMOUNT = NVLD(Rs("CURRENT_AMOUNT"), 0)
      CURRENT_PRICE = NVLD(Rs("CURRENT_PRICE"), 0)
      NEED_TOTAL_PRICE = NVLD(Rs("NEED_TOTAL_PRICE"), 0)
      NEED_TOTAL_AMOUNT = NVLD(Rs("NEED_TOTAL_AMOUNT"), 0)
      NEED_AVG_PRICE = NVLD(Rs("NEED_AVG_PRICE"), 0)
      MANUAL_PRICE = NVLS(Rs("MANUAL_PRICE"), "")
      EXPENSE_TYPE = NVLI(Rs("EXPENSE_TYPE"), 0)
      AUTO_PRICE = NVLS(Rs("AUTO_PRICE"), "")
      ITEM_DESC_ID = NVLI(Rs("ITEM_DESC_ID"), -1)
      
      PO_ID = NVLI(Rs("PO_ID"), -1)
      PO_NO = NVLS(Rs("PO_NO"), "")
   ElseIf Ind = 103 Then
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      DOCUMENT_SUBTYPE = NVLI(Rs("DOCUMENT_SUBTYPE"), 0)
      RECEIPT_TYPE = NVLI(Rs("RECEIPT_TYPE"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
   ElseIf Ind = 104 Then
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
   ElseIf Ind = 105 Then
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), 0)
      SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
      SUPPLIER_NAME = NVLS(Rs("SUPPLIER_NAME"), "")
      
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), 0)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
   ElseIf Ind = 106 Then
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), 0)
      PO_ID = NVLI(Rs("PO_ID"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
   ElseIf Ind = 107 Then
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), 0)
      SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
      
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DUE_DATE = InternalDateToDate(NVLS(Rs("DUE_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      DISCOUNT_AMT = NVLD(Rs("DISCOUNT_AMT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
     ElseIf Ind = 108 Then
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), 0)
      SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
      SUPPLIER_NAME = NVLS(Rs("SUPPLIER_NAME"), "")
      
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), 0)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      ACTUAL_UNIT_PRICE = NVLD(Rs("ACTUAL_UNIT_PRICE"), 0)
      NOTE = NVLS(Rs("NOTE"), "")
       CLOSE_FLAG = NVLS(Rs("CLOSE_FLAG"), "")
        
     ElseIf Ind = 109 Then
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), 0)
      SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
      SUPPLIER_NAME = NVLS(Rs("SUPPLIER_NAME"), "")
      
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), 0)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      ACTUAL_UNIT_PRICE = NVLD(Rs("ACTUAL_UNIT_PRICE"), 0)
      Credit = NVLD(Rs("CREDIT"), 0)
      DUE_DATE = InternalDateToDate(NVLS(Rs("DUE_DATE"), ""))
       NOTE = NVLS(Rs("NOTE"), "")
         
      
      PAID_TYPE_NAME = NVLS(Rs("PAID_TYPE_NAME"), "")
      
   ElseIf Ind = 110 Then
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), 0)
      PO_ID = NVLI(Rs("PO_ID"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      WEIGHT_AMOUNT = NVLD(Rs("WEIGHT_AMOUNT"), 0)
      SUPPLIER_WEIGHT = NVLD(Rs("SUPPLIER_WEIGHT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0) 'CALCULATE_TYPE
      'CALCULATE_TYPE = NVLD(Rs("CALCULATE_TYPE"), 0)
   ElseIf Ind = 111 Then
      PO_ID = NVLI(Rs("PO_ID"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
     ElseIf Ind = 112 Then
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), 0)
      SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
      SUPPLIER_NAME = NVLS(Rs("SUPPLIER_NAME"), "")
      
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), 0)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      ACTUAL_UNIT_PRICE = NVLD(Rs("ACTUAL_UNIT_PRICE"), 0)
      CALCULATE_TYPE = NVLD(Rs("CALCULATE_TYPE"), 0)
      NOTE = NVLS(Rs("NOTE"), "")
      CLOSE_FLAG = NVLS(Rs("CLOSE_FLAG"), "")
     ElseIf Ind = 113 Then
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), 0)
      PART_DESC = NVLS(Rs("PART_DESC"), 0)
      PART_NO = NVLS(Rs("PART_NO"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      ACTUAL_UNIT_PRICE = NVLD(Rs("ACTUAL_UNIT_PRICE"), 0)
      INCLUDE_UNIT_PRICE = NVLD(Rs("INCLUDE_UNIT_PRICE"), 0)
      CALCULATE_TYPE = NVLD(Rs("CALCULATE_TYPE"), 0)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      WEIGHT_AMOUNT = NVLD(Rs("WEIGHT_AMOUNT"), 0)
      SUPPLIER_WEIGHT = NVLD(Rs("SUPPLIER_WEIGHT"), 0)
      TOTAL_ACTUAL_PRICE = NVLD(Rs("TOTAL_ACTUAL_PRICE"), 0)
      ACTUAL_PRICE = NVLD(Rs("ACTUAL_PRICE"), 0)
   ElseIf Ind = 114 Then
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), 0)
      SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
      SUPPLIER_NAME = NVLS(Rs("LONG_NAME"), "")
      
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DUE_DATE = InternalDateToDate(NVLS(Rs("DUE_DATE"), ""))
      PO_DOCUMENT_NO = NVLS(Rs("PO_DOCUMENT_NO"), "")
      RO_DOCUMENT_NO = NVLS(Rs("RO_DOCUMENT_NO"), "")
      
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      
      CALCULATE_TYPE = NVLI(Rs("CALCULATE_TYPE"), 0)
      SUPPLIER_WEIGHT = NVLD(Rs("SUPPLIER_WEIGHT"), 0)
      WEIGHT_AMOUNT = NVLD(Rs("WEIGHT_AMOUNT"), 0)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      DISCOUNT_AMT = NVLD(Rs("DISCOUNT_AMT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      ACTUAL_UNIT_PRICE = NVLD(Rs("ACTUAL_UNIT_PRICE"), 0)
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      PACKAGING_NAME = NVLS(Rs("PACKAGING_NAME"), "")
   ElseIf Ind = 115 Then
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), 0)
      SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
      SUPPLIER_NAME = NVLS(Rs("LONG_NAME"), "")
      
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      DISCOUNT_AMT = NVLD(Rs("DISCOUNT_AMT"), 0)
      ACTUAL_UNIT_PRICE = NVLD(Rs("ACTUAL_UNIT_PRICE"), 0)
    ElseIf Ind = 116 Then
     ' DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      WEIGHT_AMOUNT = NVLD(Rs("WEIGHT_AMOUNT"), 0)
      SUPPLIER_WEIGHT = NVLD(Rs("SUPPLIER_WEIGHT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      CALCULATE_TYPE = NVLI(Rs("CALCULATE_TYPE"), 0)
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim ID2 As Long
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


   If Trim(BAG_RETURN) = "" Then
      BAG_RETURN = "N"
   End If
   
   If Trim(AUTO_PRICE) = "" Then
      AUTO_PRICE = "N"
   End If
   
   TempID1 = LAYOUT_ID
   If LAYOUT_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = PACKAGING_ID
   If PACKAGING_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID3 = PUREXP_ID1
   If PUREXP_ID1 <= 0 Then
      TempID3 = "NULL"
   End If
   
   TempID4 = PUREXP_ID2
   If PUREXP_ID2 <= 0 Then
      TempID4 = "NULL"
   End If
   
   TempID5 = TO_DEPARTMENT
   If TO_DEPARTMENT <= 0 Then
      TempID5 = "NULL"
   End If
   
   TempID6 = EXPENSE_TYPE
   If EXPENSE_TYPE <= 0 Then
      TempID6 = "NULL"
   End If
   
   TempID7 = ITEM_DESC_ID
   If ITEM_DESC_ID <= 0 Then
      TempID7 = "NULL"
   End If
      
   TempID8 = PO_ID
   If PO_ID <= 0 Then
      TempID8 = "NULL"
   End If
      
   TempID9 = LOCATION_ID
   If LOCATION_ID <= 0 Then
      TempID9 = "NULL"
   End If
   
   TempID10 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID10 = "NULL"
   End If
     
   TempID11 = PROJECT_NAME_ID
   If PROJECT_NAME_ID <= 0 Then
     TempID11 = "NULL"
   End If
   

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If TX_TYPE = "E" Then
      MULTIPLIER = -1
   ElseIf TX_TYPE = "I" Then
      MULTIPLIER = 1
   End If
   
   If MANUAL_PRICE = "" Then
      MANUAL_PRICE = "N"
   End If
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("SUP_ITEM_SEQ", ID, glbErrorLog)
      Call glbDatabaseMngr.GetSeqID("TRANSACTION_SEQUENCE_SEQ", ID2, glbErrorLog)
      
      SQL1 = "INSERT INTO SUP_ITEM " & _
                     "( " & _
                     "   SUP_ITEM_ID, " & _
                     "   DO_ID, " & _
                     "   PART_ITEM_ID, " & _
                     "   LOCATION_ID, " & _
                     "   ACTUAL_UNIT_PRICE, " & _
                     "   INCLUDE_UNIT_PRICE, " & _
                     "   PREVIOUS_PRICE, " & _
                     "   PREVIOUS_AMOUNT, " & _
                     "   NEW_PRICE, " & _
                     "   TX_AMOUNT, " & _
                     "   NEW_AMOUNT, "
            SQL1 = SQL1 & _
                     "   TRANSACTION_SEQ, " & _
                     "   GUI_ID, " & _
                     "   CALCULATE_FLAG, " & _
                     "   TOTAL_ACTUAL_PRICE, " & _
                     "   TOTAL_INCLUDE_PRICE, "
               SQL1 = SQL1 & _
                     "   TX_TYPE, " & _
                     "   LEFT_AMOUNT, " & _
                     "   MULTIPLIER, " & _
                     "   LAYOUT_ID, " & _
                     "   LINK_ID, " & _
                     "   PACKAGING_AMT, " & _
                     "   ENTRY_WEIGHT, " & _
                     "   EXIT_WEIGHT, "
            SQL1 = SQL1 & _
                     "   WEIGHT_AMOUNT, " & _
                     "   PACKAGE_WEIGHT, " & _
                     "   OTHER_WEIGHT, " & _
                     "   PERCENT_HUMID, " & _
                     "   HUMID_WEIGHT, " & _
                     "   PACKAGING_ID, " & _
                     "   SUPPLIER_WEIGHT, " & _
                     "   CALCULATE_TYPE, "
            SQL1 = SQL1 & _
                     "   PACKAGE_AMOUNT, " & _
                     "   ACTUAL_PKG_AMOUNT, " & _
                     "   PUREXP_ID1, " & _
                     "   PUREXP_ID2, " & _
                     "   EXPENSE1, " & _
                     "   EXPENSE2, " & _
                     "   TOTAL_WEIGHT, "
         SQL1 = SQL1 & _
                     "   RAW_PRICE, " & _
                     "   DISCOUNT_AMT, " & _
                     "   RAW_TOT_PRICE, " & _
                     "   ITEM_DESC, " & _
                     "   TO_DEPARTMENT, " & _
                     "   EXTRA_NAME, " & _
                     "   SALE_TOT_PRICE, " & _
                     "   CALCULATE_WEIGHT, "
         SQL1 = SQL1 & _
                     "   RAW_COST, " & _
                     "   EXPENSE_COST, " & _
                     "   TOTAL_NEW_PRICE, " & _
                     "   BAG_RETURN, " & _
                     "   ACTUAL_AMOUNT, " & _
                     "   ACTUAL_PRICE, " & _
                     "   CURRENT_AMOUNT, " & _
                     "   CURRENT_PRICE, "
         SQL1 = SQL1 & _
                     "   NEED_TOTAL_AMOUNT, " & _
                     "   NEED_TOTAL_PRICE, " & _
                     "   NEED_AVG_PRICE, " & _
                     "   MANUAL_PRICE, " & _
                     "   EXPENSE_TYPE, " & _
                     "   AUTO_PRICE, " & _
                     "   ITEM_DESC_ID, " & _
                     "   PO_ID, " & _
                     "  PROJECT_NAME_ID, " & _
                     "   PO_NO, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY ," & _
                     "  PART_SUPPLIER_NAME ," & _
                     "  PART_NOTE " & _
                     ") "
         SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "" & DO_ID & "," & _
                         TempID10 & "," & _
                         TempID9 & "," & _
                         "" & ACTUAL_UNIT_PRICE & "," & _
                         "" & INCLUDE_UNIT_PRICE & "," & _
                         "" & PREVIOUS_PRICE & "," & _
                         "" & PREVIOUS_AMOUNT & "," & _
                         "" & NEW_PRICE & ","
               SQL1 = SQL1 & _
                         "" & TX_AMOUNT & "," & _
                         "" & NEW_AMOUNT & "," & _
                         "" & ID2 & "," & _
                         "" & GUI_ID & "," & _
                         "'" & CALCULATE_FLAG & "'," & _
                         "" & TOTAL_ACTUAL_PRICE & "," & _
                         "" & TOTAL_INCLUDE_PRICE & "," & _
                         "'" & TX_TYPE & "'," & _
                         "" & LEFT_AMOUNT & "," & _
                         "" & MULTIPLIER & "," & _
                         TempID1 & "," & _
                         LINK_ID & "," & _
                         PACKAGING_AMT & "," & _
                         ENTRY_WEIGHT & "," & _
                         EXIT_WEIGHT & ","
               SQL1 = SQL1 & _
                        WEIGHT_AMOUNT & "," & _
                        PACKAGE_WEIGHT & "," & _
                        OTHER_WEIGHT & "," & _
                        PERCENT_HUMID & "," & _
                        HUMID_WEIGHT & "," & _
                        TempID2 & "," & _
                        SUPPLIER_WEIGHT & "," & _
                        CALCULATE_TYPE & "," & _
                        PACKAGE_AMOUNT & "," & _
                        ACTUAL_PKG_AMOUNT & "," & _
                        TempID3 & "," & _
                        TempID4 & "," & _
                        EXPENSE1 & "," & _
                        EXPENSE2 & "," & _
                        TOTAL_WEIGHT & ","
               SQL1 = SQL1 & _
                        RAW_PRICE & "," & _
                        DISCOUNT_AMT & "," & _
                        RAW_TOT_PRICE & "," & _
                        "'" & Trim(ChangeQuote(ITEM_DESC)) & "'," & _
                        TempID5 & "," & _
                        "'" & Trim(ChangeQuote(EXTRA_NAME)) & "'," & _
                        SALE_TOT_PRICE & "," & _
                        CALCULATE_WEIGHT & "," & _
                        RAW_COST & "," & _
                        EXPENSE_COST & "," & _
                        TOTAL_NEW_PRICE & "," & _
                        "'" & Trim(ChangeQuote(BAG_RETURN)) & "',"
               SQL1 = SQL1 & _
                        ACTUAL_AMOUNT & "," & _
                        ACTUAL_PRICE & "," & _
                        CURRENT_AMOUNT & "," & _
                        CURRENT_PRICE & "," & _
                        NEED_TOTAL_AMOUNT & "," & _
                        NEED_TOTAL_PRICE & "," & _
                        NEED_AVG_PRICE & "," & _
                        "'" & Trim(ChangeQuote(MANUAL_PRICE)) & "'," & _
                        TempID6 & "," & _
                        "'" & Trim(ChangeQuote(AUTO_PRICE)) & "'," & _
                        TempID7 & "," & _
                         TempID8 & "," & _
                         TempID11 & "," & _
                        "'" & Trim(ChangeQuote(PO_NO)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & ",'" & _
                         PART_SUPPLIER_NAME & "','" & _
                         PART_NOTE & _
                     "') "
         SUP_ITEM_ID = ID
   Else 'Show edit
      Call glbDatabaseMngr.GetSeqID("TRANSACTION_SEQUENCE_SEQ", ID2, glbErrorLog)
      
      SQL1 = "UPDATE SUP_ITEM SET " & _
                     "   DO_ID = " & DO_ID & ", " & _
                     "   PART_ITEM_ID = " & TempID10 & ", " & _
                     "   LOCATION_ID = " & TempID9 & ", " & _
                     "   ACTUAL_UNIT_PRICE = " & ACTUAL_UNIT_PRICE & ", " & _
                     "   INCLUDE_UNIT_PRICE = " & INCLUDE_UNIT_PRICE & ", " & _
                     "   PREVIOUS_PRICE = " & PREVIOUS_PRICE & ", " & _
                     "   PREVIOUS_AMOUNT = " & PREVIOUS_AMOUNT & ", " & _
                     "   NEW_PRICE = " & NEW_PRICE & ", " & _
                     "   TX_AMOUNT = " & TX_AMOUNT & ", " & _
                     "   NEW_AMOUNT = " & NEW_AMOUNT & ", " & _
                     "   TRANSACTION_SEQ = " & ID2 & ", " & _
                     "   GUI_ID = " & GUI_ID & ", " & _
                     "   CALCULATE_FLAG = '" & CALCULATE_FLAG & "', " & _
                     "   TOTAL_ACTUAL_PRICE = " & TOTAL_ACTUAL_PRICE & ", " & _
                     "   TOTAL_INCLUDE_PRICE = " & TOTAL_INCLUDE_PRICE & ", " & _
                     "   TX_TYPE = '" & TX_TYPE & "', " & _
                     "   LEFT_AMOUNT = " & LEFT_AMOUNT & ", " & _
                     "   MULTIPLIER = " & MULTIPLIER & ", " & _
                     "   LAYOUT_ID = " & TempID1 & ", " & _
                     "   LINK_ID = " & LINK_ID & ", " & _
                     "   PACKAGING_AMT = " & PACKAGING_AMT & ", "
            SQL1 = SQL1 & _
                     "   ENTRY_WEIGHT = " & ENTRY_WEIGHT & ", " & _
                     "   EXIT_WEIGHT = " & EXIT_WEIGHT & ", " & _
                     "   WEIGHT_AMOUNT = " & WEIGHT_AMOUNT & ", " & _
                     "   PACKAGE_WEIGHT = " & PACKAGE_WEIGHT & ", " & _
                     "   OTHER_WEIGHT = " & OTHER_WEIGHT & ", " & _
                     "   PERCENT_HUMID = " & PERCENT_HUMID & ", " & _
                     "   HUMID_WEIGHT = " & HUMID_WEIGHT & ", " & _
                     "   PACKAGING_ID = " & TempID2 & ", " & _
                     "   SUPPLIER_WEIGHT = " & SUPPLIER_WEIGHT & ", " & _
                     "   CALCULATE_TYPE = " & CALCULATE_TYPE & ", " & _
                     "   PACKAGE_AMOUNT = " & PACKAGE_AMOUNT & ", " & _
                     "   ACTUAL_PKG_AMOUNT = " & ACTUAL_PKG_AMOUNT & ", "
            SQL1 = SQL1 & _
                     "   PUREXP_ID1 = " & TempID3 & ", " & _
                     "   PUREXP_ID2 = " & TempID4 & ", " & _
                     "   EXPENSE1 = " & EXPENSE1 & ", " & _
                     "   EXPENSE2 = " & EXPENSE2 & ", " & _
                     "   TOTAL_WEIGHT = " & TOTAL_WEIGHT & ", " & _
                     "   RAW_PRICE = " & RAW_PRICE & ", " & _
                     "   DISCOUNT_AMT = " & DISCOUNT_AMT & ", " & _
                     "   RAW_TOT_PRICE = " & RAW_TOT_PRICE & ", " & _
                     "    ITEM_DESC = '" & Trim(ChangeQuote(ITEM_DESC)) & "', " & _
                     "   TO_DEPARTMENT = " & TempID5 & ", " & _
                     "    EXTRA_NAME = '" & Trim(ChangeQuote(EXTRA_NAME)) & "', " & _
                     "   SALE_TOT_PRICE = " & SALE_TOT_PRICE & ", " & _
                     "   CALCULATE_WEIGHT = " & CALCULATE_WEIGHT & ", " & _
                     "   RAW_COST = " & RAW_COST & ", " & _
                     "   EXPENSE_COST = " & EXPENSE_COST & ", " & _
                     "   TOTAL_NEW_PRICE = " & TOTAL_NEW_PRICE & ", " & _
                     "    BAG_RETURN = '" & Trim(ChangeQuote(BAG_RETURN)) & "', "
            SQL1 = SQL1 & _
                     "   ACTUAL_AMOUNT = " & ACTUAL_AMOUNT & ", " & _
                     "   ACTUAL_PRICE = " & ACTUAL_PRICE & ", " & _
                     "   CURRENT_AMOUNT = " & CURRENT_AMOUNT & ", " & _
                     "   CURRENT_PRICE = " & CURRENT_PRICE & ", " & _
                     "   NEED_TOTAL_AMOUNT = " & NEED_TOTAL_AMOUNT & ", " & _
                     "   NEED_TOTAL_PRICE = " & NEED_TOTAL_PRICE & ", " & _
                     "   NEED_AVG_PRICE = " & NEED_AVG_PRICE & ", " & _
                     "   MANUAL_PRICE = '" & Trim(ChangeQuote(MANUAL_PRICE)) & "', " & _
                     "   EXPENSE_TYPE = " & TempID6 & ", " & _
                     "    AUTO_PRICE = '" & Trim(ChangeQuote(AUTO_PRICE)) & "', " & _
                     "   ITEM_DESC_ID = " & TempID7 & ", " & _
                     "   PO_ID = " & TempID8 & ", " & _
                     "   PROJECT_NAME_ID = " & TempID11 & ", " & _
                     "    PO_NO = '" & Trim(ChangeQuote(PO_NO)) & "', " & _
                     "    PART_SUPPLIER_NAME = '" & Trim(ChangeQuote(PART_SUPPLIER_NAME)) & "', " & _
                     "    PART_NOTE = '" & Trim(ChangeQuote(PART_NOTE)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE SUP_ITEM_ID = " & SUP_ITEM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM SUP_ITEM WHERE SUP_ITEM_ID = " & SUP_ITEM_ID
      m_Conn.Execute (SQL1)
      
End Function

Public Function DecreaseLeftAmount(DecAmount As Double)
Dim SQL1 As String
      SQL1 = "UPDATE SUP_ITEM SET LEFT_AMOUNT = LEFT_AMOUNT - " & DecAmount & " WHERE SUP_ITEM_ID = " & SUP_ITEM_ID
      m_Conn.Execute (SQL1)
End Function

Public Function PatchDocDate() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE BILLING_DOC SET " & _
                  "   DOCUMENT_DATE = '" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "', " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE BILLING_DOC_ID = " & DO_ID
                  
   m_Conn.Execute (SQL1)
End Function

Public Function PatchAvgPriceExp(AvgPrice As Double, PrevAmount As Double, CurrentAmount As Double, ExportTotalPrice As Double, NewTotalPrice As Double) As Boolean
On Error GoTo ErrorHandler
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL1 = "UPDATE SUP_ITEM SET " & _
                  "   INCLUDE_UNIT_PRICE = " & AvgPrice & ", " & _
                  "   PREVIOUS_AMOUNT = " & PrevAmount & ", " & _
                  "   PREVIOUS_PRICE = " & AvgPrice & ", " & _
                  "   NEW_AMOUNT = " & CurrentAmount & ", " & _
                  "   NEW_PRICE = " & AvgPrice & ", " & _
                  "   TOTAL_INCLUDE_PRICE = " & ExportTotalPrice & ", " & _
                  "   TOTAL_NEW_PRICE = " & NewTotalPrice & ", " & _
                  "   TX_AMOUNT = " & PrevAmount - CurrentAmount & ", " & _
                  "   TX_TYPE = '" & TX_TYPE & "', " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE SUP_ITEM_ID = " & SUP_ITEM_ID

   m_Conn.Execute (SQL1)
   Exit Function
   
ErrorHandler:
   SQL1 = "UPDATE SUP_ITEM SET " & _
                  "   INCLUDE_UNIT_PRICE = " & AvgPrice & ", " & _
                  "   PREVIOUS_AMOUNT = " & PrevAmount & ", " & _
                  "   PREVIOUS_PRICE = " & 0 & ", " & _
                  "   NEW_AMOUNT = " & CurrentAmount & ", " & _
                  "   NEW_PRICE = " & 0 & ", " & _
                  "   TOTAL_INCLUDE_PRICE = " & 0 & ", " & _
                  "   TOTAL_NEW_PRICE = " & 0 & ", " & _
                  "   TX_AMOUNT = " & PrevAmount - CurrentAmount & ", " & _
                  "   TX_TYPE = '" & TX_TYPE & "', " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE SUP_ITEM_ID = " & SUP_ITEM_ID

   m_Conn.Execute (SQL1)
End Function

Public Function PatchAvgPrice(AvgPrice As Double, PrevAmount As Double, CurrentAmount As Double, NewPrice As Double, ImportAmount As Double, DocumentType As Long, TotalIncludePrice As Double, NewTotalPrice As Double) As Boolean
On Error GoTo ErrorHandler
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   If (DocumentType = 3) Or (DocumentType = 11) Or (DocumentType = 22) Or (DocumentType = 4) Or (DocumentType = 5) Then
      SQL1 = "UPDATE SUP_ITEM SET " & _
                     "   ACTUAL_UNIT_PRICE = " & AvgPrice & ", " & _
                     "   INCLUDE_UNIT_PRICE = " & AvgPrice & ", " & _
                     "   TOTAL_INCLUDE_PRICE = " & TotalIncludePrice & ", " & _
                     "   TOTAL_ACTUAL_PRICE = " & TotalIncludePrice & ", " & _
                     "   PREVIOUS_AMOUNT = " & PrevAmount & ", " & _
                     "   NEW_AMOUNT = " & CurrentAmount & ", " & _
                     "   NEW_PRICE = " & NewPrice & ", " & _
                     "   RAW_COST = " & RAW_COST & ", " & _
                     "   TX_AMOUNT = " & ImportAmount & ", " & _
                     "   EXPENSE_COST = " & EXPENSE_COST & ", " & _
                     "   TOTAL_NEW_PRICE = " & NewTotalPrice & ", " & _
                     "   TX_TYPE = '" & TX_TYPE & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE SUP_ITEM_ID = " & SUP_ITEM_ID
   Else
      SQL1 = "UPDATE SUP_ITEM SET " & _
                     "   PREVIOUS_AMOUNT = " & PrevAmount & ", " & _
                     "   NEW_AMOUNT = " & CurrentAmount & ", " & _
                     "   NEW_PRICE = " & NewPrice & ", " & _
                     "   RAW_COST = " & RAW_COST & ", " & _
                     "   EXPENSE_COST = " & EXPENSE_COST & ", " & _
                     "   TOTAL_NEW_PRICE = " & NewTotalPrice & ", " & _
                     "   TX_TYPE = '" & TX_TYPE & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE SUP_ITEM_ID = " & SUP_ITEM_ID
   End If
   m_Conn.Execute (SQL1)
   Exit Function
   
ErrorHandler:

   If (DocumentType = 3) Or (DocumentType = 11) Then
      SQL1 = "UPDATE SUP_ITEM SET " & _
                     "   ACTUAL_UNIT_PRICE = " & 0 & ", " & _
                     "   INCLUDE_UNIT_PRICE = " & 0 & ", " & _
                     "   TOTAL_INCLUDE_PRICE = " & 0 & ", " & _
                     "   TOTAL_ACTUAL_PRICE = " & 0 & ", " & _
                     "   PREVIOUS_AMOUNT = " & PrevAmount & ", " & _
                     "   NEW_AMOUNT = " & CurrentAmount & ", " & _
                     "   NEW_PRICE = " & 0 & ", " & _
                     "   RAW_COST = " & 0 & ", " & _
                     "   EXPENSE_COST = " & 0 & ", " & _
                     "   TOTAL_NEW_PRICE = " & 0 & ", " & _
                     "   TX_TYPE = '" & TX_TYPE & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE SUP_ITEM_ID = " & SUP_ITEM_ID
   Else
      SQL1 = "UPDATE SUP_ITEM SET " & _
                     "   PREVIOUS_AMOUNT = " & PrevAmount & ", " & _
                     "   NEW_AMOUNT = " & CurrentAmount & ", " & _
                     "   NEW_PRICE = " & 0 & ", " & _
                     "   RAW_COST = " & 0 & ", " & _
                     "   EXPENSE_COST = " & 0 & ", " & _
                     "   TOTAL_NEW_PRICE = " & 0 & ", " & _
                     "   TX_TYPE = '" & TX_TYPE & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE SUP_ITEM_ID = " & SUP_ITEM_ID
   End If
   m_Conn.Execute (SQL1)
End Function
Private Sub Class_Terminate()
   Set SubLotItems = Nothing
End Sub
Public Sub GenerateBDHeader(FileID As Long)
Dim TempStr As String
   
   TempStr = "BD;"
   TempStr = TempStr & DOCUMENT_NO & ";"
   TempStr = TempStr & DateToStringInt(DOCUMENT_DATE) & ";"
   TempStr = TempStr & DOCUMENT_TYPE & ";"
   TempStr = TempStr & DateToStringInt(DUE_DATE) & ";"
   TempStr = TempStr & NOTE & ";"
   
   TempStr = TempStr & BILLING_ADDRESS_ID & ";"
   TempStr = TempStr & ENTERPRISE_ID & ";"
   TempStr = TempStr & ENTERPRISE_ADDRESS_ID & ";"
   TempStr = TempStr & TOTAL_AMOUNT & ";"
   TempStr = TempStr & VAT_PERCENT & ";"
   
   TempStr = TempStr & VAT_AMOUNT & ";"
   TempStr = TempStr & WH_PERCENT & ";"
   TempStr = TempStr & WH_AMOUNT & ";"
   TempStr = TempStr & DISCOUNT_AMOUNT & ";"
   TempStr = TempStr & PAYMENT_TERM & ";"
   
   TempStr = TempStr & ACCEPT_BY & ";"
   TempStr = TempStr & RECEIVE_BY & ";"
   TempStr = TempStr & EXCEPTION_FLAG & ";"
   TempStr = TempStr & PAYEE_NAME & ";"
   TempStr = TempStr & COMMIT_FLAG & ";"
   
   'TempStr = TempStr & SUPPLIER_ID & ";"
   TempStr = TempStr & SUPPLIER_CODE & ";"
   TempStr = TempStr & RECEIPT_TYPE & ";"
   TempStr = TempStr & ACCOUNT_ID & ";"
   TempStr = TempStr & DEPOSIT_AMOUNT & ";"
   TempStr = TempStr & TOTAL_PRICE & ";"
   
   TempStr = TempStr & DateToStringInt(APPROVE_DATE) & ";"
   TempStr = TempStr & DateToStringInt(ESTIMATE_DATE) & ";"
   
   TempStr = TempStr & RESOURCE_ID & ";"
   TempStr = TempStr & PAYMENT_TYPE & ";"
   TempStr = TempStr & BANK_ID & ";"
   TempStr = TempStr & BBRANCH_ID & ";"
   TempStr = TempStr & CHECK_NO & ";"
   
   TempStr = TempStr & DateToStringInt(CHECK_DATE) & ";"
   TempStr = TempStr & VADILITY & ";"
   TempStr = TempStr & DELIVERY & ";"
   TempStr = TempStr & PAYMENT & ";"
   TempStr = TempStr & PAYMENT_DESC & ";"
   
   TempStr = TempStr & SHIPMENT & ";"
   TempStr = TempStr & REFER_INV & ";"
   TempStr = TempStr & PACKING_OF & ";"
   TempStr = TempStr & DateToStringInt(SHIPON_DATE) & ";"
   TempStr = TempStr & REF & ";"
   
   TempStr = TempStr & Credit & ";"
   TempStr = TempStr & AREA_CODE & ";"
   TempStr = TempStr & SHIP_NAME & ";"
   TempStr = TempStr & SHIPPING_MARKS & ";"
   TempStr = TempStr & CD_PERCENT & ";"
   
   TempStr = TempStr & CD_AMOUNT & ";"
   TempStr = TempStr & PACKAGE_TYPE & ";"
   TempStr = TempStr & TEMP_DO_NO & ";"
   TempStr = TempStr & PAID_AMOUNT & ";"
   TempStr = TempStr & DEBIT_AMOUNT & ";"
      
   TempStr = TempStr & CREDIT_AMOUNT & ";"
   TempStr = TempStr & DO_TOTAL_PRICE & ";"
   TempStr = TempStr & REVENUE_TOTAL_PRICE & ";"
   TempStr = TempStr & PAYMENT_ID & ";"
   TempStr = TempStr & BANK_BRANCH_ID & ";"
   
   TempStr = TempStr & BANK_NOTE & ";"
   TempStr = TempStr & TOTAL_RCP & ";"
   TempStr = TempStr & RUNNING_NO & ";"
   TempStr = TempStr & AGREEMENT_DATA & ";"
   TempStr = TempStr & AGREEMENT_FINANCE & ";"
   
   TempStr = TempStr & OLD_CREDIT_AMOUNT & ";"
   TempStr = TempStr & DateToStringInt(ENTRY_DATE) & ";"
   TempStr = TempStr & DateToStringInt(EXIT_DATE) & ";"
   TempStr = TempStr & DO_NO & ";"
   TempStr = TempStr & TRUCK_NO & ";"
   
   TempStr = TempStr & DELIVERY_FEE & ";"
   TempStr = TempStr & SENDER_NAME & ";"
   TempStr = TempStr & RECEIVE_NAME & ";"
   TempStr = TempStr & DEPARTMENT_ID & ";"
   TempStr = TempStr & QUE_NO & ";"
   TempStr = TempStr & PR_NO & ";"
      
   Print #FileID, TempStr
End Sub
Public Sub GenerateSupTailer(FileID As Long)
Dim TempStr As String
   
   TempStr = "SI;"
   TempStr = TempStr & DO_ID & ";"
   
   TempStr = TempStr & PART_NO & ";"
   TempStr = TempStr & LOCATION_NO & ";"
   TempStr = TempStr & ACTUAL_UNIT_PRICE & ";"
   TempStr = TempStr & INCLUDE_UNIT_PRICE & ";"
   TempStr = TempStr & PREVIOUS_PRICE & ";"
   
   TempStr = TempStr & PREVIOUS_AMOUNT & ";"
   TempStr = TempStr & NEW_PRICE & ";"
   TempStr = TempStr & TX_AMOUNT & ";"
   TempStr = TempStr & NEW_AMOUNT & ";"
   TempStr = TempStr & TRANSACTION_SEQ & ";"
   
   TempStr = TempStr & GUI_ID & ";"
   TempStr = TempStr & CALCULATE_FLAG & ";"
   TempStr = TempStr & TOTAL_ACTUAL_PRICE & ";"
   TempStr = TempStr & TOTAL_INCLUDE_PRICE & ";"
   TempStr = TempStr & TX_TYPE & ";"
   
   TempStr = TempStr & LEFT_AMOUNT & ";"
   TempStr = TempStr & LAYOUT_ID & ";"
   TempStr = TempStr & LINK_ID & ";"
   TempStr = TempStr & PACKAGING_AMT & ";"
   TempStr = TempStr & ENTRY_WEIGHT & ";"
   
   TempStr = TempStr & EXIT_WEIGHT & ";"
   TempStr = TempStr & WEIGHT_AMOUNT & ";"
   TempStr = TempStr & PACKAGE_WEIGHT & ";"
   TempStr = TempStr & OTHER_WEIGHT & ";"
   TempStr = TempStr & PERCENT_HUMID & ";"
   
   TempStr = TempStr & HUMID_WEIGHT & ";"
   TempStr = TempStr & PACKAGING_ID & ";"
   TempStr = TempStr & SUPPLIER_WEIGHT & ";"
   TempStr = TempStr & CALCULATE_TYPE & ";"
   TempStr = TempStr & PACKAGE_AMOUNT & ";"
   
   TempStr = TempStr & ACTUAL_PKG_AMOUNT & ";"
   TempStr = TempStr & PUREXP_ID1 & ";"
   TempStr = TempStr & PUREXP_ID2 & ";"
   TempStr = TempStr & EXPENSE1 & ";"
   TempStr = TempStr & EXPENSE2 & ";"
         
   TempStr = TempStr & TOTAL_WEIGHT & ";"
   TempStr = TempStr & RAW_PRICE & ";"
   TempStr = TempStr & DISCOUNT_AMT & ";"
   TempStr = TempStr & RAW_TOT_PRICE & ";"
   TempStr = TempStr & TO_DEPARTMENT & ";"
   
   TempStr = TempStr & ITEM_DESC & ";"
   TempStr = TempStr & EXTRA_NAME & ";"
   TempStr = TempStr & SALE_TOT_PRICE & ";"
   TempStr = TempStr & CALCULATE_WEIGHT & ";"
   TempStr = TempStr & RAW_COST & ";"
   
   TempStr = TempStr & EXPENSE_COST & ";"
   TempStr = TempStr & TOTAL_NEW_PRICE & ";"
   TempStr = TempStr & BAG_RETURN & ";"
   TempStr = TempStr & ACTUAL_AMOUNT & ";"
   TempStr = TempStr & ACTUAL_PRICE & ";"
   
   TempStr = TempStr & CURRENT_AMOUNT & ";"
   TempStr = TempStr & CURRENT_PRICE & ";"
   TempStr = TempStr & NEED_TOTAL_PRICE & ";"
   TempStr = TempStr & NEED_TOTAL_AMOUNT & ";"
   TempStr = TempStr & NEED_AVG_PRICE & ";"
      
   TempStr = TempStr & MANUAL_PRICE & ";"
   TempStr = TempStr & EXPENSE_TYPE & ";"
   TempStr = TempStr & AUTO_PRICE & ";"
   TempStr = TempStr & ITEM_DESC_ID & ";"
   
   TempStr = TempStr & PO_NO & ";"
   
   Print #FileID, TempStr
End Sub
Public Function DeleteFromBillInv()
Dim SQL1 As String

      SQL1 = "DELETE FROM SUP_ITEM WHERE DO_ID = " & DO_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM LOT_ITEM WHERE INVENTORY_DOC_ID = " & INVENTORY_DOC_ID
      m_Conn.Execute (SQL1)
End Function
Public Function CopyObject(Ind As Long, TempObj As CSupItem)
Dim TempStr As String
   
   PART_NO = TempObj.PART_NO
   LOCATION_NO = TempObj.LOCATION_NO
   ACTUAL_UNIT_PRICE = TempObj.ACTUAL_UNIT_PRICE
   INCLUDE_UNIT_PRICE = TempObj.INCLUDE_UNIT_PRICE
   TX_AMOUNT = TempObj.TX_AMOUNT
   CALCULATE_FLAG = TempObj.CALCULATE_FLAG
   TOTAL_ACTUAL_PRICE = TempObj.TOTAL_ACTUAL_PRICE
   TOTAL_INCLUDE_PRICE = TempObj.TOTAL_INCLUDE_PRICE
   TX_TYPE = TempObj.TX_TYPE
   LINK_ID = TempObj.LINK_ID
   
   PACKAGING_AMT = TempObj.PACKAGING_AMT
   ENTRY_WEIGHT = TempObj.ENTRY_WEIGHT
   
   EXIT_WEIGHT = TempObj.EXIT_WEIGHT
   WEIGHT_AMOUNT = TempObj.WEIGHT_AMOUNT
   PACKAGE_WEIGHT = TempObj.PACKAGE_WEIGHT
   OTHER_WEIGHT = TempObj.OTHER_WEIGHT
   PERCENT_HUMID = TempObj.PERCENT_HUMID
   
   HUMID_WEIGHT = TempObj.HUMID_WEIGHT
   PACKAGING_ID = TempObj.PACKAGING_ID
   SUPPLIER_WEIGHT = TempObj.SUPPLIER_WEIGHT
   CALCULATE_TYPE = TempObj.CALCULATE_TYPE
   PACKAGE_AMOUNT = TempObj.PACKAGE_AMOUNT
   
   ACTUAL_PKG_AMOUNT = TempObj.ACTUAL_PKG_AMOUNT
   PUREXP_ID1 = TempObj.PUREXP_ID1
   PUREXP_ID2 = TempObj.PUREXP_ID2
   EXPENSE1 = TempObj.EXPENSE1
   EXPENSE2 = TempObj.EXPENSE2
   
   TOTAL_WEIGHT = TempObj.TOTAL_WEIGHT
   RAW_PRICE = TempObj.RAW_PRICE
   DISCOUNT_AMT = TempObj.DISCOUNT_AMT
   RAW_TOT_PRICE = TempObj.RAW_TOT_PRICE
   TO_DEPARTMENT = TempObj.TO_DEPARTMENT
   
   ITEM_DESC = TempObj.ITEM_DESC
   EXTRA_NAME = TempObj.EXTRA_NAME
   SALE_TOT_PRICE = TempObj.SALE_TOT_PRICE
   CALCULATE_WEIGHT = TempObj.CALCULATE_WEIGHT
   RAW_COST = TempObj.RAW_COST
   
   EXPENSE_COST = TempObj.EXPENSE_COST
   TOTAL_NEW_PRICE = TempObj.TOTAL_NEW_PRICE
   BAG_RETURN = TempObj.BAG_RETURN
   ACTUAL_AMOUNT = TempObj.ACTUAL_AMOUNT
   ACTUAL_PRICE = TempObj.ACTUAL_PRICE
   
   CURRENT_AMOUNT = TempObj.CURRENT_AMOUNT
   CURRENT_PRICE = TempObj.CURRENT_PRICE
   NEED_TOTAL_PRICE = TempObj.NEED_TOTAL_PRICE
   NEED_TOTAL_AMOUNT = TempObj.NEED_TOTAL_AMOUNT
   NEED_AVG_PRICE = TempObj.NEED_AVG_PRICE
   
   MANUAL_PRICE = TempObj.MANUAL_PRICE
   EXPENSE_TYPE = TempObj.EXPENSE_TYPE
   AUTO_PRICE = TempObj.AUTO_PRICE
   ITEM_DESC_ID = TempObj.ITEM_DESC_ID
   
   PACKAGING_CODE = TempObj.PACKAGING_CODE
   PRIORITY1 = TempObj.PRIORITY1
   EXPORT_AVG_PRICE = TempObj.EXPORT_AVG_PRICE
   EXPORT_TOTAL_PRICE = TempObj.EXPORT_TOTAL_PRICE
   IMPORT_AMOUNT = TempObj.IMPORT_AMOUNT
   EXPORT_AMOUNT = TempObj.EXPORT_AMOUNT
      
   PART_GROUP_ID = TempObj.PART_GROUP_ID
   PART_TYPE = TempObj.PART_TYPE
   PART_TYPE_NAME = TempObj.PART_TYPE_NAME
   PART_SUPPLIER_NAME = TempObj.PART_SUPPLIER_NAME
   PART_NOTE = TempObj.PART_NOTE
   PART_TYPE_NO = TempObj.PART_TYPE_NO
   LOCATION_NO = TempObj.LOCATION_NO
   
   LOCATION_ID = TempObj.LOCATION_ID
   LOCATION_NAME = TempObj.LOCATION_NAME
   PART_ITEM_ID = TempObj.PART_ITEM_ID
   PART_NO = TempObj.PART_NO
   PART_DESC = TempObj.PART_DESC
   PIG_FLAG = TempObj.PIG_FLAG
   PIG_TYPE = TempObj.PIG_TYPE
   
   UNIT_NAME = TempObj.UNIT_NAME
   
   PO_ID = TempObj.PO_ID
   
End Function
Public Sub AddSearchSupplierCode(FromStr As String)
   If InStr(1, FromStr, "SP") <= 0 Then
        FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD.SUPPLIER_ID) "
   End If
End Sub

