VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLotItem"
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

Public LOT_ITEM_ID As Long
Public INVENTORY_DOC_ID As Long
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
Public ADJUST_FLAG As String
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
Public ENTRY_DATE As Date
Public EXIT_DATE As Date
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
Public PROJECT_ID  As Long
Public PROJECT_NAME As String

Public DEPARTMENT_NAME As String
Public CSTTYPE_NAME As String
Public PR_NO As String
Public YYYYMM  As String
Public HiFlag As Boolean

Public SubLotItems As Collection
Public IncludeFlag As Boolean

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_DATE = -1
   TO_DATE = -1
   Set SubLotItems = New Collection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional MustCount As Boolean = True)
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
      SelectStr1 = "SELECT LI.*, PT.*, LC.*, PI.*, UN.*, IVD.ADJUST_FLAG, IVD.SUPPLIER_ID, IVD.DOCUMENT_DATE, IVD.DOCUMENT_NO, IVD.COMMIT_FLAG, IVD.DOCUMENT_TYPE, IVD.QUE_NO, IVD.RECEIVE_NAME, IVD.TRUCK_NO, IVD.DO_NO, IVD.SENDER_NAME, SP.SUPPLIER_CODE, NM1.LONG_NAME SUPPLIER_NAME, DP.PRIORITY1, PG.PACKAGING_NO, PE1.PUREXP_NO PUREXP_NO1, PE1.PUREXP_NAME PUREXP_NAME1, PE1.EXPENSE_RATE EXPENSE_RATE1, PE2.PUREXP_NO PUREXP_NO2, PE2.PUREXP_NAME PUREXP_NAME2, PE2.EXPENSE_RATE EXPENSE_RATE2, PT.PART_GROUP_ID PART_GROUP_ID1 "
      SelectStr1 = SelectStr1 & ",MR1.KEY_CODE ITEM_DESC_NO,MR1.KEY_NAME ITEM_DESC_NAME,IVD.PR_NO "
      SelectStr1 = SelectStr1 & ",LO.LAY_OUT_NAME DEPARTMENT_NAME "
      SelectStr1 = SelectStr1 & ",IVD.ENTRY_DATE,IVD.EXIT_DATE  "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (SN.NAME_ID = NM1.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN DOCUMENT_PRIORITY DP ON ((IVD.DOCUMENT_TYPE = DP.DOCUMENT_TYPE) AND (LI.TX_TYPE = DP.TX_TYPE) AND (DP.AREA = 1)) "
      FromStr = FromStr & "LEFT OUTER JOIN PACKAGING PG ON (LI.PACKAGING_ID = PG.PACKAGING_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PURCHASE_EXPENSE PE1 ON (LI.PUREXP_ID1 = PE1.PUREXP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PURCHASE_EXPENSE PE2 ON (LI.PUREXP_ID2 = PE2.PUREXP_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (MR1.KEY_ID = LI.ITEM_DESC_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN LAY_OUT LO ON (IVD.DEPARTMENT_ID = LO.LAY_OUT_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT LI.*, PT.*, LC.*, PI.*, UN.*, IVD.DOCUMENT_DATE, IVD.DOCUMENT_NO, IVD.COMMIT_FLAG, IVD.DOCUMENT_TYPE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      
      WhereStr = "WHERE (LI.LEFT_AMOUNT > 0) "
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT LI.PART_ITEM_ID, SUM(LI.MULTIPLIER * LI.TX_AMOUNT) LEFT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      
      WhereStr = ""
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT LI.PART_ITEM_ID, LI.LOCATION_ID, LI.NEW_PRICE, LI.PREVIOUS_PRICE, LI.NEW_AMOUNT, LI.PREVIOUS_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM LI "

      WhereStr = "WHERE LI.TRANSACTION_SEQ IN "
      WhereStr = WhereStr & "( "
      WhereStr = WhereStr & "   SELECT MAX(LI2.TRANSACTION_SEQ) "
      WhereStr = WhereStr & "   FROM LOT_ITEM LI2 "
      WhereStr = WhereStr & "   LEFT OUTER JOIN INVENTORY_DOC IVD2 ON (LI.INVENTORY_DOC_ID = IVD2.INVENTORY_DOC_ID) "
      If HiFlag Then
         WhereStr = WhereStr & "   WHERE (IVD2.DOCUMENT_DATE <= '" & DateToStringIntHi(TO_DATE1) & "') "
      Else
         WhereStr = WhereStr & "   WHERE (IVD2.DOCUMENT_DATE <= '" & DateToStringIntLow(TO_DATE1) & "') "
      End If
      WhereStr = WhereStr & "   AND (LI2.PART_ITEM_ID = LI.PART_ITEM_ID) "
      WhereStr = WhereStr & "   AND (LI2.LOCATION_ID = LI.LOCATION_ID) "
      WhereStr = WhereStr & ") "
   ElseIf Ind = 5 Then
      SelectStr1 = "SELECT DISTINCT PT.PART_TYPE_ID "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "

      WhereStr = ""
   ElseIf Ind = 6 Then
      SelectStr1 = "SELECT DISTINCT LI.PART_ITEM_ID "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "

      WhereStr = ""
   ElseIf (Ind = 7) Then
      SelectStr1 = "SELECT IVD.SUPPLIER_ID, SUM(LI.TX_AMOUNT) TX_AMOUNT, SUM(TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "

      WhereStr = ""
   ElseIf Ind = 8 Then
      SelectStr1 = "SELECT LI.PART_ITEM_ID, LI.LOCATION_ID, LI.LAYOUT_ID, SUM(LI.LEFT_AMOUNT) LEFT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "

      WhereStr = "WHERE (LI.LAYOUT_ID IS NOT NULL) "
   ElseIf Ind = 9 Then
      SelectStr1 = "SELECT LI.*, PT.*, LC.*, PI.*, UN.*, IVD.DOCUMENT_DATE, IVD.DOCUMENT_NO, IVD.COMMIT_FLAG, IVD.DOCUMENT_TYPE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      
      WhereStr = "WHERE (LI.LEFT_AMOUNT > 0) "
   ElseIf Ind = 10 Then
      SelectStr1 = "SELECT DISTINCT LI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "
      
      WhereStr = ""
   ElseIf Ind = 11 Then
      SelectStr1 = "SELECT LI.*, PT.*, LC1.LOCATION_NAME FROM_LOCATION, LC2.LOCATION_NAME TO_LOCATION, PI.*, UN.*, IVD.DOCUMENT_DATE, IVD.DOCUMENT_NO "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN LOT_ITEM LI2 ON ((LI.GUI_ID = LI2.GUI_ID) AND (LI2.TX_TYPE = 'I')) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC1 ON (LI.LOCATION_ID = LC1.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC2 ON (LI2.LOCATION_ID = LC2.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      
      WhereStr = ""
   ElseIf Ind = 12 Then
      SelectStr1 = "SELECT LI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, LI.LOCATION_ID, SUM(LI.WEIGHT_AMOUNT) WEIGHT_AMOUNT, SUM(LI.PACKAGE_AMOUNT) PACKAGE_AMOUNT, SUM(LI.ACTUAL_PKG_AMOUNT) ACTUAL_PKG_AMOUNT "
      SelectStr1 = SelectStr1 & ", SUM(LI.PACKAGE_WEIGHT) PACKAGE_WEIGHT, SUM(LI.HUMID_WEIGHT) HUMID_WEIGHT, SUM(LI.OTHER_WEIGHT) OTHER_WEIGHT "
      SelectStr1 = SelectStr1 & ", SUM(LI.TOTAL_WEIGHT) TOTAL_WEIGHT,  SUM(LI.TX_AMOUNT) TX_AMOUNT,  SUM(LI.TOTAL_ACTUAL_PRICE) TOTAL_ACTUAL_PRICE,  SUM(LI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE  "
      SelectStr1 = SelectStr1 & ", SUM(LI.CALCULATE_WEIGHT) CALCULATE_WEIGHT , SUM(LI.EXPENSE1) EXPENSE1, SUM(LI.EXPENSE2) EXPENSE2 "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      
      WhereStr = ""
   ElseIf Ind = 13 Then
      SelectStr1 = "SELECT IVD.SUPPLIER_ID,SP.SUPPLIER_CODE, SUM(LI.WEIGHT_AMOUNT) WEIGHT_AMOUNT, SUM(LI.PACKAGE_AMOUNT) PACKAGE_AMOUNT "
      SelectStr1 = SelectStr1 & ", SUM(LI.PACKAGE_WEIGHT) PACKAGE_WEIGHT, SUM(LI.HUMID_WEIGHT) HUMID_WEIGHT, SUM(LI.OTHER_WEIGHT) OTHER_WEIGHT "
      SelectStr1 = SelectStr1 & ", SUM(LI.TOTAL_WEIGHT) TOTAL_WEIGHT,  SUM(LI.TX_AMOUNT) TX_AMOUNT,  SUM(LI.TOTAL_ACTUAL_PRICE) TOTAL_ACTUAL_PRICE,  SUM(LI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE  "
      SelectStr1 = SelectStr1 & ", SUM(LI.CALCULATE_WEIGHT) CALCULATE_WEIGHT , SUM(LI.EXPENSE1) EXPENSE1, SUM(LI.EXPENSE2) EXPENSE2 "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (IVD.SUPPLIER_ID = SP.SUPPLIER_ID) "
      
      WhereStr = ""
   ElseIf Ind = 14 Then
      SelectStr1 = "SELECT DISTINCT LI.LOCATION_ID, LC.LOCATION_NO, LC.LOCATION_NAME, LI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
      SelectStr1 = SelectStr1 & ",PT.PART_TYPE_NO, PT.PART_TYPE_NAME "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "

      WhereStr = ""
   ElseIf Ind = 15 Then
      SelectStr1 = "SELECT LI.LOCATION_ID, LI.PART_ITEM_ID, LI.TX_TYPE, SUM(LI.TX_AMOUNT) TX_AMOUNT, SUM(LI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "

      WhereStr = ""
   ElseIf Ind = 16 Then
      SelectStr1 = "SELECT LI.PART_ITEM_ID, LI.TX_TYPE, SUM(LI.TX_AMOUNT) TX_AMOUNT, SUM(LI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
'      SelectStr1 = "SELECT LI.PART_ITEM_ID, LI.TX_TYPE, SUM(LI.TX_AMOUNT) TX_AMOUNT, SUM(cast(LI.TOTAL_INCLUDE_PRICE as numeric(12,6))) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "

      WhereStr = ""
   ElseIf Ind = 17 Then
      SelectStr1 = "SELECT LI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, SUM(LI.TX_AMOUNT) TX_AMOUNT, SUM(LI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN JOB JB ON (IVD.INVENTORY_DOC_ID = JB.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "
      
      WhereStr = ""
   ElseIf Ind = 18 Then
      SelectStr1 = "SELECT LI.PART_ITEM_ID, LI.TX_TYPE, IVD.DOCUMENT_TYPE, SUM(LI.TX_AMOUNT) TX_AMOUNT, SUM(LI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "

      WhereStr = ""
   ElseIf Ind = 19 Then
      SelectStr1 = "SELECT LI.PART_ITEM_ID PART_ITEM_ID, PI.PART_NO PART_NO, PI.PART_DESC PART_DESC,  JB.PART_ITEM_ID PRODUCT_ID, PI2.PART_NO PRODUCT_NO, PI2.PART_DESC PRODUCT_NAME, SUM(LI.TX_AMOUNT) TX_AMOUNT, SUM(LI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN JOB JB ON (IVD.INVENTORY_DOC_ID = JB.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI2 ON (JB.PART_ITEM_ID = PI2.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "

      WhereStr = ""
   ElseIf Ind = 20 Then
      SelectStr1 = "SELECT LI.LOCATION_ID, LI.PART_ITEM_ID, LI.TX_TYPE, IVD.DOCUMENT_TYPE, SUM(LI.TX_AMOUNT) TX_AMOUNT, SUM(LI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "

      WhereStr = ""
   ElseIf Ind = 21 Then
      SelectStr1 = "SELECT LI.*, PI.PART_NO, PI.PART_DESC, MR1.KEY_NAME EXPENSE_NAME, LT.LAY_OUT_NAME DEP_NAME, LT.LAY_OUT_NO DEP_NO, IVD.DOCUMENT_DATE, IVD.DOCUMENT_NO, UN.UNIT_NAME  "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME ITEM_DESC_NAME "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (LI.EXPENSE_TYPE = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LAY_OUT LT ON (LI.TO_DEPARTMENT = LT.LAY_OUT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
   
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (MR2.KEY_ID = LI.ITEM_DESC_ID) "
      
      WhereStr = ""
   ElseIf Ind = 22 Then
      SelectStr1 = "SELECT DISTINCT PI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, PI.PART_TYPE, PT.PART_TYPE_NO, PT.PART_TYPE_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "
      
      WhereStr = "WHERE (LI.LOCATION_ID = " & LOCATION_ID & ") "
   ElseIf Ind = 23 Then
      SelectStr1 = "SELECT LI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, UT.UNIT_NAME, LI.EXPENSE_TYPE, LI.TO_DEPARTMENT, MR1.KEY_NAME EXPENSE_NAME, LT.LAY_OUT_NAME DEP_NAME, SUM(LI.TX_AMOUNT) TX_AMOUNT, SUM(LI.TOTAL_INCLUDE_PRICE)  TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (LI.EXPENSE_TYPE = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LAY_OUT LT ON (LI.TO_DEPARTMENT = LT.LAY_OUT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UT ON (PI.UNIT_COUNT = UT.UNIT_ID) "

      WhereStr = ""
   ElseIf Ind = 24 Then
      SelectStr1 = "SELECT LI.PART_ITEM_ID, LI.TX_TYPE, SUM(LI.TX_AMOUNT) TX_AMOUNT, SUM(LI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr1 = SelectStr1 & ",CM.CUSTOMER_TYPE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (BD.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (BD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      
      WhereStr = " WHERE (IVD.DOCUMENT_TYPE IN ( 10,21 ))  "
   ElseIf Ind = 25 Then
      SelectStr1 = "SELECT DISTINCT CM.CUSTOMER_TYPE, CT.CSTTYPE_NAME "
      SelectStr2 = "SELECT COUNT(DISTINCT CM.CUSTOMER_TYPE)  "
      
      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (BD.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (BD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_TYPE CT ON (CM.CUSTOMER_TYPE = CT.CSTTYPE_ID) "
      
      WhereStr = " WHERE (IVD.DOCUMENT_TYPE IN ( 10,21 ) AND TX_AMOUNT > 0 )  "
   ElseIf Ind = 26 Then
      SelectStr1 = "SELECT IVD.DOCUMENT_DATE,IVD.DOCUMENT_NO,NM1.LONG_NAME SUPPLIER_NAME, SUM(LI.TX_AMOUNT) TX_AMOUNT, SUM(LI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (SN.NAME_ID = NM1.NAME_ID) "
      
      WhereStr = ""
   ElseIf Ind = 27 Then
      SelectStr1 = "SELECT LI.PART_ITEM_ID, SUM(LI.TX_AMOUNT) TX_AMOUNT, SUM(LI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr1 = SelectStr1 & ",IVD.YYYYMM "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "
      
      WhereStr = ""
   ElseIf Ind = 28 Then
      SelectStr1 = "SELECT LI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, SUM(LI.WEIGHT_AMOUNT) WEIGHT_AMOUNT, SUM(LI.PACKAGE_AMOUNT) PACKAGE_AMOUNT, SUM(LI.ACTUAL_PKG_AMOUNT) ACTUAL_PKG_AMOUNT "
      SelectStr1 = SelectStr1 & ", SUM(LI.PACKAGE_WEIGHT) PACKAGE_WEIGHT, SUM(LI.HUMID_WEIGHT) HUMID_WEIGHT, SUM(LI.OTHER_WEIGHT) OTHER_WEIGHT "
      SelectStr1 = SelectStr1 & ", SUM(LI.TOTAL_WEIGHT) TOTAL_WEIGHT,  SUM(LI.TX_AMOUNT) TX_AMOUNT,  SUM(LI.TOTAL_ACTUAL_PRICE) TOTAL_ACTUAL_PRICE,  SUM(LI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE  "
      SelectStr1 = SelectStr1 & ", SUM(LI.CALCULATE_WEIGHT) CALCULATE_WEIGHT , SUM(LI.EXPENSE1) EXPENSE1, SUM(LI.EXPENSE2) EXPENSE2 "
      SelectStr1 = SelectStr1 & ",IVD.SUPPLIER_ID,SP.SUPPLIER_CODE, NM1.LONG_NAME SUPPLIER_NAME  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = IVD.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (SN.NAME_ID = NM1.NAME_ID) "
      
      WhereStr = ""
   ElseIf Ind = 29 Then
      SelectStr1 = "SELECT LI.PART_ITEM_ID, LI.TX_TYPE,IVD.DOCUMENT_DATE, SUM(LI.TX_AMOUNT) TX_AMOUNT, SUM(LI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "

      WhereStr = ""
      
       ElseIf Ind = 30 Then
      SelectStr1 = "SELECT LI.*, PI.PART_NO, PI.PART_DESC, MR1.KEY_NAME EXPENSE_NAME, LT.LAY_OUT_NAME DEP_NAME, LT.LAY_OUT_NO DEP_NO, IVD.DOCUMENT_DATE, IVD.DOCUMENT_NO, UN.UNIT_NAME  "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME ITEM_DESC_NAME "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (LI.EXPENSE_TYPE = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LAY_OUT LT ON (LI.TO_DEPARTMENT = LT.LAY_OUT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
   
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (MR2.KEY_ID = LI.ITEM_DESC_ID) "
      
      WhereStr = ""
      
      
        ElseIf Ind = 31 Then
      SelectStr1 = "SELECT LI.*, PI.PART_NO, PI.PART_DESC, MR1.KEY_NAME EXPENSE_NAME, LT.LAY_OUT_NAME DEP_NAME, LT.LAY_OUT_NO DEP_NO, IVD.DOCUMENT_DATE, IVD.DOCUMENT_NO, UN.UNIT_NAME  "
      SelectStr1 = SelectStr1 & ",MR2.KEY_NAME ITEM_DESC_NAME,MR3.KEY_NAME PROJECT_NAME,MR3.KEY_ID PROJECT_ID  "

      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR1 ON (LI.EXPENSE_TYPE = MR1.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LAY_OUT LT ON (LI.TO_DEPARTMENT = LT.LAY_OUT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
   
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR2 ON (MR2.KEY_ID = LI.ITEM_DESC_ID) "
     FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MR3 ON(MR3.KEY_ID = LI.PROJECT_NAME_ID) "
      
      WhereStr = ""
      
   ElseIf Ind = 32 Then
      SelectStr1 = "SELECT LI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, LI.LOCATION_ID, SUM(LI.WEIGHT_AMOUNT) WEIGHT_AMOUNT, SUM(LI.PACKAGE_AMOUNT) PACKAGE_AMOUNT, SUM(LI.ACTUAL_PKG_AMOUNT) ACTUAL_PKG_AMOUNT "
      SelectStr1 = SelectStr1 & ", SUM(LI.PACKAGE_WEIGHT) PACKAGE_WEIGHT, SUM(LI.HUMID_WEIGHT) HUMID_WEIGHT, SUM(LI.OTHER_WEIGHT) OTHER_WEIGHT "
      SelectStr1 = SelectStr1 & ", SUM(LI.TOTAL_WEIGHT) TOTAL_WEIGHT,  SUM(LI.TX_AMOUNT) TX_AMOUNT,  SUM(LI.TOTAL_ACTUAL_PRICE) TOTAL_ACTUAL_PRICE,  SUM(LI.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE  "
      SelectStr1 = SelectStr1 & ", SUM(LI.CALCULATE_WEIGHT) CALCULATE_WEIGHT , SUM(LI.EXPENSE1) EXPENSE1, SUM(LI.EXPENSE2) EXPENSE2 "
       SelectStr1 = SelectStr1 & ",UN.UNIT_NAME   "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LI.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = IVD.SUPPLIER_ID) "
        FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID)"
        
      WhereStr = ""
     ElseIf Ind = 33 Then ' มาจาก  Ind = 32
      SelectStr1 = "SELECT LI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC    "
      SelectStr1 = SelectStr1 & ", SUM(LI.TX_AMOUNT) TX_AMOUNT   "
   
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
'      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      
      WhereStr = ""
   ElseIf Ind = 34 Then 'LoadSumJobInputViaLotItem
      SelectStr1 = "SELECT LI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC,IVD.DOCUMENT_DATE,SUM(LI.TX_AMOUNT) TX_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      
      FromStr = "FROM LOT_ITEM LI "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (LI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (LI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
            
      WhereStr = " WHERE (IVD.DOCUMENT_TYPE = 12 OR IVD.DOCUMENT_TYPE = 13 OR IVD.DOCUMENT_TYPE = 14 ) "
      
   End If
   
   If LOT_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LOT_ITEM_ID = " & LOT_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LOT_ITEM_ID = " & LOT_ITEM_ID & ") "
      End If
   End If
   
   If INVENTORY_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LI.INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LI.INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") "
      End If
   End If
   
   If LOCATION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LI.LOCATION_ID = " & LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LI.LOCATION_ID = " & LOCATION_ID & ") "
      End If
   End If
   
   If FROM_LOCATION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LI.LOCATION_ID = " & FROM_LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LI.LOCATION_ID = " & FROM_LOCATION_ID & ") "
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
         WhereStr = "WHERE (LI.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LI.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   If Len(COMMIT_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
   End If
   
   If Len(TX_TYPE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LI.TX_TYPE = '" & TX_TYPE & "') "
      Else
         WhereStr = WhereStr & "AND (LI.TX_TYPE = '" & TX_TYPE & "') "
      End If
   End If
   
   If Len(FROM_TX_TYPE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LI.TX_TYPE = '" & FROM_TX_TYPE & "') "
      Else
         WhereStr = WhereStr & "AND (LI.TX_TYPE = '" & FROM_TX_TYPE & "') "
      End If
   End If
   
   If DOCUMENT_TYPE > 0 Then
      If Ind = 4 Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (IVD2.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
         Else
            WhereStr = WhereStr & "AND (IVD2.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
         End If
      Else
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (IVD.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
         Else
            WhereStr = WhereStr & "AND (IVD.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
         End If
      End If
   End If
   
   If Len(DOCUMENT_TYPE_SET) > 0 Then
      If Ind = 4 Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (IVD2.DOCUMENT_TYPE IN " & DOCUMENT_TYPE_SET & ") "
         Else
            WhereStr = WhereStr & "AND (IVD2.DOCUMENT_TYPE IN " & DOCUMENT_TYPE_SET & ") "
         End If
      Else
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (IVD.DOCUMENT_TYPE IN " & DOCUMENT_TYPE_SET & ") "
         Else
            WhereStr = WhereStr & "AND (IVD.DOCUMENT_TYPE IN " & DOCUMENT_TYPE_SET & ") "
         End If
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
   
   If SUPPLIER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUPPLIER_ID = " & SUPPLIER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SUPPLIER_ID = " & SUPPLIER_ID & ") "
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
         WhereStr = "WHERE (LI.TO_DEPARTMENT = " & LAYOUT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LI.TO_DEPARTMENT = " & LAYOUT_ID & ") "
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
      GroupByStr = "GROUP BY LI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, LI.LOCATION_ID "
   ElseIf Ind = 13 Then
      GroupByStr = "GROUP BY IVD.SUPPLIER_ID,SP.SUPPLIER_CODE "
   ElseIf Ind = 15 Then
      GroupByStr = "GROUP BY LI.PART_ITEM_ID, LI.LOCATION_ID, LI.TX_TYPE "
   ElseIf Ind = 16 Then
      GroupByStr = "GROUP BY LI.PART_ITEM_ID, LI.TX_TYPE "
   ElseIf Ind = 17 Then
      GroupByStr = "GROUP BY LI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
   ElseIf Ind = 18 Then
      GroupByStr = "GROUP BY LI.PART_ITEM_ID, LI.TX_TYPE, IVD.DOCUMENT_TYPE "
   ElseIf Ind = 19 Then
      GroupByStr = "GROUP BY LI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, JB.PART_ITEM_ID, PI2.PART_NO, PI2.PART_DESC "
   ElseIf Ind = 20 Then
      GroupByStr = "GROUP BY LI.LOCATION_ID, LI.PART_ITEM_ID, LI.TX_TYPE, IVD.DOCUMENT_TYPE "
   ElseIf Ind = 23 Then
      GroupByStr = "GROUP BY LI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, UT.UNIT_NAME, LI.EXPENSE_TYPE, LI.TO_DEPARTMENT, MR1.KEY_NAME, LT.LAY_OUT_NAME "
   ElseIf Ind = 24 Then
      GroupByStr = "GROUP BY LI.PART_ITEM_ID, LI.TX_TYPE, IVD.DOCUMENT_TYPE,CM.CUSTOMER_TYPE "
   ElseIf Ind = 26 Then
      GroupByStr = "GROUP BY IVD.DOCUMENT_DATE , IVD.DOCUMENT_NO, NM1.LONG_NAME "
   ElseIf Ind = 27 Then
      GroupByStr = "GROUP BY LI.PART_ITEM_ID,IVD.YYYYMM "
   ElseIf Ind = 28 Then
      GroupByStr = "GROUP BY LI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC,IVD.SUPPLIER_ID,SP.SUPPLIER_CODE,NM1.LONG_NAME "
   ElseIf Ind = 29 Then
      GroupByStr = "GROUP BY LI.PART_ITEM_ID, LI.TX_TYPE,IVD.DOCUMENT_DATE "
    ElseIf Ind = 32 Then
      GroupByStr = "GROUP BY LI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, LI.LOCATION_ID , UN.UNIT_NAME "
   ElseIf Ind = 33 Then
      GroupByStr = "GROUP BY LI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC   "
   ElseIf Ind = 34 Then
      GroupByStr = "GROUP BY LI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC,IVD.DOCUMENT_DATE"
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY LOT_ITEM_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY LOT_ITEM_ID "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY DOCUMENT_DATE DESC, TRANSACTION_SEQ "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY LOCATION_ID DESC, DOCUMENT_DATE DESC, TRANSACTION_SEQ "
      ElseIf OrderBy = 5 Then
         OrderByStr = "ORDER BY PT.PART_TYPE_NO "
      ElseIf OrderBy = 6 Then
         OrderByStr = "ORDER BY PI.PART_NO ASC, IVD.DOCUMENT_NO "
      ElseIf OrderBy = 7 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE,IVD.DOCUMENT_NO "
      ElseIf OrderBy = 8 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_NO "
      ElseIf OrderBy = 9 Then
         OrderByStr = "ORDER BY LC.LOCATION_NO "
      ElseIf OrderBy = 10 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE ASC, LI.TRANSACTION_SEQ  "
      ElseIf OrderBy = 11 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE ASC, DP.PRIORITY1 ASC, IVD.DOCUMENT_NO ASC, LI.TRANSACTION_SEQ  "
      ElseIf OrderBy = 12 Then
         OrderByStr = "ORDER BY SP.SUPPLIER_CODE ASC, IVD.DOCUMENT_NO "
      ElseIf OrderBy = 13 Then
         OrderByStr = "ORDER BY SN.LONG_NAME "
      ElseIf OrderBy = 14 Then
         OrderByStr = "ORDER BY PT.PART_TYPE_NO ASC, PI.PART_NO ASC, PI.PART_ITEM_ID, IVD.DOCUMENT_DATE ASC, DP.PRIORITY1 ASC, IVD.DOCUMENT_NO ASC, LI.TRANSACTION_SEQ  "
      ElseIf OrderBy = 15 Then
         OrderByStr = "ORDER BY PI.PART_NO ASC, IVD.DOCUMENT_DATE ASC, IVD.DOCUMENT_NO "
      ElseIf OrderBy = 16 Then
         OrderByStr = "ORDER BY PI.PART_DESC ASC, PI.PART_NO ASC, IVD.DOCUMENT_DATE ASC, IVD.DOCUMENT_NO "
      ElseIf OrderBy = 17 Then
         OrderByStr = "ORDER BY PT.PART_TYPE_NO ASC, SP.SUPPLIER_CODE ASC,PI.PART_NO ASC, IVD.DOCUMENT_DATE ASC, IVD.DOCUMENT_NO  "
      Else
         OrderByStr = "ORDER BY LOT_ITEM_ID "
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
         OrderByStr = "ORDER BY LOT_ITEM_ID "
      End If
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY LI.PART_ITEM_ID "
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY LI.PART_ITEM_ID, LOCATION_ID "
   ElseIf Ind = 5 Then
      OrderByStr = "ORDER BY PI.PART_TYPE "
   ElseIf Ind = 6 Then
      OrderByStr = "ORDER BY PT.PART_TYPE_NAME ASC, PI.PART_NO "
   ElseIf Ind = 7 Then
      OrderByStr = "ORDER BY IVD.SUPPLIER_ID "
   ElseIf Ind = 8 Then
      OrderByStr = "ORDER BY LI.LAYOUT_ID "
   ElseIf Ind = 9 Then
      OrderByStr = "ORDER BY LI.PART_ITEM_ID "
   ElseIf Ind = 10 Then
      OrderByStr = "ORDER BY PI.PART_NO "
   ElseIf Ind = 11 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY LOT_ITEM_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY LOT_ITEM_ID "
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
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE ASC, LI.TRANSACTION_SEQ  "
      ElseIf OrderBy = 11 Then
         OrderByStr = "ORDER BY LI.TRANSACTION_SEQ  "
      Else
         OrderByStr = "ORDER BY LOT_ITEM_ID "
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
      OrderByStr = "ORDER BY LI.LOCATION_ID, LI.PART_ITEM_ID "
   ElseIf Ind = 16 Then
      OrderByStr = "ORDER BY LI.PART_ITEM_ID "
   ElseIf Ind = 17 Then
      OrderByStr = "ORDER BY PI.PART_NO "
   ElseIf Ind = 18 Or Ind = 27 Then
      OrderByStr = "ORDER BY LI.PART_ITEM_ID "
   ElseIf Ind = 19 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PI2.PART_NO ASC, PI.PART_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PI.PART_NO ASC, PI2.PART_NO "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY PI2.PART_NO "
      End If
   ElseIf Ind = 20 Then
      OrderByStr = "ORDER BY LI.PART_ITEM_ID "
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
   ElseIf Ind = 26 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_NO "
      Else
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE ASC,IVD.DOCUMENT_NO "
      End If
   ElseIf Ind = 28 Then
      OrderByStr = "ORDER BY Sp.SUPPLIER_CODE ASC,PI.PART_NO "
   ElseIf Ind = 29 Then
      OrderByStr = "ORDER BY LI.PART_ITEM_ID "
      
   ElseIf Ind = 30 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY LT.LAY_OUT_NAME ASC, MR1.KEY_NAME "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE ASC, IVD.DOCUMENT_NO ASC, LT.LAY_OUT_NAME ASC, MR1.KEY_NAME "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY PI.PART_NO ASC, IVD.DOCUMENT_DATE ASC, IVD.DOCUMENT_NO "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY LT.LAY_OUT_NAME ASC,ITEM_DESC ASC,MR1.KEY_NAME ASC, PI.PART_NO ASC, IVD.DOCUMENT_DATE ASC, IVD.DOCUMENT_NO "
      Else
         OrderByStr = "ORDER BY MR1.KEY_NAME ASC, LT.LAY_OUT_NAME "
      End If

   ElseIf Ind = 31 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY LT.LAY_OUT_NAME ASC, MR1.KEY_NAME "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE ASC, IVD.DOCUMENT_NO ASC, LT.LAY_OUT_NAME ASC, MR1.KEY_NAME "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY PI.PART_NO ASC, IVD.DOCUMENT_DATE ASC, IVD.DOCUMENT_NO "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY LT.LAY_OUT_NAME ASC, MR3.KEY_ID ASC,MR1.KEY_NAME ASC, PI.PART_NO ASC, IVD.DOCUMENT_DATE ASC, IVD.DOCUMENT_NO "
      Else
         OrderByStr = "ORDER BY MR1.KEY_NAME ASC, LT.LAY_OUT_NAME "
      End If
      
      
    ElseIf Ind = 32 Then
      OrderByStr = "ORDER BY PI.PART_NO "
    ElseIf Ind = 33 Then
      OrderByStr = "ORDER BY PI.PART_NO "
   End If
   
   If Not (Ind = 7 Or Ind = 24 Or Ind = 25) Then
      If OrderType = 1 Then
         OrderTypeStr = "ASC "
      ElseIf OrderType = 2 Then
         OrderTypeStr = "DESC "
      Else
         OrderTypeStr = "ASC "
      End If
   End If
   
     If Ind = 33 Then  'สำหรับไม่ต้องการให้ เรียงเพื่อทำให้เร็วขึ้น
      OrderByStr = ""
      OrderTypeStr = ""
   End If
   
   If Len(OrderByStr) <= 0 Then
      OrderTypeStr = ""
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      LOT_ITEM_ID = NVLI(Rs("LOT_ITEM_ID"), -1)
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
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
      ADJUST_FLAG = NVLS(Rs("ADJUST_FLAG"), "")
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
      
      
       PROJECT_NAME_ID = NVLI(Rs("PROJECT_NAME_ID"), 0)
'       PROJECT_NAME = NVLS(Rs("PROJECT_NAME "), "")
'       PROJECT_ID = NVLI(Rs("PROJECT_ID "), 0)
      
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
      QUE_NO = NVLS(Rs("QUE_NO"), 0)
      TRUCK_NO = NVLS(Rs("TRUCK_NO"), "")
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), -1)
      SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
      SUPPLIER_NAME = NVLS(Rs("SUPPLIER_NAME"), "")
      
      PR_NO = NVLS(Rs("PR_NO"), "")
      
      DEPARTMENT_NAME = NVLS(Rs("DEPARTMENT_NAME"), "")
      
      ENTRY_DATE = InternalDateToDate(NVLS(Rs("ENTRY_DATE"), ""))
      EXIT_DATE = InternalDateToDate(NVLS(Rs("EXIT_DATE"), ""))
      
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
      LOT_ITEM_ID = NVLI(Rs("LOT_ITEM_ID"), -1)
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
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
      LOT_ITEM_ID = NVLI(Rs("LOT_ITEM_ID"), -1)
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      PACKAGING_AMT = NVLD(Rs("PACKAGING_AMT"), 0)
     
      PART_TYPE = NVLI(Rs("PART_TYPE"), -1)
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
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
    '  PROJECT_NAME = NVLS(Rs("PROJECT_NAME"), "")
     ' PROJECT_ID = NVLI(Rs("PROJECT_ID"), 0)
     
     
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
   ElseIf Ind = 26 Then
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      SUPPLIER_NAME = NVLS(Rs("SUPPLIER_NAME"), "")
      
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
   ElseIf Ind = 27 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
   ElseIf Ind = 28 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
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
      
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), -1)
      SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
      SUPPLIER_NAME = NVLS(Rs("SUPPLIER_NAME"), "")
   ElseIf Ind = 29 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      
      
     ElseIf Ind = 30 Then
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
      PROJECT_NAME = NVLS(Rs("PROJECT_NAME"), "")
      PROJECT_ID = NVLI(Rs("PROJECT_ID"), 0)
      ElseIf Ind = 32 Then
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
       UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
   ElseIf Ind = 33 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
   ElseIf Ind = 34 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
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
   
   TempID8 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID8 = "NULL"
   End If
   
   TempID9 = PROJECT_NAME_ID
   If PROJECT_NAME_ID <= 0 Then
     TempID9 = "NULL"
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
      Call glbDatabaseMngr.GetSeqID("LOT_ITEM_SEQ", ID, glbErrorLog)
      Call glbDatabaseMngr.GetSeqID("TRANSACTION_SEQUENCE_SEQ", ID2, glbErrorLog)
      
      SQL1 = "INSERT INTO LOT_ITEM " & _
                     "( " & _
                     "   LOT_ITEM_ID, " & _
                     "   INVENTORY_DOC_ID, " & _
                     "   PART_ITEM_ID, " & _
                     "   PROJECT_NAME_ID, " & _
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
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "" & INVENTORY_DOC_ID & "," & _
                         TempID8 & "," & _
                          TempID9 & "," & _
                         "" & LOCATION_ID & "," & _
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
                        CURRENT_PRICE & "," & _
                        CURRENT_AMOUNT & "," & _
                        NEED_TOTAL_AMOUNT & "," & _
                        NEED_TOTAL_PRICE & "," & _
                        NEED_AVG_PRICE & "," & _
                        "'" & Trim(ChangeQuote(MANUAL_PRICE)) & "'," & _
                        TempID6 & "," & _
                        "'" & Trim(ChangeQuote(AUTO_PRICE)) & "'," & _
                        TempID7 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         LOT_ITEM_ID = ID
   Else 'Show edit
      Call glbDatabaseMngr.GetSeqID("TRANSACTION_SEQUENCE_SEQ", ID2, glbErrorLog)
      
      SQL1 = "UPDATE LOT_ITEM SET " & _
                     "   INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ", " & _
                     "   PART_ITEM_ID = " & TempID8 & ", " & _
                      "   PROJECT_NAME_ID = " & TempID9 & ", " & _
                     "   LOCATION_ID = " & LOCATION_ID & ", " & _
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
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE LOT_ITEM_ID = " & LOT_ITEM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM SUBLOT_ITEM WHERE LOT_ITEM_ID = " & LOT_ITEM_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM LOT_ITEM WHERE LOT_ITEM_ID = " & LOT_ITEM_ID
      m_Conn.Execute (SQL1)
End Function

Public Function DecreaseLeftAmount(DecAmount As Double)
Dim SQL1 As String
      SQL1 = "UPDATE LOT_ITEM SET LEFT_AMOUNT = LEFT_AMOUNT - " & DecAmount & " WHERE LOT_ITEM_ID = " & LOT_ITEM_ID
      m_Conn.Execute (SQL1)
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

Public Function PatchAvgPriceExp(AvgPrice As Double, PrevAmount As Double, CurrentAmount As Double, ExportTotalPrice As Double, NewTotalPrice As Double) As Boolean
On Error GoTo ErrorHandler
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL1 = "UPDATE LOT_ITEM SET " & _
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
                  "WHERE LOT_ITEM_ID = " & LOT_ITEM_ID

   m_Conn.Execute (SQL1)
   Exit Function
   
ErrorHandler:
   SQL1 = "UPDATE LOT_ITEM SET " & _
                  "   INCLUDE_UNIT_PRICE = " & 0 & ", " & _
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
                  "WHERE LOT_ITEM_ID = " & LOT_ITEM_ID

   m_Conn.Execute (SQL1)
End Function

Public Function PatchAvgPrice(AvgPrice As Double, PrevAmount As Double, CurrentAmount As Double, NewPrice As Double, ImportAmount As Double, DocumentType As Long, TotalIncludePrice As Double, NewTotalPrice As Double) As Boolean
On Error GoTo ErrorHandler
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   If (DocumentType = 12) Or (DocumentType = 13) Or (DocumentType = 14) Or (DocumentType = 3) Or (DocumentType = 22) Or (DocumentType = 4) Or (DocumentType = 5) Then
      SQL1 = "UPDATE LOT_ITEM SET " & _
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
                     "WHERE LOT_ITEM_ID = " & LOT_ITEM_ID
   Else
      SQL1 = "UPDATE LOT_ITEM SET " & _
                     "   PREVIOUS_AMOUNT = " & PrevAmount & ", " & _
                     "   NEW_AMOUNT = " & CurrentAmount & ", " & _
                     "   NEW_PRICE = " & NewPrice & ", " & _
                     "   RAW_COST = " & RAW_COST & ", " & _
                     "   EXPENSE_COST = " & EXPENSE_COST & ", " & _
                     "   TOTAL_NEW_PRICE = " & NewTotalPrice & ", " & _
                     "   TX_TYPE = '" & TX_TYPE & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE LOT_ITEM_ID = " & LOT_ITEM_ID
   End If
   m_Conn.Execute (SQL1)
   Exit Function
   
ErrorHandler:

   If (DocumentType = 3) Or (DocumentType = 22) Or (DocumentType = 12) Or (DocumentType = 13) Or (DocumentType = 14) Then
      SQL1 = "UPDATE LOT_ITEM SET " & _
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
                     "WHERE LOT_ITEM_ID = " & LOT_ITEM_ID
   Else
      SQL1 = "UPDATE LOT_ITEM SET " & _
                     "   PREVIOUS_AMOUNT = " & PrevAmount & ", " & _
                     "   NEW_AMOUNT = " & CurrentAmount & ", " & _
                     "   NEW_PRICE = " & 0 & ", " & _
                     "   RAW_COST = " & 0 & ", " & _
                     "   EXPENSE_COST = " & 0 & ", " & _
                     "   TOTAL_NEW_PRICE = " & 0 & ", " & _
                     "   TX_TYPE = '" & TX_TYPE & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE LOT_ITEM_ID = " & LOT_ITEM_ID
   End If
   m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set SubLotItems = Nothing
End Sub
Public Function CopyObjectFromSup(Ind As Long, TempObj As CSupItem)

      PART_ITEM_ID = TempObj.PART_ITEM_ID                                                             '2
      LOCATION_ID = TempObj.LOCATION_ID                                                             '2
      
      ACTUAL_UNIT_PRICE = TempObj.ACTUAL_UNIT_PRICE                                                             '2
      INCLUDE_UNIT_PRICE = TempObj.INCLUDE_UNIT_PRICE                                                             '2
      PREVIOUS_PRICE = TempObj.PREVIOUS_PRICE                                                             '2
      PREVIOUS_AMOUNT = TempObj.PREVIOUS_AMOUNT                                                             '2
      NEW_PRICE = TempObj.NEW_PRICE                                                             '2
      TX_AMOUNT = TempObj.TX_AMOUNT                                                             '2
      TRANSACTION_SEQ = TempObj.TRANSACTION_SEQ                                                             '2
      GUI_ID = TempObj.GUI_ID                                                             '2
      
      CALCULATE_FLAG = TempObj.CALCULATE_FLAG                                                             '2
      TOTAL_ACTUAL_PRICE = TempObj.TOTAL_ACTUAL_PRICE                                                             '2
      TOTAL_INCLUDE_PRICE = TempObj.TOTAL_INCLUDE_PRICE                                                               '2
      TX_TYPE = TempObj.TX_TYPE                                                             '2
      LEFT_AMOUNT = TempObj.LEFT_AMOUNT                                                             '2
      LAYOUT_ID = TempObj.LAYOUT_ID                                                             '2
      LINK_ID = TempObj.LINK_ID                                                             '2
      PACKAGING_AMT = TempObj.PACKAGING_AMT                                                             '2
      ENTRY_WEIGHT = TempObj.ENTRY_WEIGHT                                                             '2
      EXIT_WEIGHT = TempObj.EXIT_WEIGHT                                                             '2
      
      WEIGHT_AMOUNT = TempObj.WEIGHT_AMOUNT                                                             '2
      PACKAGE_WEIGHT = TempObj.PACKAGE_WEIGHT                                                             '2
      OTHER_WEIGHT = TempObj.OTHER_WEIGHT                                                             '2
      PERCENT_HUMID = TempObj.PERCENT_HUMID                                                             '2
      HUMID_WEIGHT = TempObj.HUMID_WEIGHT                                                             '2
      PACKAGING_ID = TempObj.PACKAGING_ID                                                             '2
      
      SUPPLIER_WEIGHT = TempObj.SUPPLIER_WEIGHT                                                             '2
      CALCULATE_TYPE = TempObj.CALCULATE_TYPE                                                             '2
      PACKAGE_AMOUNT = TempObj.PACKAGE_AMOUNT                                                             '2
      ACTUAL_PKG_AMOUNT = TempObj.ACTUAL_PKG_AMOUNT                                                             '2
      PUREXP_ID1 = TempObj.PUREXP_ID1                                                             '2
      PUREXP_ID2 = TempObj.PUREXP_ID2                                                             '2
      EXPENSE1 = TempObj.EXPENSE1                                                             '2
      EXPENSE2 = TempObj.EXPENSE2                                                             '2
      TOTAL_WEIGHT = TempObj.TOTAL_WEIGHT                                                             '2
      RAW_PRICE = TempObj.RAW_PRICE                                                             '2
      
      DISCOUNT_AMT = TempObj.DISCOUNT_AMT                                                             '2
      RAW_TOT_PRICE = TempObj.RAW_TOT_PRICE                                                             '2
      TO_DEPARTMENT = TempObj.TO_DEPARTMENT                                                             '2
      ITEM_DESC = TempObj.ITEM_DESC                                                             '2
      EXTRA_NAME = TempObj.EXTRA_NAME                                                             '2
      SALE_TOT_PRICE = TempObj.SALE_TOT_PRICE                                                             '2
      CALCULATE_WEIGHT = TempObj.CALCULATE_WEIGHT                                                             '2
      RAW_COST = TempObj.RAW_COST                                                             '2
      EXPENSE_COST = TempObj.EXPENSE_COST                                                             '2
      TOTAL_NEW_PRICE = TempObj.TOTAL_NEW_PRICE                                                             '2
      BAG_RETURN = TempObj.BAG_RETURN                                                             '2
   
      ACTUAL_AMOUNT = TempObj.ACTUAL_AMOUNT                                                             '2
      ACTUAL_PRICE = TempObj.ACTUAL_PRICE                                                             '2
      CURRENT_AMOUNT = TempObj.CURRENT_AMOUNT                                                             '2
      CURRENT_PRICE = TempObj.CURRENT_PRICE                                                             '2
      NEED_TOTAL_PRICE = TempObj.NEED_TOTAL_PRICE                                                             '2
      NEED_TOTAL_AMOUNT = TempObj.NEED_TOTAL_AMOUNT                                                             '2
      NEED_AVG_PRICE = TempObj.NEED_AVG_PRICE                                                             '2
      MANUAL_PRICE = TempObj.MANUAL_PRICE                                                             '2
      EXPENSE_TYPE = TempObj.EXPENSE_TYPE                                                             '2
      AUTO_PRICE = TempObj.AUTO_PRICE                                                             '2
      ITEM_DESC_ID = TempObj.ITEM_DESC_ID                                                             '2
      ITEM_DESC_NO = TempObj.ITEM_DESC_NO                                                             '2
      
      ITEM_DESC_NAME = TempObj.ITEM_DESC_NAME                                                             '2
      PUREXP_NO1 = TempObj.PUREXP_NO1                                                             '2
      PUREXP_NO2 = TempObj.PUREXP_NO2                                                             '2
      PUREXP_NAME1 = TempObj.PUREXP_NAME1                                                             '2
      PUREXP_NAME2 = TempObj.PUREXP_NAME2                                                             '2
      EXPENSE_RATE1 = TempObj.EXPENSE_RATE1                                                             '2
      EXPENSE_RATE2 = TempObj.EXPENSE_RATE2                                                             '2
      PACKAGING_CODE = TempObj.PACKAGING_CODE                                                             '2
      
      PRIORITY1 = TempObj.PRIORITY1                                                             '2
      EXPORT_AVG_PRICE = TempObj.EXPORT_AVG_PRICE                                                             '2
      EXPORT_TOTAL_PRICE = TempObj.EXPORT_TOTAL_PRICE                                                             '2
      IMPORT_AMOUNT = TempObj.IMPORT_AMOUNT                                                             '2
      OLD_LOCATION_ID = TempObj.OLD_LOCATION_ID                                                             '2
      OLD_TX_AMOUNT = TempObj.OLD_TX_AMOUNT                                                             '2
      OLD_LOCATION_ID = TempObj.OLD_LOCATION_ID                                                             '2
      PACKAGING_CODE = TempObj.PACKAGING_CODE                                                             '2
      PACKAGING_CODE = TempObj.PACKAGING_CODE                                                             '2
      PACKAGING_CODE = TempObj.PACKAGING_CODE                                                             '2
      
      PART_GROUP_ID = TempObj.PART_GROUP_ID                                                             '2
      PART_TYPE = TempObj.PART_TYPE                                                             '2
      PART_TYPE_NAME = TempObj.PART_TYPE_NAME                                                             '2
      PART_TYPE_NO = TempObj.PART_TYPE_NO                                                             '2
      LOCATION_NO = TempObj.LOCATION_NO                                                             '2
      LOCATION_NAME = TempObj.LOCATION_NAME                                                             '2
      PART_NO = TempObj.PART_NO                                                             '2
      PART_DESC = TempObj.PART_DESC                                                             '2
      PIG_FLAG = TempObj.PIG_FLAG                                                             '2
      PIG_TYPE = TempObj.PIG_TYPE                                                             '2
      UNIT_NAME = TempObj.UNIT_NAME                                                             '2
      
      DOCUMENT_DATE = TempObj.DOCUMENT_DATE                                                             '2
      DOCUMENT_NO = TempObj.DOCUMENT_NO                                                             '2
      DO_NO = TempObj.DO_NO                                                             '2
      SENDER_NAME = TempObj.SENDER_NAME                                                             '2
      DOCUMENT_TYPE = TempObj.DOCUMENT_TYPE                                                             '2
      RECEIVE_NAME = TempObj.RECEIVE_NAME                                                             '2
      QUE_NO = TempObj.QUE_NO                                                             '2
      TRUCK_NO = TempObj.TRUCK_NO                                                             '2
      SUPPLIER_ID = TempObj.SUPPLIER_ID                                                             '2
      SUPPLIER_CODE = TempObj.SUPPLIER_CODE                                                             '2
      SUPPLIER_NAME = TempObj.SUPPLIER_NAME                                                             '2
      PR_NO = TempObj.PR_NO                                                             '2
      DEPARTMENT_NAME = TempObj.DEPARTMENT_NAME                                                             '2
   
End Function

