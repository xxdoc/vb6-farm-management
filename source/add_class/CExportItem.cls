VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExportItem"
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

Public EXPORT_ITEM_ID As Long
Public INVENTORY_DOC_ID As Long
Public PART_ITEM_ID As Long
Public LOCATION_ID As Long
Public HOUSE_ID As Long
Public PIG_ID As Long
Public PIG_IN_HOUSE As Double
Public EXPORT_AMOUNT As Double
Public EXPORT_AVG_PRICE As Double
Public EXPORT_LAST_PRICE As Double
Public PREVIOUS_AMOUNT As Double
Public CURRENT_AMOUNT As Double
Public TRANSACTION_SEQ As Long
Public GUI_ID As Long
Public CALCULATE_FLAG As String
Public PIG_STATUS As Long
Public TOTAL_WEIGHT As Double
Public LINK_ID As Long
Public TOTAL_PRICE As Double
Public EXPORT_TOTAL_PRICE As Double
Public REPLACE_FLAG As String
Public PIG_AGE As Double
Public AGE_CODE As String

Public PIG_STATUS_NO As String
Public PIG_STATUS_NAME As String
Public YYYYMM As String
Public BATCH_ID As Long
Public SIMULATE_FLAG As String
Public TX_TYPE As String
Public PIG_TYPE As String
Public PART_NO As String
Public PART_DESC As String
Public PART_TYPE As Long
Public PART_TYPE_ID As String
Public PART_TYPE_NO As String
Public PART_TYPE_NAME As String
Public LOCATION_NAME As String
Public LOCATION_NO As String
Public UNIT_NAME As String
Public OLD_EXPORT_AMOUNT As Double
Public OLD_LOCATION_ID As Long
Public PART_PIG_TYPE As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public COMMIT_FLAG As String
Public DOCUMENT_TYPE As Long
Public PIG_FLAG As String
Public PART_GROUP_ID As Long
Public HOUSE_GROUP_ID As Long
Public HOUSE_GROUP_NAME As String
Public TO_HOUSE_GROUP_ID As Long
Public STATUS_GROUP_ID As Long
Public FROM_DATE1 As Date
Public TO_DATE1 As Date
Public HOUSE_ID1 As Long
Public COMMIT_FLAG1 As String
Public PIG_NO As String
Public PIG_NAME As String
Public HOUSE_NAME As String
Public SUM_ALL_EXPORT_AMT As Double
Public SUM_ALL_EXPORT As Double
Public SALE_FLAG As String
Public HOUSE_SALE_FLAG As String
Public DOCUMENT_DATE As Date
Public DOCUMENT_NO As String
Public FROM_TX_SEQ As Long
Public TO_TX_SEQ As Long
Public INTAKE_FLAG As String
Public TO_PIG_TYPE As String
Public PARENT_FLAG As String
Public DOCUMENT_SUBTYPE As Long

Public PRIORITY1 As Long
Public PART_ITEM_ID1 As Long
Public LOCATION_ID1 As Long
Public CAPITAL_MOVE_FLAG As String
Public EXCEPTION_FLAG As String
Public FROM_GUI_ID As Long
Public TO_GUI_ID As Long
Public PARENT_ID As Long
Public OLD_PARENT_ID As Long
Public DocTypeSet As String
Public CUS_ID As Long
Public FEED_GROUP As Long

Public DISCOUNT_AMOUNT As Double
Public YEAR_SEQ_ID As Long
Public WEEK_NO As String

'Public EXPOSE_TYPE_ID As Long
'Public MFG_DATE As Date
'Public EXP_DATE As Date
Public EMP_ID As Long
Public EMP_NAME As String
Public EMP_LNAME As String
Public SURETY_NAME As String
Public SURETY_LNAME As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   TX_TYPE = "E"
   FROM_DATE = -1
   TO_DATE = -1
'   MFG_DATE = -1
'   EXP_DATE = -1
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
Dim GroupByStr As String
Dim OrderTypeStr As String
Dim TempDate As String
Dim SubLen As Long
Dim NewStr As String
Dim FromDate1 As String
Dim ToDate1 As String
Dim MustCount As Boolean
   
   MustCount = True
   SIMULATE_FLAG = glbUser.SIMULATE_FLAG
   
   If Ind = 1 Or Ind = 9999 Then
      SelectStr1 = "SELECT EI.*, PT.*, LC.*, HS.LOCATION_NAME HOUSE_NAME, PI.*, PG.PART_NO PIG_NO, PG.PIG_TYPE PIG_TYPE, PI.PIG_TYPE PIG_TYPE1, UN.*, IVD.*, PI.PARENT_ID PART_PARENT_ID"
      SelectStr1 = SelectStr1 & ", ENM.LONG_NAME EMP_NAME, ENM.LAST_NAME EMP_LNAME,ENS.LONG_NAME SURETY_NAME, ENS.LAST_NAME SURETY_LNAME "
      SelectStr1 = SelectStr1 & " ,SNS.LONG_NAME SUP_NAME ,SC.SUPPLIER_CODE SUP_NUM "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (EI.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION HS ON (EI.HOUSE_ID = HS.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PG ON (EI.PIG_ID = PG.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE ED ON (IVD.EMP_ID = ED.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (ED.EMP_ID = EN.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME ENM ON (EN.NAME_ID = ENM.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EDS ON (IVD.SURETY_ID = EDS.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN2 ON (EDS.EMP_ID = EN2.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME ENS ON (EN2.NAME_ID = ENS.NAME_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (PI.SUPPLIER_ID = SN.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME SNS ON (SN.NAME_ID = SNS.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SC ON (PI.SUPPLIER_ID = SC.SUPPLIER_ID) "

   
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT PTT.PART_TYPE_NO, PTT.PART_TYPE_NAME, PI.PART_NO, PI.PART_DESC, "
      SelectStr1 = SelectStr1 & "IVD.DOCUMENT_NO, IVD.DOCUMENT_DATE, EI.EXPORT_AMOUNT, UN.UNIT_NAME, "
      SelectStr1 = SelectStr1 & "LC1.LOCATION_NAME LC_NAME1, LC1.LOCATION_NO LC_NO1, LC2.LOCATION_NAME LC_NAME2, LC2.LOCATION_NO LC_NO2 "

      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI, INVENTORY_DOC IVD, PART_TYPE PTT, PART_ITEM PI, UNIT UN, LOCATION LC1, LOCATION LC2, IMPORT_ITEM II "

      WhereStr = "WHERE (Ei.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID) "
      WhereStr = WhereStr & "AND (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      WhereStr = WhereStr & "AND (PI.PART_TYPE = PTT.PART_TYPE_ID) "
      WhereStr = WhereStr & "AND (PI.UNIT_COUNT = UN.UNIT_ID) "
      WhereStr = WhereStr & "AND (EI.GUI_ID = II.GUI_ID) "
      WhereStr = WhereStr & "AND (EI.LOCATION_ID = LC1.LOCATION_ID) "
      WhereStr = WhereStr & "AND (II.LOCATION_ID = LC2.LOCATION_ID) "
      WhereStr = WhereStr & "AND (IVD.DOCUMENT_TYPE = 3) "
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT DISTINCT EI.HOUSE_ID, EI.PIG_ID "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (EI.HOUSE_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PIG_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_TYPE PDT ON (PI.PIG_TYPE = PDT.PRODUCT_TYPE_NO) "
   
      WhereStr = "WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      WhereStr = WhereStr & "AND (EI.PIG_ID IS NOT NULL) "
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT EI.HOUSE_ID, EI.PIG_ID, PT.PART_GROUP_ID, SUM(EI.EXPORT_TOTAL_PRICE) EXPORT_TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "
      
      WhereStr = "WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      WhereStr = WhereStr & "AND (IVD.DOCUMENT_TYPE = 2) "
      WhereStr = WhereStr & "AND (EI.PIG_ID IS NOT NULL) "
   ElseIf Ind = 5 Then
      SelectStr1 = "SELECT PT1.*, "
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM (EI.EXPORT_TOTAL_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI "
      SelectStr1 = SelectStr1 & "   LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   LEFT OUTER JOIN PART_TYPE PT2 ON (PI.PART_TYPE = PT2.PART_TYPE_ID) "
      SelectStr1 = SelectStr1 & "   WHERE (EI.HOUSE_ID IN (SELECT HGI.LOCATION_ID FROM HGROUP_ITEM HGI WHERE (HGI.HOUSE_GROUP_ID = " & HOUSE_GROUP_ID & ") AND (HGI.SELECT_FLAG = 'Y'))) "
      SelectStr1 = SelectStr1 & "   AND (PI.PART_TYPE = PT1.PART_TYPE_ID) "
      If Len(COMMIT_FLAG1) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG1 & "') "
      End If
      
      If BATCH_ID > 0 Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (IVD.BATCH_ID = " & BATCH_ID & ") "
         Else
            WhereStr = WhereStr & "AND (IVD.BATCH_ID = " & BATCH_ID & ") "
         End If
      Else
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE ((IVD.BATCH_ID <= 0) OR (IVD.BATCH_ID IS NULL)) "
         Else
            WhereStr = WhereStr & "AND ((IVD.BATCH_ID <= 0) OR (IVD.BATCH_ID IS NULL)) "
         End If
      End If
      
      If Len(SIMULATE_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.SIMULATE_FLAG = '" & SIMULATE_FLAG & "') "
         SIMULATE_FLAG = ""
      End If
      
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 2) "
      If FROM_DATE1 > 0 Then
         FromDate1 = DateToStringIntLow(Trim(FROM_DATE1))
         SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate1 & "') "
      End If
      If TO_DATE1 > 0 Then
         ToDate1 = DateToStringIntLow(Trim(TO_DATE1))
         SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate1 & "') "
      End If
      SelectStr1 = SelectStr1 & ") EXPORT_TOTAL_PRICE "

      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM PART_TYPE PT1 "

      WhereStr = ""
   ElseIf Ind = 6 Then
      SelectStr1 = "SELECT PT1.*, "
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM (EI.EXPORT_TOTAL_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI "
      SelectStr1 = SelectStr1 & "   LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   LEFT OUTER JOIN PART_TYPE PT2 ON (PI.PART_TYPE = PT2.PART_TYPE_ID) "
      SelectStr1 = SelectStr1 & "   WHERE (EI.HOUSE_ID = " & HOUSE_ID1 & ") "
      SelectStr1 = SelectStr1 & "   AND (PI.PART_TYPE = PT1.PART_TYPE_ID) "
      If Len(COMMIT_FLAG1) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG1 & "') "
      End If
      If Len(SIMULATE_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.SIMULATE_FLAG = '" & SIMULATE_FLAG & "') "
         SIMULATE_FLAG = ""
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 2) "
      If FROM_DATE1 > 0 Then
         FromDate1 = DateToStringIntLow(Trim(FROM_DATE1))
         SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate1 & "') "
      End If
      If TO_DATE1 > 0 Then
         ToDate1 = DateToStringIntLow(Trim(TO_DATE1))
         SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate1 & "') "
      End If
      SelectStr1 = SelectStr1 & ") EXPORT_TOTAL_PRICE "

      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM PART_TYPE PT1 "

      WhereStr = ""
   ElseIf Ind = 7 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, EI.HOUSE_ID, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT, SUM(EI.EXPORT_TOTAL_PRICE) EXPORT_TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = "WHERE (PI.PIG_FLAG = 'N') "
      WhereStr = WhereStr & "AND (IVD.DOCUMENT_TYPE = 2) "
   ElseIf Ind = 8 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, EI.HOUSE_ID, EI.PIG_ID, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT, SUM(EI.EXPORT_TOTAL_PRICE) EXPORT_TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "

      WhereStr = "WHERE (PI.PIG_FLAG = 'N') "
      WhereStr = WhereStr & "AND (EI.PIG_ID IS NOT NULL) "
      WhereStr = WhereStr & "AND (IVD.DOCUMENT_TYPE = 2) "
   ElseIf Ind = 9 Then
      SelectStr1 = "SELECT DISTINCT EI.PART_ITEM_ID, EI.PIG_ID "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "

      WhereStr = "WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      WhereStr = WhereStr & "AND (PI.PIG_FLAG = 'N') "
      WhereStr = WhereStr & "AND (EI.PIG_ID IS NOT NULL) "
      WhereStr = WhereStr & "AND (EI.HOUSE_ID IS NOT NULL) "
      WhereStr = WhereStr & "AND (EI.HOUSE_ID IN (SELECT HGI.LOCATION_ID FROM HGROUP_ITEM HGI WHERE (HGI.HOUSE_GROUP_ID = " & HOUSE_GROUP_ID & ") AND (HGI.SELECT_FLAG = 'Y'))) "
   ElseIf Ind = 10 Then
      SelectStr1 = "SELECT PI.PART_NO, PI.PIG_TYPE, EI.LOCATION_ID, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = "WHERE (PI.PIG_FLAG = 'Y') "
   ElseIf Ind = 11 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, SUM(EI.EXPORT_TOTAL_PRICE) SUM_ALL_EXPORT_AMT, SUM(EI.EXPORT_AMOUNT) SUM_ALL_EXPORT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = "WHERE (PI.PIG_FLAG = 'N') "
      WhereStr = WhereStr & "AND (IVD.DOCUMENT_DATE >= '" & DateToStringIntLow(FROM_DATE) & "')"
      WhereStr = WhereStr & "AND (IVD.DOCUMENT_DATE < '" & DateToStringIntLow(TO_DATE) & "')"
   ElseIf Ind = 12 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, IVD.DOCUMENT_TYPE, IVD.SALE_FLAG, SUM(EI.EXPORT_TOTAL_PRICE) SUM_EXPORT_AMT, SUM(EI.EXPORT_AMOUNT) SUM_EXPORT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = "WHERE (PI.PIG_FLAG = 'N') "
   ElseIf (Ind = 13) Or (Ind = 17) Then
'      SelectStr1 = "SELECT PG.PIG_TYPE TO_PIG_TYPE, UN.UNIT_NAME, EI.*, PT.PART_TYPE_NO, PT.PART_TYPE_NAME, IVD.COMMIT_FLAG, IVD.DOCUMENT_TYPE, IVD.DOCUMENT_DATE, IVD.DOCUMENT_NO, PI.PART_NO, PI.PART_DESC, PR.PRIORITY1, PR.PRIORITY2, LC.SALE_FLAG HOUSE_SALE_FLAG, "
'      SelectStr1 = SelectStr1 & "PI.PIG_FLAG, PI.PIG_TYPE, PI.PART_TYPE, EI.EXPORT_AMOUNT TX_AMOUNT, EI.EXPORT_TOTAL_PRICE TX_PRICE, "
'      SelectStr1 = SelectStr1 & "EI.PIG_STATUS, EI.TRANSACTION_SEQ, IVD.SALE_FLAG DOC_SALE_FLAG "
'      SelectStr2 = "SELECT COUNT(*)  "
      
      SelectStr1 = "SELECT PG.PIG_TYPE TO_PIG_TYPE, UN.UNIT_NAME, EI.*, PT.PART_TYPE_NO, PT.PART_TYPE_NAME, IVD.COMMIT_FLAG, IVD.DOCUMENT_SUBTYPE, IVD.DOCUMENT_TYPE, IVD.DOCUMENT_DATE, IVD.DOCUMENT_NO,IVD.CUS_ID, PI.PART_NO, PI.PART_DESC, PR.PRIORITY1, PR.PRIORITY2, LC.SALE_FLAG HOUSE_SALE_FLAG, "
      SelectStr1 = SelectStr1 & "PI.PIG_FLAG, PI.PIG_TYPE, PI.PART_TYPE, EI.EXPORT_AMOUNT TX_AMOUNT, 0 TX_PRICE, "
      SelectStr1 = SelectStr1 & "EI.PIG_STATUS, EI.TRANSACTION_SEQ, IVD.SALE_FLAG DOC_SALE_FLAG "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN DOCUMENT_PRIORITY PR ON (IVD.DOCUMENT_TYPE = PR.DOCUMENT_TYPE) AND (PR.AREA = 1)"
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (EI.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PG ON (EI.PIG_ID = PG.PART_ITEM_ID) "
      
      WhereStr = ""
   ElseIf Ind = 14 Then
      SelectStr1 = "SELECT DISTINCT EI.PART_ITEM_ID "
      SelectStr2 = "SELECT COUNT(*)  "

     FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "

      WhereStr = "WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      WhereStr = WhereStr & "AND (PI.PIG_FLAG = 'N') "
      WhereStr = WhereStr & "AND (EI.PIG_ID IS NOT NULL) "
      WhereStr = WhereStr & "AND (EI.HOUSE_ID IS NOT NULL) "
      WhereStr = WhereStr & "AND (EI.HOUSE_ID IN (SELECT HGI.LOCATION_ID FROM HGROUP_ITEM HGI WHERE (HGI.HOUSE_GROUP_ID = " & HOUSE_GROUP_ID & ") AND (HGI.SELECT_FLAG = 'Y'))) "
   ElseIf Ind = 15 Then
      SelectStr1 = "SELECT DISTINCT EI.PIG_ID, PG.PART_NO, PG.PART_DESC, PG.PIG_TYPE "
      SelectStr2 = "SELECT COUNT(*)  "
        
      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PG ON (EI.PIG_ID = PG.PART_ITEM_ID) "

      WhereStr = WhereStr & "WHERE (PG.PIG_FLAG = 'Y') "
   ElseIf Ind = 16 Then
      SelectStr1 = "SELECT EI.PIG_ID, PG.PIG_TYPE, PI.PART_TYPE, SUM(EI.EXPORT_AMOUNT) SUM_EXPORT_AMOUNT, SUM(EI.EXPORT_TOTAL_PRICE) SUM_EXPORT_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
        
      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PG ON (EI.PIG_ID = PG.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "

      WhereStr = "WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      WhereStr = WhereStr & "AND (EI.PIG_ID IS NOT NULL) "
   ElseIf Ind = 18 Then '17 is already used
      SelectStr1 = "SELECT IVD.DOCUMENT_DATE, EI.TRANSACTION_SEQ, EI.LOCATION_ID, EI.PART_ITEM_ID, EI.CURRENT_AMOUNT, EI.EXPORT_AVG_PRICE , 'E' "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "

      WhereStr = "WHERE EI.TRANSACTION_SEQ IN "
      WhereStr = WhereStr & "( "
      WhereStr = WhereStr & "   SELECT MAX(EI1.TRANSACTION_SEQ) "
      WhereStr = WhereStr & "   FROM EXPORT_ITEM EI1 "
      WhereStr = WhereStr & "   LEFT OUTER JOIN INVENTORY_DOC IVD1 ON (EI1.INVENTORY_DOC_ID = IVD1.INVENTORY_DOC_ID) "
      WhereStr = WhereStr & "   WHERE (IVD1.DOCUMENT_DATE = IVD.DOCUMENT_DATE) "
      WhereStr = WhereStr & "   AND (EI1.PART_ITEM_ID = EI.PART_ITEM_ID) "
      WhereStr = WhereStr & "   AND (EI1.LOCATION_ID = EI.LOCATION_ID) "
      WhereStr = WhereStr & ") "
      
      MustCount = False
   ElseIf Ind = 19 Then
      SelectStr1 = "SELECT EI.LOCATION_ID, EI.PIG_STATUS, SUM(EI.EXPORT_AMOUNT) SUM_EXPORT_AMOUNT, SUM(EI.EXPORT_TOTAL_PRICE) SUM_EXPORT_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "

      WhereStr = "WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
   ElseIf Ind = 20 Then
      SelectStr1 = "SELECT EI.LOCATION_ID, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "

      WhereStr = "WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      WhereStr = WhereStr & "AND (PI.PIG_FLAG = 'Y') "
      WhereStr = WhereStr & "AND (EI.PIG_STATUS NOT IN (SELECT HGI.STATUS_ID FROM SGROUP_ITEM HGI WHERE (HGI.STATUS_GROUP_ID = " & STATUS_GROUP_ID & ") AND (HGI.SELECT_FLAG = 'Y'))) "
   ElseIf Ind = 21 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, EI.PIG_STATUS, SUM(EI.EXPORT_AMOUNT) SUM_EXPORT_AMOUNT, SUM(EI.EXPORT_TOTAL_PRICE) SUM_EXPORT_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "

      WhereStr = "WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
   ElseIf Ind = 22 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "

      WhereStr = "WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      WhereStr = WhereStr & "AND (PI.PIG_FLAG = 'Y') "
      WhereStr = WhereStr & "AND (EI.PIG_STATUS NOT IN (SELECT HGI.STATUS_ID FROM SGROUP_ITEM HGI WHERE (HGI.STATUS_GROUP_ID = " & STATUS_GROUP_ID & ") AND (HGI.SELECT_FLAG = 'Y'))) "
   ElseIf Ind = 23 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_TYPE PDT ON (PI.PIG_TYPE = PDT.PRODUCT_TYPE_NO) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION HS ON (EI.LOCATION_ID = HS.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (EI.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
      
      WhereStr = ""
      
      If STATUS_GROUP_ID > 0 Then
         If Len(WhereStr) <= 0 Then
            WhereStr = WhereStr & "WHERE (EI.PIG_STATUS IN (SELECT HGI.STATUS_ID FROM SGROUP_ITEM HGI WHERE (HGI.STATUS_GROUP_ID = " & STATUS_GROUP_ID & ") AND (HGI.SELECT_FLAG = 'Y'))) "
         Else
            WhereStr = WhereStr & "AND (EI.PIG_STATUS IN (SELECT HGI.STATUS_ID FROM SGROUP_ITEM HGI WHERE (HGI.STATUS_GROUP_ID = " & STATUS_GROUP_ID & ") AND (HGI.SELECT_FLAG = 'Y'))) "
         End If
      End If
      
   ElseIf Ind = 24 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "

      WhereStr = "WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      WhereStr = WhereStr & "AND (PI.PIG_FLAG = 'Y') "
      WhereStr = WhereStr & "AND (EI.PIG_STATUS IN (SELECT HGI.STATUS_ID FROM SGROUP_ITEM HGI WHERE (HGI.STATUS_GROUP_ID = " & STATUS_GROUP_ID & ") AND (HGI.SELECT_FLAG = 'Y'))) "
   ElseIf Ind = 25 Then
      SelectStr1 = "SELECT EI.PIG_ID, PT.PART_GROUP_ID, IVD.SALE_FLAG, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "

      WhereStr = "WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      WhereStr = WhereStr & "AND (IVD.DOCUMENT_TYPE = 2) "
      WhereStr = WhereStr & "AND (EI.PIG_ID IS NOT NULL) "
   ElseIf Ind = 26 Then
      SelectStr1 = "SELECT EI.PIG_STATUS, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT, SUM(TOTAL_PRICE) TOTAL_PRICE,SUM(DISCOUNT_AMOUNT) DISCOUNT_AMOUNT, SUM(TOTAL_WEIGHT) TOTAL_WEIGHT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
   
      WhereStr = ""
   ElseIf Ind = 27 Then
      SelectStr1 = "SELECT EI.*, PS.PRODUCT_STATUS_NAME, PI.PART_NO, PI.PART_DESC, PI.PIG_TYPE, IVD.*, PT.PART_TYPE_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (EI.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
      
      WhereStr = ""
   ElseIf Ind = 28 Then
      SelectStr1 = "SELECT EI.*,  IVD.* "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "

      WhereStr = "WHERE EI.TRANSACTION_SEQ IN "
      WhereStr = WhereStr & "( "
      WhereStr = WhereStr & "   SELECT MAX(EI1.TRANSACTION_SEQ) "
      WhereStr = WhereStr & "   FROM EXPORT_ITEM EI1, INVENTORY_DOC IVD1 "
      WhereStr = WhereStr & "   WHERE (EI1.INVENTORY_DOC_ID = IVD1.INVENTORY_DOC_ID) "
      WhereStr = WhereStr & "   AND (EI1.TRANSACTION_SEQ < " & TRANSACTION_SEQ & " ) "
      WhereStr = WhereStr & "   AND (EI1.PART_ITEM_ID = " & PART_ITEM_ID1 & ") "
      WhereStr = WhereStr & "   AND (EI1.LOCATION_ID = " & LOCATION_ID1 & ") "
      WhereStr = WhereStr & "   AND (IVD1.COMMIT_FLAG = 'Y') "
      WhereStr = WhereStr & ") "
   ElseIf Ind = 29 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, EI.HOUSE_ID, EI.PIG_ID, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT, SUM(EI.EXPORT_TOTAL_PRICE) EXPORT_TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PG ON (EI.PIG_ID = PG.PART_ITEM_ID) "
'('G', 'L', 'B')
      WhereStr = "WHERE (PG.PIG_TYPE IN (SELECT PRODUCT_TYPE_NO FROM PRODUCT_TYPE WHERE CAPITAL_FLAG = 'Y')) "
      WhereStr = WhereStr & "AND (EI.PIG_ID IS NOT NULL) "
   ElseIf Ind = 30 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, EI.LOCATION_ID, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_TYPE PDT ON (PI.PIG_TYPE = PDT.PRODUCT_TYPE_NO) "

      WhereStr = ""
   ElseIf Ind = 31 Then
      SelectStr1 = "SELECT  EI.LOCATION_ID, EI.PART_ITEM_ID, EI.PIG_STATUS, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "

      WhereStr = ""
   ElseIf Ind = 32 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, EI.PIG_STATUS, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
   
      WhereStr = ""
   ElseIf Ind = 33 Then
      SelectStr1 = "SELECT DISTINCT EI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, PI.PART_TYPE, PT.PART_TYPE_NAME, PT.PART_TYPE_NO "
      SelectStr2 = "SELECT COUNT(DISTINCT EI.PART_ITEM_ID)  "
      
      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
   
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PG ON (EI.PIG_ID = PG.PART_ITEM_ID) "
      
      WhereStr = ""
   ElseIf Ind = 34 Then
      SelectStr1 = "SELECT EI.LOCATION_ID, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_TYPE PDT ON (PI.PIG_TYPE = PDT.PRODUCT_TYPE_NO) "
      
      WhereStr = ""
   ElseIf Ind = 35 Then
      SelectStr1 = "SELECT PT.PART_GROUP_ID, PG.PIG_TYPE, SUM(EI.EXPORT_TOTAL_PRICE) EXPORT_TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PG ON (EI.PIG_ID = PG.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PT.PART_TYPE_ID = PI.PART_TYPE) "
      
      WhereStr = "WHERE (PG.PIG_TYPE IN (SELECT PRODUCT_TYPE_NO FROM PRODUCT_TYPE WHERE CAPITAL_FLAG = 'Y')) "
   ElseIf Ind = 36 Then
      SelectStr1 = "SELECT PT.PART_GROUP_ID, EI.HOUSE_ID, SUM(EI.EXPORT_TOTAL_PRICE) EXPORT_TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PG ON (EI.PIG_ID = PG.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PT.PART_TYPE_ID = PI.PART_TYPE) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (EI.HOUSE_ID = LC.LOCATION_ID) "
         
      WhereStr = "WHERE (PG.PIG_TYPE IN (SELECT PRODUCT_TYPE_NO FROM PRODUCT_TYPE WHERE CAPITAL_FLAG = 'Y')) "
      WhereStr = WhereStr & "AND (LC.CAPITAL_BIRTH_FLAG = 'Y') "
   ElseIf Ind = 37 Then
      SelectStr1 = "SELECT PI.PIG_TYPE, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = ""
   ElseIf Ind = 38 Then
      SelectStr1 = "SELECT IVD.DOCUMENT_TYPE, PI.PIG_TYPE, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = ""
   ElseIf Ind = 39 Then
      SelectStr1 = "SELECT PI.PART_ITEM_ID, SUM(EI.EXPORT_TOTAL_PRICE) EXPORT_TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PG ON (EI.PIG_ID = PG.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (EI.HOUSE_ID = LC.LOCATION_ID) "

      WhereStr = "WHERE (PG.PIG_TYPE IN (SELECT PRODUCT_TYPE_NO FROM PRODUCT_TYPE WHERE CAPITAL_FLAG = 'Y')) "
      WhereStr = WhereStr & "AND (LC.CAPITAL_BIRTH_FLAG = 'Y') "
   ElseIf Ind = 40 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, EI.PIG_STATUS, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT, SUM(EI.TOTAL_PRICE) TOTAL_PRICE, SUM(EI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT, SUM(EI.TOTAL_WEIGHT) TOTAL_WEIGHT "
      SelectStr1 = SelectStr1 & ",PI.PART_NO "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PST ON (EI.PIG_STATUS = PST.PRODUCT_STATUS_ID) "
   
      WhereStr = ""
   ElseIf Ind = 41 Then
      SelectStr1 = "SELECT EI.AGE_CODE, EI.PIG_STATUS, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT, SUM(EI.TOTAL_PRICE) TOTAL_PRICE,SUM(EI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT, SUM(EI.TOTAL_WEIGHT) TOTAL_WEIGHT "
      SelectStr1 = SelectStr1 & ",PST.PRODUCT_STATUS_NO,AR.FROM_WEEK "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PST ON (EI.PIG_STATUS = PST.PRODUCT_STATUS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN AGE_RANGE AR ON (EI.AGE_CODE = AR.AGE_RANGE_NO) "
   
      WhereStr = ""
   ElseIf Ind = 42 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, IVD.YYYYMM, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT, SUM(EI.EXPORT_TOTAL_PRICE) EXPORT_TOTAL_PRICE, SUM(EI.TOTAL_WEIGHT) TOTAL_WEIGHT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
   
      WhereStr = ""
   ElseIf Ind = 43 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, EI.PIG_STATUS TR_PIG_STATUS, IVD.YYYYMM, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT, SUM(EI.EXPORT_TOTAL_PRICE) EXPORT_TOTAL_PRICE, SUM(EI.TOTAL_WEIGHT) TOTAL_WEIGHT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
   
      WhereStr = ""
   ElseIf Ind = 44 Then
      SelectStr1 = "SELECT DISTINCT EI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, PI.PIG_TYPE, EI.PIG_STATUS, PS.PRODUCT_STATUS_NO, PS.PRODUCT_STATUS_NAME "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (EI.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
   
      WhereStr = ""
   ElseIf Ind = 45 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, IVD.YYYYMM, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PG ON (EI.PIG_ID = PG.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (EI.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
   
      WhereStr = ""
   ElseIf Ind = 46 Then
      SelectStr1 = "SELECT EI.PIG_ID, IVD.YYYYMM, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (EI.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
   
      WhereStr = ""
   ElseIf Ind = 47 Then
      SelectStr1 = "SELECT DISTINCT EI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "

      WhereStr = "WHERE (EI.PIG_ID IS NOT NULL) "
      WhereStr = WhereStr & "AND (EI.HOUSE_ID IS NOT NULL) "
   ElseIf Ind = 48 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT, SUM(EI.EXPORT_TOTAL_PRICE) TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "

      WhereStr = ""
   ElseIf Ind = 49 Then
      SelectStr1 = "SELECT DISTINCT EI.PIG_ID PIG_ID, PG.PART_NO PIG_NO, PG.PART_DESC PIG_NAME, PG.PIG_TYPE PIG_TYPE, PI.PART_NO, PI.PART_DESC, PI.PART_ITEM_ID PART_ITEM_ID "
      SelectStr2 = "SELECT COUNT(DISTINCT EI.PIG_ID)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PG ON (EI.PIG_ID = PG.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
   
      WhereStr = ""
   ElseIf Ind = 50 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, EI.PIG_ID, IVD.YYYYMM, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (EI.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
   
      WhereStr = ""
   ElseIf Ind = 51 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, EI.PIG_ID, IVD.YYYYMM, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (EI.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
   
      WhereStr = ""
   ElseIf Ind = 52 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, EI.PIG_ID, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT, SUM(EI.EXPORT_TOTAL_PRICE) EXPORT_TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (EI.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
   
      WhereStr = ""
   ElseIf Ind = 53 Then
      'SelectStr1 = "SELECT EI.HOUSE_ID, EI.PIG_ID, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT, SUM(EI.EXPORT_TOTAL_PRICE) EXPORT_TOTAL_PRICE "
      SelectStr1 = "SELECT EI.HOUSE_ID, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT, SUM(EI.EXPORT_TOTAL_PRICE) EXPORT_TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*) "
      
      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PG ON (EI.PIG_ID = PG.PART_ITEM_ID) "
'('G', 'L', 'B')
      WhereStr = "WHERE (PG.PIG_TYPE IN (SELECT PRODUCT_TYPE_NO FROM PRODUCT_TYPE WHERE CAPITAL_FLAG = 'Y')) "
      WhereStr = WhereStr & "AND (EI.PIG_ID IS NOT NULL) "
   ElseIf Ind = 54 Then
      SelectStr1 = "SELECT DISTINCT EI.PIG_ID PIG_ID, PG.PART_NO PIG_NO, PG.PART_DESC PIG_NAME, PG.PIG_TYPE PIG_TYPE, PI.PART_NO, PI.PART_DESC, PI.PART_ITEM_ID PART_ITEM_ID,EI.PIG_AGE "
      SelectStr2 = "SELECT COUNT(DISTINCT EI.PIG_ID)  "
      
      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PG ON (EI.PIG_ID = PG.PART_ITEM_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_TYPE PDT ON (PG.PIG_TYPE = PDT.PRODUCT_TYPE_NO) "
      
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN PRTITEM_MAP PM ON (PG.PART_ITEM_ID = PM.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN YEAR_WEEK YW ON (YW.YEAR_WEEK_ID = PM.YEAR_WEEK_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN YEAR_SEQ YQ ON (YQ.YEAR_SEQ_ID = YW.YEAR_SEQ_ID) "
         
      WhereStr = ""
   ElseIf Ind = 55 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, EI.PIG_ID,EI.PIG_AGE, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT, SUM(EI.EXPORT_TOTAL_PRICE) EXPORT_TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (EI.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
   
      WhereStr = ""
   ElseIf Ind = 56 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID,Ei.PIG_STATUS, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_TYPE PDT ON (PI.PIG_TYPE = PDT.PRODUCT_TYPE_NO) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION HS ON (EI.LOCATION_ID = HS.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (EI.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
      
      WhereStr = ""
   ElseIf Ind = 57 Then
      SelectStr1 = "SELECT PS.PRODUCT_STATUS_NAME, PI.PART_NO, PI.PART_DESC, PI.PIG_TYPE, PT.PART_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",IVD.DOCUMENT_NO,IVD.DOCUMENT_DATE "
      SelectStr1 = SelectStr1 & ",SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT,SUM(EI.TOTAL_WEIGHT) TOTAL_WEIGHT "
      SelectStr1 = SelectStr1 & ",SUM(EI.TOTAL_PRICE) TOTAL_PRICE,SUM(EI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (EI.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
      
      WhereStr = ""
   ElseIf Ind = 58 Then
      SelectStr1 = "SELECT PI.FEED_GROUP, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_TYPE PDT ON (PI.PIG_TYPE = PDT.PRODUCT_TYPE_NO) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION HS ON (EI.LOCATION_ID = HS.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (EI.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
      
      WhereStr = ""
      
   ElseIf Ind = 59 Then
      SelectStr1 = "SELECT IVD.DOCUMENT_DATE,HGI.HOUSE_GROUP_ID, SUM(EI.EXPORT_TOTAL_PRICE) EXPORT_TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN HGROUP_ITEM HGI ON (EI.HOUSE_ID = HGI.LOCATION_ID) "
      
      WhereStr = " WHERE (HGI.SELECT_FLAG = 'Y')  "
      
      If PART_GROUP_ID > 0 Then
         FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
         FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      End If
      
  ElseIf Ind = 60 Then
       SelectStr1 = "SELECT EI.*,IM.*, PT.*, LC.*, HS.LOCATION_NAME HOUSE_NAME, PI.*, PG.PART_NO PIG_NO, PG.PIG_TYPE PIG_TYPE, PI.PIG_TYPE PIG_TYPE1, UN.*, IVD.*, PI.PARENT_ID PART_PARENT_ID"
      SelectStr1 = SelectStr1 & ", ENM.LONG_NAME EMP_NAME, ENM.LAST_NAME EMP_LNAME,ENS.LONG_NAME SURETY_NAME, ENS.LAST_NAME SURETY_LNAME "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN IMPORT_ITEM IM "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (EI.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION HS ON (EI.HOUSE_ID = HS.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PG ON (EI.PIG_ID = PG.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE ED ON (IVD.EMP_ID = ED.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (ED.EMP_ID = EN.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME ENM ON (EN.NAME_ID = ENM.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EDS ON (IVD.SURETY_ID = EDS.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN2 ON (EDS.EMP_ID = EN2.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME ENS ON (EN2.NAME_ID = ENS.NAME_ID) "
   
      WhereStr = ""
      
   ElseIf Ind = 61 Then
'      SelectStr1 = "SELECT PS.PRODUCT_STATUS_NAME  ,IVD.EMP_ID , PI.PART_NO, PI.PART_DESC, PI.PIG_TYPE, PT.PART_TYPE_NAME "
'      SelectStr1 = SelectStr1 & ",IVD.DOCUMENT_NO,IVD.DOCUMENT_DATE "
      SelectStr1 = "SELECT IVD.EMP_ID, PI.PART_ITEM_ID , IVD.CUS_ID " ', PS.PRODUCT_STATUS_NAME "
      SelectStr1 = SelectStr1 & ",SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT,SUM(EI.TOTAL_WEIGHT) TOTAL_WEIGHT "
      SelectStr1 = SelectStr1 & ",SUM(EI.TOTAL_PRICE) TOTAL_PRICE,SUM(EI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (EI.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
      
      WhereStr = ""
      
   ElseIf Ind = 62 Then
         SelectStr1 = "SELECT DISTINCT HG.HOUSE_GROUP_ID , HG.HOUSE_GROUP_NAME "
  
         SelectStr2 = "SELECT COUNT(*)  "
  
         FromStr = "FROM EXPORT_ITEM EI "
         FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
         FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
'         FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (EI.LOCATION_ID = LC.LOCATION_ID) "
'         FromStr = FromStr & "LEFT OUTER JOIN LOCATION HS ON (EI.HOUSE_ID = HS.LOCATION_ID) "
'         FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PG ON (EI.PIG_ID = PG.PART_ITEM_ID) "
'         FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
         FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
         FromStr = FromStr & "LEFT OUTER JOIN HGROUP_ITEM HGI ON (EI.HOUSE_ID = HGI.LOCATION_ID) "
         FromStr = FromStr & "LEFT OUTER JOIN HOUSE_GROUP HG ON (HGI.HOUSE_GROUP_ID = HG.HOUSE_GROUP_ID) "
   
         WhereStr = "WHERE (HG.EXTRA_FLAG = 'N') AND (HGI.SELECT_FLAG='Y') "
      ElseIf Ind = 63 Then
      SelectStr1 = "SELECT SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT, DOCUMENT_DATE"
      SelectStr1 = SelectStr1 & ", PI.PART_NO PART_NO, PI.PART_DESC PART_DESC, PI.PART_ITEM_ID PART_ITEM_ID, HOUSE_ID "
      SelectStr1 = SelectStr1 & ", ENM.LONG_NAME EMP_NAME, ENM.LAST_NAME EMP_LNAME, DOCUMENT_TYPE "
       
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (EI.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION HS ON (EI.HOUSE_ID = HS.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PG ON (EI.PIG_ID = PG.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE ED ON (IVD.EMP_ID = ED.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (ED.EMP_ID = EN.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME ENM ON (EN.NAME_ID = ENM.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE EDS ON (IVD.SURETY_ID = EDS.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN2 ON (EDS.EMP_ID = EN2.EMP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME ENS ON (EN2.NAME_ID = ENS.NAME_ID) "
   
      WhereStr = ""
   ElseIf Ind = 64 Then
      SelectStr1 = "SELECT IVD.DOCUMENT_DATE,EI.PART_ITEM_ID, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_TYPE PDT ON (PI.PIG_TYPE = PDT.PRODUCT_TYPE_NO) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION HS ON (EI.LOCATION_ID = HS.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (EI.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
      
      WhereStr = ""
   ElseIf Ind = 100 Then
      SelectStr1 = "SELECT EI.PART_ITEM_ID, EI.LOCATION_ID, SUM(EI.EXPORT_AMOUNT) EXPORT_AMOUNT,SUM(EI.EXPORT_TOTAL_PRICE) EXPORT_TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM EXPORT_ITEM EI "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      
   End If
   
   If EXPORT_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EXPORT_ITEM_ID = " & EXPORT_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EXPORT_ITEM_ID = " & EXPORT_ITEM_ID & ") "
      End If
   End If
   
   If INVENTORY_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EI.INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EI.INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") "
      End If
   End If
   
   If LOCATION_ID > 0 Then
        If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (EI.LOCATION_ID = " & LOCATION_ID & ") "
        Else
            WhereStr = WhereStr & "AND (EI.LOCATION_ID = " & LOCATION_ID & ") "
        End If
    End If
   
   If PART_ITEM_ID > 0 Then
        If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (EI.PART_ITEM_ID = " & PART_ITEM_ID & ") "
        Else
            WhereStr = WhereStr & "AND (EI.PART_ITEM_ID = " & PART_ITEM_ID & ") "
        End If
    End If
   
   If PIG_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EI.PIG_ID = " & PIG_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EI.PIG_ID = " & PIG_ID & ") "
      End If
   End If
   
   If YEAR_SEQ_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (YQ.YEAR_SEQ_ID = " & YEAR_SEQ_ID & ") "
      Else
         WhereStr = WhereStr & "AND (YQ.YEAR_SEQ_ID = " & YEAR_SEQ_ID & ") "
      End If
   End If
   
   If GUI_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (GUI_ID = " & GUI_ID & ") "
      Else
         WhereStr = WhereStr & "AND (GUI_ID = " & GUI_ID & ") "
      End If
   End If
   
   If HOUSE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EI.HOUSE_ID = " & HOUSE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EI.HOUSE_ID = " & HOUSE_ID & ") "
      End If
   End If
   
   If PIG_STATUS > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PIG_STATUS = " & PIG_STATUS & ") "
      Else
         WhereStr = WhereStr & "AND (PIG_STATUS = " & PIG_STATUS & ") "
      End If
   End If
    
   If Len(COMMIT_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
   End If
   
   If Len(SALE_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (IVD.SALE_FLAG = '" & SALE_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (IVD.SALE_FLAG = '" & SALE_FLAG & "') "
      End If
   End If
   
   If Ind <> 5 Then
      If SIMULATE_FLAG = "Y" Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (SIMULATE_FLAG = '" & SIMULATE_FLAG & "') "
         Else
            WhereStr = WhereStr & "AND (SIMULATE_FLAG = '" & SIMULATE_FLAG & "') "
         End If
      ElseIf SIMULATE_FLAG = "N" Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE ((SIMULATE_FLAG <> 'Y') OR (SIMULATE_FLAG IS NULL)) "
         Else
            WhereStr = WhereStr & "AND ((SIMULATE_FLAG <> 'Y') OR (SIMULATE_FLAG IS NULL)) "
         End If
      End If
   End If
   
   If Len(PIG_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PI.PIG_FLAG = '" & PIG_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (PI.PIG_FLAG = '" & PIG_FLAG & "') "
      End If
   End If
   
   If Len(PIG_TYPE) > 0 Then
      If Len(WhereStr) <= 0 Then
         If Ind = 23 Or Ind = 64 Then
            WhereStr = "WHERE (PI.PIG_TYPE = '" & PIG_TYPE & "') "
         Else
            WhereStr = "WHERE (PG.PIG_TYPE = '" & PIG_TYPE & "') "
         End If
      Else
         If Ind = 23 Or Ind = 64 Then
            WhereStr = WhereStr & "AND (PI.PIG_TYPE = '" & PIG_TYPE & "') "
         Else
            WhereStr = WhereStr & "AND (PG.PIG_TYPE = '" & PIG_TYPE & "') "
         End If
      End If
   End If
   
   If Len(TO_PIG_TYPE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PG.PIG_TYPE = '" & TO_PIG_TYPE & "') "
      Else
         WhereStr = WhereStr & "AND (PG.PIG_TYPE = '" & TO_PIG_TYPE & "') "
      End If
   End If
   
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      End If
   End If
   
   If Len(INTAKE_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (PI.INTAKE_FLAG = '" & INTAKE_FLAG & "') "
      Else
            WhereStr = WhereStr & "AND (PI.INTAKE_FLAG = '" & INTAKE_FLAG & "') "
      End If
   End If
   
'   If Ind = 2 Then
      If PART_TYPE > 0 Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (PI.PART_TYPE = " & PART_TYPE & ") "
         Else
            WhereStr = WhereStr & "AND (PI.PART_TYPE = " & PART_TYPE & ") "
         End If
      End If
      
      If PART_GROUP_ID > 0 Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (PT.PART_GROUP_ID = " & PART_GROUP_ID & ") "
         Else
            WhereStr = WhereStr & "AND (PT.PART_GROUP_ID = " & PART_GROUP_ID & ") "
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
'   End If

'   If MFG_DATE > -1 Then
'      TempDate = Trim(DateToStringInt(MFG_DATE))
'      If WhereStr = "" Then
'         WhereStr = "WHERE (MFG_DATE = '" & ChangeQuote(TempDate) & "')"
'      Else
'         WhereStr = WhereStr & "AND (MFG_DATE = '" & ChangeQuote(TempDate) & "')"
'      End If
'   End If
'
'   If EXP_DATE > -1 Then
'      TempDate = Trim(DateToStringInt(EXP_DATE))
'      If WhereStr = "" Then
'         WhereStr = "WHERE (EXP_DATE = '" & ChangeQuote(TempDate) & "')"
'      Else
'         WhereStr = WhereStr & "AND (EXP_DATE = '" & ChangeQuote(TempDate) & "')"
'      End If
'   End If
   
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
   
   If Trim(WEEK_NO) <> "" Then
      If WildCard(WEEK_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(WEEK_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(WEEK_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (WEEK_NO = '" & ChangeQuote(Trim(WEEK_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (WEEK_NO = '" & ChangeQuote(Trim(WEEK_NO)) & "')"
         End If
      End If
   End If
   
   If Trim(DOCUMENT_NO) <> "" Then
      If WildCard(DOCUMENT_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(IVD.DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(IVD.DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (IVD.DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (IVD.DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "')"
         End If
      End If
   End If
   
   
   If FROM_TX_SEQ > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (TRANSACTION_SEQ > " & FROM_TX_SEQ & ") "
      Else
         WhereStr = WhereStr & "AND (TRANSACTION_SEQ > " & FROM_TX_SEQ & ") "
      End If
   End If
   
   If TO_TX_SEQ > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (TRANSACTION_SEQ < " & TO_TX_SEQ & ") "
      Else
         WhereStr = WhereStr & "AND (TRANSACTION_SEQ < " & TO_TX_SEQ & ") "
      End If
   End If
   
   If Len(PARENT_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CAPITAL_FLAG = '" & PARENT_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (CAPITAL_FLAG = '" & PARENT_FLAG & "') "
      End If
   End If
   
   If Len(DocTypeSet) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DOCUMENT_TYPE IN " & DocTypeSet & ") "
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_TYPE IN " & DocTypeSet & ") "
      End If
   End If
   
   If Len(REPLACE_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (REPLACE_FLAG = '" & REPLACE_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (REPLACE_FLAG = '" & REPLACE_FLAG & "') "
      End If
   End If
   
   If (Ind = 23) Or (Ind = 24) Or (Ind = 25) Or (Ind = 56) Then
      If Len(HOUSE_SALE_FLAG) > 0 Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (HS.SALE_FLAG = '" & HOUSE_SALE_FLAG & "') "
         Else
            WhereStr = WhereStr & "AND (HS.SALE_FLAG = '" & HOUSE_SALE_FLAG & "') "
         End If
      End If
   
      If HOUSE_GROUP_ID > 0 Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (EI.LOCATION_ID IN (SELECT LOCATION_ID FROM HGROUP_ITEM WHERE (HOUSE_GROUP_ID = " & HOUSE_GROUP_ID & ") AND (SELECT_FLAG = 'Y')))"
         Else
            WhereStr = WhereStr & "AND (EI.LOCATION_ID IN (SELECT LOCATION_ID FROM HGROUP_ITEM WHERE (HOUSE_GROUP_ID = " & HOUSE_GROUP_ID & ") AND (SELECT_FLAG = 'Y')))"
         End If
      End If
   End If
   
      If TO_HOUSE_GROUP_ID > 0 Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (EI.HOUSE_ID IN (SELECT LOCATION_ID FROM HGROUP_ITEM WHERE (HOUSE_GROUP_ID = " & TO_HOUSE_GROUP_ID & ") AND (SELECT_FLAG = 'Y')))"
         Else
            WhereStr = WhereStr & "AND (EI.HOUSE_ID IN (SELECT LOCATION_ID FROM HGROUP_ITEM WHERE (HOUSE_GROUP_ID = " & TO_HOUSE_GROUP_ID & ") AND (SELECT_FLAG = 'Y')))"
         End If
      End If
   
   If Len(EXCEPTION_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EXCEPTION_FLAG = '" & EXCEPTION_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (EXCEPTION_FLAG = '" & EXCEPTION_FLAG & "') "
      End If
   End If
   
   If Len(CAPITAL_MOVE_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CAPITAL_MOVE_FLAG = '" & CAPITAL_MOVE_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (CAPITAL_MOVE_FLAG = '" & CAPITAL_MOVE_FLAG & "') "
      End If
   End If
   
   If FROM_GUI_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (GUI_ID >= " & FROM_GUI_ID & ") "
      Else
         WhereStr = WhereStr & "AND (GUI_ID >= " & FROM_GUI_ID & ") "
      End If
   End If
   
   If TO_GUI_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (GUI_ID <= " & TO_GUI_ID & ") "
      Else
         WhereStr = WhereStr & "AND (GUI_ID <= " & TO_GUI_ID & ") "
      End If
   End If
   
   If (Ind <> 5 And Ind <> 6) Then
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
   End If
   
   GroupByStr = ""
   
   If Ind = 4 Then
      GroupByStr = "GROUP BY EI.HOUSE_ID, EI.PIG_ID, PT.PART_GROUP_ID "
   ElseIf Ind = 7 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID, EI.HOUSE_ID "
   ElseIf Ind = 8 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID, EI.HOUSE_ID, EI.PIG_ID "
   ElseIf Ind = 10 Then
      GroupByStr = "GROUP BY PI.PART_NO, EI.LOCATION_ID, PI.PIG_TYPE "
   ElseIf Ind = 11 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID "
   ElseIf Ind = 12 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID, IVD.DOCUMENT_TYPE, IVD.SALE_FLAG "
   ElseIf Ind = 16 Then
      GroupByStr = "GROUP BY EI.PIG_ID, PG.PIG_TYPE, PI.PART_TYPE "
   ElseIf Ind = 19 Then
      GroupByStr = "GROUP BY EI.LOCATION_ID, EI.PIG_STATUS "
   ElseIf Ind = 20 Then
      GroupByStr = "GROUP BY EI.LOCATION_ID "
   ElseIf Ind = 21 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID, EI.PIG_STATUS "
   ElseIf Ind = 22 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID "
   ElseIf Ind = 23 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID "
   ElseIf Ind = 24 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID "
   ElseIf Ind = 25 Then
      GroupByStr = "GROUP BY EI.PIG_ID , PT.PART_GROUP_ID, IVD.SALE_FLAG "
   ElseIf Ind = 26 Then
      GroupByStr = "GROUP BY EI.PIG_STATUS "
   ElseIf Ind = 29 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID, EI.HOUSE_ID, EI.PIG_ID "
   ElseIf Ind = 30 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID, EI.LOCATION_ID "
   ElseIf Ind = 31 Then
      GroupByStr = "GROUP BY EI.LOCATION_ID, EI.PART_ITEM_ID, EI.PIG_STATUS "
   ElseIf Ind = 32 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID, EI.PIG_STATUS "
   ElseIf Ind = 34 Then
      GroupByStr = "GROUP BY EI.LOCATION_ID "
   ElseIf Ind = 35 Then
      GroupByStr = "GROUP BY PT.PART_GROUP_ID , PG.PIG_TYPE "
   ElseIf Ind = 36 Then
      GroupByStr = "GROUP BY PT.PART_GROUP_ID , EI.HOUSE_ID "
   ElseIf Ind = 37 Then
      GroupByStr = "GROUP BY PI.PIG_TYPE "
   ElseIf Ind = 38 Then
      GroupByStr = "GROUP BY IVD.DOCUMENT_TYPE, PI.PIG_TYPE "
   ElseIf Ind = 39 Then
      GroupByStr = "GROUP BY PI.PART_ITEM_ID "
   ElseIf Ind = 40 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID,PI.PART_NO, EI.PIG_STATUS "
   ElseIf Ind = 41 Then
      GroupByStr = "GROUP BY EI.AGE_CODE, EI.PIG_STATUS,PST.PRODUCT_STATUS_NO,AR.FROM_WEEK "
   ElseIf Ind = 42 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID, IVD.YYYYMM "
   ElseIf Ind = 43 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID, EI.PIG_STATUS, IVD.YYYYMM "
   ElseIf Ind = 45 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID , IVD.YYYYMM "
   ElseIf Ind = 46 Then
      GroupByStr = "GROUP BY EI.PIG_ID , IVD.YYYYMM "
   ElseIf Ind = 48 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID "
   ElseIf Ind = 50 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID, EI.PIG_ID, IVD.YYYYMM "
   ElseIf Ind = 52 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID, EI.PIG_ID "
   ElseIf Ind = 53 Then
      'GroupByStr = "GROUP BY  EI.HOUSE_ID, EI.PIG_ID "
      GroupByStr = "GROUP BY  EI.HOUSE_ID "
   ElseIf Ind = 55 Then
      GroupByStr = "GROUP BY EI.PART_ITEM_ID, EI.PIG_ID,EI.PIG_AGE "
   ElseIf Ind = 56 Then
      GroupByStr = "GROUP BY EI.PIG_STATUS,EI.PART_ITEM_ID "
   ElseIf Ind = 57 Then '  ,IVD.EMP_ID
      GroupByStr = "GROUP BY PS.PRODUCT_STATUS_NAME, PI.PART_NO, PI.PART_DESC, PI.PIG_TYPE, IVD.DOCUMENT_NO,IVD.DOCUMENT_DATE, PT.PART_TYPE_NAME "
   ElseIf Ind = 58 Then
      GroupByStr = "GROUP BY PI.FEED_GROUP "
   ElseIf Ind = 59 Then
      GroupByStr = "GROUP BY IVD.DOCUMENT_DATE , HGI.HOUSE_GROUP_ID "
   ElseIf Ind = 61 Then '  ,IVD.EMP_ID
      GroupByStr = "GROUP BY IVD.EMP_ID, PI.PART_ITEM_ID , IVD.CUS_ID "      ', PS.PRODUCT_STATUS_NAME "
   ElseIf Ind = 63 Then
      GroupByStr = "GROUP BY DOCUMENT_DATE, PI.PART_NO, PI.PART_DESC, PI.PART_ITEM_ID, HOUSE_ID, ENM.LONG_NAME, ENM.LAST_NAME, DOCUMENT_TYPE "
   ElseIf Ind = 64 Then
      GroupByStr = "GROUP BY IVD.DOCUMENT_DATE,EI.PART_ITEM_ID "
   ElseIf Ind = 100 Then
      GroupByStr = " GROUP BY EI.PART_ITEM_ID, EI.LOCATION_ID "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY EXPORT_ITEM_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY EXPORT_ITEM_ID "
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
         OrderByStr = "ORDER BY DOCUMENT_DATE, DOCUMENT_NO "
      ElseIf OrderBy = 11 Then
         OrderByStr = "ORDER BY TRANSACTION_SEQ "
      ElseIf OrderBy = 12 Then
         OrderByStr = "ORDER BY PT.PART_GROUP_ID ASC , PI.PART_NO  ASC , DOCUMENT_DATE "
      ElseIf OrderBy = 13 Then
         OrderByStr = "ORDER BY PT.PART_GROUP_ID ASC , PT.PART_TYPE_NO ASC , PI.PART_NO  ASC , DOCUMENT_DATE "
      Else
         OrderByStr = "ORDER BY EXPORT_ITEM_ID "
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
         OrderByStr = "ORDER BY EXPORT_ITEM_ID "
      End If
   ElseIf Ind = 3 Then
      If OrderBy = 1 Then
         If OrderType = 1 Then
            OrderByStr = "ORDER BY LOCATION_NO ASC, PART_NO "
         Else
            OrderByStr = "ORDER BY LOCATION_NO DESC, PART_NO "
         End If
      Else
         If OrderType = 1 Then
            OrderByStr = "ORDER BY LOCATION_NAME ASC, PART_NO "
         Else
            OrderByStr = "ORDER BY LOCATION_NAME DESC, PART_NO "
         End If
      End If
   ElseIf Ind = 4 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY EI.HOUSE_ID "
      Else
         OrderByStr = "ORDER BY EI.HOUSE_ID "
      End If
   ElseIf Ind = 5 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PT1.PART_TYPE_NO "
      Else
         OrderByStr = "ORDER BY PT1.PART_TYPE_NO "
      End If
   ElseIf Ind = 6 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PT1.PART_TYPE_NO "
      Else
         OrderByStr = "ORDER BY PT1.PART_TYPE_NO "
      End If
   ElseIf Ind = 7 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY EI.HOUSE_ID "
      Else
         OrderByStr = "ORDER BY EI.HOUSE_ID "
      End If
   ElseIf Ind = 8 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY EI.PART_ITEM_ID "
      Else
         OrderByStr = "ORDER BY EI.PART_ITEM_ID "
      End If
   ElseIf Ind = 9 Then
      Dim TempOrderType As String
         
      If OrderType = 1 Then
         TempOrderType = "ASC"
      Else
         TempOrderType = "DESC"
      End If
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PART_TYPE_NO " & TempOrderType & ", PART_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PART_TYPE_NO " & TempOrderType & ", PART_NO "
      Else
         OrderByStr = "ORDER BY PART_TYPE_NO " & TempOrderType & ", PART_NO "
      End If
   ElseIf Ind = 10 Then
      OrderByStr = "ORDER BY PI.PART_NO "
   ElseIf Ind = 11 Then
      OrderByStr = "ORDER BY PI.PART_TYPE "
   ElseIf Ind = 13 Then
      If OrderType = 1 Then
         If OrderBy = 1 Then
            OrderByStr = "ORDER BY PT.PART_TYPE_NO ASC, PI.PART_NO ASC, IVD.DOCUMENT_DATE ASC, PR.PRIORITY1 ASC, IVD.DOCUMENT_NO "
         ElseIf OrderBy = 2 Then
            OrderByStr = "ORDER BY IVD.DOCUMENT_DATE ASC, PR.PRIORITY1 ASC, IVD.DOCUMENT_NO ASC, EI.TRANSACTION_SEQ "
         End If
      Else
         If OrderBy = 1 Then
            OrderByStr = "ORDER BY PT.PART_TYPE_NO DESC, PI.PART_NO ASC, IVD.DOCUMENT_DATE ASC, PR.PRIORITY1 ASC, IVD.DOCUMENT_NO "
         ElseIf OrderBy = 2 Then
            OrderByStr = "ORDER BY IVD.DOCUMENT_DATE ASC, PR.PRIORITY1 DESC, IVD.DOCUMENT_NO ASC, EI.TRANSACTION_SEQ "
         End If
      End If
   ElseIf Ind = 14 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PT.PART_TYPE_NO "
      Else
         OrderByStr = "ORDER BY PT.PART_TYPE_NO "
      End If
   ElseIf Ind = 15 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PG.PART_NO "
      Else
         OrderByStr = "ORDER BY PG.PART_NO  "
      End If
   ElseIf Ind = 17 Then
      If OrderType = 1 Then
         OrderByStr = "ORDER BY PT.PART_TYPE_NO ASC, PI.PART_NO ASC, IVD.DOCUMENT_DATE ASC, IVD.DOCUMENT_NO  "
      Else
         OrderByStr = "ORDER BY PT.PART_TYPE_NO DESC, PI.PART_NO ASC, IVD.DOCUMENT_DATE ASC, IVD.DOCUMENT_NO "
      End If

   ElseIf Ind = 27 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE ASC, PRODUCT_STATUS_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE ASC, PRODUCT_STATUS_NAME "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY PRODUCT_STATUS_NAME ASC, IVD.DOCUMENT_DATE "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY PRODUCT_STATUS_NAME ASC, IVD.DOCUMENT_DATE "
      ElseIf OrderBy = 5 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE ASC, IVD.DOCUMENT_NO ASC, PI.PART_NO "
      ElseIf OrderBy = 6 Then
         OrderByStr = "ORDER BY PIG_AGE "
      End If
   ElseIf Ind = 33 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PART_TYPE_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PART_TYPE_NO "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY PART_TYPE_NO ASC, PI.PART_NO "
      Else
         OrderByStr = "ORDER BY PART_TYPE_NO "
      End If
   ElseIf Ind = 40 Then
      OrderByStr = "ORDER BY PI.PART_NO, EI.PIG_STATUS "
   ElseIf Ind = 41 Then
      OrderByStr = "ORDER BY PST.PRODUCT_STATUS_NO,AR.FROM_WEEK  "
   ElseIf Ind = 44 Then
      OrderByStr = "ORDER BY PS.PRODUCT_STATUS_NO ASC, PI.PART_NO ASC, PI.PIG_TYPE "
   ElseIf Ind = 45 Then
      OrderByStr = "ORDER BY IVD.YYYYMM "
   ElseIf Ind = 46 Then
      OrderByStr = "ORDER BY IVD.YYYYMM "
   ElseIf Ind = 47 Then
      OrderByStr = "ORDER BY PI.PART_NO "
   ElseIf Ind = 49 Then
      OrderByStr = "ORDER BY PG.PART_NO ASC, PG.PIG_TYPE ASC, PI.PART_DESC "
   ElseIf Ind = 50 Then
      OrderByStr = "ORDER BY EI.PART_ITEM_ID , EI.PIG_ID, IVD.YYYYMM "
   ElseIf Ind = 51 Then
      OrderByStr = "ORDER BY EI.PART_ITEM_ID , EI.PIG_ID, IVD.YYYYMM "
   ElseIf Ind = 54 Then
      OrderByStr = "ORDER BY PG.PART_NO ASC,PG.PIG_TYPE,EI.PIG_AGE ASC,PI.PART_NO "
   ElseIf Ind = 55 Then
      OrderByStr = "ORDER BY EI.PART_ITEM_ID , EI.PIG_ID,EI.PIG_AGE "
   ElseIf Ind = 56 Then
      OrderByStr = "ORDER BY EI.PIG_STATUS,EI.PART_ITEM_ID "
   ElseIf Ind = 57 Or Ind = 61 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE ASC, PRODUCT_STATUS_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE ASC, PRODUCT_STATUS_NAME "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY PRODUCT_STATUS_NAME ASC, IVD.DOCUMENT_DATE "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY PRODUCT_STATUS_NAME ASC, IVD.DOCUMENT_DATE "
      ElseIf OrderBy = 5 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE ASC, IVD.DOCUMENT_NO ASC, PI.PART_NO "
      ElseIf OrderBy = 6 Then
         OrderByStr = "ORDER BY PIG_AGE "
      End If
   ElseIf Ind = 58 Then
      OrderByStr = "ORDER BY PI.FEED_GROUP "
   ElseIf Ind = 60 Then
      OrderByStr = "ORDER BY PT.PART_GROUP_ID ASC , PI.PART_NO  ASC , DOCUMENT_DATE "
   End If
   
   If OrderByStr <> "" Then
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, itemcount, Rs, glbErrorLog)
End Function

Public Function CopyObject(O As CExportItem)
   EXPORT_ITEM_ID = O.EXPORT_ITEM_ID
   INVENTORY_DOC_ID = O.INVENTORY_DOC_ID
   PART_ITEM_ID = O.PART_ITEM_ID
   LOCATION_ID = O.LOCATION_ID
   HOUSE_ID = O.HOUSE_ID
   PIG_ID = O.PIG_ID
   PIG_IN_HOUSE = O.PIG_IN_HOUSE
   EXPORT_AMOUNT = O.EXPORT_AMOUNT
   EXPORT_AVG_PRICE = O.EXPORT_AVG_PRICE
   EXPORT_LAST_PRICE = O.EXPORT_LAST_PRICE
   PREVIOUS_AMOUNT = O.PREVIOUS_AMOUNT
   CURRENT_AMOUNT = O.CURRENT_AMOUNT
   TRANSACTION_SEQ = O.TRANSACTION_SEQ
   GUI_ID = O.GUI_ID
   CALCULATE_FLAG = O.CALCULATE_FLAG
   PIG_STATUS = O.PIG_STATUS
   TOTAL_WEIGHT = O.TOTAL_WEIGHT
   LINK_ID = O.LINK_ID
   TOTAL_PRICE = O.TOTAL_PRICE
   
   OLD_EXPORT_AMOUNT = O.EXPORT_AMOUNT
   OLD_LOCATION_ID = O.LOCATION_ID
   PART_TYPE = O.PART_TYPE
   PART_TYPE_NAME = O.PART_TYPE_NAME
   PART_TYPE_NO = O.PART_TYPE_NO
   LOCATION_NAME = O.LOCATION_NAME
   LOCATION_NO = O.LOCATION_NO
   UNIT_NAME = O.UNIT_NAME
   PART_NO = O.PART_NO
   PART_DESC = O.PART_DESC
   PART_PIG_TYPE = O.PART_PIG_TYPE
   PIG_TYPE = O.PIG_TYPE
   TX_TYPE = O.TX_TYPE
   
   PIG_NO = O.PIG_NO
   HOUSE_NAME = O.HOUSE_NAME
   DOCUMENT_NO = O.DOCUMENT_NO
   DOCUMENT_DATE = O.DOCUMENT_DATE
   DOCUMENT_TYPE = O.DOCUMENT_TYPE
   DISCOUNT_AMOUNT = O.DISCOUNT_AMOUNT
   
'   EXPOSE_TYPE_ID = O.EXPOSE_TYPE_ID
'   MFG_DATE = O.MFG_DATE
'   EXP_DATE = O.EXP_DATE
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Or Ind = 9999 Then
      EXPORT_ITEM_ID = NVLI(Rs("EXPORT_ITEM_ID"), -1)
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      HOUSE_ID = NVLI(Rs("HOUSE_ID"), -1)
      PIG_ID = NVLI(Rs("PIG_ID"), -1)
      PIG_IN_HOUSE = NVLD(Rs("PIG_IN_HOUSE"), 0)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      EXPORT_AVG_PRICE = NVLD(Rs("EXPORT_AVG_PRICE"), 0)
      EXPORT_LAST_PRICE = NVLD(Rs("EXPORT_LAST_PRICE"), 0)
      PREVIOUS_AMOUNT = NVLD(Rs("PREVIOUS_AMOUNT"), 0)
      CURRENT_AMOUNT = NVLD(Rs("CURRENT_AMOUNT"), 0)
      TRANSACTION_SEQ = NVLI(Rs("TRANSACTION_SEQ"), 0)
      GUI_ID = NVLI(Rs("GUI_ID"), -1)
      CALCULATE_FLAG = NVLS(Rs("CALCULATE_FLAG"), "")
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), -1)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      LINK_ID = NVLI(Rs("LINK_ID"), -1)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      EXPORT_TOTAL_PRICE = NVLD(Rs("EXPORT_TOTAL_PRICE"), 0)
      REPLACE_FLAG = NVLS(Rs("REPLACE_FLAG"), "")
      PIG_AGE = NVLD(Rs("PIG_AGE"), 0)
      AGE_CODE = NVLS(Rs("AGE_CODE"), "")
      
      OLD_EXPORT_AMOUNT = EXPORT_AMOUNT
      OLD_LOCATION_ID = LOCATION_ID
      PART_TYPE = NVLI(Rs("PART_TYPE"), -1)
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
      PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      PART_PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      PIG_TYPE = NVLS(Rs("PIG_TYPE1"), "")
      TX_TYPE = "E" 'Export
      
      PIG_NO = NVLS(Rs("PIG_NO"), "")
      HOUSE_NAME = NVLS(Rs("HOUSE_NAME"), "")
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      COMMIT_FLAG = NVLS(Rs("COMMIT_FLAG"), "")
      PARENT_ID = NVLI(Rs("PART_PARENT_ID"), -1)
      
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      OLD_PARENT_ID = PARENT_ID
      
'      EXPOSE_TYPE_ID = NVLI(Rs("EXPOSE_TYPE_ID"), -1)
'      MFG_DATE = InternalDateToDate(NVLS(Rs("MFG_DATE"), ""))
'      EXP_DATE = InternalDateToDate(NVLS(Rs("EXP_DATE"), ""))
      
      EMP_NAME = NVLS(Rs("EMP_NAME"), "")
      EMP_LNAME = NVLS(Rs("EMP_LNAME"), "")
      EMP_NAME = EMP_NAME & " " & EMP_LNAME
      
      SURETY_NAME = NVLS(Rs("SURETY_NAME"), "")
      SURETY_LNAME = NVLS(Rs("SURETY_LNAME"), "")
      SURETY_NAME = SURETY_NAME & " " & SURETY_LNAME
      
   ElseIf Ind = 4 Then
      HOUSE_ID = NVLI(Rs("HOUSE_ID"), -1)
      PIG_ID = NVLI(Rs("PIG_ID"), -1)
      PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), -1)
      EXPORT_TOTAL_PRICE = NVLD(Rs("EXPORT_TOTAL_PRICE"), 0)
   ElseIf Ind = 5 Then
      PART_TYPE_ID = NVLI(Rs("PART_TYPE_ID"), -1)
      PART_TYPE = NVLI(Rs("PART_TYPE_ID"), -1)
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
      PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
      EXPORT_TOTAL_PRICE = NVLD(Rs("EXPORT_TOTAL_PRICE"), 0)
   ElseIf Ind = 6 Then
      PART_TYPE_ID = NVLI(Rs("PART_TYPE_ID"), -1)
      PART_TYPE = NVLI(Rs("PART_TYPE_ID"), -1)
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
      PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
      EXPORT_TOTAL_PRICE = NVLD(Rs("EXPORT_TOTAL_PRICE"), 0)
   ElseIf Ind = 7 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      HOUSE_ID = NVLI(Rs("HOUSE_ID"), -1)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      EXPORT_TOTAL_PRICE = NVLD(Rs("EXPORT_TOTAL_PRICE"), 0)
   ElseIf Ind = 8 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      HOUSE_ID = NVLI(Rs("HOUSE_ID"), -1)
      PIG_ID = NVLI(Rs("PIG_ID"), -1)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      EXPORT_TOTAL_PRICE = NVLD(Rs("EXPORT_TOTAL_PRICE"), 0)
   ElseIf Ind = 10 Then
      PART_NO = NVLS(Rs("PART_NO"), "")
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
   ElseIf Ind = 11 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      SUM_ALL_EXPORT_AMT = NVLD(Rs("SUM_ALL_EXPORT_AMT"), 0)
      SUM_ALL_EXPORT = NVLD(Rs("SUM_ALL_EXPORT"), 0)
   ElseIf Ind = 12 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      SALE_FLAG = NVLS(Rs("SALE_FLAG"), "")
      SUM_ALL_EXPORT_AMT = NVLD(Rs("SUM_EXPORT_AMT"), 0)
      SUM_ALL_EXPORT = NVLD(Rs("SUM_EXPORT"), 0)
   ElseIf (Ind = 13) Or (Ind = 17) Then
      EXPORT_ITEM_ID = NVLI(Rs("EXPORT_ITEM_ID"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      PIG_FLAG = NVLS(Rs("PIG_FLAG"), "")
      PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      EXPORT_AVG_PRICE = NVLD(Rs("EXPORT_AVG_PRICE"), 0)
      EXPORT_TOTAL_PRICE = EXPORT_AVG_PRICE * EXPORT_AMOUNT
      CURRENT_AMOUNT = NVLD(Rs("CURRENT_AMOUNT"), 0)
      SALE_FLAG = NVLS(Rs("DOC_SALE_FLAG"), "")
      HOUSE_SALE_FLAG = NVLS(Rs("HOUSE_SALE_FLAG"), "")
      TRANSACTION_SEQ = NVLI(Rs("TRANSACTION_SEQ"), 0)
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), 0)
      PRIORITY1 = NVLI(Rs("PRIORITY1"), -1)
      GUI_ID = NVLI(Rs("GUI_ID"), -1)
      COMMIT_FLAG = NVLS(Rs("COMMIT_FLAG"), "")
      HOUSE_ID = NVLI(Rs("HOUSE_ID"), -1)
      PIG_ID = NVLI(Rs("PIG_ID"), -1)
      TO_PIG_TYPE = NVLS(Rs("TO_PIG_TYPE"), "")
      REPLACE_FLAG = NVLS(Rs("REPLACE_FLAG"), "")
      DOCUMENT_SUBTYPE = NVLI(Rs("DOCUMENT_SUBTYPE"), -1)
      CUS_ID = NVLI(Rs("CUS_ID"), -1)
'      EXPOSE_TYPE_ID = NVLI(Rs("EXPOSE_TYPE_ID"), -1)
'      MFG_DATE = InternalDateToDate(NVLS(Rs("MFG_DATE"), ""))
'      EXP_DATE = InternalDateToDate(NVLS(Rs("EXP_DATE"), ""))
   ElseIf Ind = 15 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLI(Rs("PART_DESC"), 0)
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
   ElseIf Ind = 16 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      PART_TYPE = NVLI(Rs("PART_TYPE"), 0)
      SUM_ALL_EXPORT = NVLD(Rs("SUM_EXPORT_AMOUNT"), 0)
      SUM_ALL_EXPORT_AMT = NVLD(Rs("SUM_EXPORT_PRICE"), 0)
   ElseIf Ind = 18 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      CURRENT_AMOUNT = NVLD(Rs("CURRENT_AMOUNT"), 0)
      EXPORT_AVG_PRICE = NVLD(Rs("EXPORT_AVG_PRICE"), 0)
      TRANSACTION_SEQ = NVLI(Rs("TRANSACTION_SEQ"), -1)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
   ElseIf Ind = 19 Then
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), 0)
      SUM_ALL_EXPORT_AMT = NVLD(Rs("SUM_EXPORT_PRICE"), 0)
      SUM_ALL_EXPORT = NVLD(Rs("SUM_EXPORT_AMOUNT"), 0)
   ElseIf Ind = 20 Then
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      SUM_ALL_EXPORT = NVLD(Rs("EXPORT_AMOUNT"), 0)
   ElseIf Ind = 21 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), 0)
      SUM_ALL_EXPORT_AMT = NVLD(Rs("SUM_EXPORT_PRICE"), 0)
      SUM_ALL_EXPORT = NVLD(Rs("SUM_EXPORT_AMOUNT"), 0)
   ElseIf Ind = 22 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      SUM_ALL_EXPORT = NVLD(Rs("EXPORT_AMOUNT"), 0)
   ElseIf Ind = 23 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
   ElseIf Ind = 24 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
   ElseIf Ind = 25 Then
      PIG_ID = NVLI(Rs("PIG_ID"), -1)
      PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), -1)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      SALE_FLAG = NVLS(Rs("SALE_FLAG"), "")
   ElseIf Ind = 26 Then
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), -1)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      SUM_ALL_EXPORT_AMT = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
   ElseIf Ind = 28 Then
      EXPORT_ITEM_ID = NVLI(Rs("EXPORT_ITEM_ID"), -1)
      INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      HOUSE_ID = NVLI(Rs("HOUSE_ID"), -1)
      PIG_ID = NVLI(Rs("PIG_ID"), -1)
      PIG_IN_HOUSE = NVLD(Rs("PIG_IN_HOUSE"), 0)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      EXPORT_AVG_PRICE = NVLD(Rs("EXPORT_AVG_PRICE"), 0)
      EXPORT_LAST_PRICE = NVLD(Rs("EXPORT_LAST_PRICE"), 0)
      PREVIOUS_AMOUNT = NVLD(Rs("PREVIOUS_AMOUNT"), 0)
      CURRENT_AMOUNT = NVLD(Rs("CURRENT_AMOUNT"), 0)
      TRANSACTION_SEQ = NVLI(Rs("TRANSACTION_SEQ"), 0)
      GUI_ID = NVLI(Rs("GUI_ID"), -1)
      CALCULATE_FLAG = NVLS(Rs("CALCULATE_FLAG"), "")
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), -1)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      LINK_ID = NVLI(Rs("LINK_ID"), -1)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
'      EXPOSE_TYPE_ID = NVLI(Rs("EXPOSE_TYPE_ID"), -1)
'      MFG_DATE = InternalDateToDate(NVLS(Rs("MFG_DATE"), ""))
'      EXP_DATE = InternalDateToDate(NVLS(Rs("EXP_DATE"), ""))
   ElseIf Ind = 29 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      HOUSE_ID = NVLI(Rs("HOUSE_ID"), -1)
      PIG_ID = NVLI(Rs("PIG_ID"), -1)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      EXPORT_TOTAL_PRICE = NVLD(Rs("EXPORT_TOTAL_PRICE"), 0)
   ElseIf Ind = 30 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
   ElseIf Ind = 31 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), -1)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
   ElseIf Ind = 32 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), -1)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
   ElseIf Ind = 33 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      PART_TYPE = NVLI(Rs("PART_TYPE"), -1)
      PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
   ElseIf Ind = 34 Then
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
   ElseIf Ind = 35 Then
      PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), 0)
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      EXPORT_TOTAL_PRICE = NVLD(Rs("EXPORT_TOTAL_PRICE"), 0)
   ElseIf Ind = 36 Then
      PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), 0)
      HOUSE_ID = NVLI(Rs("HOUSE_ID"), 0)
      EXPORT_TOTAL_PRICE = NVLD(Rs("EXPORT_TOTAL_PRICE"), 0)
   ElseIf Ind = 37 Then
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
   ElseIf Ind = 38 Then
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
   ElseIf Ind = 39 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      EXPORT_TOTAL_PRICE = NVLD(Rs("EXPORT_TOTAL_PRICE"), 0)
   ElseIf Ind = 40 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
   ElseIf Ind = 41 Then
      AGE_CODE = NVLS(Rs("AGE_CODE"), "")
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
   ElseIf Ind = 42 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      EXPORT_TOTAL_PRICE = NVLD(Rs("EXPORT_TOTAL_PRICE"), 0)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
   ElseIf Ind = 43 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PIG_STATUS = NVLI(Rs("TR_PIG_STATUS"), 0)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      EXPORT_TOTAL_PRICE = NVLD(Rs("EXPORT_TOTAL_PRICE"), 0)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
   ElseIf Ind = 44 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      PIG_STATUS_NO = NVLS(Rs("PRODUCT_STATUS_NO"), "")
      PIG_STATUS_NAME = NVLS(Rs("PRODUCT_STATUS_NAME"), "")
   ElseIf Ind = 45 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
   ElseIf Ind = 46 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
   ElseIf Ind = 47 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
   ElseIf Ind = 48 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
   ElseIf Ind = 49 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      PIG_NO = NVLS(Rs("PIG_NO"), "")
      PIG_NAME = NVLS(Rs("PIG_NAME"), "")
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
   ElseIf Ind = 50 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
   ElseIf Ind = 52 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      EXPORT_TOTAL_PRICE = NVLD(Rs("EXPORT_TOTAL_PRICE"), 0)
   ElseIf Ind = 53 Then
      HOUSE_ID = NVLI(Rs("HOUSE_ID"), -1)
      'PIG_ID = NVLI(Rs("PIG_ID"), -1)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      EXPORT_TOTAL_PRICE = NVLD(Rs("EXPORT_TOTAL_PRICE"), 0)
   ElseIf Ind = 54 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      PIG_NO = NVLS(Rs("PIG_NO"), "")
      PIG_NAME = NVLS(Rs("PIG_NAME"), "")
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      PIG_AGE = NVLD(Rs("PIG_AGE"), 0)
   ElseIf Ind = 55 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      PIG_AGE = NVLD(Rs("PIG_AGE"), 0)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      EXPORT_TOTAL_PRICE = NVLD(Rs("EXPORT_TOTAL_PRICE"), 0)
   ElseIf Ind = 56 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), -1)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
   ElseIf Ind = 58 Then
      FEED_GROUP = NVLI(Rs("FEED_GROUP"), -1)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
   ElseIf Ind = 59 Then
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      HOUSE_GROUP_ID = NVLI(Rs("HOUSE_GROUP_ID"), -1)
      EXPORT_TOTAL_PRICE = NVLD(Rs("EXPORT_TOTAL_PRICE"), 0)
    ElseIf Ind = 61 Then
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      SUM_ALL_EXPORT_AMT = NVLD(Rs("TOTAL_PRICE"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      CUS_ID = NVLI(Rs("CUS_ID"), 0)
      EMP_ID = NVLI(Rs("EMP_ID"), 0)
    ElseIf Ind = 62 Then
      HOUSE_GROUP_ID = NVLI(Rs("HOUSE_GROUP_ID"), -1)
      HOUSE_GROUP_NAME = NVLS(Rs("HOUSE_GROUP_NAME"), "")
    ElseIf Ind = 63 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      HOUSE_ID = NVLI(Rs("HOUSE_ID"), -1)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      EMP_NAME = NVLS(Rs("EMP_NAME"), "")
      EMP_LNAME = NVLS(Rs("EMP_LNAME"), "")
      EMP_NAME = EMP_NAME & " " & EMP_LNAME
   ElseIf Ind = 64 Then
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
   ElseIf Ind = 100 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      EXPORT_TOTAL_PRICE = NVLD(Rs("EXPORT_TOTAL_PRICE"), 0)
   End If
End Function

Public Function AddEditData(Optional NewSeq As Boolean = True)
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim ID2 As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String

   TempID1 = PIG_ID
   If PIG_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = HOUSE_ID
   If HOUSE_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID3 = PIG_STATUS
   If PIG_STATUS <= 0 Then
      TempID3 = "NULL"
   End If
   
   If REPLACE_FLAG = "" Then
      REPLACE_FLAG = "N"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("EXPORT_ITEM_SEQ", ID, glbErrorLog)
      If NewSeq Then
         Call glbDatabaseMngr.GetSeqID("TRANSACTION_SEQUENCE_SEQ", ID2, glbErrorLog)
      Else
         ID2 = TRANSACTION_SEQ
      End If
      
      SQL1 = "INSERT INTO EXPORT_ITEM " & _
                     "( " & _
                     "   EXPORT_ITEM_ID, " & _
                     "   INVENTORY_DOC_ID, " & _
                     "   PART_ITEM_ID, " & _
                     "   LOCATION_ID, " & _
                     "   HOUSE_ID, " & _
                     "   PIG_ID, " & _
                     "   PIG_IN_HOUSE, " & _
                     "   EXPORT_AMOUNT, " & _
                     "   EXPORT_AVG_PRICE, " & _
                     "   EXPORT_LAST_PRICE, " & _
                     "   PREVIOUS_AMOUNT, " & _
                     "   CURRENT_AMOUNT, "
            SQL1 = SQL1 & _
                     "   TRANSACTION_SEQ, " & _
                     "   GUI_ID, " & _
                     "   CALCULATE_FLAG, " & _
                     "   PIG_STATUS, " & _
                     "   TOTAL_WEIGHT, " & _
                     "   LINK_ID, "
            SQL1 = SQL1 & _
                     "   TOTAL_PRICE, " & _
                     "   REPLACE_FLAG, " & _
                     "   PIG_AGE, " & _
                     "   AGE_CODE, " & _
                     "   DISCOUNT_AMOUNT, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "" & INVENTORY_DOC_ID & "," & _
                         "" & PART_ITEM_ID & "," & _
                         "" & LOCATION_ID & "," & _
                         TempID2 & "," & _
                         TempID1 & "," & _
                         "" & PIG_IN_HOUSE & "," & _
                         "" & EXPORT_AMOUNT & "," & _
                         "" & EXPORT_AVG_PRICE & "," & _
                         "" & EXPORT_LAST_PRICE & "," & _
                         "" & PREVIOUS_AMOUNT & ","
               SQL1 = SQL1 & _
                         "" & CURRENT_AMOUNT & "," & _
                         "" & ID2 & "," & _
                         "" & GUI_ID & "," & _
                         "'" & CALCULATE_FLAG & "'," & _
                         TempID3 & "," & _
                         TOTAL_WEIGHT & "," & _
                         "" & LINK_ID & "," & _
                         "" & TOTAL_PRICE & "," & _
                         "'" & REPLACE_FLAG & "'," & _
                         PIG_AGE & "," & _
                         "'" & Trim(ChangeQuote(AGE_CODE)) & "'," & _
                         DISCOUNT_AMOUNT & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         EXPORT_ITEM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE EXPORT_ITEM SET " & _
                     "   INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ", " & _
                     "   PART_ITEM_ID = " & PART_ITEM_ID & ", " & _
                     "   LOCATION_ID = " & LOCATION_ID & ", " & _
                     "   HOUSE_ID = " & TempID2 & ", " & _
                     "   PIG_ID = " & TempID1 & ", " & _
                     "   PIG_IN_HOUSE = " & PIG_IN_HOUSE & ", " & _
                     "   EXPORT_AMOUNT = " & EXPORT_AMOUNT & ", " & _
                     "   EXPORT_AVG_PRICE = " & EXPORT_AVG_PRICE & ", " & _
                     "   EXPORT_LAST_PRICE = " & EXPORT_LAST_PRICE & ", " & _
                     "   PREVIOUS_AMOUNT = " & PREVIOUS_AMOUNT & ", " & _
                     "   CURRENT_AMOUNT = " & CURRENT_AMOUNT & ", " & _
                     "   TRANSACTION_SEQ = " & TRANSACTION_SEQ & ", " & _
                     "   GUI_ID = " & GUI_ID & ", " & _
                     "   CALCULATE_FLAG = '" & CALCULATE_FLAG & "', " & _
                     "   PIG_STATUS = " & TempID3 & ", " & _
                     "   TOTAL_WEIGHT = " & TOTAL_WEIGHT & ", " & _
                     "   LINK_ID = " & LINK_ID & ", " & _
                     "   TOTAL_PRICE = " & TOTAL_PRICE & ", " & _
                     "   REPLACE_FLAG = '" & REPLACE_FLAG & "', " & _
                     "   PIG_AGE = " & PIG_AGE & ", " & _
                     "   AGE_CODE = '" & Trim(ChangeQuote(AGE_CODE)) & "', "
            SQL1 = SQL1 & _
                     "   DISCOUNT_AMOUNT = " & DISCOUNT_AMOUNT & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE EXPORT_ITEM_ID = " & EXPORT_ITEM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM LOSS_ITEM WHERE CAPITAL_LOSS_ID  IN (SELECT CAPITAL_LOSS_ID FROM CAPITAL_LOSS WHERE EXPORT_ITEM_ID = " & EXPORT_ITEM_ID & ") "
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM CAPITAL_LOSS WHERE EXPORT_ITEM_ID = " & EXPORT_ITEM_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM MOVEMENT_ITEM WHERE CAPITAL_MOVEMENT_ID  IN (SELECT CAPITAL_MOVEMENT_ID FROM CAPITAL_MOVEMENT WHERE EXPORT_ITEM_ID = " & EXPORT_ITEM_ID & ") "
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM CAPITAL_MOVEMENT WHERE EXPORT_ITEM_ID = " & EXPORT_ITEM_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM EXPORT_ITEM WHERE EXPORT_ITEM_ID = " & EXPORT_ITEM_ID
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

Public Function PatchTxSeq() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE EXPORT_ITEM SET " & _
                  "   TRANSACTION_SEQ = " & TRANSACTION_SEQ & ", " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE EXPORT_ITEM_ID = " & EXPORT_ITEM_ID
                  
   m_Conn.Execute (SQL1)
End Function

Public Function PatchAvgPrice(AvgPrice As Double, PrevAmount As Double, CurrentAmount As Double, ExportTotalPrice As Double) As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL1 = "UPDATE EXPORT_ITEM SET " & _
                  "   EXPORT_AVG_PRICE = " & AvgPrice & ", " & _
                  "   PREVIOUS_AMOUNT = " & PrevAmount & ", " & _
                  "   CURRENT_AMOUNT = " & CurrentAmount & ", " & _
                  "   EXPORT_TOTAL_PRICE = " & ExportTotalPrice & ", " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE EXPORT_ITEM_ID = " & EXPORT_ITEM_ID
                  
   m_Conn.Execute (SQL1)
End Function


Public Function PatchPigAge() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE EXPORT_ITEM SET " & _
                  "   PIG_AGE = " & PIG_AGE & ", " & _
                  "   AGE_CODE = '" & Trim(ChangeQuote(AGE_CODE)) & "', " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE EXPORT_ITEM_ID = " & EXPORT_ITEM_ID
                  
   m_Conn.Execute (SQL1)
End Function
