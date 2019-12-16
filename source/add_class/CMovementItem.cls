VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMovementItem"
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

Public MOVEMENT_ITEM_ID As Long
Public CAPITAL_MOVEMENT_ID As Long
Public EXPENSE_TYPE As Long
Public PART_ITEM_ID As Long
Public CAPITAL_AMOUNT As Double

Public BATCH_ID As Long
Public SIMULATE_FLAG As String
Public DOCUMENT_DATE As Date
Public DOCUMENT_NO As String
Public DOCUMENT_TYPE As Long
Public DOCUMENT_CATEGORY As Long
Public FROM_DATE As Date
Public TO_DATE As Date
Public PART_GROUP_ID As Long
Public TX_TYPE As String
Public EXPORT_ITEM_ID As Long
Public FROM_HOUSE_ID As Long
Public PIG_ID As Long
Public TO_PIG_ID As Long
Public TO_HOUSE_ID As Long
Public PIG_STATUS As Long
Public PIG_TYPE As String
Public COMMIT_FLAG As String
Public SALE_FLAG As String
Public TX_AMOUNT As Double
Public HOUSE_GROUP_ID As Long
Public REPLACE_FLAG As String
Public EXPENSE_TYPE_NAME As String
Public PART_GROUP_NAME As String
Public DocTypeSet As String
Public YYYYMM As String
Public STATUS_GROUP_ID As Long

Public PIG_NO  As String
Public PIG_STATUS_NO As String
Public PIG_AGE As Double
Public PART_TYPE_NO  As String
Public EXPENSE_TYPE_NO As String
Public TO_PIG_COUNT As Double
Public PART_GROUP_NO As String
Public AGE_CODE  As String
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

   SIMULATE_FLAG = glbUser.SIMULATE_FLAG
   If Ind = 1 Then
      SelectStr1 = "SELECT MI.*, CM.*, PG.PART_GROUP_ID, PG.PART_GROUP_NAME, ET.EXPENSE_TYPE_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PI.PART_ITEM_ID = MI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EXPENSE_TYPE ET ON (MI.EXPENSE_TYPE = ET.EXPENSE_TYPE_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT DISTINCT FROM_HOUSE_ID  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "

      WhereStr = ""
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT DISTINCT PIG_ID  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "

      WhereStr = ""
   ElseIf Ind = 4 Then
       SelectStr1 = "SELECT PG.PART_GROUP_ID, MI.EXPENSE_TYPE, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "

      WhereStr = ""
   ElseIf Ind = 5 Then
       SelectStr1 = "SELECT CM.DOCUMENT_DATE, CM.DOCUMENT_NO, CM.DOCUMENT_CATEGORY, CM.DOCUMENT_TYPE, CM.TX_TYPE, MI.PIG_ID, MI.FROM_HOUSE_ID, MI.TO_HOUSE_ID, CM.PIG_STATUS, SUM(CM.TX_AMOUNT) TX_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "

      WhereStr = ""
   ElseIf Ind = 6 Then
       SelectStr1 = "SELECT MI.EXPENSE_TYPE, PG.PART_GROUP_ID, CM.DOCUMENT_DATE, CM.DOCUMENT_NO, CM.DOCUMENT_CATEGORY, CM.DOCUMENT_TYPE, CM.TX_TYPE, CM.PIG_ID, CM.FROM_HOUSE_ID, CM.TO_HOUSE_ID, CM.PIG_STATUS, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "

      WhereStr = ""
   ElseIf Ind = 7 Then
       SelectStr1 = "SELECT PT.PART_GROUP_ID, MI.EXPENSE_TYPE, CM.PIG_ID, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (CM.FROM_HOUSE_ID = LC.LOCATION_ID) "

      WhereStr = ""
   ElseIf Ind = 8 Then      'LoadSumPigHouseBalance
       SelectStr1 = "SELECT PT.PART_GROUP_ID, MI.EXPENSE_TYPE, CM.PIG_ID, CM.FROM_HOUSE_ID, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI2 ON (MI.PART_ITEM_ID = PI2.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI2.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (CM.FROM_HOUSE_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CM.PIG_ID = PI.PART_ITEM_ID) "
      
      WhereStr = ""
   ElseIf Ind = 9 Then
       SelectStr1 = "SELECT PT.PART_GROUP_ID, MI.EXPENSE_TYPE, CM.PIG_ID, CM.PIG_STATUS, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI2 ON (MI.PART_ITEM_ID = PI2.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI2.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (CM.FROM_HOUSE_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CM.PIG_ID = PI.PART_ITEM_ID) "
      
      WhereStr = ""
   ElseIf Ind = 10 Then
       SelectStr1 = "SELECT CM.PIG_STATUS, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "
      
      WhereStr = ""
   ElseIf Ind = 11 Then
       SelectStr1 = "SELECT CM.FROM_HOUSE_ID, CM.PIG_ID, MI.PART_ITEM_ID, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "
      
      WhereStr = "WHERE (MI.PART_ITEM_ID IS NOT NULL) "
   ElseIf Ind = 12 Then
       SelectStr1 = "SELECT CM.FROM_HOUSE_ID, CM.PIG_ID, MI.EXPENSE_TYPE, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "

      WhereStr = "WHERE (MI.EXPENSE_TYPE IS NOT NULL) "
   ElseIf Ind = 13 Then
       SelectStr1 = "SELECT CM.PIG_ID, CM.FROM_HOUSE_ID, PG.PART_GROUP_ID, MI.EXPENSE_TYPE, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT, SUM(CM.TX_AMOUNT) TX_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "

      WhereStr = ""
   ElseIf Ind = 14 Then
       SelectStr1 = "SELECT CM.PIG_ID, CM.FROM_HOUSE_ID, MI.EXPENSE_TYPE, PG.PART_GROUP_ID, CM.DOCUMENT_DATE, CM.DOCUMENT_NO, CM.DOCUMENT_CATEGORY, CM.DOCUMENT_TYPE, CM.TX_TYPE, CM.PIG_ID, CM.FROM_HOUSE_ID, CM.TO_HOUSE_ID, CM.PIG_STATUS, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "

      WhereStr = ""
   ElseIf Ind = 15 Then
       SelectStr1 = "SELECT CM.FROM_HOUSE_ID, MI.EXPENSE_TYPE, PG.PART_GROUP_ID, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT, SUM(CM.TO_PIG_COUNT) TX_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "

      WhereStr = ""
   ElseIf Ind = 16 Then
       SelectStr1 = "SELECT CM.PIG_ID, CM.PIG_STATUS, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "

      WhereStr = ""
   ElseIf Ind = 17 Then
       SelectStr1 = "SELECT CM.PIG_ID, CM.FROM_HOUSE_ID, CM.TO_HOUSE_ID, PG.PART_GROUP_ID, MI.EXPENSE_TYPE, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT, SUM(CM.TX_AMOUNT) TX_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "

      WhereStr = ""
   ElseIf Ind = 18 Then
       SelectStr1 = "SELECT CM.PIG_ID, CM.TO_PIG_ID, CM.FROM_HOUSE_ID, CM.TO_HOUSE_ID, PG.PART_GROUP_ID, MI.EXPENSE_TYPE, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT, SUM(CM.TX_AMOUNT) TX_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "

      WhereStr = ""
   ElseIf Ind = 19 Then
       SelectStr1 = "SELECT CM.YYYYMM, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "

      WhereStr = "WHERE (MI.PART_ITEM_ID IS NOT NULL) "
   ElseIf Ind = 20 Then
       SelectStr1 = "SELECT CM.FROM_HOUSE_ID, CM.PIG_ID, MI.PART_ITEM_ID, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
       SelectStr1 = SelectStr1 & ",CM.DOCUMENT_DATE,CM.DOCUMENT_TYPE,CM.DOCUMENT_CATEGORY "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "

      'WhereStr = "WHERE (MI.PART_ITEM_ID IS NOT NULL) "
   ElseIf Ind = 21 Then
       SelectStr1 = "SELECT PT.PART_GROUP_ID, MI.EXPENSE_TYPE, CM.PIG_ID,CM.PIG_STATUS, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (CM.FROM_HOUSE_ID = LC.LOCATION_ID) "
      
      WhereStr = ""
   ElseIf Ind = 22 Then 'frmExportToSumFarm              -->>      COST_ITEM
      SelectStr1 = "SELECT PI1.PART_NO PIG_NO,PI1.PIG_TYPE,PS.PRODUCT_STATUS_NO PIG_STATUS_NO,EI.PIG_AGE "
      SelectStr1 = SelectStr1 & ",PT.PART_TYPE_NO,CM.TX_TYPE,ET.EXPENSE_TYPE_NO,CM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",CM.DOCUMENT_DATE,SUM(TX_AMOUNT) TX_AMOUNT "
      SelectStr1 = SelectStr1 & ",SUM(CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI1 ON (CM.PIG_ID = PI1.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EXPORT_ITEM EI ON (EI.EXPORT_ITEM_ID = CM.EXPORT_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EXPENSE_TYPE ET ON (MI.EXPENSE_TYPE = ET.EXPENSE_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (PS.PRODUCT_STATUS_ID = CM.PIG_STATUS) "
      
      WhereStr = "WHERE (CM.DOCUMENT_TYPE = 5 AND CM.DOCUMENT_CATEGORY = 1) OR CM.DOCUMENT_TYPE = 10 OR DOCUMENT_TYPE = 13 "           ' 5 ใบหมูเกิด 10 ใบขายเชื่อ 13 ใบขายสด
   ElseIf Ind = 23 Then       'frmExportToSumFarm              -->>           COST                    เป็น Table ที่ใช้บอกจำนวน อย่างเดี่ยว
      SelectStr1 = "SELECT PI1.PART_NO PIG_NO,PI1.PIG_TYPE,PS.PRODUCT_STATUS_NO PIG_STATUS_NO,EI.PIG_AGE "
      SelectStr1 = SelectStr1 & ",CM.TX_TYPE,CM.DOCUMENT_TYPE "
      SelectStr1 = SelectStr1 & ",CM.DOCUMENT_DATE,SUM(TX_AMOUNT) TX_AMOUNT "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM CAPITAL_MOVEMENT CM "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI1 ON (CM.PIG_ID = PI1.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EXPORT_ITEM EI ON (EI.EXPORT_ITEM_ID = CM.EXPORT_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (PS.PRODUCT_STATUS_ID = CM.PIG_STATUS) "
      
      WhereStr = "WHERE (CM.DOCUMENT_TYPE = 5 AND CM.DOCUMENT_CATEGORY = 1) OR CM.DOCUMENT_TYPE = 10 OR DOCUMENT_TYPE = 13 "           ' 5 ใบหมูเกิด 10 ใบขายเชื่อ 13 ใบขายสด
   ElseIf Ind = 24 Then          'frmExportToSumFarm           -->>   COST_BALANCE           เป็น Table ต้นทุนคงเหลือของสุกร
      SelectStr1 = "SELECT PG.PART_GROUP_NO, ET.EXPENSE_TYPE_NO, PI1.PART_NO PIG_NO,PI1.PIG_TYPE, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr1 = SelectStr1 & ",CM.YYYYMM,CM.DOCUMENT_DATE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI1 ON (CM.PIG_ID = PI1.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EXPENSE_TYPE ET ON (ET.EXPENSE_TYPE_ID = MI.EXPENSE_TYPE) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (CM.FROM_HOUSE_ID = LC.LOCATION_ID) "
      
      WhereStr = ""
   ElseIf Ind = 25 Then          '
      SelectStr1 = "SELECT EI.AGE_CODE,CM.PIG_STATUS, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
            
      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EXPORT_ITEM EI ON (EI.EXPORT_ITEM_ID = CM.EXPORT_ITEM_ID) "
      
      WhereStr = ""
   ElseIf Ind = 26 Then          '
      SelectStr1 = "SELECT EI.PIG_AGE,CM.PIG_ID,CM.PIG_STATUS, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MOVEMENT_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN EXPORT_ITEM EI ON (EI.EXPORT_ITEM_ID = CM.EXPORT_ITEM_ID) "
      
      WhereStr = ""
   End If
   
   If MOVEMENT_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MOVEMENT_ITEM_ID = " & MOVEMENT_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (MOVEMENT_ITEM_ID = " & MOVEMENT_ITEM_ID & ") "
      End If
   End If
   
   If CAPITAL_MOVEMENT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CAPITAL_MOVEMENT_ID = " & CAPITAL_MOVEMENT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CAPITAL_MOVEMENT_ID = " & CAPITAL_MOVEMENT_ID & ") "
      End If
   End If
   
   If EXPORT_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EXPORT_ITEM_ID = " & EXPORT_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EXPORT_ITEM_ID = " & EXPORT_ITEM_ID & ") "
      End If
   End If
   
   If FROM_HOUSE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FROM_HOUSE_ID = " & FROM_HOUSE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (FROM_HOUSE_ID = " & FROM_HOUSE_ID & ") "
      End If
   End If

   If PIG_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PIG_ID = " & PIG_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PIG_ID = " & PIG_ID & ") "
      End If
   End If

   If PIG_STATUS > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.PIG_STATUS = " & PIG_STATUS & ") "
      Else
         WhereStr = WhereStr & "AND (CM.PIG_STATUS = " & PIG_STATUS & ") "
      End If
   End If

   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      End If
   End If
   
   If DOCUMENT_CATEGORY > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DOCUMENT_CATEGORY = " & DOCUMENT_CATEGORY & ") "
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_CATEGORY = " & DOCUMENT_CATEGORY & ") "
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
   
   If Len(PIG_TYPE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PI.PIG_TYPE = '" & PIG_TYPE & "') "
      Else
         WhereStr = WhereStr & "AND (PI.PIG_TYPE = '" & PIG_TYPE & "') "
      End If
   End If
   
   If Len(COMMIT_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (CM.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
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
   
   If Len(SALE_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SALE_FLAG = '" & SALE_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (SALE_FLAG = '" & SALE_FLAG & "') "
      End If
   End If
   
   If Len(TX_TYPE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (TX_TYPE = '" & TX_TYPE & "') "
      Else
         WhereStr = WhereStr & "AND (TX_TYPE = '" & TX_TYPE & "') "
      End If
   End If
   
   If HOUSE_GROUP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FROM_HOUSE_ID IN (SELECT LOCATION_ID FROM HGROUP_ITEM WHERE (HOUSE_GROUP_ID = " & HOUSE_GROUP_ID & ") AND (SELECT_FLAG = 'Y')))"
      Else
         WhereStr = WhereStr & "AND (FROM_HOUSE_ID IN (SELECT LOCATION_ID FROM HGROUP_ITEM WHERE (HOUSE_GROUP_ID = " & HOUSE_GROUP_ID & ") AND (SELECT_FLAG = 'Y')))"
      End If
   End If
   
   If Len(REPLACE_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.REPLACE_FLAG = '" & REPLACE_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (CM.REPLACE_FLAG = '" & REPLACE_FLAG & "') "
      End If
   End If
   
   If Len(DocTypeSet) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.DOCUMENT_TYPE IN " & DocTypeSet & ") "
      Else
         WhereStr = WhereStr & "AND (CM.DOCUMENT_TYPE IN " & DocTypeSet & ") "
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

   If STATUS_GROUP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = WhereStr & "WHERE (CM.PIG_STATUS IN (SELECT HGI.STATUS_ID FROM SGROUP_ITEM HGI WHERE (HGI.STATUS_GROUP_ID = " & STATUS_GROUP_ID & ") AND (HGI.SELECT_FLAG = 'Y'))) "
      Else
         WhereStr = WhereStr & "AND (CM.PIG_STATUS IN (SELECT HGI.STATUS_ID FROM SGROUP_ITEM HGI WHERE (HGI.STATUS_GROUP_ID = " & STATUS_GROUP_ID & ") AND (HGI.SELECT_FLAG = 'Y'))) "
      End If
   End If
         
   GroupByStr = ""
   If Ind = 4 Then
      GroupByStr = "GROUP BY PG.PART_GROUP_ID, MI.EXPENSE_TYPE "
   ElseIf Ind = 5 Then
      GroupByStr = "GROUP BY CM.DOCUMENT_DATE, CM.DOCUMENT_NO, CM.DOCUMENT_CATEGORY, CM.DOCUMENT_TYPE, CM.TX_TYPE, MI.PIG_ID, MI.FROM_HOUSE_ID, MI.TO_HOUSE_ID, CM.PIG_STATUS "
   ElseIf Ind = 6 Then
      GroupByStr = "GROUP BY MI.EXPENSE_TYPE , PG.PART_GROUP_ID, CM.DOCUMENT_DATE, CM.DOCUMENT_NO, CM.DOCUMENT_CATEGORY, CM.DOCUMENT_TYPE, CM.TX_TYPE, CM.PIG_ID, CM.FROM_HOUSE_ID, CM.TO_HOUSE_ID, CM.PIG_STATUS "
   ElseIf Ind = 7 Then
      GroupByStr = "GROUP BY PT.PART_GROUP_ID, MI.EXPENSE_TYPE, CM.PIG_ID "
   ElseIf Ind = 8 Then
      GroupByStr = "GROUP BY PT.PART_GROUP_ID, MI.EXPENSE_TYPE, CM.PIG_ID, CM.FROM_HOUSE_ID "
   ElseIf Ind = 9 Then
      GroupByStr = "GROUP BY PT.PART_GROUP_ID, MI.EXPENSE_TYPE, CM.PIG_ID, CM.PIG_STATUS "
   ElseIf Ind = 10 Then
      GroupByStr = "GROUP BY CM.PIG_STATUS "
   ElseIf Ind = 11 Then
      GroupByStr = "GROUP BY CM.FROM_HOUSE_ID, CM.PIG_ID, MI.PART_ITEM_ID "
   ElseIf Ind = 12 Then
      GroupByStr = "GROUP BY CM.FROM_HOUSE_ID, CM.PIG_ID, MI.EXPENSE_TYPE "
   ElseIf Ind = 13 Then
      GroupByStr = "GROUP BY CM.PIG_ID, CM.FROM_HOUSE_ID, PG.PART_GROUP_ID, MI.EXPENSE_TYPE "
   ElseIf Ind = 14 Then
      GroupByStr = "GROUP BY CM.PIG_ID, CM.FROM_HOUSE_ID, MI.EXPENSE_TYPE , PG.PART_GROUP_ID, CM.DOCUMENT_DATE, CM.DOCUMENT_NO, CM.DOCUMENT_CATEGORY, CM.DOCUMENT_TYPE, CM.TX_TYPE, CM.PIG_ID, CM.FROM_HOUSE_ID, CM.TO_HOUSE_ID, CM.PIG_STATUS "
   ElseIf Ind = 15 Then
      GroupByStr = "GROUP BY CM.FROM_HOUSE_ID, MI.EXPENSE_TYPE , PG.PART_GROUP_ID "
   ElseIf Ind = 16 Then
      GroupByStr = "GROUP BY CM.PIG_ID, CM.PIG_STATUS "
   ElseIf Ind = 17 Then
      GroupByStr = "GROUP BY CM.PIG_ID, CM.FROM_HOUSE_ID, CM.TO_HOUSE_ID,PG.PART_GROUP_ID, MI.EXPENSE_TYPE "
   ElseIf Ind = 18 Then
      GroupByStr = "GROUP BY CM.PIG_ID, CM.TO_PIG_ID, CM.FROM_HOUSE_ID, CM.TO_HOUSE_ID,PG.PART_GROUP_ID, MI.EXPENSE_TYPE "
   ElseIf Ind = 19 Then
      GroupByStr = "GROUP BY CM.YYYYMM "
   ElseIf Ind = 20 Then
      GroupByStr = "GROUP BY CM.FROM_HOUSE_ID, CM.PIG_ID, MI.PART_ITEM_ID,CM.DOCUMENT_DATE,CM.DOCUMENT_TYPE,CM.DOCUMENT_CATEGORY "
   ElseIf Ind = 21 Then
      GroupByStr = "GROUP BY PT.PART_GROUP_ID, MI.EXPENSE_TYPE, CM.PIG_ID,CM.PIG_STATUS "
   ElseIf Ind = 22 Then
      GroupByStr = "GROUP BY PI1.PART_NO,PI1.PIG_TYPE,PS.PRODUCT_STATUS_NO,EI.PIG_AGE,PT.PART_TYPE_NO,CM.TX_TYPE,ET.EXPENSE_TYPE_NO,CM.DOCUMENT_TYPE,CM.DOCUMENT_DATE "
   ElseIf Ind = 23 Then       'frmExportToSumFarm
      GroupByStr = "GROUP BY PI1.PART_NO,PI1.PIG_TYPE,PS.PRODUCT_STATUS_NO,EI.PIG_AGE,CM.TX_TYPE,CM.DOCUMENT_TYPE,CM.DOCUMENT_DATE "
   ElseIf Ind = 24 Then       'frmExportToSumFarm
      GroupByStr = "GROUP BY PG.PART_GROUP_NO, ET.EXPENSE_TYPE_NO, PI1.PART_NO,PI1.PIG_TYPE,CM.YYYYMM,CM.DOCUMENT_DATE "
   ElseIf Ind = 25 Then
      GroupByStr = "GROUP BY EI.AGE_CODE,CM.PIG_STATUS "
   ElseIf Ind = 26 Then
      GroupByStr = "GROUP BY EI.PIG_AGE , CM.PIG_ID, CM.PIG_STATUS "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY MI.CAPITAL_MOVEMENT_ID "
      Else
         OrderByStr = "ORDER BY MI.CAPITAL_MOVEMENT_ID "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY FROM_HOUSE_ID "
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY PIG_ID "
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY PIG_ID "
   ElseIf Ind = 5 Then
      OrderByStr = "ORDER BY MI.MOVEMENT_ITEM_ID "
   ElseIf Ind = 6 Then
      OrderByStr = "ORDER BY MI.MOVEMENT_ITEM_ID "
   ElseIf Ind = 7 Then
      OrderByStr = "ORDER BY PT.PART_GROUP_ID, MI.EXPENSE_TYPE, CM.PIG_ID "
   ElseIf Ind = 8 Then
      'Pt.PART_GROUP_ID , Mi.EXPENSE_TYPE, Cm.PIG_ID, Cm.FROM_HOUSE_ID
'      If OrderBy = 1 Then
'         OrderByStr = "ORDER BY PI.PART_NO "
'      ElseIf OrderBy = 2 Then
'         OrderByStr = "ORDER BY PI.PART_NO, LC.HOUSE_NAME "
'      Else
         OrderByStr = "ORDER BY CM.PIG_ID,CM.FROM_HOUSE_ID "
'      End If
   ElseIf Ind = 9 Then
      OrderByStr = "ORDER BY PI.PART_NO "
   ElseIf Ind = 10 Then
      OrderByStr = "ORDER BY CM.PIG_STATUS "
   ElseIf Ind = 11 Then
      OrderByStr = "ORDER BY CM.PIG_ID "
   ElseIf Ind = 12 Then
      OrderByStr = "ORDER BY CM.PIG_ID "
   ElseIf Ind = 13 Then
      OrderByStr = "ORDER BY CM.PIG_ID "
   ElseIf Ind = 14 Then
      OrderByStr = "ORDER BY CM.PIG_ID "
   ElseIf Ind = 15 Then
      OrderByStr = "ORDER BY CM.PIG_ID "
   ElseIf Ind = 16 Then
      OrderByStr = "ORDER BY CM.PIG_ID, CM.PIG_STATUS "
   ElseIf Ind = 17 Then
      OrderByStr = "ORDER BY CM.PIG_ID "
   ElseIf Ind = 18 Then
      OrderByStr = "ORDER BY CM.PIG_ID "
   ElseIf Ind = 19 Then
      OrderByStr = "ORDER BY CM.YYYYMM "
   ElseIf Ind = 20 Then
      OrderByStr = "ORDER BY CM.PIG_ID "
   ElseIf Ind = 21 Then
      OrderByStr = "ORDER BY PT.PART_GROUP_ID, MI.EXPENSE_TYPE, CM.PIG_ID,CM.PIG_STATUS "
   ElseIf Ind = 22 Then
      OrderByStr = "ORDER BY CM.DOCUMENT_DATE "
   ElseIf Ind = 23 Then
      OrderByStr = "ORDER BY CM.DOCUMENT_DATE "
   ElseIf Ind = 24 Then
   
   ElseIf Ind = 25 Then
      OrderByStr = "ORDER BY CM.PIG_STATUS,EI.AGE_CODE "
   ElseIf Ind = 26 Then
      OrderByStr = "ORDER BY EI.PIG_AGE , CM.PIG_ID, CM.PIG_STATUS "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   If Ind = 24 Then
      OrderByStr = ""
      OrderTypeStr = ""
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      MOVEMENT_ITEM_ID = NVLI(Rs("MOVEMENT_ITEM_ID"), 0)
      CAPITAL_MOVEMENT_ID = NVLI(Rs("CAPITAL_MOVEMENT_ID"), 0)
      EXPENSE_TYPE = NVLI(Rs("EXPENSE_TYPE"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
      
      SIMULATE_FLAG = NVLS(Rs("SIMULATE_FLAG"), "")
      EXPENSE_TYPE_NAME = NVLS(Rs("EXPENSE_TYPE_NAME"), "")
      PART_GROUP_NAME = NVLS(Rs("PART_GROUP_NAME"), "")
      PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), 0)
      BATCH_ID = NVLI(Rs("BATCH_ID"), 0)
   ElseIf Ind = 2 Then
      FROM_HOUSE_ID = NVLI(Rs("FROM_HOUSE_ID"), 0)
   ElseIf Ind = 3 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
   ElseIf Ind = 4 Then
      EXPENSE_TYPE = NVLI(Rs("EXPENSE_TYPE"), 0)
      PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), 0)
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
   ElseIf Ind = 5 Then
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_CATEGORY = NVLI(Rs("DOCUMENT_CATEGORY"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      FROM_HOUSE_ID = NVLI(Rs("FROM_HOUSE_ID"), 0)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
   ElseIf Ind = 6 Then
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_CATEGORY = NVLI(Rs("DOCUMENT_CATEGORY"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), 0)
      FROM_HOUSE_ID = NVLI(Rs("FROM_HOUSE_ID"), 0)
      TO_HOUSE_ID = NVLI(Rs("TO_HOUSE_ID"), 0)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      EXPENSE_TYPE = NVLI(Rs("EXPENSE_TYPE"), 0)
      PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), 0)
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
   ElseIf Ind = 7 Then
      PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), 0)
      EXPENSE_TYPE = NVLI(Rs("EXPENSE_TYPE"), 0)
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
   ElseIf Ind = 8 Then
      PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), 0)
      EXPENSE_TYPE = NVLI(Rs("EXPENSE_TYPE"), 0)
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      FROM_HOUSE_ID = NVLI(Rs("FROM_HOUSE_ID"), 0)
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
   ElseIf Ind = 9 Then
      PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), 0)
      EXPENSE_TYPE = NVLI(Rs("EXPENSE_TYPE"), 0)
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), 0)
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
   ElseIf Ind = 10 Then
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), 0)
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
   ElseIf Ind = 11 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      FROM_HOUSE_ID = NVLI(Rs("FROM_HOUSE_ID"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
   ElseIf Ind = 12 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      FROM_HOUSE_ID = NVLI(Rs("FROM_HOUSE_ID"), 0)
      EXPENSE_TYPE = NVLI(Rs("EXPENSE_TYPE"), 0)
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
   ElseIf Ind = 13 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      FROM_HOUSE_ID = NVLI(Rs("FROM_HOUSE_ID"), 0)
      EXPENSE_TYPE = NVLI(Rs("EXPENSE_TYPE"), 0)
      PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), 0)
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
   ElseIf Ind = 14 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      FROM_HOUSE_ID = NVLI(Rs("FROM_HOUSE_ID"), 0)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_CATEGORY = NVLI(Rs("DOCUMENT_CATEGORY"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), 0)
      FROM_HOUSE_ID = NVLI(Rs("FROM_HOUSE_ID"), 0)
      TO_HOUSE_ID = NVLI(Rs("TO_HOUSE_ID"), 0)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      EXPENSE_TYPE = NVLI(Rs("EXPENSE_TYPE"), 0)
      PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), 0)
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
   ElseIf Ind = 15 Then
      FROM_HOUSE_ID = NVLI(Rs("FROM_HOUSE_ID"), 0)
      EXPENSE_TYPE = NVLI(Rs("EXPENSE_TYPE"), 0)
      PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), 0)
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
   ElseIf Ind = 16 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), 0)
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
   ElseIf Ind = 17 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      FROM_HOUSE_ID = NVLI(Rs("FROM_HOUSE_ID"), 0)
      TO_HOUSE_ID = NVLI(Rs("TO_HOUSE_ID"), 0)
      EXPENSE_TYPE = NVLI(Rs("EXPENSE_TYPE"), 0)
      PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), 0)
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
   ElseIf Ind = 18 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      TO_PIG_ID = NVLI(Rs("TO_PIG_ID"), 0)
      FROM_HOUSE_ID = NVLI(Rs("FROM_HOUSE_ID"), 0)
      TO_HOUSE_ID = NVLI(Rs("TO_HOUSE_ID"), 0)
      EXPENSE_TYPE = NVLI(Rs("EXPENSE_TYPE"), 0)
      PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), 0)
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
   ElseIf Ind = 19 Then
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
   ElseIf Ind = 20 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      FROM_HOUSE_ID = NVLI(Rs("FROM_HOUSE_ID"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_CATEGORY = NVLI(Rs("DOCUMENT_CATEGORY"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
   ElseIf Ind = 21 Then
      PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), 0)
      EXPENSE_TYPE = NVLI(Rs("EXPENSE_TYPE"), 0)
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), 0)
   ElseIf Ind = 22 Then
      PIG_NO = NVLS(Rs("PIG_NO"), "")
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      PIG_STATUS_NO = NVLS(Rs("PIG_STATUS_NO"), "")
      PIG_AGE = NVLD(Rs("PIG_AGE"), 0)
      
      PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      EXPENSE_TYPE_NO = NVLS(Rs("EXPENSE_TYPE_NO"), "")
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
   ElseIf Ind = 23 Then
      PIG_NO = NVLS(Rs("PIG_NO"), "")
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      PIG_STATUS_NO = NVLS(Rs("PIG_STATUS_NO"), "")
      PIG_AGE = NVLD(Rs("PIG_AGE"), 0)
      
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
   ElseIf Ind = 24 Then
      PIG_NO = NVLS(Rs("PIG_NO"), "")
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      EXPENSE_TYPE_NO = NVLS(Rs("EXPENSE_TYPE_NO"), "")
      PART_GROUP_NO = NVLS(Rs("PART_GROUP_NO"), "")
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
   ElseIf Ind = 25 Then
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), -1)
      AGE_CODE = NVLS(Rs("AGE_CODE"), "")
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
   ElseIf Ind = 26 Then
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), -1)
      PIG_AGE = NVLI(Rs("PIG_AGE"), -1)
      PIG_ID = NVLI(Rs("PIG_ID"), -1)
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
   End If
End Function
Public Function GetKey1() As String
Dim TempKey As String
   TempKey = EXPENSE_TYPE & "-" & _
                           PART_GROUP_ID & "-" & _
                           DOCUMENT_NO & "-" & _
                           DOCUMENT_CATEGORY & "-" & _
                           DOCUMENT_TYPE & "-" & _
                           TX_TYPE & "-" & _
                           TO_HOUSE_ID & "-" & _
                           PIG_STATUS
      GetKey1 = TempKey
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String

   TempID1 = EXPENSE_TYPE
   If EXPENSE_TYPE <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("MOVEMENT_ITEM_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO MOVEMENT_ITEM " & _
                     "( " & _
                     "   MOVEMENT_ITEM_ID, " & _
                     "   CAPITAL_MOVEMENT_ID, " & _
                     "   EXPENSE_TYPE, " & _
                     "   PART_ITEM_ID, " & _
                     "   CAPITAL_AMOUNT, "
         SQL1 = SQL1 & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                        ID & "," & _
                        CAPITAL_MOVEMENT_ID & "," & _
                        TempID1 & "," & _
                        TempID2 & "," & _
                        CAPITAL_AMOUNT & ","
               SQL1 = SQL1 & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         MOVEMENT_ITEM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE MOVEMENT_ITEM SET " & _
                     "    CAPITAL_MOVEMENT_ID = " & CAPITAL_MOVEMENT_ID & ", " & _
                     "    PART_ITEM_ID = " & PART_ITEM_ID & ", " & _
                     "    EXPENSE_TYPE = " & EXPENSE_TYPE & ", " & _
                     "    CAPITAL_AMOUNT = " & CAPITAL_AMOUNT & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE MOVEMENT_ITEM_ID = " & MOVEMENT_ITEM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM MOVEMENT_ITEM WHERE MOVEMENT_ITEM_ID = " & MOVEMENT_ITEM_ID
      m_Conn.Execute (SQL1)
End Function

Public Function ClearData()
Dim SQL1 As String
      SQL1 = "DELETE FROM MOVEMENT_ITEM "
      m_Conn.Execute (SQL1)
End Function


