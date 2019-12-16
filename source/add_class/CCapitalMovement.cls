VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCapitalMovement"
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

Public CAPITAL_MOVEMENT_ID As Long
Public DOCUMENT_NO As String
Public DOCUMENT_DATE As Date
Public IVD_ID As Long
Public BL_ID As Long
Public DOCUMENT_CATEGORY As Long
Public DOCUMENT_TYPE As Long
Public TX_AMOUNT As Double
Public TX_TYPE As String
Public PIG_ID As Long
Public FROM_HOUSE_ID As Long
Public TO_HOUSE_ID As Long
Public PIG_STATUS As Long
Public EXPORT_ITEM_ID As Long
Public IMPORT_ITEM_ID As Long
Public COMMIT_FLAG As String
Public TX_SEQ As Long
Public FOR_PIG_COUNT As Long
Public REPLACE_FLAG As String
Public BILLING_DOC_ID As Long
Public TO_PIG_ID As Long
Public SIMULATE_FLAG As String
Public BATCH_ID As Long

Public YYYYMM As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public EXPENSE_TYPE As Long
Public PART_GROUP_ID As Long
Public PIG_TYPE As String
Public HOUSE_SALE_FLAG As String
Public TOTAL_PRICE As Double
Public TOTAL_WEIGHT As Double
Public DELETE_BALANCE_FLAG As String
Public TO_PIG_COUNT As Long
Public PARENT_FLAG As String
Public HOUSE_NAME As String
Public PIG_NO As String
Public TOTAL_CAPITAL As Double
Public EXCLUDE_STATUS_ID As Long
Public PIG_FLAG As String
Public DocTypeSet As String

Public DISCOUNT_AMOUNT As Double

Public PRODUCT_STATUS_NO As String
Public PRODUCT_STATUS_NAME As String
Public PART_NO As String

Public YEAR_SEQ_ID  As Long
Public WEEK_NO As String
Public MovementItems As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set MovementItems = Nothing
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
Dim GroupByStr As String
Dim SubLen As Long
Dim NewStr As String

   SIMULATE_FLAG = glbUser.SIMULATE_FLAG
   If Ind = 1 Then
      SelectStr1 = "SELECT CM.*, PI.PART_NO, PI.PIG_TYPE,  LC.LOCATION_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM CAPITAL_MOVEMENT CM "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CM.PIG_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (CM.FROM_HOUSE_ID = LC.LOCATION_ID) "

      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT DISTINCT FROM_HOUSE_ID  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM CAPITAL_MOVEMENT CM "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CM.PIG_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRTITEM_MAP PM ON (PI.PART_ITEM_ID = PM.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN YEAR_WEEK YW ON (YW.YEAR_WEEK_ID = PM.YEAR_WEEK_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN YEAR_SEQ YQ ON (YQ.YEAR_SEQ_ID = YW.YEAR_SEQ_ID) "
      
      WhereStr = ""
   ElseIf Ind = 3 Then      'LoadMovementPig
      SelectStr1 = "SELECT DISTINCT PIG_ID  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM CAPITAL_MOVEMENT CM "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CM.PIG_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_TYPE PDT ON (PI.PIG_TYPE = PDT.PRODUCT_TYPE_NO) "
      
      FromStr = FromStr & "LEFT OUTER JOIN PRTITEM_MAP PM ON (PI.PART_ITEM_ID = PM.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN YEAR_WEEK YW ON (YW.YEAR_WEEK_ID = PM.YEAR_WEEK_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN YEAR_SEQ YQ ON (YQ.YEAR_SEQ_ID = YW.YEAR_SEQ_ID) "
      
      WhereStr = ""
   ElseIf Ind = 4 Then  'CReportInventory033
       'SelectStr1 = "SELECT CM.DOCUMENT_DATE, CM.DOCUMENT_NO, CM.DOCUMENT_CATEGORY, CM.DOCUMENT_TYPE, CM.TX_TYPE, CM.PIG_ID, CM.FROM_HOUSE_ID, CM.TO_HOUSE_ID, CM.PIG_STATUS, SUM(CM.TX_AMOUNT) TX_AMOUNT,DP.PRIORITY1, CM.TX_SEQ "
       SelectStr1 = "SELECT CM.DOCUMENT_DATE, CM.DOCUMENT_NO, CM.DOCUMENT_CATEGORY, CM.DOCUMENT_TYPE, CM.TX_TYPE, CM.PIG_ID, CM.FROM_HOUSE_ID, CM.TO_HOUSE_ID, CM.PIG_STATUS, SUM(CM.TX_AMOUNT) TX_AMOUNT,DP.PRIORITY1 "
       
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM CAPITAL_MOVEMENT CM "
      FromStr = FromStr & "LEFT OUTER JOIN DOCUMENT_PRIORITY DP ON ((CM.DOCUMENT_TYPE = DP.DOCUMENT_TYPE) AND (CM.DOCUMENT_CATEGORY = DP.AREA))"

      WhereStr = ""
   ElseIf Ind = 5 Then
       SelectStr1 = "SELECT DISTINCT CM.FROM_HOUSE_ID, CM.PIG_ID "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM CAPITAL_MOVEMENT CM "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (CM.FROM_HOUSE_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CM.PIG_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_TYPE PDT ON (PI.PIG_TYPE = PDT.PRODUCT_TYPE_NO) "

      WhereStr = ""
   ElseIf Ind = 6 Then
       SelectStr1 = "SELECT DISTINCT CM.PIG_STATUS, CM.PIG_ID "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM CAPITAL_MOVEMENT CM "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (CM.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (CM.FROM_HOUSE_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CM.PIG_ID = PI.PART_ITEM_ID) "
      
      WhereStr = ""
   ElseIf Ind = 7 Then
       SelectStr1 = "SELECT CM.PIG_STATUS,PS.PRODUCT_STATUS_NO, SUM(CM.TX_AMOUNT) XXXX, SUM(DI.TOTAL_WEIGHT) TOTAL_WEIGHT, SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM CAPITAL_MOVEMENT CM "
      FromStr = FromStr & "LEFT OUTER JOIN EXPORT_ITEM EI ON (CM.EXPORT_ITEM_ID = EI.EXPORT_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BLD ON (BLD.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN DO_ITEM DI ON (DI.DO_ID = BLD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CM.PIG_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (CM.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
      
      WhereStr = WhereStr & "WHERE (EI.LINK_ID = DI.LINK_ID) "
   ElseIf Ind = 8 Then
       SelectStr1 = "SELECT MIN(CM.DOCUMENT_DATE) DOCUMENT_DATE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM CAPITAL_MOVEMENT CM "
      WhereStr = ""
   ElseIf Ind = 9 Then
       SelectStr1 = "SELECT CM.PIG_ID, CM.PIG_STATUS, SUM(CM.TX_AMOUNT) TX_AMOUNT, SUM(DI.TOTAL_WEIGHT) TOTAL_WEIGHT, SUM(DI.TOTAL_PRICE) TOTAL_PRICE,SUM(DI.DISCOUNT_AMOUNT) DISCOUNT_AMOUNT "
       SelectStr1 = SelectStr1 & ",PS.PRODUCT_STATUS_NO "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM CAPITAL_MOVEMENT CM "
      FromStr = FromStr & "LEFT OUTER JOIN EXPORT_ITEM EI ON (CM.EXPORT_ITEM_ID = EI.EXPORT_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BLD ON (BLD.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN DO_ITEM DI ON (DI.DO_ID = BLD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (CM.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CM.PIG_ID = PI.PART_ITEM_ID) "
      
      'FromStr = FromStr & "LEFT OUTER JOIN SGROUP_ITEM SG ON (CM.PIG_ID = PI.PART_ITEM_ID) "
      
      
      WhereStr = WhereStr & "WHERE (EI.LINK_ID = DI.LINK_ID) "
   ElseIf Ind = 10 Then
       SelectStr1 = "SELECT DISTINCT CM.FROM_HOUSE_ID, CM.TO_HOUSE_ID, CM.PIG_ID "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM CAPITAL_MOVEMENT CM "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (CM.FROM_HOUSE_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CM.PIG_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_TYPE PDT ON (PI.PIG_TYPE = PDT.PRODUCT_TYPE_NO) "

      WhereStr = ""
   ElseIf Ind = 11 Then
       SelectStr1 = "SELECT CM.PIG_ID, CM.FROM_HOUSE_ID, CM.TO_HOUSE_ID, SUM(CM.TX_AMOUNT) TX_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM CAPITAL_MOVEMENT CM "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (CM.FROM_HOUSE_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CM.PIG_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_TYPE PDT ON (PI.PIG_TYPE = PDT.PRODUCT_TYPE_NO) "

      WhereStr = ""
   ElseIf Ind = 12 Then
       SelectStr1 = "SELECT DISTINCT CM.FROM_HOUSE_ID, CM.TO_HOUSE_ID, CM.PIG_ID, CM.TO_PIG_ID "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM CAPITAL_MOVEMENT CM "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (CM.FROM_HOUSE_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CM.PIG_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_TYPE PDT ON (PI.PIG_TYPE = PDT.PRODUCT_TYPE_NO) "

      WhereStr = ""
   ElseIf Ind = 13 Then
       SelectStr1 = "SELECT CM.PIG_ID, CM.TO_PIG_ID, CM.FROM_HOUSE_ID, CM.TO_HOUSE_ID, SUM(CM.TX_AMOUNT) TX_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM CAPITAL_MOVEMENT CM "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (CM.FROM_HOUSE_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CM.PIG_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_TYPE PDT ON (PI.PIG_TYPE = PDT.PRODUCT_TYPE_NO) "

      WhereStr = ""
   ElseIf Ind = 14 Then
      SelectStr1 = "SELECT DISTINCT CM.PIG_ID,PI.PART_NO,PI.PIG_TYPE,CM.PIG_STATUS,PDS.PRODUCT_STATUS_NO,PDS.PRODUCT_STATUS_NAME  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM CAPITAL_MOVEMENT CM "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CM.PIG_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_TYPE PDT ON (PI.PIG_TYPE = PDT.PRODUCT_TYPE_NO) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PDS ON (CM.PIG_STATUS = PDS.PRODUCT_STATUS_ID) "
      
      WhereStr = ""
   End If
   
   If CAPITAL_MOVEMENT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CAPITAL_MOVEMENT_ID = " & CAPITAL_MOVEMENT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CAPITAL_MOVEMENT_ID = " & CAPITAL_MOVEMENT_ID & ") "
      End If
   End If
   
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (CM.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      End If
   End If
   
   If Len(DocTypeSet) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.DOCUMENT_TYPE IN " & DocTypeSet & ") "
      Else
         WhereStr = WhereStr & "AND (CM.DOCUMENT_TYPE IN " & DocTypeSet & ") "
      End If
   End If
   
   If DOCUMENT_CATEGORY > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.DOCUMENT_CATEGORY = " & DOCUMENT_CATEGORY & ") "
      Else
         WhereStr = WhereStr & "AND (CM.DOCUMENT_CATEGORY = " & DOCUMENT_CATEGORY & ") "
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
   
   If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (CM.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (CM.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (CM.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (CM.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If Len(PIG_TYPE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PIG_TYPE = '" & PIG_TYPE & "') "
      Else
         WhereStr = WhereStr & "AND (PIG_TYPE = '" & PIG_TYPE & "') "
      End If
   End If
   
   If SIMULATE_FLAG = "Y" Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.SIMULATE_FLAG = '" & SIMULATE_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (CM.SIMULATE_FLAG = '" & SIMULATE_FLAG & "') "
      End If
   Else
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE ((CM.SIMULATE_FLAG <> 'Y') OR (CM.SIMULATE_FLAG IS NULL)) "
      Else
         WhereStr = WhereStr & "AND ((CM.SIMULATE_FLAG <> 'Y') OR (CM.SIMULATE_FLAG IS NULL)) "
      End If
   End If
   
   If Len(HOUSE_SALE_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SALE_FLAG = '" & HOUSE_SALE_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (SALE_FLAG = '" & HOUSE_SALE_FLAG & "') "
      End If
   End If
   
   If Len(PIG_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PIG_FLAG = '" & PIG_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (PIG_FLAG = '" & PIG_FLAG & "') "
      End If
   End If
   
   If Len(TX_TYPE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.TX_TYPE = '" & TX_TYPE & "') "
      Else
         WhereStr = WhereStr & "AND (CM.TX_TYPE = '" & TX_TYPE & "') "
      End If
   End If
   
   If Len(COMMIT_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (CM.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
   End If
   
   If Len(PARENT_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CAPITAL_FLAG = '" & PARENT_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (CAPITAL_FLAG = '" & PARENT_FLAG & "') "
      End If
   End If
   
   If Len(REPLACE_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (REPLACE_FLAG = '" & REPLACE_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (REPLACE_FLAG = '" & REPLACE_FLAG & "') "
      End If
   End If
   
   If BILLING_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      End If
   End If
   
   If EXCLUDE_STATUS_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.PIG_STATUS NOT IN (SELECT STATUS_ID FROM SGROUP_ITEM WHERE (STATUS_GROUP_ID = " & EXCLUDE_STATUS_ID & ") AND (SELECT_FLAG = 'Y'))) "
      Else
         WhereStr = WhereStr & "AND (CM.PIG_STATUS NOT IN (SELECT STATUS_ID FROM SGROUP_ITEM WHERE (STATUS_GROUP_ID = " & EXCLUDE_STATUS_ID & ") AND (SELECT_FLAG = 'Y'))) "
      End If
   End If
   
   If BATCH_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.BATCH_ID = " & BATCH_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CM.BATCH_ID = " & BATCH_ID & ") "
      End If
   Else
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE ((CM.BATCH_ID <= 0) OR (CM.BATCH_ID IS NULL)) "
      Else
         WhereStr = WhereStr & "AND ((CM.BATCH_ID <= 0) OR (CM.BATCH_ID IS NULL)) "
      End If
   End If

   If YEAR_SEQ_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (YW.YEAR_SEQ_ID = " & YEAR_SEQ_ID & ") "
      Else
         WhereStr = WhereStr & "AND (YW.YEAR_SEQ_ID = " & YEAR_SEQ_ID & ") "
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
   
   GroupByStr = ""
   If Ind = 4 Then
      'GroupByStr = "GROUP BY CM.DOCUMENT_DATE, CM.DOCUMENT_NO, CM.DOCUMENT_CATEGORY, CM.DOCUMENT_TYPE, CM.TX_TYPE, CM.PIG_ID, CM.FROM_HOUSE_ID, CM.TO_HOUSE_ID, CM.PIG_STATUS,DP.PRIORITY1, CM.TX_SEQ "
      GroupByStr = "GROUP BY CM.DOCUMENT_DATE, CM.DOCUMENT_NO, CM.DOCUMENT_CATEGORY, CM.DOCUMENT_TYPE, CM.TX_TYPE, CM.PIG_ID, CM.FROM_HOUSE_ID, CM.TO_HOUSE_ID, CM.PIG_STATUS,DP.PRIORITY1 "
   ElseIf Ind = 7 Then
      GroupByStr = "GROUP BY CM.PIG_STATUS,PS.PRODUCT_STATUS_NO "
   ElseIf Ind = 9 Then
      GroupByStr = "GROUP BY PS.PRODUCT_STATUS_NO,CM.PIG_STATUS,CM.PIG_ID  "
   ElseIf Ind = 11 Then
      GroupByStr = "GROUP BY CM.PIG_ID, CM.FROM_HOUSE_ID, CM.TO_HOUSE_ID "
   ElseIf Ind = 13 Then
      GroupByStr = "GROUP BY CM.PIG_ID, CM.TO_PIG_ID, CM.FROM_HOUSE_ID, CM.TO_HOUSE_ID "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY CAPITAL_MOVEMENT_ID "
      Else
         OrderByStr = "ORDER BY CAPITAL_MOVEMENT_ID "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY FROM_HOUSE_ID "
   ElseIf Ind = 3 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PART_NO "
      Else
         OrderByStr = "ORDER BY PIG_ID "
      End If
   ElseIf Ind = 4 Then
      'OrderByStr = "ORDER BY CAPITAL_MOVEMENT_ID "
      'OrderByStr = "ORDER BY CM.DOCUMENT_DATE ASC, DP.PRIORITY1 ASC, CM.DOCUMENT_NO ASC, CM.TX_SEQ "
      OrderByStr = "ORDER BY CM.DOCUMENT_DATE ASC, DP.PRIORITY1 ASC, CM.DOCUMENT_NO "
   ElseIf Ind = 5 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY LC.LOCATION_NAME, PI.PART_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY LC.LOCATION_NAME, PI.PART_NO "
      Else
         OrderByStr = "ORDER BY LC.LOCATION_NO, PI.PART_NO "
      End If
   ElseIf Ind = 6 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PI.PART_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PI.PART_NO, PS.PRODUCT_STATUS_NAME "
      Else
         OrderByStr = "ORDER BY PS.PRODUCT_STATUS_NAME, PI.PART_NO "
      End If
   ElseIf Ind = 7 Then
      OrderByStr = "ORDER BY PS.PRODUCT_STATUS_NO "
   ElseIf Ind = 9 Then
      OrderByStr = "ORDER BY PS.PRODUCT_STATUS_NO,CM.PIG_STATUS,CM.PIG_ID "
   ElseIf Ind = 10 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY LC.LOCATION_NAME, PI.PART_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY LC.LOCATION_NAME, PI.PART_NO "
      Else
         OrderByStr = "ORDER BY LC.LOCATION_NO, PI.PART_NO "
      End If
   ElseIf Ind = 11 Then
      OrderByStr = "ORDER BY CM.CAPITAL_MOVEMENT_ID "
   ElseIf Ind = 12 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY LC.LOCATION_NAME, PI.PART_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY LC.LOCATION_NAME, PI.PART_NO "
      Else
         OrderByStr = "ORDER BY LC.LOCATION_NO, PI.PART_NO "
      End If
   ElseIf Ind = 14 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY CM.PIG_STATUS,PART_NO "
      Else
         OrderByStr = "ORDER BY CM.PIG_STATUS,CM.PIG_ID "
      End If
   End If
   
   If OrderByStr <> "" Then
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
      CAPITAL_MOVEMENT_ID = NVLI(Rs("CAPITAL_MOVEMENT_ID"), 0)
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      IVD_ID = NVLI(Rs("IVD_ID"), 0)
      BL_ID = NVLI(Rs("BL_ID"), 0)
      DOCUMENT_CATEGORY = NVLI(Rs("DOCUMENT_CATEGORY"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      FROM_HOUSE_ID = NVLI(Rs("FROM_HOUSE_ID"), 0)
      TO_HOUSE_ID = NVLI(Rs("TO_HOUSE_ID"), 0)
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), 0)
      IMPORT_ITEM_ID = NVLI(Rs("IMPORT_ITEM_ID"), 0)
      EXPORT_ITEM_ID = NVLI(Rs("EXPORT_ITEM_ID"), 0)
      COMMIT_FLAG = NVLS(Rs("COMMIT_FLAG"), "")
      TX_SEQ = NVLI(Rs("TX_SEQ"), 0)
      REPLACE_FLAG = NVLS(Rs("REPLACE_FLAG"), "")
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), 0)
      TO_PIG_ID = NVLI(Rs("TO_PIG_ID"), 0)
      SIMULATE_FLAG = NVLS(Rs("SIMULATE_FLAG"), "")
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      BATCH_ID = NVLI(Rs("BATCH_ID"), 0)
      
      PIG_NO = NVLS(Rs("PART_NO"), "")
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      HOUSE_NAME = NVLS(Rs("LOCATION_NAME"), "")
   ElseIf Ind = 2 Then
      FROM_HOUSE_ID = NVLI(Rs("FROM_HOUSE_ID"), 0)
   ElseIf Ind = 3 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
   ElseIf Ind = 4 Then
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      DOCUMENT_CATEGORY = NVLI(Rs("DOCUMENT_CATEGORY"), 0)
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      FROM_HOUSE_ID = NVLI(Rs("FROM_HOUSE_ID"), 0)
      TO_HOUSE_ID = NVLI(Rs("TO_HOUSE_ID"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      TX_TYPE = NVLS(Rs("TX_TYPE"), "")
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), 0)
   ElseIf Ind = 5 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      FROM_HOUSE_ID = NVLI(Rs("FROM_HOUSE_ID"), 0)
   ElseIf Ind = 6 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), 0)
   ElseIf Ind = 7 Then
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      TX_AMOUNT = NVLD(Rs("XXXX"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      
   ElseIf Ind = 8 Then
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
   ElseIf Ind = 9 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
      DISCOUNT_AMOUNT = NVLD(Rs("DISCOUNT_AMOUNT"), 0)
      
   ElseIf Ind = 10 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      FROM_HOUSE_ID = NVLI(Rs("FROM_HOUSE_ID"), 0)
      TO_HOUSE_ID = NVLI(Rs("TO_HOUSE_ID"), 0)
   ElseIf Ind = 11 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      FROM_HOUSE_ID = NVLI(Rs("FROM_HOUSE_ID"), 0)
      TO_HOUSE_ID = NVLI(Rs("TO_HOUSE_ID"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
   ElseIf Ind = 12 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      TO_PIG_ID = NVLI(Rs("TO_PIG_ID"), 0)
      FROM_HOUSE_ID = NVLI(Rs("FROM_HOUSE_ID"), 0)
      TO_HOUSE_ID = NVLI(Rs("TO_HOUSE_ID"), 0)
   ElseIf Ind = 13 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      TO_PIG_ID = NVLI(Rs("TO_PIG_ID"), 0)
      FROM_HOUSE_ID = NVLI(Rs("FROM_HOUSE_ID"), 0)
      TO_HOUSE_ID = NVLI(Rs("TO_HOUSE_ID"), 0)
      TX_AMOUNT = NVLD(Rs("TX_AMOUNT"), 0)
   ElseIf Ind = 14 Then
      PIG_ID = NVLI(Rs("PIG_ID"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), 0)
      PRODUCT_STATUS_NO = NVLS(Rs("PRODUCT_STATUS_NO"), "")
      PRODUCT_STATUS_NAME = NVLS(Rs("PRODUCT_STATUS_NAME"), "")
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
Dim TempID3 As String
Dim TempID4 As String
Dim TempID5 As String
Dim TempID6 As String
Dim TempID7 As String
Dim TempID8 As String
Dim TempID9 As String
Dim TempID10 As String
Dim TempID11 As String

   TempID1 = IVD_ID
   If IVD_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = BL_ID
   If BL_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID3 = PIG_ID
   If PIG_ID <= 0 Then
      TempID3 = "NULL"
   End If
   
   TempID4 = FROM_HOUSE_ID
   If FROM_HOUSE_ID <= 0 Then
      TempID4 = "NULL"
   End If
   
   TempID5 = TO_HOUSE_ID
   If TO_HOUSE_ID <= 0 Then
      TempID5 = "NULL"
   End If
   
   TempID6 = PIG_STATUS
   If PIG_STATUS <= 0 Then
      TempID6 = "NULL"
   End If
   
   TempID7 = EXPORT_ITEM_ID
   If EXPORT_ITEM_ID <= 0 Then
      TempID7 = "NULL"
   End If
   
   TempID8 = IMPORT_ITEM_ID
   If IMPORT_ITEM_ID <= 0 Then
      TempID8 = "NULL"
   End If
   
   TempID9 = BILLING_DOC_ID
   If BILLING_DOC_ID <= 0 Then
      TempID9 = "NULL"
   End If
   
   TempID10 = TO_PIG_ID
   If TO_PIG_ID <= 0 Then
      TempID10 = "NULL"
   End If
   
   TempID11 = BATCH_ID
   If BATCH_ID <= 0 Then
      TempID11 = "NULL"
   End If
   
   If REPLACE_FLAG = "" Then
      REPLACE_FLAG = "N"
   End If
   
   SIMULATE_FLAG = glbUser.SIMULATE_FLAG
   YYYYMM = Mid(DateToStringInt(DOCUMENT_DATE), 1, 7)
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("CAPITAL_MOVEMENT_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO CAPITAL_MOVEMENT " & _
                     "( " & _
                     "   CAPITAL_MOVEMENT_ID, " & _
                     "   DOCUMENT_NO, " & _
                     "   DOCUMENT_DATE, " & _
                     "   IVD_ID, " & _
                     "   BL_ID, " & _
                     "   DOCUMENT_CATEGORY, " & _
                     "   DOCUMENT_TYPE, "
         SQL1 = SQL1 & _
                     "   TX_AMOUNT, " & _
                     "   TX_TYPE, " & _
                     "   PIG_ID, " & _
                     "   FROM_HOUSE_ID, " & _
                     "   TO_HOUSE_ID, " & _
                     "   PIG_STATUS, " & _
                     "   EXPORT_ITEM_ID, " & _
                     "   IMPORT_ITEM_ID, " & _
                     "   COMMIT_FLAG, " & _
                     "   TX_SEQ, "
            SQL1 = SQL1 & _
                     "   TO_PIG_COUNT, " & _
                     "   REPLACE_FLAG, " & _
                     "   BILLING_DOC_ID, " & _
                     "   TO_PIG_ID, " & _
                     "   SIMULATE_FLAG, " & _
                     "   YYYYMM, " & _
                     "   BATCH_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                        "'" & ChangeQuote(Trim(DOCUMENT_NO)) & "'," & _
                        "'" & ChangeQuote(DateToStringInt(Trim(DOCUMENT_DATE))) & "'," & _
                        TempID1 & "," & _
                        TempID2 & "," & _
                        DOCUMENT_CATEGORY & "," & _
                        DOCUMENT_TYPE & ","
               SQL1 = SQL1 & _
                        TX_AMOUNT & "," & _
                        "'" & ChangeQuote(Trim(TX_TYPE)) & "'," & _
                        TempID3 & ", " & _
                        TempID4 & ", " & _
                        TempID5 & ", " & _
                        TempID6 & ", " & _
                        TempID7 & ", " & _
                        TempID8 & ", " & _
                        "'" & ChangeQuote(Trim(COMMIT_FLAG)) & "'," & _
                        TX_SEQ & "," & _
                        TO_PIG_COUNT & "," & _
                        "'" & ChangeQuote(Trim(REPLACE_FLAG)) & "'," & _
                        TempID9 & ", " & _
                        TempID10 & ", " & _
                        "'" & ChangeQuote(Trim(SIMULATE_FLAG)) & "'," & _
                        "'" & ChangeQuote(Trim(YYYYMM)) & "'," & _
                        TempID11 & ", " & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         CAPITAL_MOVEMENT_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE CAPITAL_MOVEMENT SET " & _
                     "    DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "', " & _
                     "    DOCUMENT_DATE = '" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "', " & _
                     "    IVD_ID = " & TempID1 & ", " & _
                     "    BL_ID = " & TempID2 & ", " & _
                     "    DOCUMENT_CATEGORY = " & DOCUMENT_CATEGORY & ", " & _
                     "    DOCUMENT_TYPE = " & DOCUMENT_TYPE & ", " & _
                     "    TX_AMOUNT = " & TX_AMOUNT & ", " & _
                     "    TX_TYPE = '" & TX_TYPE & "', " & _
                     "    PIG_ID = " & TempID3 & ", " & _
                     "    FROM_HOUSE_ID = " & TempID4 & ", " & _
                     "    TO_HOUSE_ID = " & TempID5 & ", " & _
                     "    PIG_STATUS = " & TempID6 & ", " & _
                     "    EXPORT_ITEM_ID = " & TempID7 & ", " & _
                     "    IMPORT_ITEM_ID = " & TempID8 & ", " & _
                     "    COMMIT_FLAG = '" & ChangeQuote(Trim(COMMIT_FLAG)) & "', " & _
                     "    TX_SEQ = " & TX_SEQ & ", " & _
                     "    TO_PIG_COUNT = " & TO_PIG_COUNT & ", " & _
                     "    REPLACE_FLAG = '" & ChangeQuote(Trim(REPLACE_FLAG)) & "', " & _
                     "    BILLING_DOC_ID = " & TempID9 & ", " & _
                     "    TO_PIG_ID = " & TempID10 & ", " & _
                     "    SIMULATE_FLAG = '" & ChangeQuote(Trim(SIMULATE_FLAG))
            SQL1 = SQL1 & _
                     "    YYYYMM = '" & ChangeQuote(Trim(YYYYMM)) & "', " & _
                     "    BATCH_ID = " & TempID11 & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE CAPITAL_MOVEMENT_ID = " & CAPITAL_MOVEMENT_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM MOVEMENT_ITEM WHERE CAPITAL_MOVEMENT_ID = " & CAPITAL_MOVEMENT_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM CAPITAL_MOVEMENT WHERE CAPITAL_MOVEMENT_ID = " & CAPITAL_MOVEMENT_ID
      m_Conn.Execute (SQL1)
End Function

Public Function DeleteDataFromBillingDoc()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM MOVEMENT_ITEM WHERE CAPITAL_MOVEMENT_ID IN (SELECT CAPITAL_MOVEMENT_ID FROM CAPITAL_MOVEMENT WHERE BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM CAPITAL_MOVEMENT WHERE BILLING_DOC_ID = " & BILLING_DOC_ID
      m_Conn.Execute (SQL1)
End Function

Public Function DeleteAllData()
Dim SQL1 As String
      
      SIMULATE_FLAG = glbUser.SIMULATE_FLAG
      
      SQL1 = "DELETE FROM MOVEMENT_ITEM WHERE (CAPITAL_MOVEMENT_ID IN (SELECT CAPITAL_MOVEMENT_ID FROM CAPITAL_MOVEMENT WHERE (SIMULATE_FLAG = '" & SIMULATE_FLAG & "') OR (SIMULATE_FLAG IS NULL))) "
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM CAPITAL_MOVEMENT WHERE (SIMULATE_FLAG = '" & SIMULATE_FLAG & "') OR (SIMULATE_FLAG IS NULL) "
      m_Conn.Execute (SQL1)

      SQL1 = "DELETE FROM LOSS_ITEM WHERE (CAPITAL_LOSS_ID IN (SELECT CAPITAL_LOSS_ID FROM CAPITAL_LOSS WHERE (SIMULATE_FLAG = '" & SIMULATE_FLAG & "') OR (SIMULATE_FLAG IS NULL))) "
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM CAPITAL_LOSS WHERE (SIMULATE_FLAG = '" & SIMULATE_FLAG & "') OR (SIMULATE_FLAG IS NULL) "
      m_Conn.Execute (SQL1)
End Function

Public Function ClearData()
Dim SQL1 As String
Dim TempDate As String
 Dim WhereStr As String
 
   SIMULATE_FLAG = glbUser.SIMULATE_FLAG
   WhereStr = " "
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (CM.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (CM.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (CM.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (CM.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
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
   
   If DELETE_BALANCE_FLAG = "Y" Then
      SQL1 = "DELETE FROM MOVEMENT_ITEM WHERE CAPITAL_MOVEMENT_ID IN (SELECT CAPITAL_MOVEMENT_ID FROM CAPITAL_MOVEMENT CM WHERE (CAPITAL_MOVEMENT_ID > 0) " & WhereStr & ") "
      m_Conn.Execute (SQL1)

      SQL1 = "DELETE FROM CAPITAL_MOVEMENT CM WHERE (CAPITAL_MOVEMENT_ID > 0) " 'ยอดยกมา
      SQL1 = SQL1 & WhereStr
      m_Conn.Execute (SQL1)
   Else
      SQL1 = "DELETE FROM MOVEMENT_ITEM WHERE CAPITAL_MOVEMENT_ID IN (SELECT CAPITAL_MOVEMENT_ID FROM CAPITAL_MOVEMENT CM WHERE (DOCUMENT_CATEGORY <> 3) " & WhereStr & ") "
      m_Conn.Execute (SQL1)

      SQL1 = "DELETE FROM CAPITAL_MOVEMENT CM WHERE (DOCUMENT_CATEGORY <> 3) " 'ยอดยกมา
      SQL1 = SQL1 & WhereStr
      m_Conn.Execute (SQL1)
   End If
End Function

Private Sub Class_Terminate()
   Set MovementItems = Nothing
End Sub
