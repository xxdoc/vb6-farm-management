VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLossItem"
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

Public LOSS_ITEM_ID As Long
Public CAPITAL_LOSS_ID As Long
Public EXPENSE_TYPE As Long
Public PART_ITEM_ID As Long
Public CAPITAL_AMOUNT As Double

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
Public SIMULATE_FLAG As String
Public YYYYMM As String

Public PARENT_FLAG  As String
Public YEAR_SEQ_ID As Long
Public WEEK_NO As String

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
Dim SubLen As Long
Dim NewStr As String

   SIMULATE_FLAG = glbUser.SIMULATE_FLAG
   If Ind = 1 Then
      SelectStr1 = "SELECT MI.*, CM.*, PG.PART_GROUP_ID, PG.PART_GROUP_NAME, ET.EXPENSE_TYPE_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOSS_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_LOSS CM ON (MI.CAPITAL_LOSS_ID = CM.CAPITAL_LOSS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PI.PART_ITEM_ID = MI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN EXPENSE_TYPE ET ON (MI.EXPENSE_TYPE = ET.EXPENSE_TYPE_ID) "

      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT DISTINCT FROM_HOUSE_ID  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOSS_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_LOSS CM ON (MI.CAPITAL_LOSS_ID = CM.CAPITAL_LOSS_ID) "

      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CM.PIG_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_TYPE PDT ON (PI.PIG_TYPE = PDT.PRODUCT_TYPE_NO) "

      FromStr = FromStr & "LEFT OUTER JOIN PRTITEM_MAP PM ON (PI.PART_ITEM_ID = PM.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN YEAR_WEEK YW ON (YW.YEAR_WEEK_ID = PM.YEAR_WEEK_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN YEAR_SEQ YQ ON (YQ.YEAR_SEQ_ID = YW.YEAR_SEQ_ID) "
      WhereStr = ""
      
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT DISTINCT PIG_ID  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOSS_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_LOSS CM ON (MI.CAPITAL_LOSS_ID = CM.CAPITAL_LOSS_ID) "
   
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CM.PIG_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_TYPE PDT ON (PI.PIG_TYPE = PDT.PRODUCT_TYPE_NO) "

      FromStr = FromStr & "LEFT OUTER JOIN PRTITEM_MAP PM ON (PI.PART_ITEM_ID = PM.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN YEAR_WEEK YW ON (YW.YEAR_WEEK_ID = PM.YEAR_WEEK_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN YEAR_SEQ YQ ON (YQ.YEAR_SEQ_ID = YW.YEAR_SEQ_ID) "
      WhereStr = ""
   ElseIf Ind = 4 Then
       SelectStr1 = "SELECT PG.PART_GROUP_ID, MI.EXPENSE_TYPE, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOSS_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_LOSS_ID = CM.CAPITAL_LOSS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "

      WhereStr = ""
   ElseIf Ind = 5 Then
       SelectStr1 = "SELECT CM.DOCUMENT_DATE, CM.DOCUMENT_NO, CM.DOCUMENT_CATEGORY, CM.DOCUMENT_TYPE, CM.TX_TYPE, MI.PIG_ID, MI.FROM_HOUSE_ID, MI.TO_HOUSE_ID, MI.PIG_STATUS, SUM(CM.TX_AMOUNT) TX_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOSS_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_LOSS_ID = CM.CAPITAL_LOSS_ID) "

      WhereStr = ""
   ElseIf Ind = 6 Then
       SelectStr1 = "SELECT MI.EXPENSE_TYPE, PG.PART_GROUP_ID, CM.DOCUMENT_DATE, CM.DOCUMENT_NO, CM.DOCUMENT_CATEGORY, CM.DOCUMENT_TYPE, CM.TX_TYPE, CM.PIG_ID, CM.FROM_HOUSE_ID, CM.TO_HOUSE_ID, CM.PIG_STATUS, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOSS_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_LOSS_ID = CM.CAPITAL_LOSS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "

      WhereStr = ""
   ElseIf Ind = 7 Then
       SelectStr1 = "SELECT PT.PART_GROUP_ID, MI.EXPENSE_TYPE, CM.PIG_ID, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOSS_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_LOSS_ID = CM.CAPITAL_LOSS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (CM.FROM_HOUSE_ID = LC.LOCATION_ID) "

      WhereStr = ""
   ElseIf Ind = 8 Then
       SelectStr1 = "SELECT PT.PART_GROUP_ID, MI.EXPENSE_TYPE, CM.PIG_ID, CM.FROM_HOUSE_ID, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOSS_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_LOSS CM ON (MI.CAPITAL_LOSS_ID = CM.CAPITAL_LOSS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI2 ON (MI.PART_ITEM_ID = PI2.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI2.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (CM.FROM_HOUSE_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CM.PIG_ID = PI.PART_ITEM_ID) "
      
      'FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_TYPE PDT ON (PI.PIG_TYPE = PDT.PRODUCT_TYPE_NO) "
      
      WhereStr = ""
   ElseIf Ind = 9 Then
       SelectStr1 = "SELECT PT.PART_GROUP_ID, MI.EXPENSE_TYPE, CM.PIG_ID, CM.PIG_STATUS, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOSS_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_LOSS_ID = CM.CAPITAL_LOSS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI2 ON (MI.PART_ITEM_ID = PI2.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI2.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (CM.FROM_HOUSE_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CM.PIG_ID = PI.PART_ITEM_ID) "

      WhereStr = ""
   ElseIf Ind = 10 Then
       SelectStr1 = "SELECT CM.PIG_STATUS, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOSS_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_LOSS_ID = CM.CAPITAL_LOSS_ID) "

      WhereStr = ""
   ElseIf Ind = 11 Then
       SelectStr1 = "SELECT CM.FROM_HOUSE_ID, CM.PIG_ID, MI.PART_ITEM_ID, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOSS_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_LOSS_ID = CM.CAPITAL_LOSS_ID) "

      WhereStr = "WHERE (MI.PART_ITEM_ID IS NOT NULL) "
   ElseIf Ind = 12 Then
       SelectStr1 = "SELECT CM.FROM_HOUSE_ID, CM.PIG_ID, MI.EXPENSE_TYPE, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOSS_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_LOSS_ID = CM.CAPITAL_LOSS_ID) "

      WhereStr = "WHERE (MI.EXPENSE_TYPE IS NOT NULL) "
   ElseIf Ind = 13 Then
       SelectStr1 = "SELECT CM.PIG_ID, CM.FROM_HOUSE_ID, PG.PART_GROUP_ID, MI.EXPENSE_TYPE, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT, SUM(CM.TX_AMOUNT) TX_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOSS_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_LOSS_ID = CM.CAPITAL_LOSS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "

      WhereStr = ""
   ElseIf Ind = 14 Then
       SelectStr1 = "SELECT CM.PIG_ID, CM.FROM_HOUSE_ID, MI.EXPENSE_TYPE, PG.PART_GROUP_ID, CM.DOCUMENT_DATE, CM.DOCUMENT_NO, CM.DOCUMENT_CATEGORY, CM.DOCUMENT_TYPE, CM.TX_TYPE, CM.PIG_ID, CM.FROM_HOUSE_ID, CM.TO_HOUSE_ID, CM.PIG_STATUS, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOSS_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_LOSS_ID = CM.CAPITAL_LOSS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "

      WhereStr = ""
   ElseIf Ind = 15 Then
       SelectStr1 = "SELECT CM.FROM_HOUSE_ID, MI.EXPENSE_TYPE, PG.PART_GROUP_ID, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT, SUM(CM.TO_PIG_COUNT) TX_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOSS_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_LOSS_ID = CM.CAPITAL_LOSS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "

      WhereStr = ""
   ElseIf Ind = 16 Then
       SelectStr1 = "SELECT CM.PIG_ID, CM.PIG_STATUS, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOSS_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_LOSS_ID = CM.CAPITAL_LOSS_ID) "

      WhereStr = ""
   ElseIf Ind = 17 Then
       SelectStr1 = "SELECT CM.PIG_ID, CM.FROM_HOUSE_ID, CM.TO_HOUSE_ID, PG.PART_GROUP_ID, MI.EXPENSE_TYPE, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT, SUM(CM.TX_AMOUNT) TX_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOSS_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_LOSS_ID = CM.CAPITAL_LOSS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "

      WhereStr = ""
   ElseIf Ind = 18 Then
       SelectStr1 = "SELECT CM.PIG_ID, CM.TO_PIG_ID, CM.FROM_HOUSE_ID, CM.TO_HOUSE_ID, PG.PART_GROUP_ID, MI.EXPENSE_TYPE, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT, SUM(CM.TX_AMOUNT) TX_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM LOSS_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_LOSS_ID = CM.CAPITAL_LOSS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "

      WhereStr = ""
   ElseIf Ind = 19 Then
       SelectStr1 = "SELECT CM.YYYYMM, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOSS_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_LOSS CM ON (MI.CAPITAL_LOSS_ID = CM.CAPITAL_LOSS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "

   ElseIf Ind = 20 Then
       SelectStr1 = "SELECT SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LOSS_ITEM MI "
      FromStr = FromStr & "LEFT OUTER JOIN CAPITAL_LOSS CM ON (MI.CAPITAL_LOSS_ID = CM.CAPITAL_LOSS_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (MI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "
      
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI1 ON (CM.PIG_ID = PI1.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_TYPE PDT ON (PI1.PIG_TYPE = PDT.PRODUCT_TYPE_NO) "
      
      'WhereStr = " WHERE (PI1.PART_NO IS NOT NULL) "
      
   End If
   
   If LOSS_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LOSS_ITEM_ID = " & LOSS_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LOSS_ITEM_ID = " & LOSS_ITEM_ID & ") "
      End If
   End If
   
   If CAPITAL_LOSS_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CAPITAL_LOSS_ID = " & CAPITAL_LOSS_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CAPITAL_LOSS_ID = " & CAPITAL_LOSS_ID & ") "
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
         WhereStr = "WHERE (PIG_STATUS = " & PIG_STATUS & ") "
      Else
         WhereStr = WhereStr & "AND (PIG_STATUS = " & PIG_STATUS & ") "
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
   
   If Len(PARENT_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CAPITAL_FLAG = '" & PARENT_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (CAPITAL_FLAG = '" & PARENT_FLAG & "') "
      End If
   End If
   
   If YEAR_SEQ_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (YEAR_SEQ_ID = " & YEAR_SEQ_ID & ") "
      Else
         WhereStr = WhereStr & "AND (YEAR_SEQ_ID = " & YEAR_SEQ_ID & ") "
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
      GroupByStr = "GROUP BY PG.PART_GROUP_ID, MI.EXPENSE_TYPE "
   ElseIf Ind = 5 Then
      GroupByStr = "GROUP BY CM.DOCUMENT_DATE, CM.DOCUMENT_NO, CM.DOCUMENT_CATEGORY, CM.DOCUMENT_TYPE, CM.TX_TYPE, MI.PIG_ID, MI.FROM_HOUSE_ID, MI.TO_HOUSE_ID, MI.PIG_STATUS "
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
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY CAPITAL_LOSS_ID "
      Else
         OrderByStr = "ORDER BY CAPITAL_LOSS_ID "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY FROM_HOUSE_ID "
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY PIG_ID "
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY PIG_ID "
   ElseIf Ind = 5 Then
      OrderByStr = "ORDER BY MI.LOSS_ITEM_ID "
   ElseIf Ind = 6 Then
      OrderByStr = "ORDER BY MI.LOSS_ITEM_ID "
   ElseIf Ind = 7 Then
      OrderByStr = "ORDER BY PI.PART_NO "
   ElseIf Ind = 8 Then
         OrderByStr = "ORDER BY CM.PIG_ID "
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
   End If
      
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   If Ind = 20 Then
      OrderByStr = ""
      OrderTypeStr = ""
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      LOSS_ITEM_ID = NVLI(Rs("LOSS_ITEM_ID"), 0)
      CAPITAL_LOSS_ID = NVLI(Rs("CAPITAL_LOSS_ID"), 0)
      EXPENSE_TYPE = NVLI(Rs("EXPENSE_TYPE"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      CAPITAL_AMOUNT = NVLD(Rs("CAPITAL_AMOUNT"), 0)
      
      EXPENSE_TYPE_NAME = NVLS(Rs("EXPENSE_TYPE_NAME"), "")
      PART_GROUP_NAME = NVLS(Rs("PART_GROUP_NAME"), "")
      PART_GROUP_ID = NVLI(Rs("PART_GROUP_ID"), 0)
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
      Call glbDatabaseMngr.GetSeqID("LOSS_ITEM_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO LOSS_ITEM " & _
                     "( " & _
                     "   LOSS_ITEM_ID, " & _
                     "   CAPITAL_LOSS_ID, " & _
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
                        CAPITAL_LOSS_ID & "," & _
                        TempID1 & "," & _
                        TempID2 & "," & _
                        CAPITAL_AMOUNT & ","
               SQL1 = SQL1 & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         LOSS_ITEM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE LOSS_ITEM SET " & _
                     "    CAPITAL_LOSS_ID = " & CAPITAL_LOSS_ID & ", " & _
                     "    PART_ITEM_ID = " & PART_ITEM_ID & ", " & _
                     "    EXPENSE_TYPE = " & EXPENSE_TYPE & ", " & _
                     "    CAPITAL_AMOUNT = " & CAPITAL_AMOUNT & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE LOSS_ITEM_ID = " & LOSS_ITEM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM LOSS_ITEM WHERE LOSS_ITEM_ID = " & LOSS_ITEM_ID
      m_Conn.Execute (SQL1)
End Function

Public Sub CopyFromMovementItem(Mi As CMovementItem)
   PART_ITEM_ID = Mi.PART_ITEM_ID
   EXPENSE_TYPE = Mi.EXPENSE_TYPE
   CAPITAL_AMOUNT = Mi.CAPITAL_AMOUNT
End Sub

Public Function ClearData()
Dim SQL1 As String
      SQL1 = "DELETE FROM LOSS_ITEM "
      m_Conn.Execute (SQL1)
End Function




