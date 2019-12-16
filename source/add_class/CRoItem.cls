VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CROItem"
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

Public RO_ITEM_ID As Long
Public BILLING_DOC_ID As Long
Public PART_ITEM_ID As Long
Public LOCATION_ID As Long
Public ITEM_AMOUNT As Double
Public TOTAL_WEIGHT As Double
Public AVG_WEIGHT As Double
Public TOTAL_PRICE As Double
Public AVG_PRICE As Double
Public GUI_ID As Long
Public LINK_ID As Long
Public PIG_STATUS As Long
Public EXPENSE_DESC As String
Public EXPENSE_TYPE As Long
Public TRANSACTION_SEQ As Long

Public BATCH_ID As Long
Public EXPENSE_TYPE_NO As String
Public EXPENSE_TYPE_NAME
Public FROM_DATE As Date
Public TO_DATE As Date
Public COMMIT_FLAG As String
Public PIG_TYPE As String
Public PART_NO As String
Public PIG_STATUS_NAME As String
Public LOCATION_NAME As String
Public DOCUMENT_NO As String
Public DOCUMENT_DATE As Date
Public TX_TYPE As String
Public DOCUMENT_TYPE As Long
Public PRIORITY1 As Long
Public SIMULATE_FLAG As String
Public YYYYMM As String
Public YYYYMM2 As String
Public DEPRECIATION_GOOD_FLAG As String

Public ExpenseRatios As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_DATE = -1
   TO_DATE = -1
   TX_TYPE = "X"
   Set ExpenseRatios = New Collection
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

   SIMULATE_FLAG = glbUser.SIMULATE_FLAG
   If Ind = 1 Then
      SelectStr1 = "SELECT ET.*, UG.*, PT.*, LC.*, PI.*, UN.*, PS.*, IVD.BATCH_ID, IVD.DOCUMENT_DATE, IVD.DOCUMENT_NO, IVD.COMMIT_FLAG, IVD.DOCUMENT_TYPE, IVD.SIMULATE_FLAG, PR.PRIORITY1 "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RO_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN EXPENSE_TYPE ET ON (UG.EXPENSE_TYPE = ET.EXPENSE_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.BILLING_DOC_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN DOCUMENT_PRIORITY PR ON (IVD.DOCUMENT_TYPE = PR.DOCUMENT_TYPE) AND (PR.AREA = 2)"
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (UG.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT IVD.YYYYMM, UG.EXPENSE_TYPE, UG.EXPENSE_DESC, SUM(UG.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RO_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN EXPENSE_TYPE ET ON (UG.EXPENSE_TYPE = ET.EXPENSE_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.BILLING_DOC_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN DOCUMENT_PRIORITY PR ON (IVD.DOCUMENT_TYPE = PR.DOCUMENT_TYPE) AND (PR.AREA = 2)"
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (UG.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
      
      WhereStr = ""
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT IVD.YYYYMM, SUM(UG.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RO_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN EXPENSE_TYPE ET ON (UG.EXPENSE_TYPE = ET.EXPENSE_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.BILLING_DOC_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN DOCUMENT_PRIORITY PR ON (IVD.DOCUMENT_TYPE = PR.DOCUMENT_TYPE) AND (PR.AREA = 2)"
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (UG.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
      
      WhereStr = ""
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT IVD.YYYYMM2, SUM(UG.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM RO_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN EXPENSE_TYPE ET ON (UG.EXPENSE_TYPE = ET.EXPENSE_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.BILLING_DOC_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN DOCUMENT_PRIORITY PR ON (IVD.DOCUMENT_TYPE = PR.DOCUMENT_TYPE) AND (PR.AREA = 2)"
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (UG.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
      
      WhereStr = ""
   ElseIf Ind = 5 Then
      SelectStr1 = "SELECT DISTINCT UG.EXPENSE_TYPE, ET.EXPENSE_TYPE_NO, ET.EXPENSE_TYPE_NAME, UG.EXPENSE_DESC "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RO_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN EXPENSE_TYPE ET ON (UG.EXPENSE_TYPE = ET.EXPENSE_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.BILLING_DOC_ID = IVD.BILLING_DOC_ID) "
      
      WhereStr = ""
   ElseIf Ind = 6 Then
      SelectStr1 = "SELECT IVD.YYYYMM, UG.EXPENSE_TYPE, SUM(UG.TOTAL_PRICE) TOTAL_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM RO_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN EXPENSE_TYPE ET ON (UG.EXPENSE_TYPE = ET.EXPENSE_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (UG.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (UG.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC IVD ON (UG.BILLING_DOC_ID = IVD.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN DOCUMENT_PRIORITY PR ON (IVD.DOCUMENT_TYPE = PR.DOCUMENT_TYPE) AND (PR.AREA = 2)"
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (UG.PIG_STATUS = PS.PRODUCT_STATUS_ID) "
      
      WhereStr = ""
   End If
   
   If RO_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (RO_ITEM_ID = " & RO_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (RO_ITEM_ID = " & RO_ITEM_ID & ") "
      End If
   End If
   
   If BILLING_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UG.BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (UG.BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      End If
   End If
   
   If LOCATION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UG.LOCATION_ID = " & LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (UG.LOCATION_ID = " & LOCATION_ID & ") "
      End If
   End If
   
'   If BATCH_ID > 0 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (BATCH_ID = " & BATCH_ID & ") "
'      Else
'         WhereStr = WhereStr & "AND (BATCH_ID = " & BATCH_ID & ") "
'      End If
'   Else
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE ((BATCH_ID <= 0) OR (BATCH_ID IS NULL)) "
'      Else
'         WhereStr = WhereStr & "AND ((BATCH_ID <= 0) OR (BATCH_ID IS NULL)) "
'      End If
'   End If

   
   If Len(COMMIT_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
   End If
   
   If Len(DEPRECIATION_GOOD_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DEPRECIATION_GOOD_FLAG = '" & DEPRECIATION_GOOD_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (DEPRECIATION_GOOD_FLAG = '" & DEPRECIATION_GOOD_FLAG & "') "
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
   
''   If Ind = 2 Then
'      If PART_TYPE > 0 Then
'         If Len(WhereStr) <= 0 Then
'            WhereStr = "WHERE (PART_TYPE = " & PART_TYPE & ") "
'         Else
'            WhereStr = WhereStr & "AND (PART_TYPE = " & PART_TYPE & ") "
'         End If
'      End If
'
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

'
'   If Trim(PART_NO) <> "" Then
'      If WildCard(PART_NO, SubLen, NewStr) Then
'         If WhereStr = "" Then
'            WhereStr = "WHERE (SUBSTR(PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
'         Else
'            WhereStr = WhereStr & "AND (SUBSTR(PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
'         End If
'      Else
'         If WhereStr = "" Then
'            WhereStr = "WHERE (PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')"
'         Else
'            WhereStr = WhereStr & "AND (PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')"
'         End If
'      End If
'   End If
   
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

   
   GroupByStr = ""
   If Ind = 2 Then
      GroupByStr = "GROUP BY IVD.YYYYMM, UG.EXPENSE_TYPE, UG.EXPENSE_DESC "
   ElseIf Ind = 3 Then
      GroupByStr = "GROUP BY IVD.YYYYMM "
   ElseIf Ind = 4 Then
      GroupByStr = "GROUP BY IVD.YYYYMM2 "
   ElseIf Ind = 6 Then
      GroupByStr = "GROUP BY IVD.YYYYMM, UG.EXPENSE_TYPE "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY RO_ITEM_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY RO_ITEM_ID "
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
         OrderByStr = "ORDER BY DOCUMENT_DATE, DOCUMENT_NO "
      ElseIf OrderBy = 11 Then
         OrderByStr = "ORDER BY IVD.DOCUMENT_DATE ASC, PR.PRIORITY1 ASC, IVD.DOCUMENT_NO ASC, UG.TRANSACTION_SEQ "
      ElseIf OrderBy = 12 Then
         OrderByStr = "ORDER BY ET.EXPENSE_TYPE_NO ASC, UG.EXPENSE_DESC "
      Else
         OrderByStr = "ORDER BY RO_ITEM_ID "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY IVD.YYYYMM "
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY IVD.YYYYMM "
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY IVD.YYYYMM2 "
   ElseIf Ind = 5 Then
      OrderByStr = "ORDER BY EXPENSE_TYPE_NO ASC, EXPENSE_DESC "
   ElseIf Ind = 6 Then
      OrderByStr = "ORDER BY IVD.YYYYMM "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      RO_ITEM_ID = NVLI(Rs("RO_ITEM_ID"), -1)
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
      ITEM_AMOUNT = NVLD(Rs("ITEM_AMOUNT"), 0)
      TOTAL_WEIGHT = NVLD(Rs("TOTAL_WEIGHT"), 0)
      AVG_WEIGHT = NVLD(Rs("AVG_WEIGHT"), 0)
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
      GUI_ID = NVLI(Rs("GUI_ID"), -1)
      PIG_STATUS = NVLI(Rs("PIG_STATUS"), -1)
      LINK_ID = NVLI(Rs("LINK_ID"), -1)
      EXPENSE_DESC = NVLS(Rs("EXPENSE_DESC"), "")
      EXPENSE_TYPE = NVLI(Rs("EXPENSE_TYPE"), -1)
      TRANSACTION_SEQ = NVLI(Rs("TRANSACTION_SEQ"), -1)
      SIMULATE_FLAG = NVLS(Rs("SIMULATE_FLAG"), "")
      BATCH_ID = NVLI(Rs("BATCH_ID"), -1)
      
      EXPENSE_TYPE_NO = NVLS(Rs("EXPENSE_TYPE_NO"), "")
      EXPENSE_TYPE_NAME = NVLS(Rs("EXPENSE_TYPE_NAME"), "")
      COMMIT_FLAG = NVLS(Rs("COMMIT_FLAG"), "")
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      PART_NO = NVLS(Rs("PART_NO"), "")
      PIG_STATUS_NAME = NVLS(Rs("PRODUCT_STATUS_NAME"), "")
      LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      
      DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
      COMMIT_FLAG = NVLS(Rs("COMMIT_FLAG"), "")
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
      PRIORITY1 = NVLI(Rs("PRIORITY1"), -1)
   ElseIf Ind = 2 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      EXPENSE_TYPE = NVLI(Rs("EXPENSE_TYPE"), -1)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      EXPENSE_DESC = NVLS(Rs("EXPENSE_DESC"), "")
   ElseIf Ind = 3 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
   ElseIf Ind = 4 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      YYYYMM2 = NVLS(Rs("YYYYMM2"), "")
   ElseIf Ind = 5 Then
      EXPENSE_TYPE = NVLI(Rs("EXPENSE_TYPE"), -1)
      EXPENSE_DESC = NVLS(Rs("EXPENSE_DESC"), "")
      EXPENSE_TYPE_NO = NVLS(Rs("EXPENSE_TYPE_NO"), "")
      EXPENSE_TYPE_NAME = NVLS(Rs("EXPENSE_TYPE_NAME"), "")
   ElseIf Ind = 6 Then
      TOTAL_PRICE = NVLD(Rs("TOTAL_PRICE"), 0)
      EXPENSE_TYPE = NVLI(Rs("EXPENSE_TYPE"), -1)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
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

   TempID1 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = LOCATION_ID
   If LOCATION_ID <= 0 Then
      TempID2 = "NULL"
   End If
      
   TempID3 = PIG_STATUS
   If PIG_STATUS <= 0 Then
      TempID3 = "NULL"
   End If

   TempID4 = EXPENSE_TYPE
   If EXPENSE_TYPE <= 0 Then
      TempID4 = "NULL"
   End If

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("RO_ITEM_SEQ", ID, glbErrorLog)
      Call glbDatabaseMngr.GetSeqID("TRANSACTION_SEQUENCE_SEQ", ID2, glbErrorLog)
      
      SQL1 = "INSERT INTO RO_ITEM " & _
                     "( " & _
                     "   RO_ITEM_ID, " & _
                     "   BILLING_DOC_ID, " & _
                     "   PART_ITEM_ID, " & _
                     "   LOCATION_ID, " & _
                     "   ITEM_AMOUNT, " & _
                     "   TOTAL_WEIGHT, " & _
                     "   AVG_WEIGHT, " & _
                     "   TOTAL_PRICE, " & _
                     "   AVG_PRICE, "
            SQL1 = SQL1 & _
                     "   GUI_ID, " & _
                     "   LINK_ID, " & _
                     "   PIG_STATUS, " & _
                     "   EXPENSE_DESC, " & _
                     "   EXPENSE_TYPE, "
            SQL1 = SQL1 & _
                     "   TRANSACTION_SEQ, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "" & BILLING_DOC_ID & "," & _
                         TempID1 & "," & _
                         TempID2 & "," & _
                         "" & ITEM_AMOUNT & "," & _
                         "" & TOTAL_WEIGHT & "," & _
                         "" & AVG_WEIGHT & "," & _
                         "" & TOTAL_PRICE & "," & _
                         "" & AVG_PRICE & "," & _
                         "" & GUI_ID & "," & _
                         "" & LINK_ID & "," & _
                         TempID3 & ","
            SQL1 = SQL1 & _
                        "'" & Trim(ChangeQuote(EXPENSE_DESC)) & "'," & _
                        TempID4 & "," & _
                        ID2 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         RO_ITEM_ID = ID
   Else 'Show edit
      Call glbDatabaseMngr.GetSeqID("TRANSACTION_SEQUENCE_SEQ", ID2, glbErrorLog)
      
      SQL1 = "UPDATE RO_ITEM SET " & _
                     "   BILLING_DOC_ID = " & BILLING_DOC_ID & ", " & _
                     "   PART_ITEM_ID = " & TempID1 & ", " & _
                     "   LOCATION_ID = " & TempID2 & ", " & _
                     "   ITEM_AMOUNT = " & ITEM_AMOUNT & ", " & _
                     "   TOTAL_WEIGHT = " & TOTAL_WEIGHT & ", " & _
                     "   AVG_WEIGHT = " & AVG_WEIGHT & ", " & _
                     "   TOTAL_PRICE = " & TOTAL_PRICE & ", " & _
                     "   AVG_PRICE = " & AVG_PRICE & ", " & _
                     "   GUI_ID = " & GUI_ID & ", " & _
                     "   LINK_ID = " & LINK_ID & ", " & _
                     "   PIG_STATUS = " & TempID3 & ", " & _
                     "   EXPENSE_DESC = '" & Trim(ChangeQuote(EXPENSE_DESC)) & "', " & _
                     "   EXPENSE_TYPE = " & TempID4 & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE RO_ITEM_ID = " & RO_ITEM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM EXPENSE_RATIO WHERE RO_ITEM_ID = " & RO_ITEM_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM RO_ITEM WHERE RO_ITEM_ID = " & RO_ITEM_ID
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
                  "WHERE BILLING_DOC_ID = " & BILLING_DOC_ID
                  
   m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set ExpenseRatios = Nothing
End Sub
