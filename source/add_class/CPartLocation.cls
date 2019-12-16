VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPartLocation"
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

Public PART_LOCATION_ID As Long
Public PART_ITEM_ID As Long
Public LOCATION_ID As Long
Public AVG_PRICE As Double
Public LAST_PRICE As Double
Public MINIMUM_ALLOW As Double
Public CURRENT_AMOUNT As Double

Public LOCATION_NAME As String
Public PART_NAME As String
Public PART_NO As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public COMMIT_FLAG As String
Public PIG_STATUS As Long

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
Dim SubLen As Long
Dim NewStr As String
Dim TempDate As String
Dim FromDate1 As String
Dim ToDate1 As String
Dim FromDate2 As String
Dim ToDate2 As String

   If Ind = 1 Then
      SelectStr1 = "SELECT LT.*, LC.*  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PART_LOCATION LT "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (LT.LOCATION_ID = LC.LOCATION_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      FromDate1 = DateToStringIntLow(-1)
      ToDate1 = DateToStringIntLow(FROM_DATE)
      
      FromDate2 = DateToStringIntLow(FROM_DATE)
      If TO_DATE <= 0 Then
         ToDate2 = "9999-99-99 23:59:59"
      Else
         ToDate2 = DateToStringIntHi(TO_DATE)
      End If
      
      SelectStr1 = "SELECT PTT.*, PL.*, PI.*, LC.LOCATION_NO, LC.LOCATION_NAME, UN.UNIT_NAME, "
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(II.IMPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM IMPORT_ITEM II, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (II.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (II.PART_ITEM_ID = PL.PART_ITEM_ID) "
      If LOCATION_ID > 0 Then
         SelectStr1 = SelectStr1 & "   AND (II.LOCATION_ID = PL.LOCATION_ID) "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate1 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate1 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & ") SUM_ALL_IMPORT, "
      
      '===
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(II.TOTAL_INCLUDE_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM IMPORT_ITEM II, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (II.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (II.PART_ITEM_ID = PL.PART_ITEM_ID) "
      If LOCATION_ID > 0 Then
         SelectStr1 = SelectStr1 & "   AND (II.LOCATION_ID = PL.LOCATION_ID) "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate1 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate1 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & ") SUM_ALL_IMPORT_AMT, "
      '===

      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PL.PART_ITEM_ID) "
      If LOCATION_ID > 0 Then
         SelectStr1 = SelectStr1 & "   AND (EI.LOCATION_ID = PL.LOCATION_ID) "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate1 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate1 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & ") SUM_ALL_EXPORT, "
      
      '===
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT * EI.EXPORT_AVG_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PL.PART_ITEM_ID) "
      If LOCATION_ID > 0 Then
         SelectStr1 = SelectStr1 & "   AND (EI.LOCATION_ID = PL.LOCATION_ID) "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate1 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate1 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & ") SUM_ALL_EXPORT_AMT, "
      '===
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(II.IMPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM IMPORT_ITEM II, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (II.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (II.PART_ITEM_ID = PL.PART_ITEM_ID) "
      If LOCATION_ID > 0 Then
         SelectStr1 = SelectStr1 & "   AND (II.LOCATION_ID = PL.LOCATION_ID) "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 1) " 'ใบนำเข้า
      SelectStr1 = SelectStr1 & ") SUM_IMPORT1, "
      
      '===
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(II.TOTAL_INCLUDE_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM IMPORT_ITEM II, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (II.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (II.PART_ITEM_ID = PL.PART_ITEM_ID) "
      If LOCATION_ID > 0 Then
         SelectStr1 = SelectStr1 & "   AND (II.LOCATION_ID = PL.LOCATION_ID) "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 1) " 'ใบนำเข้า
      SelectStr1 = SelectStr1 & ") SUM_IMPORT1_AMT, "
      '===
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PL.PART_ITEM_ID) "
      If LOCATION_ID > 0 Then
         SelectStr1 = SelectStr1 & "   AND (EI.LOCATION_ID = PL.LOCATION_ID) "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate1 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate1 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & ") SUM_ALL_EXPORT, "
      
      '===
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT * EI.EXPORT_AVG_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PL.PART_ITEM_ID) "
      If LOCATION_ID > 0 Then
         SelectStr1 = SelectStr1 & "   AND (EI.LOCATION_ID = PL.LOCATION_ID) "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate1 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate1 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & ") SUM_ALL_EXPORT_AMT, "
      '===
      
'      SelectStr1 = SelectStr1 & "( "
'      SelectStr1 = SelectStr1 & "   SELECT SUM(II.IMPORT_AMOUNT) "
'      SelectStr1 = SelectStr1 & "   FROM IMPORT_ITEM II, INVENTORY_DOC IVD "
'      SelectStr1 = SelectStr1 & "   WHERE (II.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID) "
'      SelectStr1 = SelectStr1 & "   AND (II.PART_ITEM_ID = PL.PART_ITEM_ID) "
'      If LOCATION_ID > 0 Then
'         SelectStr1 = SelectStr1 & "   AND (II.LOCATION_ID = PL.LOCATION_ID) "
'      End If
'      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
'      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
'      If Len(COMMIT_FLAG) > 0 Then
'         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
'      End If
'      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 1) " 'ใบนำเข้า
'      SelectStr1 = SelectStr1 & ") SUM_IMPORT1, "
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(II.IMPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM IMPORT_ITEM II, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (II.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (II.PART_ITEM_ID = PL.PART_ITEM_ID) "
      If LOCATION_ID > 0 Then
         SelectStr1 = SelectStr1 & "   AND (II.LOCATION_ID = PL.LOCATION_ID) "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 3) " 'โอนเข้า
      SelectStr1 = SelectStr1 & ") SUM_IMPORT2, "
      
      '===
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(II.TOTAL_INCLUDE_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM IMPORT_ITEM II, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (II.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (II.PART_ITEM_ID = PL.PART_ITEM_ID) "
      If LOCATION_ID > 0 Then
         SelectStr1 = SelectStr1 & "   AND (II.LOCATION_ID = PL.LOCATION_ID) "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 3) " 'โอนเข้า
      SelectStr1 = SelectStr1 & ") SUM_IMPORT2_AMT, "
      '===
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(II.IMPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM IMPORT_ITEM II, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (II.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (II.PART_ITEM_ID = PL.PART_ITEM_ID) "
      If LOCATION_ID > 0 Then
         SelectStr1 = SelectStr1 & "   AND (II.LOCATION_ID = PL.LOCATION_ID) "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 4) " 'ปรับเพิ่ม
      SelectStr1 = SelectStr1 & ") SUM_IMPORT3, "
      
      '===
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(II.TOTAL_INCLUDE_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM IMPORT_ITEM II, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (II.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (II.PART_ITEM_ID = PL.PART_ITEM_ID) "
      If LOCATION_ID > 0 Then
         SelectStr1 = SelectStr1 & "   AND (II.LOCATION_ID = PL.LOCATION_ID) "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 4) " 'ปรับเพิ่ม
      SelectStr1 = SelectStr1 & ") SUM_IMPORT3_AMT, "
      '===
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PL.PART_ITEM_ID) "
      If LOCATION_ID > 0 Then
         SelectStr1 = SelectStr1 & "   AND (EI.LOCATION_ID = PL.LOCATION_ID) "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.SALE_FLAG = 'N') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 2) " 'เบิกใช้
      SelectStr1 = SelectStr1 & ") SUM_EXPORT1_1, "
      
      '===
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT * EI.EXPORT_AVG_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PL.PART_ITEM_ID) "
      If LOCATION_ID > 0 Then
         SelectStr1 = SelectStr1 & "   AND (EI.LOCATION_ID = PL.LOCATION_ID) "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.SALE_FLAG = 'N') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 2) " 'เบิกใช้
      SelectStr1 = SelectStr1 & ") SUM_EXPORT1_1_AMT, "
      '===
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PL.PART_ITEM_ID) "
      If LOCATION_ID > 0 Then
         SelectStr1 = SelectStr1 & "   AND (EI.LOCATION_ID = PL.LOCATION_ID) "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 3) " 'โอนออก
      SelectStr1 = SelectStr1 & ") SUM_EXPORT2, "
      
      '===
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT * EI.EXPORT_AVG_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PL.PART_ITEM_ID) "
      If LOCATION_ID > 0 Then
         SelectStr1 = SelectStr1 & "   AND (EI.LOCATION_ID = PL.LOCATION_ID) "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 3) " 'โอนออก
      SelectStr1 = SelectStr1 & ") SUM_EXPORT2_AMT, "
      '===
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PL.PART_ITEM_ID) "
      If LOCATION_ID > 0 Then
         SelectStr1 = SelectStr1 & "   AND (EI.LOCATION_ID = PL.LOCATION_ID) "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.SALE_FLAG = 'N') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 2) " 'เบิกใช้
      SelectStr1 = SelectStr1 & ") SUM_EXPORT1_1, "
      
      '===
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT * EI.EXPORT_AVG_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PL.PART_ITEM_ID) "
      If LOCATION_ID > 0 Then
         SelectStr1 = SelectStr1 & "   AND (EI.LOCATION_ID = PL.LOCATION_ID) "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.SALE_FLAG = 'N') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 2) " 'เบิกใช้
      SelectStr1 = SelectStr1 & ") SUM_EXPORT1_1_AMT, "
      '===
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PL.PART_ITEM_ID) "
      If LOCATION_ID > 0 Then
         SelectStr1 = SelectStr1 & "   AND (EI.LOCATION_ID = PL.LOCATION_ID) "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.SALE_FLAG = 'Y') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 2) " 'เบิกใช้
      SelectStr1 = SelectStr1 & ") SUM_EXPORT1_2, "
      
      '===
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT * EI.EXPORT_AVG_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PL.PART_ITEM_ID) "
      If LOCATION_ID > 0 Then
         SelectStr1 = SelectStr1 & "   AND (EI.LOCATION_ID = PL.LOCATION_ID) "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.SALE_FLAG = 'Y') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 2) " 'เบิกใช้
      SelectStr1 = SelectStr1 & ") SUM_EXPORT1_2_AMT, "
      '===
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PL.PART_ITEM_ID) "
      If LOCATION_ID > 0 Then
         SelectStr1 = SelectStr1 & "   AND (EI.LOCATION_ID = PL.LOCATION_ID) "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 3) " 'โอนออก
      SelectStr1 = SelectStr1 & ") SUM_EXPORT2, "
      
      '===
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT * EI.EXPORT_AVG_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PL.PART_ITEM_ID) "
      If LOCATION_ID > 0 Then
         SelectStr1 = SelectStr1 & "   AND (EI.LOCATION_ID = PL.LOCATION_ID) "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 3) " 'โอนออก
      SelectStr1 = SelectStr1 & ") SUM_EXPORT2_AMT, "
      '===
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PL.PART_ITEM_ID) "
      If LOCATION_ID > 0 Then
         SelectStr1 = SelectStr1 & "   AND (EI.LOCATION_ID = PL.LOCATION_ID) "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 4) " 'ปรับลด
      SelectStr1 = SelectStr1 & ") SUM_EXPORT3, "
      
      '===
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT * EI.EXPORT_AVG_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PL.PART_ITEM_ID) "
      If LOCATION_ID > 0 Then
         SelectStr1 = SelectStr1 & "   AND (EI.LOCATION_ID = PL.LOCATION_ID) "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 4) " 'ปรับลด
      SelectStr1 = SelectStr1 & ") SUM_EXPORT3_AMT "
      '===
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PART_LOCATION PL "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (PL.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PL.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PTT ON (PI.PART_TYPE = PTT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      
      WhereStr = "WHERE PI.PIG_FLAG = 'N' "
   ElseIf Ind = 3 Then
      FromDate1 = DateToStringIntLow(-1)
      ToDate1 = DateToStringIntLow(FROM_DATE)
      
      FromDate2 = DateToStringIntLow(FROM_DATE)
      If TO_DATE <= 0 Then
         ToDate2 = "9999-99-99 23:59:59"
      Else
         ToDate2 = DateToStringIntHi(TO_DATE)
      End If
      
      SelectStr1 = "SELECT PI.*, PTT.*, UN.UNIT_NAME, "
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(II.IMPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM IMPORT_ITEM II, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (II.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (II.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate1 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate1 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & ") SUM_ALL_IMPORT, "
      
      '===
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(II.TOTAL_INCLUDE_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM IMPORT_ITEM II, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (II.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (II.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate1 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate1 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & ") SUM_ALL_IMPORT_AMT, "
      '===
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate1 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate1 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & ") SUM_ALL_EXPORT, "
      
      '===
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT * EI.EXPORT_AVG_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate1 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate1 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & ") SUM_ALL_EXPORT_AMT, "
      '===
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(II.IMPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM IMPORT_ITEM II, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (II.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (II.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 1) " 'ใบนำเข้า
      SelectStr1 = SelectStr1 & ") SUM_IMPORT1, "

      '===
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(II.TOTAL_INCLUDE_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM IMPORT_ITEM II, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (II.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (II.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 1) " 'ใบนำเข้า
      SelectStr1 = SelectStr1 & ") SUM_IMPORT1_AMT, "
      '===
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(II.IMPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM IMPORT_ITEM II, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (II.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (II.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 3) " 'โอนเข้า
      SelectStr1 = SelectStr1 & ") SUM_IMPORT2, "
      
      '===
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(II.TOTAL_INCLUDE_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM IMPORT_ITEM II, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (II.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (II.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 3) " 'โอนเข้า
      SelectStr1 = SelectStr1 & ") SUM_IMPORT2_AMT, "
      '===
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(II.IMPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM IMPORT_ITEM II, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (II.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (II.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 4) " 'ปรับเพิ่ม
      SelectStr1 = SelectStr1 & ") SUM_IMPORT3, "
      
      '===
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(II.TOTAL_INCLUDE_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM IMPORT_ITEM II, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (II.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (II.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 4) " 'ปรับเพิ่ม
      SelectStr1 = SelectStr1 & ") SUM_IMPORT3_AMT, "
      '===
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.SALE_FLAG = 'N') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 2) " 'เบิกใช้
      SelectStr1 = SelectStr1 & ") SUM_EXPORT1_1, "
      
      '===
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT * EI.EXPORT_AVG_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.SALE_FLAG = 'N') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 2) " 'เบิกใช้
      SelectStr1 = SelectStr1 & ") SUM_EXPORT1_1_AMT, "
      '===
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.SALE_FLAG = 'Y') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 2) " 'เบิกใช้ ขาย
      SelectStr1 = SelectStr1 & ") SUM_EXPORT1_2, "
      
      '===
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT * EI.EXPORT_AVG_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.SALE_FLAG = 'Y') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 2) " 'เบิกใช้ ขาย
      SelectStr1 = SelectStr1 & ") SUM_EXPORT1_2_AMT, "
      '===
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 3) " 'โอนออก
      SelectStr1 = SelectStr1 & ") SUM_EXPORT2, "
      
      '===
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT * EI.EXPORT_AVG_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 3) " 'โอนออก
      SelectStr1 = SelectStr1 & ") SUM_EXPORT2_AMT, "
      '===
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 4) " 'ปรับลด
      SelectStr1 = SelectStr1 & ") SUM_EXPORT3, "
      
      '===
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT * EI.EXPORT_AVG_PRICE) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 4) " 'ปรับลด
      SelectStr1 = SelectStr1 & ") SUM_EXPORT3_AMT "
      '===
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PART_ITEM PI "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PTT ON (PI.PART_TYPE = PTT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      
      WhereStr = "WHERE (PI.PIG_FLAG = 'N') "
      WhereStr = WhereStr & "AND (PI.PART_ITEM_ID IN (SELECT DISTINCT(PART_ITEM_ID) FROM PART_LOCATION)) "
   ElseIf Ind = 4 Then
      FromDate1 = DateToStringIntLow(-1)
      ToDate1 = DateToStringIntLow(FROM_DATE)
      
      FromDate2 = DateToStringIntLow(FROM_DATE)
      If TO_DATE <= 0 Then
         ToDate2 = "9999-99-99 23:59:59"
      Else
         ToDate2 = DateToStringIntHi(TO_DATE)
      End If
      
      SelectStr1 = "SELECT PTT.*, "
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(II.IMPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM IMPORT_ITEM II, INVENTORY_DOC IVD, PART_ITEM PI "
      SelectStr1 = SelectStr1 & "   WHERE (II.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (II.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (PI.PART_TYPE = PTT.PART_TYPE_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate1 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate1 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & ") SUM_ALL_IMPORT, "
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD, PART_ITEM PI "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (PI.PART_TYPE = PTT.PART_TYPE_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate1 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate1 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & ") SUM_ALL_EXPORT, "
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(II.IMPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM IMPORT_ITEM II, INVENTORY_DOC IVD, PART_ITEM PI "
      SelectStr1 = SelectStr1 & "   WHERE (II.INVENTORY_DOC_ID = Ivd.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (II.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (PI.PART_TYPE = PTT.PART_TYPE_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 1) " 'ใบนำเข้า
      SelectStr1 = SelectStr1 & ") SUM_IMPORT1, "
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(II.IMPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM IMPORT_ITEM II, INVENTORY_DOC IVD, PART_ITEM PI "
      SelectStr1 = SelectStr1 & "   WHERE (II.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (II.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (PI.PART_TYPE = PTT.PART_TYPE_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 3) " 'โอนเข้า
      SelectStr1 = SelectStr1 & ") SUM_IMPORT2, "
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(II.IMPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM IMPORT_ITEM II, INVENTORY_DOC IVD, PART_ITEM PI "
      SelectStr1 = SelectStr1 & "   WHERE (II.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (II.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (PI.PART_TYPE = PTT.PART_TYPE_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 4) " 'ปรับเพิ่ม
      SelectStr1 = SelectStr1 & ") SUM_IMPORT3, "
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD, PART_ITEM PI "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (PI.PART_TYPE = PTT.PART_TYPE_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.SALE_FLAG = 'N') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 2) " 'เบิกใช้
      SelectStr1 = SelectStr1 & ") SUM_EXPORT1_1, "
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD, PART_ITEM PI "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (PI.PART_TYPE = PTT.PART_TYPE_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.SALE_FLAG = 'Y') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 2) " 'เบิกใช้ ขาย
      SelectStr1 = SelectStr1 & ") SUM_EXPORT1_2, "
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD, PART_ITEM PI "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (PI.PART_TYPE = PTT.PART_TYPE_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 3) " 'โอนออก
      SelectStr1 = SelectStr1 & ") SUM_EXPORT2, "
      
      SelectStr1 = SelectStr1 & "( "
      SelectStr1 = SelectStr1 & "   SELECT SUM(EI.EXPORT_AMOUNT) "
      SelectStr1 = SelectStr1 & "   FROM EXPORT_ITEM EI, INVENTORY_DOC IVD, PART_ITEM PI "
      SelectStr1 = SelectStr1 & "   WHERE (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) "
      SelectStr1 = SelectStr1 & "   AND (EI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      SelectStr1 = SelectStr1 & "   AND (PI.PART_TYPE = PTT.PART_TYPE_ID) "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE >= '" & FromDate2 & "') "
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_DATE <= '" & ToDate2 & "') "
      If Len(COMMIT_FLAG) > 0 Then
         SelectStr1 = SelectStr1 & "   AND (IVD.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
      SelectStr1 = SelectStr1 & "   AND (IVD.DOCUMENT_TYPE = 4) " 'ปรับลด
      SelectStr1 = SelectStr1 & ") SUM_EXPORT3 "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PART_TYPE PTT "
      
      WhereStr = ""
   ElseIf Ind = 5 Then
      SelectStr1 = "SELECT DISTINCT PL.PART_ITEM_ID "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PART_LOCATION PL "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PL.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      
      WhereStr = "WHERE (PI.PIG_FLAG = 'N') "
   End If
   
   If PART_LOCATION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_LOCATION_ID = " & PART_LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PART_LOCATION_ID = " & PART_LOCATION_ID & ") "
      End If
   End If
   
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   If LOCATION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LC.LOCATION_ID = " & LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LC.LOCATION_ID = " & LOCATION_ID & ") "
      End If
   End If
   
   If Trim(PART_NO) <> "" Then
      If WildCard(PART_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')"
         End If
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PART_LOCATION_ID "
      Else
         OrderByStr = "ORDER BY PART_LOCATION_ID "
      End If
   ElseIf (Ind = 2) Or (Ind = 3) Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PART_TYPE_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PART_TYPE_NAME "
      Else
         OrderByStr = "ORDER BY PART_TYPE_NO "
      End If
   ElseIf (Ind = 4) Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PART_TYPE_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PART_TYPE_NAME "
      Else
         OrderByStr = "ORDER BY PART_TYPE_ID "
      End If
   ElseIf (Ind = 5) Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PT.PART_TYPE_NO "
      Else
         OrderByStr = "ORDER BY PT.PART_TYPE_NO "
      End If
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

Public Function PopulateFromRS(Rs As ADODB.Recordset)
   PART_LOCATION_ID = NVLI(Rs("PART_LOCATION_ID"), -1)
   PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
   AVG_PRICE = NVLD(Rs("AVG_PRICE"), -1)
   LOCATION_ID = NVLD(Rs("LOCATION_ID"), -1)
   LAST_PRICE = NVLD(Rs("LAST_PRICE"), -1)
   MINIMUM_ALLOW = NVLD(Rs("MINIMUM_ALLOW"), -1)
   CURRENT_AMOUNT = NVLD(Rs("CURRENT_AMOUNT"), -1)
   
   LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String

   TempID1 = LOCATION_ID
   If LOCATION_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = LAST_PRICE
   If LAST_PRICE <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID3 = MINIMUM_ALLOW
   If MINIMUM_ALLOW <= 0 Then
      TempID3 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("PART_LOCATION_SEQ", ID, glbErrorLog)
      SQL1 = "INSERT INTO PART_LOCATION " & _
                     "( " & _
                     "   PART_LOCATION_ID, " & _
                     "   AVG_PRICE, " & _
                     "   PART_ITEM_ID, " & _
                     "   LOCATION_ID, " & _
                     "   LAST_PRICE, " & _
                     "   MINIMUM_ALLOW, " & _
                     "   CURRENT_AMOUNT, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
               SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "" & AVG_PRICE & "," & _
                         "'" & ChangeQuote(Trim(PART_ITEM_ID)) & "'," & _
                         "" & TempID1 & "," & _
                         "" & TempID2 & "," & _
                         "" & TempID3 & "," & _
                         CURRENT_AMOUNT & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
               PART_LOCATION_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE PART_LOCATION SET " & _
                     "   AVG_PRICE = " & AVG_PRICE & ", " & _
                     "   PART_ITEM_ID = '" & ChangeQuote(Trim(PART_ITEM_ID)) & "', " & _
                     "   LOCATION_ID = " & TempID1 & ", " & _
                     "   LAST_PRICE = " & TempID2 & ", " & _
                     "   MINIMUM_ALLOW = " & TempID3 & ", " & _
                     "   CURRENT_AMOUNT = " & CURRENT_AMOUNT & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PART_LOCATION_ID = " & PART_LOCATION_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM PART_LOCATION WHERE PART_LOCATION_ID = " & PART_LOCATION_ID
      m_Conn.Execute (SQL1)
End Function

