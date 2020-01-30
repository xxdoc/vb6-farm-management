VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBalanceAccum"
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

Public BALANCE_ACCUM_ID As Long
Public DOCUMENT_DATE As Date
Public PART_ITEM_ID As Long
Public IMPORT_AMOUNT As Double
Public EXPORT_AMOUNT As Double
Public BALANCE_AMOUNT As Double
Public TOTAL_INCLUDE_PRICE As Double
Public LOCATION_ID As Long
Public AVG_PRICE As Double
Public SIMULATE_FLAG As String
Public BATCH_ID As Long
Public YYYYMM As String
Public PIG_AGE As Long

Public HOUSE_GROUP_ID As Long
Public PIG_TYPE As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public TO_DATE1 As Date
Public MASTER_FLAG As String
Public PIG_FLAG As String
Public LOCATION_TYPE As Long
Public PART_TYPE As Long
Public FEED_GROUP As Long

Public PART_NO As String
Public PART_DESC As String
Public PART_TYPE_NO  As String
Public PART_TYPE_NAME As String
Public UNIT_NAME As String
Public LOCATION_NAME As String
 
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_DATE = -1
   TO_DATE = -1
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
Dim TempDate As String
Dim GroupByStr As String
Dim MustCount As Boolean
      
   SIMULATE_FLAG = glbUser.SIMULATE_FLAG
   MustCount = True
   
   If Ind = 1 Then
      SelectStr1 = "SELECT BA1.*, PI.PART_NO  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM BA1 "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (BA1.PART_ITEM_ID = PI.PART_ITEM_ID) "

      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT SUM(IMPORT_AMOUNT - EXPORT_AMOUNT) BALANCE_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM BA1 "
      
      WhereStr = ""
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT PART_ITEM_ID, SUM(IMPORT_AMOUNT-EXPORT_AMOUNT) BALANCE_AMOUNT, SUM(TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM BA1 "
      
      WhereStr = ""
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT BA1.PART_ITEM_ID, BA1.LOCATION_ID, SUM(IMPORT_AMOUNT - EXPORT_AMOUNT) BALANCE_AMOUNT  "
      'SelectStr1 = SelectStr1 & ",LC.LOCATION_NAME,PI.PART_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM BA1 "
      'FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (BA1.LOCATION_ID = LC.LOCATION_ID) "
      'FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (BA1.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      
      WhereStr = ""
   ElseIf Ind = 5 Then        'ยกเลิกใช้
      SelectStr1 = "SELECT BA1.PART_ITEM_ID, SUM(BA1.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE, SUM(BALANCE_AMOUNT) BALANCE_AMOUNT, AVG(AVG_PRICE) AVG_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM BA1 "
      
     WhereStr = "WHERE BALANCE_ACCUM_ID IN "
     WhereStr = WhereStr & "( "
     WhereStr = WhereStr & "   SELECT MAX(BALANCE_ACCUM_ID) FROM BALANCE_ACCUM BA2 "
     WhereStr = WhereStr & "   WHERE BA2.DOCUMENT_DATE <= '" & DateToStringIntHi(TO_DATE1) & "' "
     WhereStr = WhereStr & "   AND (BA2.PART_ITEM_ID = BA1.PART_ITEM_ID) "
     WhereStr = WhereStr & "   AND (BA2.LOCATION_ID = BA1.LOCATION_ID) "
     WhereStr = WhereStr & "   AND (BA2.BALANCE_AMOUNT >= 0) "
     WhereStr = WhereStr & ") "
   
   ElseIf Ind = 6 Then        'ยกเลิกใช้
      SelectStr1 = "SELECT BA1.BALANCE_ACCUM_ID, LC.MASTER_FLAG, BA1.DOCUMENT_DATE, BA1.LOCATION_ID, BA1.PART_ITEM_ID, BA1.BALANCE_AMOUNT, BA1.TOTAL_INCLUDE_PRICE, BA1.AVG_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM BA1 "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (BA1.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (BA1.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
     WhereStr = "WHERE BA1.DOCUMENT_DATE IN "
     WhereStr = WhereStr & "( "
     WhereStr = WhereStr & "   SELECT MAX (BA2.DOCUMENT_DATE) FROM BALANCE_ACCUM BA2 "
     WhereStr = WhereStr & "   WHERE BA2.DOCUMENT_DATE <= '" & DateToStringIntHi(TO_DATE1) & "' "
     WhereStr = WhereStr & "   AND (BA2.PART_ITEM_ID = BA1.PART_ITEM_ID) "
     WhereStr = WhereStr & "   AND (BA2.LOCATION_ID = BA1.LOCATION_ID) "
     WhereStr = WhereStr & ") "
   ElseIf Ind = 7 Then
      SelectStr1 = "SELECT BA1.DOCUMENT_DATE, BA1.LOCATION_ID, BA1.PART_ITEM_ID, BA1.AVG_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM BA1 "

     WhereStr = "WHERE BA1.DOCUMENT_DATE IN "
     WhereStr = WhereStr & "( "
     WhereStr = WhereStr & "   SELECT MAX (BA2.DOCUMENT_DATE) FROM BALANCE_ACCUM BA2 "
     WhereStr = WhereStr & "   WHERE BA2.DOCUMENT_DATE <= '" & DateToStringIntHi(TO_DATE1) & "' "
     WhereStr = WhereStr & "   AND (BA2.PART_ITEM_ID = BA1.PART_ITEM_ID) "
     WhereStr = WhereStr & "   AND (BA2.LOCATION_ID = BA1.LOCATION_ID) "
     WhereStr = WhereStr & ") "
   ElseIf Ind = 8 Then
      SelectStr1 = "SELECT DISTINCT BA1.LOCATION_ID, BA1.PART_ITEM_ID "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM BALANCE_ACCUM BA1 "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (BA1.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (BA1.LOCATION_ID = LC.LOCATION_ID) "

     WhereStr = ""
   ElseIf Ind = 9 Then
      SelectStr1 = "SELECT BA1.YYYYMM, BA1.LOCATION_ID, SUM(BALANCE_AMOUNT) BALANCE_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM BALANCE_ACCUM BA1 "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (BA1.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (BA1.LOCATION_ID = LC.LOCATION_ID) "

     WhereStr = ""
   ElseIf Ind = 10 Then
      SelectStr1 = "SELECT BA1.YYYYMM, BA1.PART_ITEM_ID, SUM(BALANCE_AMOUNT) BALANCE_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM BALANCE_ACCUM BA1 "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (BA1.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (BA1.LOCATION_ID = LC.LOCATION_ID) "

     WhereStr = ""
   ElseIf Ind = 11 Then
      SelectStr1 = "SELECT SUM(BA1.BALANCE_AMOUNT) BALANCE_AMOUNT  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM BA1 "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (BA1.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (BA1.LOCATION_ID = LC.LOCATION_ID) "
      
      WhereStr = ""
   ElseIf Ind = 12 Then
      SelectStr1 = "SELECT BA1.PART_ITEM_ID, SUM(BALANCE_AMOUNT) BALANCE_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM BALANCE_ACCUM BA1 "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (BA1.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (BA1.LOCATION_ID = LC.LOCATION_ID) "
      
     WhereStr = ""
   ElseIf Ind = 13 Then          'ยกเลิกใช้
      SelectStr1 = "SELECT  BA1.PART_ITEM_ID, SUM(BA1.BALANCE_AMOUNT) BALANCE_AMOUNT, SUM(BA1.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM BA1 "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (BA1.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (BA1.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
     WhereStr = "WHERE BA1.DOCUMENT_DATE IN "
     WhereStr = WhereStr & "( "
     WhereStr = WhereStr & "   SELECT MAX (BA2.DOCUMENT_DATE) FROM BALANCE_ACCUM BA2 "
     WhereStr = WhereStr & "   WHERE BA2.DOCUMENT_DATE <= '" & DateToStringIntHi(TO_DATE1) & "' "
     WhereStr = WhereStr & "   AND (BA2.PART_ITEM_ID = BA1.PART_ITEM_ID) "
     WhereStr = WhereStr & "   AND (BA2.LOCATION_ID = BA1.LOCATION_ID) "
     WhereStr = WhereStr & ") "
   ElseIf Ind = 14 Then          'ยกเลิกใช้
      SelectStr1 = "SELECT  BA1.LOCATION_ID,BA1.PART_ITEM_ID, SUM(BA1.BALANCE_AMOUNT) BALANCE_AMOUNT, SUM(BA1.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM BA1 "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (BA1.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (BA1.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
     WhereStr = "WHERE BA1.DOCUMENT_DATE IN "
     WhereStr = WhereStr & "( "
     WhereStr = WhereStr & "   SELECT MAX (BA2.DOCUMENT_DATE) FROM BALANCE_ACCUM BA2 "
     WhereStr = WhereStr & "   WHERE BA2.DOCUMENT_DATE <= '" & DateToStringIntHi(TO_DATE1) & "' "
     WhereStr = WhereStr & "   AND (BA2.PART_ITEM_ID = BA1.PART_ITEM_ID) "
     WhereStr = WhereStr & "   AND (BA2.LOCATION_ID = BA1.LOCATION_ID) "
     WhereStr = WhereStr & ") "
     
     'WhereStr = WhereStr & "AND (BA1.BALANCE_AMOUNT   <> 0  ) "
     MustCount = False
   ElseIf Ind = 15 Then
      SelectStr1 = "SELECT DISTINCT BA1.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, PI.PIG_TYPE,PI.PART_TYPE, PT.PART_TYPE_NO, PT.PART_TYPE_NAME "
      SelectStr1 = SelectStr1 & ",UN.UNIT_NAME "
      SelectStr2 = "SELECT COUNT(DISTINCT(BA1.PART_ITEM_ID))  "
      
      FromStr = "FROM BALANCE_ACCUM BA1 "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (BA1.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
      
      WhereStr = ""
      MustCount = False
   ElseIf Ind = 16 Then
      SelectStr1 = "SELECT BA1.PART_ITEM_ID, BA1.DOCUMENT_DATE, BA1.PIG_AGE, SUM(BA1.BALANCE_AMOUNT) BALANCE_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM BA1 "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (BA1.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = ""
   ElseIf Ind = 17 Then 'ยกเลิกใช้
      SelectStr1 = "SELECT PI.FEED_GROUP, SUM(BA1.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE, SUM(BALANCE_AMOUNT) BALANCE_AMOUNT, AVG(AVG_PRICE) AVG_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM BA1 "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (BA1.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
     WhereStr = "WHERE BALANCE_ACCUM_ID IN "
     WhereStr = WhereStr & "( "
     WhereStr = WhereStr & "   SELECT MAX(BALANCE_ACCUM_ID) FROM BALANCE_ACCUM BA2 "
     WhereStr = WhereStr & "   WHERE BA2.DOCUMENT_DATE <= '" & DateToStringIntHi(TO_DATE1) & "' "
     WhereStr = WhereStr & "   AND (BA2.PART_ITEM_ID = BA1.PART_ITEM_ID) "
     WhereStr = WhereStr & "   AND (BA2.LOCATION_ID = BA1.LOCATION_ID) "
     WhereStr = WhereStr & "   AND (BA2.BALANCE_AMOUNT >= 0) "
     WhereStr = WhereStr & ") "
   ElseIf Ind = 18 Then
      SelectStr1 = "SELECT BA1.YYYYMM,PI.PART_NO,PI.PIG_TYPE, BA1.BALANCE_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "

      FromStr = "FROM BALANCE_ACCUM BA1 "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (BA1.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = "WHERE BA1.DOCUMENT_DATE IN "
      WhereStr = WhereStr & "( "
      WhereStr = WhereStr & "   SELECT MAX (BA2.DOCUMENT_DATE) FROM BALANCE_ACCUM BA2 "
      WhereStr = WhereStr & "   WHERE BA2.DOCUMENT_DATE <= '" & DateToStringIntHi(TO_DATE1) & "' "
      WhereStr = WhereStr & "   AND (BA2.PART_ITEM_ID = BA1.PART_ITEM_ID) "
      WhereStr = WhereStr & ") "
   ElseIf Ind = 20 Then 'ปรับเป็น TempTable Inner Join แทน Ind = 5
       SelectStr1 = "SELECT BA1.PART_ITEM_ID, SUM(BA1.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE, SUM(BA1.BALANCE_AMOUNT) BALANCE_AMOUNT, AVG(BA1.AVG_PRICE) AVG_PRICE "
       SelectStr2 = "SELECT COUNT(*)  "
       
       FromStr = "FROM BALANCE_ACCUM BA1 "
       FromStr = FromStr & "INNER JOIN "
       FromStr = FromStr & " ( "
       FromStr = FromStr & " SELECT BA2.PART_ITEM_ID,BA2.LOCATION_ID,MAX(BA2.BALANCE_ACCUM_ID) BALANCE_ACCUM_ID FROM BALANCE_ACCUM BA2 "
       FromStr = FromStr & "   WHERE BA2.DOCUMENT_DATE <= '" & DateToStringIntHi(TO_DATE1) & "' "
       FromStr = FromStr & " GROUP BY BA2.PART_ITEM_ID,BA2.LOCATION_ID "
       FromStr = FromStr & " ) "
       FromStr = FromStr & " AS P ON (BA1.PART_ITEM_ID = P.PART_ITEM_ID AND BA1.LOCATION_ID = P.LOCATION_ID AND BA1.BALANCE_ACCUM_ID = P.BALANCE_ACCUM_ID) "
   ElseIf Ind = 21 Then
      SelectStr1 = "SELECT BA1.BALANCE_ACCUM_ID, LC.MASTER_FLAG, BA1.DOCUMENT_DATE, BA1.LOCATION_ID, BA1.PART_ITEM_ID, BA1.BALANCE_AMOUNT, BA1.TOTAL_INCLUDE_PRICE, BA1.AVG_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM BA1 "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (BA1.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (BA1.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "INNER JOIN "
       FromStr = FromStr & " ( "
       FromStr = FromStr & " SELECT BA2.PART_ITEM_ID,BA2.LOCATION_ID,MAX(BA2.DOCUMENT_DATE) DOCUMENT_DATE FROM BALANCE_ACCUM BA2 "
       FromStr = FromStr & "   WHERE BA2.DOCUMENT_DATE <= '" & DateToStringIntHi(TO_DATE1) & "' "
       FromStr = FromStr & " GROUP BY BA2.PART_ITEM_ID,BA2.LOCATION_ID "
       FromStr = FromStr & " ) "
       FromStr = FromStr & " AS P ON (BA1.PART_ITEM_ID = P.PART_ITEM_ID AND BA1.LOCATION_ID = P.LOCATION_ID AND BA1.DOCUMENT_DATE = P.DOCUMENT_DATE) "
   ElseIf Ind = 22 Then
      SelectStr1 = "SELECT  BA1.PART_ITEM_ID, SUM(BA1.BALANCE_AMOUNT) BALANCE_AMOUNT, SUM(BA1.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM BA1 "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (BA1.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (BA1.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "INNER JOIN "
       FromStr = FromStr & " ( "
       FromStr = FromStr & " SELECT BA2.PART_ITEM_ID,BA2.LOCATION_ID,MAX(BA2.DOCUMENT_DATE) DOCUMENT_DATE FROM BALANCE_ACCUM BA2 "
       FromStr = FromStr & "   WHERE BA2.DOCUMENT_DATE <= '" & DateToStringIntHi(TO_DATE1) & "' "
       FromStr = FromStr & " GROUP BY BA2.PART_ITEM_ID,BA2.LOCATION_ID "
       FromStr = FromStr & " ) "
       FromStr = FromStr & " AS P ON (BA1.PART_ITEM_ID = P.PART_ITEM_ID AND BA1.LOCATION_ID = P.LOCATION_ID AND BA1.DOCUMENT_DATE = P.DOCUMENT_DATE) "
   ElseIf Ind = 23 Then
      SelectStr1 = "SELECT  BA1.LOCATION_ID,BA1.PART_ITEM_ID, SUM(BA1.BALANCE_AMOUNT) BALANCE_AMOUNT, SUM(BA1.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM BA1 "
      FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (BA1.LOCATION_ID = LC.LOCATION_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (BA1.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "INNER JOIN "
       FromStr = FromStr & " ( "
       FromStr = FromStr & " SELECT BA2.PART_ITEM_ID,BA2.LOCATION_ID,MAX(BA2.DOCUMENT_DATE) DOCUMENT_DATE FROM BALANCE_ACCUM BA2 "
       FromStr = FromStr & "   WHERE BA2.DOCUMENT_DATE <= '" & DateToStringIntHi(TO_DATE1) & "' "
       FromStr = FromStr & " GROUP BY BA2.PART_ITEM_ID,BA2.LOCATION_ID "
       FromStr = FromStr & " ) "
       FromStr = FromStr & " AS P ON (BA1.PART_ITEM_ID = P.PART_ITEM_ID AND BA1.LOCATION_ID = P.LOCATION_ID AND BA1.DOCUMENT_DATE = P.DOCUMENT_DATE) "
       
     MustCount = False
   ElseIf Ind = 24 Then
      SelectStr1 = "SELECT PI.FEED_GROUP, SUM(BA1.TOTAL_INCLUDE_PRICE) TOTAL_INCLUDE_PRICE, SUM(BA1.BALANCE_AMOUNT) BALANCE_AMOUNT, AVG(BA1.AVG_PRICE) AVG_PRICE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM BA1 "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (BA1.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "INNER JOIN "
       FromStr = FromStr & " ( "
       FromStr = FromStr & " SELECT BA2.PART_ITEM_ID,BA2.LOCATION_ID,MAX(BA2.DOCUMENT_DATE) DOCUMENT_DATE FROM BALANCE_ACCUM BA2 "
       FromStr = FromStr & "   WHERE BA2.DOCUMENT_DATE <= '" & DateToStringIntHi(TO_DATE1) & "' "
       FromStr = FromStr & " GROUP BY BA2.PART_ITEM_ID,BA2.LOCATION_ID "
       FromStr = FromStr & " ) "
       FromStr = FromStr & " AS P ON (BA1.PART_ITEM_ID = P.PART_ITEM_ID AND BA1.LOCATION_ID = P.LOCATION_ID AND BA1.DOCUMENT_DATE = P.DOCUMENT_DATE) "
   
   ElseIf Ind = 25 Then
      SelectStr1 = "SELECT BA1.PART_ITEM_ID,BA1.DOCUMENT_DATE, SUM(IMPORT_AMOUNT-EXPORT_AMOUNT) BALANCE_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM BA1 "
   ElseIf Ind = 26 Then
      SelectStr1 = "SELECT BA1.PART_ITEM_ID, SUM(IMPORT_AMOUNT-EXPORT_AMOUNT) BALANCE_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM BALANCE_ACCUM BA1 "
   End If
   
   If BALANCE_ACCUM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BALANCE_ACCUM_ID = " & BALANCE_ACCUM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BALANCE_ACCUM_ID = " & BALANCE_ACCUM_ID & ") "
      End If
   End If
      
   If PART_ITEM_ID > 0 Then
        If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (BA1.PART_ITEM_ID = " & PART_ITEM_ID & ") "
        Else
            WhereStr = WhereStr & "AND (BA1.PART_ITEM_ID = " & PART_ITEM_ID & ") "
        End If
    End If
   
   If LOCATION_ID > 0 Then
        If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (BA1.LOCATION_ID = " & LOCATION_ID & ") "
         Else
            WhereStr = WhereStr & "AND (BA1.LOCATION_ID = " & LOCATION_ID & ") "
         End If
    End If
   
   If HOUSE_GROUP_ID > 0 Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (BA1.LOCATION_ID IN (SELECT LOCATION_ID FROM HGROUP_ITEM WHERE (HOUSE_GROUP_ID = " & HOUSE_GROUP_ID & ") AND (SELECT_FLAG = 'Y')))"
         Else
            WhereStr = WhereStr & "AND (BA1.LOCATION_ID IN (SELECT LOCATION_ID FROM HGROUP_ITEM WHERE (HOUSE_GROUP_ID = " & HOUSE_GROUP_ID & ") AND (SELECT_FLAG = 'Y')))"
         End If
   End If
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BA1.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BA1.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BA1.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BA1.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
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
         WhereStr = "WHERE (PI.PIG_TYPE = '" & PIG_TYPE & "') "
      Else
         WhereStr = WhereStr & "AND (PI.PIG_TYPE = '" & PIG_TYPE & "') "
      End If
   End If
   
   If Ind = 14 Then
      If SIMULATE_FLAG = "Y" Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (BA1.SIMULATE_FLAG = '" & SIMULATE_FLAG & "') "
         Else
            WhereStr = WhereStr & "AND (BA1.SIMULATE_FLAG = '" & SIMULATE_FLAG & "') "
         End If
      Else
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE ((BA1.SIMULATE_FLAG <> 'Y') OR (BA1.SIMULATE_FLAG IS NULL)) "
         Else
            WhereStr = WhereStr & "AND ((BA1.SIMULATE_FLAG <> 'Y') OR (BA1.SIMULATE_FLAG IS NULL)) "
         End If
      End If
   Else
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
   End If
   
   If LOCATION_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LOCATION_TYPE = " & LOCATION_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (LOCATION_TYPE = " & LOCATION_TYPE & ") "
      End If
   End If
      
   If PART_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_TYPE = " & PART_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (PART_TYPE = " & PART_TYPE & ") "
      End If
   End If
   
   If PIG_AGE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PIG_AGE = " & PIG_AGE & ") "
      Else
         WhereStr = WhereStr & "AND (PIG_AGE = " & PIG_AGE & ") "
      End If
   End If
   
   If Ind = 13 Then
      If BATCH_ID > 0 Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (BA1.BATCH_ID = " & BATCH_ID & ") "
         Else
            WhereStr = WhereStr & "AND (BA1.BATCH_ID = " & BATCH_ID & ") "
         End If
      Else
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE ((BA1.BATCH_ID <= 0) OR (BA1.BATCH_ID IS NULL)) "
         Else
            WhereStr = WhereStr & "AND ((BA1.BATCH_ID <= 0) OR (BA1.BATCH_ID IS NULL)) "
         End If
      End If
   Else
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
   If Ind = 3 Then
      GroupByStr = "GROUP BY BA1.PART_ITEM_ID "
   ElseIf Ind = 4 Then
      'GroupByStr = "GROUP BY PART_ITEM_ID, LOCATION_ID,LC.LOCATION_NAME,PI.PART_DESC "
      GroupByStr = "GROUP BY PART_ITEM_ID, LOCATION_ID "
   ElseIf Ind = 5 Then
      GroupByStr = "GROUP BY PART_ITEM_ID "
   ElseIf Ind = 9 Then
      GroupByStr = "GROUP BY BA1.YYYYMM , BA1.LOCATION_ID "
   ElseIf Ind = 10 Then
      GroupByStr = "GROUP BY BA1.YYYYMM , BA1.PART_ITEM_ID "
   ElseIf Ind = 12 Then
      GroupByStr = "GROUP BY BA1.PART_ITEM_ID "
   ElseIf Ind = 13 Then
      GroupByStr = "GROUP BY BA1.PART_ITEM_ID "
   ElseIf Ind = 14 Then
      GroupByStr = "GROUP BY BA1.PART_ITEM_ID, BA1.LOCATION_ID "
   ElseIf Ind = 16 Then
      GroupByStr = "GROUP BY BA1.PART_ITEM_ID, BA1.DOCUMENT_DATE, BA1.PIG_AGE "
   ElseIf Ind = 17 Then
      GroupByStr = "GROUP BY PI.FEED_GROUP "
   ElseIf Ind = 20 Then
      GroupByStr = "GROUP BY BA1.PART_ITEM_ID "
   ElseIf Ind = 22 Then
      GroupByStr = "GROUP BY BA1.PART_ITEM_ID "
   ElseIf Ind = 23 Then
      GroupByStr = "GROUP BY BA1.PART_ITEM_ID, BA1.LOCATION_ID "
   ElseIf Ind = 24 Then
      GroupByStr = "GROUP BY PI.FEED_GROUP "
   ElseIf Ind = 25 Then
      GroupByStr = "GROUP BY BA1.PART_ITEM_ID, BA1.DOCUMENT_DATE "
   ElseIf Ind = 26 Then
      GroupByStr = "GROUP BY BA1.PART_ITEM_ID "
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY DOCUMENT_DATE "
      Else
         OrderByStr = "ORDER BY BALANCE_ACCUM_ID "
      End If
   ElseIf Ind = 8 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY LOCATION_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY LOCATION_NAME "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY PART_NO "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY PART_DESC "
      Else
         OrderByStr = "ORDER BY LOCATION_NO "
      End If
   ElseIf Ind = 9 Then
      OrderByStr = "ORDER BY BA1.YYYYMM "
   ElseIf Ind = 10 Then
      OrderByStr = "ORDER BY BA1.YYYYMM "
   ElseIf Ind = 15 Then
      OrderByStr = "ORDER BY PART_TYPE_NO ASC, PART_NO "
      OrderType = 1 ' จะต้องเป็น ASC เหมือนเนื่องจาก จะ ต้อง Match กัน รายงาน การ STOCK CARD ที่เรียง แบบ ASC เหมือนกัน
   ElseIf Ind = 16 Then
      OrderByStr = "ORDER BY BA1.PART_ITEM_ID ASC, BA1.PIG_AGE "
   ElseIf Ind = 17 Then
      OrderByStr = "ORDER BY PI.FEED_GROUP "
   ElseIf Ind = 24 Then
      OrderByStr = "ORDER BY PI.FEED_GROUP "
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, itemcount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      BALANCE_ACCUM_ID = NVLI(Rs("BALANCE_ACCUM_ID"), 0)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      IMPORT_AMOUNT = NVLD(Rs("IMPORT_AMOUNT"), 0)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
      SIMULATE_FLAG = NVLS(Rs("SIMULATE_FLAG"), "")
      BATCH_ID = NVLI(Rs("BATCH_ID"), 0)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      PIG_AGE = NVLI(Rs("PIG_AGE"), 0)
   
      PART_NO = NVLS(Rs("PART_NO"), "")
   ElseIf Ind = 2 Then
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
   ElseIf Ind = 3 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
   ElseIf Ind = 4 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      'PART_DESC = NVLS(Rs("PART_DESC"), "")
      'LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
      
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
   ElseIf Ind = 5 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
   ElseIf Ind = 6 Then
      BALANCE_ACCUM_ID = NVLI(Rs("BALANCE_ACCUM_ID"), 0)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
      MASTER_FLAG = NVLS(Rs("MASTER_FLAG"), "")
   ElseIf Ind = 7 Then
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
   ElseIf Ind = 8 Then
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
   ElseIf Ind = 9 Then
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
   ElseIf Ind = 10 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
   ElseIf Ind = 11 Then
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
   ElseIf Ind = 12 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
   ElseIf Ind = 13 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
   ElseIf Ind = 14 Then
'      BALANCE_ACCUM_ID = NVLI(Rs("BALANCE_ACCUM_ID"), 0)
'      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
'      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
   ElseIf Ind = 15 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      PART_TYPE = NVLI(Rs("PART_TYPE"), 0)
      PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
      PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
   ElseIf Ind = 16 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      PIG_AGE = NVLI(Rs("PIG_AGE"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
   ElseIf Ind = 17 Then
      FEED_GROUP = NVLI(Rs("FEED_GROUP"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
   ElseIf Ind = 18 Then
      PART_NO = NVLS(Rs("PART_NO"), "")
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
   ElseIf Ind = 20 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
   ElseIf Ind = 21 Then
      BALANCE_ACCUM_ID = NVLI(Rs("BALANCE_ACCUM_ID"), 0)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
      MASTER_FLAG = NVLS(Rs("MASTER_FLAG"), "")
   ElseIf Ind = 22 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
   ElseIf Ind = 23 Then
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
   ElseIf Ind = 24 Then
      FEED_GROUP = NVLI(Rs("FEED_GROUP"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
      TOTAL_INCLUDE_PRICE = NVLD(Rs("TOTAL_INCLUDE_PRICE"), 0)
      AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
   ElseIf Ind = 25 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
   ElseIf Ind = 26 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      BALANCE_AMOUNT = NVLD(Rs("BALANCE_AMOUNT"), 0)
   End If
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

   TempID1 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID1 = "NULL"
   End If
      
   TempID2 = BATCH_ID
   If BATCH_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   SIMULATE_FLAG = glbUser.SIMULATE_FLAG
   YYYYMM = Mid(DateToStringInt(DOCUMENT_DATE), 1, 7)
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("BALANCE_ACCUM_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO BALANCE_ACCUM " & _
                     "( " & _
                     "   BALANCE_ACCUM_ID, " & _
                     "   DOCUMENT_DATE, " & _
                     "   PART_ITEM_ID, " & _
                     "   IMPORT_AMOUNT, " & _
                     "   EXPORT_AMOUNT, " & _
                     "   BALANCE_AMOUNT, " & _
                     "   TOTAL_INCLUDE_PRICE, " & _
                     "   LOCATION_ID, " & _
                     "   AVG_PRICE, " & _
                     "   SIMULATE_FLAG, " & _
                     "   BATCH_ID, " & _
                     "   YYYYMM, " & _
                     "   PIG_AGE, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
            SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                        "'" & ChangeQuote(DateToStringInt(Trim(DOCUMENT_DATE))) & "'," & _
                        PART_ITEM_ID & "," & _
                        IMPORT_AMOUNT & "," & _
                        EXPORT_AMOUNT & "," & _
                        BALANCE_AMOUNT & "," & _
                        TOTAL_INCLUDE_PRICE & "," & _
                        LOCATION_ID & "," & _
                        AVG_PRICE & "," & _
                        "'" & ChangeQuote(Trim(SIMULATE_FLAG)) & "'," & _
                        TempID2 & "," & _
                        "'" & ChangeQuote(Trim(YYYYMM)) & "'," & _
                        PIG_AGE & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         BALANCE_ACCUM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE BALANCE_ACCUM SET " & _
                     "    DOCUMENT_DATE = '" & ChangeQuote(DateToStringInt(Trim(DOCUMENT_DATE))) & "', " & _
                     "    PART_ITEM_ID = " & PART_ITEM_ID & ", " & _
                     "    IMPORT_AMOUNT = " & IMPORT_AMOUNT & ", " & _
                     "    EXPORT_AMOUNT = " & EXPORT_AMOUNT & ", " & _
                     "    BALANCE_AMOUNT = " & BALANCE_AMOUNT & ", " & _
                     "    TOTAL_INCLUDE_PRICE = " & TOTAL_INCLUDE_PRICE & ", " & _
                     "    LOCATION_ID = " & LOCATION_ID & ", " & _
                     "    AVG_PRICE = " & AVG_PRICE & ", " & _
                     "    SIMULATE_FLAG = '" & ChangeQuote(Trim(SIMULATE_FLAG)) & "', " & _
                     "    BATCH_ID = " & TempID2 & ", " & _
                     "    YYYYMM = '" & ChangeQuote(Trim(YYYYMM)) & "', " & _
                     "    PIG_AGE = " & PIG_AGE & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE BALANCE_ACCUM_ID = " & BALANCE_ACCUM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function UpdatePriceAmount()
Dim ServerDateTime As String
Dim SQL1 As String
      
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
   Else 'Show edit
      SQL1 = "UPDATE BALANCE_ACCUM SET " & _
                     "    BALANCE_AMOUNT = " & BALANCE_AMOUNT & ", " & _
                     "    TOTAL_INCLUDE_PRICE = " & BALANCE_AMOUNT * AVG_PRICE & ", " & _
                     "    AVG_PRICE = " & AVG_PRICE & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE BALANCE_ACCUM_ID = " & BALANCE_ACCUM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM BALANCE_ACCUM WHERE BALANCE_ACCUM_ID = " & BALANCE_ACCUM_ID
      m_Conn.Execute (SQL1)
End Function

Public Function ClearData()
Dim SQL1 As String
Dim TempDate As String
Dim WhereStr As String

   SIMULATE_FLAG = glbUser.SIMULATE_FLAG
   
   SQL1 = "DELETE FROM BALANCE_ACCUM "
   
   If FROM_DATE > 0 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   If TO_DATE > 0 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
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
      
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   SQL1 = SQL1 & WhereStr
   m_Conn.Execute (SQL1)
End Function

Public Function PatchPigAge()
Dim ServerDateTime As String
Dim SQL1 As String
      
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
   Else 'Show edit
      SQL1 = "UPDATE BALANCE_ACCUM SET " & _
                     "    PIG_AGE = " & PIG_AGE & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE BALANCE_ACCUM_ID = " & BALANCE_ACCUM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function


