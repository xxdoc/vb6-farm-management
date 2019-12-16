VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCustomer"
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

Public CUSTOMER_ID As Long
Public CUSTOMER_CODE As String
Public CUSTOMER_GRADE As Long
Public Credit As Long
Public TAX_ID As String
Public CUSTOMER_TYPE As Long
Public EMAIL As String
Public BIRTH_DATE As Date
Public WEBSITE As String
Public CUSTOMER_PASSWORD As String
Public BUSINESS_TYPE As Long
Public BUSINESS_DESC As String
Public NORMAL_DISCOUNT As Double
Public RESPONSE_BY As Long
Public CREDIT_AMOUNT As Double
Public REGION_ID As Long

Public FROM_DATE As Date
Public TO_DATE As Date
Public CUSTOMER_NAME As String
Public DOCUMENT_NO As String
Public DOCUMENT_TYPE As Long
Public CSTTYPE_NAME As String
Public CSTGRADE_NAME As String
Public EXCEPTION_FLAG As String
Public SALE_BY As Long
Public ACCOUNT_ID As Long

Public TAKE_AWAY_FLAG As String
Public TAKE_KEY_ACCOUNT As String

Public ADDRESS_ID As Long
Public HOME As String
Public SOI As String
Public MOO As String
Public VILLAGE As String
Public ROAD As String
Public DISTRICT As String
Public AMPHUR As String
Public PROVINCE As String
Public COUNTRY_ID As Long
Public PHONE1 As String
Public PHONE2 As String
Public FAX1 As String
Public FAX2 As String
Public ZIPCODE As String
Public BANGKOK_FLAG As String

Public EMP_CODE As String


Public KEY_LOOKUP As String
Public KEY_ID As Long

Public CstNames As Collection
Public CstDocs As Collection
Public CstAddr As Collection
Public CstContacts As Collection
Public CstAccounts As Collection
Public CstCoas As Collection

Public CstPackage As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set CstNames = New Collection
   Set CstDocs = New Collection
   Set CstAddr = New Collection
   Set CstContacts = New Collection
   Set CstAccounts = New Collection
   Set CstCoas = New Collection
   
   Set CstPackage = New Collection
   
   FROM_DATE = -1
   TO_DATE = -1
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Rs As ADODB.Recordset, itemcount As Long)
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

   SelectStr1 = "SELECT ADR.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM CUSTOMER ADR "
   
   If CUSTOMER_ID >= 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CUSTOMER_ID = " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""
   If OrderBy > 0 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY CUSTOMER_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY CUSTOMER_ID "
      Else
         OrderByStr = "ORDER BY CUSTOMER_ID "
      End If
      
      If OrderType = 1 Then
         OrderTypeStr = "ASC "
      ElseIf OrderType = 2 Then
         OrderTypeStr = "DESC "
      Else
         OrderTypeStr = "DESC "
      End If
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Function

'In this class we will not handle any error or any transaction
Public Function QueryData2(Rs As ADODB.Recordset, itemcount As Long)
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

   SelectStr1 = "SELECT CM.*, NM.LONG_NAME CUSTOMER_NAME, CT.CSTTYPE_ID, CT.CSTTYPE_NAME, CG.CSTGRADE_ID, CG.CSTGRADE_NAME "
   SelectStr2 = "SELECT COUNT(*)  "
   
   WhereStr = ""

   FromStr = "FROM CUSTOMER CM "
   FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CN.CUSTOMER_ID = CM.CUSTOMER_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_TYPE CT ON (CM.CUSTOMER_TYPE = CT.CSTTYPE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_GRADE CG ON (CM.CUSTOMER_GRADE = CG.CSTGRADE_ID) "
   
   If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CM.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If
      
   If CUSTOMER_GRADE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CUSTOMER_GRADE = " & CUSTOMER_GRADE & ") "
      Else
         WhereStr = WhereStr & "AND (CUSTOMER_GRADE = " & CUSTOMER_GRADE & ") "
      End If
   End If

   If CUSTOMER_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CUSTOMER_TYPE = " & CUSTOMER_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (CUSTOMER_TYPE = " & CUSTOMER_TYPE & ") "
      End If
   End If

   If BUSINESS_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BUSINESS_TYPE = " & BUSINESS_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (BUSINESS_TYPE = " & BUSINESS_TYPE & ") "
      End If
   End If

   If Trim(CUSTOMER_NAME) <> "" Then
      If WildCard(CUSTOMER_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (LONG_NAME = '" & ChangeQuote(Trim(CUSTOMER_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (LONG_NAME = '" & ChangeQuote(Trim(CUSTOMER_NAME)) & "')"
         End If
      End If
   End If

   If Trim(CUSTOMER_CODE) <> "" Then
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
   
   Dim TempStr As String
   If DOCUMENT_TYPE > 0 Then
      TempStr = "(SELECT COUNT(*) FROM CUSTOMER_PROOF CP, DOCUMENT DM WHERE (CP.DOCUMENT_ID = DM.DOCUMENT_ID) AND (CP.CUSTOMER_ID = CM.CUSTOMER_ID) AND (DM.DOCUMENT_NO = '" & DOCUMENT_NO & "') AND (DM.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ")) > 0 "
      If WhereStr = "" Then
         WhereStr = "WHERE (" & TempStr & ") "
      Else
         WhereStr = WhereStr & "AND (" & TempStr & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY CUSTOMER_CODE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY NM.LONG_NAME "
   Else
      OrderByStr = "ORDER BY CUSTOMER_CODE "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Function

Public Function QueryData3(Rs As ADODB.Recordset, itemcount As Long)
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
Dim TempStr As String

   SelectStr1 = "SELECT BT.BSNTYPE_NAME, CM.*, NM.LONG_NAME CUSTOMER_NAME, CT.CSTTYPE_ID, CT.CSTTYPE_NAME, CG.CSTGRADE_ID, CG.CSTGRADE_NAME, AC.ACCOUNT_NO,  "
   
   TempStr = ""
   If FROM_DATE > 0 Then
      TempStr = TempStr & "(DO1.DO_DATE >= '" & DateToStringIntLow(FROM_DATE) & "') AND "
   End If
   If TO_DATE > 0 Then
      TempStr = TempStr & "(DO1.DO_DATE <= '" & DateToStringIntHi(TO_DATE) & "') AND "
   End If
   If Len(EXCEPTION_FLAG) > 0 Then
      TempStr = TempStr & "(DO1.EXCEPTION_FLAG = '" & EXCEPTION_FLAG & "') AND "
   End If
   SelectStr1 = SelectStr1 & "(SELECT SUM(DI.AMOUNT) FROM DO_ITEM DI, DELIVERY_ORDER DO1 "
   SelectStr1 = SelectStr1 & "WHERE " & TempStr & " (DI.DO_ID = DO1.DO_ID) AND (DO1.ACCOUNT_ID = AC.ACCOUNT_ID)) SUM_DO_AMOUNT, "
   
   TempStr = ""
   If FROM_DATE > 0 Then
      TempStr = TempStr & "(DO1.DO_DATE >= '" & DateToStringIntLow(FROM_DATE) & "') AND "
   End If
   If TO_DATE > 0 Then
      TempStr = TempStr & "(DO1.DO_DATE <= '" & DateToStringIntHi(TO_DATE) & "') AND "
   End If
   If Len(EXCEPTION_FLAG) > 0 Then
      TempStr = TempStr & "(DO1.EXCEPTION_FLAG = '" & EXCEPTION_FLAG & "') AND "
   End If
   SelectStr1 = SelectStr1 & "(SELECT SUM(DI.DISCOUNT_AMOUNT) FROM DO_ITEM DI, DELIVERY_ORDER DO1 "
   SelectStr1 = SelectStr1 & "WHERE " & TempStr & " (DI.DO_ID = DO1.DO_ID) AND (DO1.ACCOUNT_ID = AC.ACCOUNT_ID)) SUM_DISCOUNT, "

   TempStr = ""
   If FROM_DATE > 0 Then
      TempStr = TempStr & "(BL.BILL_DATE >= '" & DateToStringIntLow(FROM_DATE) & "') AND "
   End If
   If TO_DATE > 0 Then
      TempStr = TempStr & "(BL.BILL_DATE <= '" & DateToStringIntHi(TO_DATE) & "') AND "
   End If
   If Len(EXCEPTION_FLAG) > 0 Then
      TempStr = TempStr & "(BL.EXCEPTION_FLAG = '" & EXCEPTION_FLAG & "') AND "
   End If
   SelectStr1 = SelectStr1 & "(SELECT SUM(BL.AVG_TAX*(BL.BILL_AMOUNT - BL.DISCOUNT_PERCENT) /100) FROM BILL BL "
   SelectStr1 = SelectStr1 & "WHERE " & TempStr & " (BL.BILLS_FLAG = 'N') AND (BL.ACCOUNT_ID = AC.ACCOUNT_ID)) SUM_TAX, "

   TempStr = ""
   If FROM_DATE > 0 Then
      TempStr = TempStr & "(PO.PO_DATE >= '" & DateToStringIntLow(FROM_DATE) & "') AND "
   End If
   If TO_DATE > 0 Then
      TempStr = TempStr & "(PO.PO_DATE <= '" & DateToStringIntHi(TO_DATE) & "') AND "
   End If
   If Len(EXCEPTION_FLAG) > 0 Then
      TempStr = TempStr & "(PO.EXCEPTION_FLAG = '" & EXCEPTION_FLAG & "') AND "
   End If
   SelectStr1 = SelectStr1 & "(SELECT SUM(PO.DEPOSIT_AMOUNT) FROM PURCHASE_ORDER PO "
   SelectStr1 = SelectStr1 & "WHERE " & TempStr & " (PO.ACCOUNT_ID = AC.ACCOUNT_ID)) SUM_DEPOSIT, "

   TempStr = ""
   If FROM_DATE > 0 Then
      TempStr = TempStr & "(BL.BILL_DATE >= '" & DateToStringIntLow(FROM_DATE) & "') AND "
   End If
   If TO_DATE > 0 Then
      TempStr = TempStr & "(BL.BILL_DATE <= '" & DateToStringIntHi(TO_DATE) & "') AND "
   End If
   If Len(EXCEPTION_FLAG) > 0 Then
      TempStr = TempStr & "(BL.EXCEPTION_FLAG = '" & EXCEPTION_FLAG & "') AND "
   End If
   SelectStr1 = SelectStr1 & "(SELECT SUM(BL.PAID_AMOUNT) FROM BILL BL "
   SelectStr1 = SelectStr1 & "WHERE " & TempStr & " (BL.BILLS_FLAG = 'N') AND (BL.ACCOUNT_ID = AC.ACCOUNT_ID)) SUM_PAID, "

   TempStr = ""
   If FROM_DATE > 0 Then
      TempStr = TempStr & "(RC.RECEIPT_DATE >= '" & DateToStringIntLow(FROM_DATE) & "') AND "
   End If
   If TO_DATE > 0 Then
      TempStr = TempStr & "(RC.RECEIPT_DATE <= '" & DateToStringIntHi(TO_DATE) & "') AND "
   End If
   If Len(EXCEPTION_FLAG) > 0 Then
      TempStr = TempStr & "(RC.EXCEPTION_FLAG = '" & EXCEPTION_FLAG & "') AND "
   End If
   SelectStr1 = SelectStr1 & "(SELECT SUM(RC.ACTUAL_RECEIVE) FROM RECEIPTS RC "
   SelectStr1 = SelectStr1 & "LEFT OUTER JOIN RECEIPT_STATUS RS ON (RC.RECEIPT_STATUS = RS.RECEIPTSTS_ID) "
   SelectStr1 = SelectStr1 & "WHERE " & TempStr & " (RC.ACCOUNT_ID = AC.ACCOUNT_ID) AND ((RS.POST_FLAG = 'Y') OR (RC.RECEIPT_STATUS IS NULL))) SUM_RECEIPT, "

   TempStr = ""
   If FROM_DATE > 0 Then
      TempStr = TempStr & "(DO1.DO_DATE >= '" & DateToStringIntLow(FROM_DATE) & "') AND "
   End If
   If TO_DATE > 0 Then
      TempStr = TempStr & "(DO1.DO_DATE <= '" & DateToStringIntHi(TO_DATE) & "') AND "
   End If
   If Len(EXCEPTION_FLAG) > 0 Then
      TempStr = TempStr & "(DO1.EXCEPTION_FLAG = '" & EXCEPTION_FLAG & "') AND "
   End If
   SelectStr1 = SelectStr1 & "(SELECT SUM(DO1.PAID_AMOUNT) FROM DELIVERY_ORDER DO1 "
   SelectStr1 = SelectStr1 & "WHERE " & TempStr & " (DO1.ACCOUNT_ID = AC.ACCOUNT_ID)) SUM_DO_PAID "

   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM ACCOUNT AC, CUSTOMER CM, NAME NM, CUSTOMER_TYPE CT, CUSTOMER_GRADE CG, CUSTOMER_NAME CN, BUSINESS_TYPE BT "
   
   WhereStr = "WHERE (AC.CUSTOMER_ID = CM.CUSTOMER_ID) AND (CN.CUSTOMER_ID = CM.CUSTOMER_ID) "
   WhereStr = WhereStr & "AND (CN.NAME_ID = NM.NAME_ID) "
   WhereStr = WhereStr & "AND (CM.CUSTOMER_TYPE = CT.CSTTYPE_ID) "
   WhereStr = WhereStr & "AND (CM.CUSTOMER_GRADE = CG.CSTGRADE_ID) "
   WhereStr = WhereStr & "AND (CM.BUSINESS_TYPE = BT.BSNTYPE_ID) "
   
   If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CUSTOMER_ID = " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If
      
   If CUSTOMER_GRADE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CUSTOMER_GRADE = " & CUSTOMER_GRADE & ") "
      Else
         WhereStr = WhereStr & "AND (CUSTOMER_GRADE = " & CUSTOMER_GRADE & ") "
      End If
   End If

   If CUSTOMER_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CUSTOMER_TYPE = " & CUSTOMER_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (CUSTOMER_TYPE = " & CUSTOMER_TYPE & ") "
      End If
   End If

   If BUSINESS_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BUSINESS_TYPE = " & BUSINESS_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (BUSINESS_TYPE = " & BUSINESS_TYPE & ") "
      End If
   End If

   If Trim(CUSTOMER_NAME) <> "" Then
      If WildCard(CUSTOMER_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (LONG_NAME = '" & ChangeQuote(Trim(CUSTOMER_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (LONG_NAME = '" & ChangeQuote(Trim(CUSTOMER_NAME)) & "')"
         End If
      End If
   End If

   If Trim(CUSTOMER_CODE) <> "" Then
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

   If DOCUMENT_TYPE > 0 Then
      TempStr = "(SELECT COUNT(*) FROM CUSTOMER_PROOF CP, DOCUMENT DM WHERE (CP.DOCUMENT_ID = DM.DOCUMENT_ID) AND (CP.CUSTOMER_ID = CM.CUSTOMER_ID) AND (DM.DOCUMENT_NO = '" & DOCUMENT_NO & "') AND (DM.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ")) > 0 "
      If WhereStr = "" Then
         WhereStr = "WHERE (" & TempStr & ") "
      Else
         WhereStr = WhereStr & "AND (" & TempStr & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY NM.LONG_NAME "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY NM.LONG_NAME "
   Else
      OrderByStr = "ORDER BY CUSTOMER_ID "
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

Public Function QueryData4(Rs As ADODB.Recordset, itemcount As Long)
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
Dim TempStr As String

   SelectStr1 = "SELECT CM.*, NM.LONG_NAME CUSTOMER_NAME, "
      
   TempStr = ""
   If FROM_DATE > 0 Then
      TempStr = TempStr & "(DO1.DO_DATE >= '" & DateToStringIntLow(FROM_DATE) & "') AND "
   End If
   If TO_DATE > 0 Then
      TempStr = TempStr & "(DO1.DO_DATE <= '" & DateToStringIntHi(TO_DATE) & "') AND "
   End If
   If Len(EXCEPTION_FLAG) > 0 Then
      TempStr = TempStr & "(DO1.EXCEPTION_FLAG = '" & EXCEPTION_FLAG & "') AND "
   End If
   If SALE_BY > 0 Then
      TempStr = TempStr & "(DO1.ACCEPT_BY = '" & SALE_BY & "') AND "
   End If
   SelectStr1 = SelectStr1 & "(SELECT SUM(DO1.DO_AMOUNT - DO1.DISCOUNT_AMOUNT) FROM DELIVERY_ORDER DO1, ACCOUNT AC "
   SelectStr1 = SelectStr1 & "WHERE " & TempStr & " (DO1.ACCOUNT_ID = AC.ACCOUNT_ID) AND (AC.CUSTOMER_ID = CM.CUSTOMER_ID)) SUM_DO_AMOUNT, "
   
   TempStr = ""
   If FROM_DATE > 0 Then
      TempStr = TempStr & "(DO1.DO_DATE >= '" & DateToStringIntLow(FROM_DATE) & "') AND "
   End If
   If TO_DATE > 0 Then
      TempStr = TempStr & "(DO1.DO_DATE <= '" & DateToStringIntHi(TO_DATE) & "') AND "
   End If
   If Len(EXCEPTION_FLAG) > 0 Then
      TempStr = TempStr & "(DO1.EXCEPTION_FLAG = '" & EXCEPTION_FLAG & "') AND "
   End If
   If SALE_BY > 0 Then
      TempStr = TempStr & "(DO1.ACCEPT_BY = '" & SALE_BY & "') AND "
   End If
   SelectStr1 = SelectStr1 & "(SELECT SUM(US.UNIT_AMOUNT) FROM USAGE US, DELIVERY_ORDER DO1, ACCOUNT AC "
   SelectStr1 = SelectStr1 & "WHERE " & TempStr & " (US.USAGE_GROUP_ID = DO1.USAGE_GROUP_ID) AND (DO1.ACCOUNT_ID = AC.ACCOUNT_ID) AND (AC.CUSTOMER_ID = CM.CUSTOMER_ID)) SUM_UNIT_AMOUNT "

   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM CUSTOMER CM "
   FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CN.CUSTOMER_ID = CM.CUSTOMER_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID) "
   
   WhereStr = ""
   
   If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CUSTOMER_ID = " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If

   If Trim(CUSTOMER_NAME) <> "" Then
      If WildCard(CUSTOMER_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (LONG_NAME = '" & ChangeQuote(Trim(CUSTOMER_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (LONG_NAME = '" & ChangeQuote(Trim(CUSTOMER_NAME)) & "')"
         End If
      End If
   End If

   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY NM.LONG_NAME "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY NM.LONG_NAME "
   Else
      OrderByStr = "ORDER BY NM.LONG_NAME "
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
   CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
   CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
   CUSTOMER_GRADE = NVLI(Rs("CUSTOMER_GRADE"), -1)
   Credit = NVLS(Rs("CREDIT"), "")
   NORMAL_DISCOUNT = NVLD(Rs("NORMAL_DISCOUNT"), 0)
   TAX_ID = NVLS(Rs("TAX_ID"), "")
   CUSTOMER_TYPE = NVLI(Rs("CUSTOMER_TYPE"), -1)
   EMAIL = NVLS(Rs("EMAIL"), "")
   WEBSITE = NVLS(Rs("WEBSITE"), "")
   BIRTH_DATE = InternalDateToDate(NVLS(Rs("BIRTH_DATE"), ""))
   CUSTOMER_PASSWORD = DecryptText(NVLS(Rs("CUSTOMER_PASSWORD"), ""))
   BUSINESS_TYPE = NVLI(Rs("BUSINESS_TYPE"), -1)
   BUSINESS_DESC = NVLS(Rs("BUSINESS_DESC"), "")
   RESPONSE_BY = NVLI(Rs("RESPONSE_BY"), 0)
   CREDIT_AMOUNT = NVLD(Rs("CREDIT_AMOUNT"), 0)
   REGION_ID = NVLI(Rs("REGION_ID"), 0)
   
   
   TAKE_AWAY_FLAG = NVLS(Rs("TAKE_AWAY_FLAG"), "")
   TAKE_KEY_ACCOUNT = NVLS(Rs("TAKE_KEY_ACCOUNT"), "")
   
   KEY_LOOKUP = CUSTOMER_CODE
   KEY_ID = CUSTOMER_ID
   
   CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
   CSTTYPE_NAME = NVLS(Rs("CSTTYPE_NAME"), "")
   CSTGRADE_NAME = NVLS(Rs("CSTGRADE_NAME"), "")
End Function
Public Function PopulateFromRSInfo(Ind As Long, Rs As ADODB.Recordset)
   CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
   CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
   CUSTOMER_GRADE = NVLI(Rs("CUSTOMER_GRADE"), -1)
   Credit = NVLS(Rs("CREDIT"), "")
   NORMAL_DISCOUNT = NVLD(Rs("NORMAL_DISCOUNT"), 0)
   TAX_ID = NVLS(Rs("TAX_ID"), "")
   CUSTOMER_TYPE = NVLI(Rs("CUSTOMER_TYPE"), -1)
   EMAIL = NVLS(Rs("EMAIL"), "")
   WEBSITE = NVLS(Rs("WEBSITE"), "")
   BIRTH_DATE = InternalDateToDate(NVLS(Rs("BIRTH_DATE"), ""))
   CUSTOMER_PASSWORD = DecryptText(NVLS(Rs("CUSTOMER_PASSWORD"), ""))
   BUSINESS_TYPE = NVLI(Rs("BUSINESS_TYPE"), -1)
   BUSINESS_DESC = NVLS(Rs("BUSINESS_DESC"), "")
   RESPONSE_BY = NVLI(Rs("RESPONSE_BY"), 0)
   CREDIT_AMOUNT = NVLD(Rs("CREDIT_AMOUNT"), 0)
   REGION_ID = NVLI(Rs("REGION_ID"), 0)
   
   TAKE_AWAY_FLAG = NVLS(Rs("TAKE_AWAY_FLAG"), "")
   TAKE_KEY_ACCOUNT = NVLS(Rs("TAKE_KEY_ACCOUNT"), "")
   
   KEY_LOOKUP = CUSTOMER_CODE
   KEY_ID = CUSTOMER_ID
   
   CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
   CSTTYPE_NAME = NVLS(Rs("CSTTYPE_NAME"), "")
   CSTGRADE_NAME = NVLS(Rs("CSTGRADE_NAME"), "")
   
   '===
      ADDRESS_ID = NVLI(Rs("ADDRESS_ID"), -1)
      HOME = NVLS(Rs("HOME"), "")
      SOI = NVLS(Rs("SOI"), "")
      MOO = NVLS(Rs("MOO"), "")
      VILLAGE = NVLS(Rs("VILLAGE"), "")
      ROAD = NVLS(Rs("ROAD"), "")
      DISTRICT = NVLS(Rs("DISTRICT"), "")
      AMPHUR = NVLS(Rs("AMPHUR"), "")
      PROVINCE = NVLS(Rs("PROVINCE"), "")
      COUNTRY_ID = NVLI(Rs("COUNTRY_ID"), -1)
      PHONE1 = NVLS(Rs("PHONE1"), "")
      PHONE2 = NVLS(Rs("PHONE2"), "")
      FAX1 = NVLS(Rs("FAX1"), "")
      FAX2 = NVLS(Rs("FAX2"), "")
      ZIPCODE = NVLS(Rs("ZIPCODE"), "")
'      COUNTRY_NAME = NVLS(Rs("COUNTRY_NAME"), "")
      BANGKOK_FLAG = NVLS(Rs("BANGKOK_FLAG"), "N")
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

   TempID1 = RESPONSE_BY
   If RESPONSE_BY <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = BUSINESS_TYPE
   If BUSINESS_TYPE <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID3 = CUSTOMER_GRADE
   If CUSTOMER_GRADE <= 0 Then
      TempID3 = "NULL"
   End If
   
   TempID4 = CUSTOMER_TYPE
   If CUSTOMER_TYPE <= 0 Then
      TempID4 = "NULL"
   End If
   
   TempID5 = REGION_ID
   If REGION_ID <= 0 Then
      TempID5 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("CUSTOMER_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO CUSTOMER " & _
                     "( " & _
                     "   CUSTOMER_ID, " & _
                     "   CUSTOMER_CODE, " & _
                     "   CUSTOMER_GRADE, " & _
                     "   CREDIT, " & _
                     "   TAX_ID, " & _
                     "   CUSTOMER_TYPE, " & _
                     "   EMAIL, " & _
                     "   WEBSITE, " & _
                     "   CUSTOMER_PASSWORD, " & _
                     "   BIRTH_DATE, " & _
                     "   BUSINESS_TYPE, " & _
                     "   BUSINESS_DESC, " & _
                     "   NORMAL_DISCOUNT, " & _
                     "   RESPONSE_BY, " & _
                     "   CREDIT_AMOUNT, " & _
                     "   REGION_ID, "
                  SQL1 = SQL1 & _
                     "   TAKE_AWAY_FLAG, " & _
                     "   TAKE_KEY_ACCOUNT, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(CUSTOMER_CODE)) & "'," & _
                         TempID3 & "," & _
                         "'" & ChangeQuote(Trim(Credit)) & "'," & _
                         "'" & ChangeQuote(Trim(TAX_ID)) & "'," & _
                         TempID4 & "," & _
                         "'" & ChangeQuote(Trim(EMAIL)) & "'," & _
                         "'" & ChangeQuote(Trim(WEBSITE)) & "'," & _
                         "'" & ChangeQuote(Trim(EncryptText(CUSTOMER_PASSWORD))) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(BIRTH_DATE))) & "'," & _
                         TempID2 & "," & _
                         "'" & ChangeQuote(Trim(BUSINESS_DESC)) & "'," & _
                         "" & NORMAL_DISCOUNT & "," & _
                         "" & TempID1 & "," & _
                         "" & CREDIT_AMOUNT & "," & _
                         "" & TempID5 & "," & _
                         "'" & ChangeQuote(Trim(TAKE_AWAY_FLAG)) & "'," & _
                        "'" & ChangeQuote(Trim(TAKE_KEY_ACCOUNT)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         CUSTOMER_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE CUSTOMER SET " & _
                     "   CUSTOMER_CODE = '" & ChangeQuote(Trim(CUSTOMER_CODE)) & "', " & _
                     "   CUSTOMER_GRADE = " & TempID3 & ", " & _
                     "   CREDIT = '" & Credit & "', " & _
                     "   TAX_ID = '" & ChangeQuote(Trim(TAX_ID)) & "', " & _
                     "   CUSTOMER_TYPE = " & TempID4 & ", " & _
                     "   EMAIL = '" & ChangeQuote(Trim(EMAIL)) & "', " & _
                     "   WEBSITE = '" & ChangeQuote(Trim(WEBSITE)) & "', " & _
                     "   CUSTOMER_PASSWORD = '" & ChangeQuote(Trim(EncryptText(CUSTOMER_PASSWORD))) & "', " & _
                     "   BIRTH_DATE = '" & ChangeQuote(Trim(DateToStringInt(BIRTH_DATE))) & "', " & _
                     "   BUSINESS_TYPE = " & TempID2 & ", " & _
                     "   BUSINESS_DESC = '" & ChangeQuote(Trim(BUSINESS_DESC)) & "', " & _
                     "   NORMAL_DISCOUNT = '" & NORMAL_DISCOUNT & "', " & _
                     "   RESPONSE_BY = " & TempID1 & ", " & _
                     "   CREDIT_AMOUNT = '" & CREDIT_AMOUNT & "', " & _
                     "   REGION_ID = " & TempID5 & ", " & _
                    "   TAKE_AWAY_FLAG = '" & ChangeQuote(Trim(TAKE_AWAY_FLAG)) & "', " & _
                     "   TAKE_KEY_ACCOUNT = '" & ChangeQuote(Trim(TAKE_KEY_ACCOUNT)) & "', " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE CUSTOMER_ID = " & CUSTOMER_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM CUSTOMER_ADDRESS WHERE CUSTOMER_ID = " & CUSTOMER_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM CUSTOMER_NAME WHERE CUSTOMER_ID = " & CUSTOMER_ID
      m_Conn.Execute (SQL1)
      
'      SQL1 = "DELETE FROM CUSTOMER_PROOF WHERE CUSTOMER_ID = " & CUSTOMER_ID
'      m_Conn.Execute (SQL1)
'
'      SQL1 = "DELETE FROM CUSTOMER_CONTACT WHERE CUSTOMER_ID = " & CUSTOMER_ID
'      m_Conn.Execute (SQL1)
'
'      SQL1 = "DELETE FROM CUSTOMER_PICTURE WHERE CUSTOMER_ID = " & CUSTOMER_ID
'      m_Conn.Execute (SQL1)
'
'      SQL1 = "DELETE FROM AGREEMENT WHERE SUBSCRIBER_ID IN (SELECT SUBSCRIBER_ID FROM SUBSCRIBER WHERE ACCOUNT_ID IN (SELECT ACCOUNT_ID FROM ACCOUNT WHERE CUSTOMER_ID = " & CUSTOMER_ID & ")) "
'      m_Conn.Execute (SQL1)
'
'      SQL1 = "DELETE FROM SUBSCRIBER WHERE ACCOUNT_ID IN (SELECT ACCOUNT_ID FROM ACCOUNT WHERE CUSTOMER_ID = " & CUSTOMER_ID & ") "
'      m_Conn.Execute (SQL1)
'
      SQL1 = "DELETE FROM ACCOUNT WHERE CUSTOMER_ID = " & CUSTOMER_ID
      m_Conn.Execute (SQL1)
'
'      SQL1 = "DELETE FROM CUSTOMER_COA WHERE CUSTOMER_ID = " & CUSTOMER_ID
'      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM CUSTOMER WHERE CUSTOMER_ID = " & CUSTOMER_ID
      m_Conn.Execute (SQL1)
End Function

Public Function UpdateARCreditByCode()
Dim SQL1 As String
            
   SQL1 = "UPDATE CUSTOMER SET CREDIT = " & Credit & " WHERE CUSTOMER_CODE = " & "'" & CUSTOMER_CODE & "'"
   m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set CstNames = Nothing
   Set CstDocs = Nothing
   Set CstAddr = Nothing
   Set CstContacts = Nothing
   Set CstAccounts = Nothing
   Set CstCoas = Nothing
   Set CstPackage = Nothing
End Sub
Public Function QueryData1(Rs As ADODB.Recordset, itemcount As Long)
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

   SelectStr1 = "SELECT CM.*, NM.LONG_NAME CUSTOMER_NAME, CT.CSTTYPE_ID, CT.CSTTYPE_NAME, CG.CSTGRADE_ID, CG.CSTGRADE_NAME, AD.* "
   SelectStr2 = "SELECT COUNT(*)  "
   
   WhereStr = ""
   
   FromStr = "FROM CUSTOMER CM "
   FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CN.CUSTOMER_ID = CM.CUSTOMER_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_ADDRESS CA ON (CA.CUSTOMER_ID = CM.CUSTOMER_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN ADDRESS AD ON (CA.ADDRESS_ID = AD.ADDRESS_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_TYPE CT ON (CM.CUSTOMER_TYPE = CT.CSTTYPE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_GRADE CG ON (CM.CUSTOMER_GRADE = CG.CSTGRADE_ID) "
   
   If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CUSTOMER_ID = " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If
      
   If CUSTOMER_GRADE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CUSTOMER_GRADE = " & CUSTOMER_GRADE & ") "
      Else
         WhereStr = WhereStr & "AND (CUSTOMER_GRADE = " & CUSTOMER_GRADE & ") "
      End If
   End If

   If CUSTOMER_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CUSTOMER_TYPE = " & CUSTOMER_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (CUSTOMER_TYPE = " & CUSTOMER_TYPE & ") "
      End If
   End If

   If BUSINESS_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BUSINESS_TYPE = " & BUSINESS_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (BUSINESS_TYPE = " & BUSINESS_TYPE & ") "
      End If
   End If

   If Trim(CUSTOMER_NAME) <> "" Then
      If WildCard(CUSTOMER_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (LONG_NAME = '" & ChangeQuote(Trim(CUSTOMER_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (LONG_NAME = '" & ChangeQuote(Trim(CUSTOMER_NAME)) & "')"
         End If
      End If
   End If

   If Trim(CUSTOMER_CODE) <> "" Then
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

   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY CUSTOMER_CODE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY NM.LONG_NAME "
   ElseIf OrderBy = 3 Then
      OrderByStr = "ORDER BY PROVINCE "
   Else
      OrderByStr = "ORDER BY CUSTOMER_CODE "
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
Public Function QueryData5(Rs As ADODB.Recordset, itemcount As Long)
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
Dim TempDate  As String
    
   SelectStr1 = "SELECT DISTINCT CM.CUSTOMER_CODE, NM.LONG_NAME CUSTOMER_NAME, CT.CSTTYPE_ID, CT.CSTTYPE_NAME, CG.CSTGRADE_ID, CG.CSTGRADE_NAME, AD.* "
   SelectStr2 = "SELECT COUNT(*)  "
   
   WhereStr = ""
   
   FromStr = "FROM CUSTOMER CM "
   FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_NAME CN ON (CN.CUSTOMER_ID = CM.CUSTOMER_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (CN.NAME_ID = NM.NAME_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_ADDRESS CA ON (CA.CUSTOMER_ID = CM.CUSTOMER_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN ADDRESS AD ON (CA.ADDRESS_ID = AD.ADDRESS_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_TYPE CT ON (CM.CUSTOMER_TYPE = CT.CSTTYPE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER_GRADE CG ON (CM.CUSTOMER_GRADE = CG.CSTGRADE_ID) "
   
   FromStr = FromStr & "LEFT OUTER JOIN ACCOUNT AC ON (CM.CUSTOMER_ID = AC.CUSTOMER_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD ON (BD.ACCOUNT_ID = AC.ACCOUNT_ID) "
      
   If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CUSTOMER_ID = " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If
      
   If CUSTOMER_GRADE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CUSTOMER_GRADE = " & CUSTOMER_GRADE & ") "
      Else
         WhereStr = WhereStr & "AND (CUSTOMER_GRADE = " & CUSTOMER_GRADE & ") "
      End If
   End If

   If CUSTOMER_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CUSTOMER_TYPE = " & CUSTOMER_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (CUSTOMER_TYPE = " & CUSTOMER_TYPE & ") "
      End If
   End If

   If BUSINESS_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BUSINESS_TYPE = " & BUSINESS_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (BUSINESS_TYPE = " & BUSINESS_TYPE & ") "
      End If
   End If

   If Trim(CUSTOMER_NAME) <> "" Then
      If WildCard(CUSTOMER_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (LONG_NAME = '" & ChangeQuote(Trim(CUSTOMER_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (LONG_NAME = '" & ChangeQuote(Trim(CUSTOMER_NAME)) & "')"
         End If
      End If
   End If

   If Trim(CUSTOMER_CODE) <> "" Then
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

   If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BD.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BD.DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (BD.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (BD.DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY CUSTOMER_CODE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY NM.LONG_NAME "
   ElseIf OrderBy = 3 Then
      OrderByStr = "ORDER BY PROVINCE "
   Else
      OrderByStr = "ORDER BY CUSTOMER_CODE "
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
Public Function PopulateFromRS5(Ind As Long, Rs As ADODB.Recordset)
   CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
   
   CUSTOMER_NAME = NVLS(Rs("CUSTOMER_NAME"), "")
   CSTTYPE_NAME = NVLS(Rs("CSTTYPE_NAME"), "")
   CSTGRADE_NAME = NVLS(Rs("CSTGRADE_NAME"), "")
   
   '===
      ADDRESS_ID = NVLI(Rs("ADDRESS_ID"), -1)
      HOME = NVLS(Rs("HOME"), "")
      SOI = NVLS(Rs("SOI"), "")
      MOO = NVLS(Rs("MOO"), "")
      VILLAGE = NVLS(Rs("VILLAGE"), "")
      ROAD = NVLS(Rs("ROAD"), "")
      DISTRICT = NVLS(Rs("DISTRICT"), "")
      AMPHUR = NVLS(Rs("AMPHUR"), "")
      PROVINCE = NVLS(Rs("PROVINCE"), "")
      COUNTRY_ID = NVLI(Rs("COUNTRY_ID"), -1)
      PHONE1 = NVLS(Rs("PHONE1"), "")
      PHONE2 = NVLS(Rs("PHONE2"), "")
      FAX1 = NVLS(Rs("FAX1"), "")
      FAX2 = NVLS(Rs("FAX2"), "")
      ZIPCODE = NVLS(Rs("ZIPCODE"), "")
'      COUNTRY_NAME = NVLS(Rs("COUNTRY_NAME"), "")
      BANGKOK_FLAG = NVLS(Rs("BANGKOK_FLAG"), "N")
End Function

Public Function PackAddress(Optional NoZipCode As Boolean = False, Optional NoProvince As Boolean = False) As String
Dim AddressStr As String
   
   AddressStr = ""
   
   If Trim(HOME) <> "" Then
      AddressStr = AddressStr & HOME & " "
   End If

   If Trim(SOI) <> "" Then
      AddressStr = AddressStr & "´." & SOI & " "
   End If

   If Trim(MOO) <> "" Then
      AddressStr = AddressStr & MapText("À¡ŸË") & MOO & " "
   End If

   If Trim(VILLAGE) <> "" Then
      AddressStr = AddressStr & MapText("À¡ŸË∫È“π") & VILLAGE & " "
   End If

   If Trim(ROAD) <> "" Then
      AddressStr = AddressStr & MapText("∂.") & ROAD & " "
   End If

   If Trim(DISTRICT) <> "" Then
      If BANGKOK_FLAG = "Y" Then
         AddressStr = AddressStr & MapText("·¢«ß") & DISTRICT & " "
      Else
         AddressStr = AddressStr & MapText("µ”∫≈") & DISTRICT & " "
      End If
   End If

   If Trim(AMPHUR) <> "" Then
      If BANGKOK_FLAG = "Y" Then
         AddressStr = AddressStr & MapText("‡¢µ") & AMPHUR & " "
      Else
         AddressStr = AddressStr & MapText("Õ”‡¿Õ") & AMPHUR & " "
      End If
   End If
   
   If Not NoProvince Then
      If Trim(PROVINCE) <> "" Then
         If BANGKOK_FLAG = "Y" Then
            AddressStr = AddressStr & "" & PROVINCE & " "
         Else
            AddressStr = AddressStr & MapText("®.") & PROVINCE & " "
         End If
      End If
   End If
   
   If Not NoZipCode Then
      If Trim(ZIPCODE) <> "" Then
         AddressStr = AddressStr & " " & ZIPCODE & " "
      End If
   End If

   PackAddress = AddressStr
End Function


