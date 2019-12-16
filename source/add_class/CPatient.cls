VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPatient"
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

Public PATIENT_ID As Long
Public PATIENT_CODE As String
Public REGISTER_DATE As Date
Public CHANNEL As String
Public NAME As String
Public LAST_NAME As String
Public SEX As String
Public AGE As Double
Public HOME_NO1 As String
Public MOO1 As String
Public SOI1 As String
Public ROAD1 As String
Public KWANG1 As String
Public KHATE1 As String
Public ZIPCODE1 As String
Public PHONE1 As String
Public REFERENTOR1 As String
Public PHONE2 As String
Public PROVINCE As String
Public CARD_NO As String
Public DOCUMENT_ALLOW1 As String
Public NOTE1 As String
Public YEAR_KNOW As Long
Public REASON_KNOW As String
Public FAMILY_STATUS As Long
Public FAMILY_STATUS_OTH As String
Public INHERIT As String
Public INHERIT_OTH As String
Public JOB_FLAG As String
Public JOB_NAME As String
Public JOB_SALARY As Double
Public JOB_OTH As String
Public MEDICAL_EXPENSE As Double
Public EXPENSE_CODE As String
Public EXPENSE_OTH As String
Public CD4_VALUE As Double
Public CD4_DATE As Date
Public VIRAL_ROAD As Double
Public EQUIVALENCE As Double
Public EQUIVALENCE_DATE As Date
Public WEIGHT As Double
Public PRE_WEIGHT
Public HEIGHT As Double
Public HEALT_OTH As String
Public RELATIVE_CODE1 As String
Public RELATIVE_OTH1 As String
Public RELATIVE_CODE2 As String
Public RELATIVE_OTH2 As String
Public RELATIVE_CODE3 As String
Public RELATIVE_OTH3 As String
Public RELATIVE_CODE4 As String
Public RELATIVE_OTH4 As String
Public RELATIVE_CODE5 As String
Public RELATIVE_OTH5 As String
Public RELATIVE_CODE6 As String
Public RELATIVE_OTH6 As String
Public RELATIVE_CODE7 As String
Public RELATIVE_OTH7 As String
Public RELATIVE_CODE8 As String
Public RELATIVE_OTH8 As String
Public RELATIVE_CODE9 As String
Public RELATIVE_OTH9 As String
Public RELATIVE_CODE10 As String
Public RELATIVE_OTH10 As String
Public PSYCO_CODE1 As String
Public PSYCO_OTH1 As String
Public PSYCO_CODE2 As String
Public PSYCO_OTH2 As String
Public REPAIR_PLACE As String
Public REPAIR_PERSON As String
Public MEDICINE1 As String
Public MEDICINE2 As String
Public MEDICINE3 As String
Public MEDICINE4 As String
Public MEDICINE5 As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   REGISTER_DATE = -1
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT PT.* "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PATIENT PT "
   
   WhereStr = ""
   If PATIENT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PATIENT_ID = " & PATIENT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PATIENT_ID = " & PATIENT_ID & ") "
      End If
   End If
      
   If Trim(PATIENT_CODE) <> "" Then
      If WildCard(PATIENT_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(PATIENT_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(PATIENT_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PATIENT_CODE = '" & ChangeQuote(Trim(PATIENT_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (PATIENT_CODE = '" & ChangeQuote(Trim(PATIENT_CODE)) & "')"
         End If
      End If
   End If
      
   If Trim(NAME) <> "" Then
      If WildCard(NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (NAME = '" & ChangeQuote(Trim(NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (NAME = '" & ChangeQuote(Trim(NAME)) & "')"
         End If
      End If
   End If
      
   If Trim(LAST_NAME) <> "" Then
      If WildCard(LAST_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(LAST_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(LAST_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (LAST_NAME = '" & ChangeQuote(Trim(LAST_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (LAST_NAME = '" & ChangeQuote(Trim(LAST_NAME)) & "')"
         End If
      End If
   End If
      
   If Trim(CARD_NO) <> "" Then
      If WildCard(CARD_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(CARD_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(CARD_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CARD_NO = '" & ChangeQuote(Trim(CARD_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (CARD_NO = '" & ChangeQuote(Trim(CARD_NO)) & "')"
         End If
      End If
   End If
      
   If Trim(HOME_NO1) <> "" Then
      If WildCard(HOME_NO1, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(HOME_NO1,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(HOME_NO1,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (HOME_NO1 = '" & ChangeQuote(Trim(HOME_NO1)) & "')"
         Else
            WhereStr = WhereStr & "AND (HOME_NO1 = '" & ChangeQuote(Trim(HOME_NO1)) & "')"
         End If
      End If
   End If
      
   If Trim(MOO1) <> "" Then
      If WildCard(MOO1, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(MOO1,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(MOO1,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (MOO1 = '" & ChangeQuote(Trim(MOO1)) & "')"
         Else
            WhereStr = WhereStr & "AND (MOO1 = '" & ChangeQuote(Trim(MOO1)) & "')"
         End If
      End If
   End If
      
   If Trim(SOI1) <> "" Then
      If WildCard(SOI1, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(SOI1,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(SOI1,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (SOI1 = '" & ChangeQuote(Trim(SOI1)) & "')"
         Else
            WhereStr = WhereStr & "AND (SOI1 = '" & ChangeQuote(Trim(SOI1)) & "')"
         End If
      End If
   End If
      
   If Trim(ROAD1) <> "" Then
      If WildCard(ROAD1, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(ROAD1,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(ROAD1,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (ROAD1 = '" & ChangeQuote(Trim(ROAD1)) & "')"
         Else
            WhereStr = WhereStr & "AND (ROAD1 = '" & ChangeQuote(Trim(ROAD1)) & "')"
         End If
      End If
   End If
      
   If Trim(KWANG1) <> "" Then
      If WildCard(KWANG1, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(KWANG1,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(KWANG1,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (KWANG1 = '" & ChangeQuote(Trim(KWANG1)) & "')"
         Else
            WhereStr = WhereStr & "AND (KWANG1 = '" & ChangeQuote(Trim(KWANG1)) & "')"
         End If
      End If
   End If
      
   If Trim(KHATE1) <> "" Then
      If WildCard(KHATE1, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(KHATE1,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(KHATE1,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (KHATE1 = '" & ChangeQuote(Trim(KHATE1)) & "')"
         Else
            WhereStr = WhereStr & "AND (KHATE1 = '" & ChangeQuote(Trim(KHATE1)) & "')"
         End If
      End If
   End If
      
   If Trim(PROVINCE) <> "" Then
      If WildCard(PROVINCE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(PROVINCE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(PROVINCE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PROVINCE = '" & ChangeQuote(Trim(PROVINCE)) & "')"
         Else
            WhereStr = WhereStr & "AND (PROVINCE = '" & ChangeQuote(Trim(PROVINCE)) & "')"
         End If
      End If
   End If
      
    If REGISTER_DATE > -1 Then
      TempDate = Trim(DateToStringInt(REGISTER_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (REGISTER_DATE = '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (REGISTER_DATE = '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PATIENT_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PATIENT_ID "
   Else
      OrderByStr = "ORDER BY PATIENT_ID "
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
   PATIENT_ID = NVLI(Rs("PATIENT_ID"), -1)
   PATIENT_CODE = NVLS(Rs("PATIENT_CODE"), "")
   REGISTER_DATE = InternalDateToDate(NVLS(Rs("REGISTER_DATE"), ""))
   CHANNEL = NVLS(Rs("CHANNEL"), "")
   NAME = NVLS(Rs("NAME"), "")
   LAST_NAME = NVLS(Rs("LAST_NAME"), "")
   SEX = NVLS(Rs("SEX"), "")
   AGE = NVLD(Rs("AGE"), "")
   HOME_NO1 = NVLS(Rs("HOME_NO1"), "")
   MOO1 = NVLS(Rs("MOO1"), "")
   SOI1 = NVLS(Rs("SOI1"), "")
   ROAD1 = NVLS(Rs("ROAD1"), "")
   KWANG1 = NVLS(Rs("KWANG1"), "")
   KHATE1 = NVLS(Rs("KHATE1"), "")
   ZIPCODE1 = NVLS(Rs("ZIPCODE1"), "")
   PROVINCE = NVLS(Rs("PROVINCE"), "")
   CARD_NO = NVLS(Rs("CARD_NO"), "")
   PHONE1 = NVLS(Rs("PHONE1"), "")
   REFERENTOR1 = NVLS(Rs("REFERENTOR1"), "")
   PHONE2 = NVLS(Rs("PHONE2"), "")
   DOCUMENT_ALLOW1 = NVLS(Rs("DOCUMENT_ALLOW1"), "")
   NOTE1 = NVLS(Rs("NOTE1"), "")
   YEAR_KNOW = NVLI(Rs("YEAR_KNOW"), -1)
   REASON_KNOW = NVLS(Rs("REASON_KNOW"), "")
   FAMILY_STATUS = NVLI(Rs("FAMILY_STATUS"), -1)
   FAMILY_STATUS_OTH = NVLS(Rs("FAMILY_STATUS_OTH"), "")
   INHERIT = NVLS(Rs("INHERIT"), "")
   INHERIT_OTH = NVLS(Rs("INHERIT_OTH"), "")
   JOB_FLAG = NVLS(Rs("JOB_FLAG"), "")
   JOB_NAME = NVLS(Rs("JOB_NAME"), "")
   JOB_SALARY = NVLD(Rs("JOB_SALARY"), 0)
   JOB_OTH = NVLS(Rs("JOB_OTH"), "")
   MEDICAL_EXPENSE = NVLD(Rs("MEDICAL_EXPENSE"), 0)
   EXPENSE_CODE = NVLS(Rs("EXPENSE_CODE"), "")
   EXPENSE_OTH = NVLS(Rs("EXPENSE_OTH"), "")
   CD4_VALUE = NVLD(Rs("CD4_VALUE"), 0)
   CD4_DATE = InternalDateToDate(NVLS(Rs("CD4_DATE"), ""))
   VIRAL_ROAD = NVLD(Rs("VIRAL_ROAD"), 0)
   EQUIVALENCE = NVLD(Rs("EQUIVALENCE"), 0)
   EQUIVALENCE_DATE = InternalDateToDate(NVLS(Rs("EQUIVALENCE_DATE"), ""))
   WEIGHT = NVLD(Rs("WEIGHT"), 0)
   PRE_WEIGHT = NVLD(Rs("PRE_WEIGHT"), 0)
   HEIGHT = NVLD(Rs("HEIGHT"), 0)
   HEALT_OTH = NVLS(Rs("HEALT_OTH"), "")
   RELATIVE_CODE1 = NVLS(Rs("RELATIVE_CODE1"), "")
   RELATIVE_OTH1 = NVLS(Rs("RELATIVE_OTH1"), "")
   RELATIVE_CODE2 = NVLS(Rs("RELATIVE_CODE2"), "")
   RELATIVE_OTH2 = NVLS(Rs("RELATIVE_OTH2"), "")
   RELATIVE_CODE3 = NVLS(Rs("RELATIVE_CODE3"), "")
   RELATIVE_OTH3 = NVLS(Rs("RELATIVE_OTH3"), "")
   RELATIVE_CODE4 = NVLS(Rs("RELATIVE_CODE4"), "")
   RELATIVE_OTH4 = NVLS(Rs("RELATIVE_OTH4"), "")
   RELATIVE_CODE5 = NVLS(Rs("RELATIVE_CODE5"), "")
   RELATIVE_OTH5 = NVLS(Rs("RELATIVE_OTH5"), "")
   RELATIVE_CODE6 = NVLS(Rs("RELATIVE_CODE6"), "")
   RELATIVE_OTH6 = NVLS(Rs("RELATIVE_OTH6"), "")
   RELATIVE_CODE7 = NVLS(Rs("RELATIVE_CODE7"), "")
   RELATIVE_OTH7 = NVLS(Rs("RELATIVE_OTH7"), "")
   RELATIVE_CODE8 = NVLS(Rs("RELATIVE_CODE8"), "")
   RELATIVE_OTH8 = NVLS(Rs("RELATIVE_OTH8"), "")
   RELATIVE_CODE9 = NVLS(Rs("RELATIVE_CODE9"), "")
   RELATIVE_OTH9 = NVLS(Rs("RELATIVE_OTH9"), "")
   RELATIVE_CODE10 = NVLS(Rs("RELATIVE_CODE10"), "")
   RELATIVE_OTH10 = NVLS(Rs("RELATIVE_OTH10"), "")
   PSYCO_CODE1 = NVLS(Rs("PSYCO_CODE1"), "")
   PSYCO_OTH1 = NVLS(Rs("PSYCO_OTH1"), "")
   PSYCO_CODE2 = NVLS(Rs("PSYCO_CODE2"), "")
   PSYCO_OTH2 = NVLS(Rs("PSYCO_OTH2"), "")
   REPAIR_PLACE = NVLS(Rs("REPAIR_PLACE"), "")
   REPAIR_PERSON = NVLS(Rs("REPAIR_PERSON"), "")
   MEDICINE1 = NVLS(Rs("MEDICINE1"), "")
   MEDICINE2 = NVLS(Rs("MEDICINE2"), "")
   MEDICINE3 = NVLS(Rs("MEDICINE3"), "")
   MEDICINE4 = NVLS(Rs("MEDICINE4"), "")
   MEDICINE5 = NVLS(Rs("MEDICINE5"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("PATIENT_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO PATIENT "
      SQL1 = SQL1 & "( "
      SQL1 = SQL1 & "       PATIENT_ID,"
      SQL1 = SQL1 & "       PATIENT_CODE,"
      SQL1 = SQL1 & "       REGISTER_DATE,"
      SQL1 = SQL1 & "       CHANNEL,"
      SQL1 = SQL1 & "       NAME,"
      SQL1 = SQL1 & "       LAST_NAME,"
      SQL1 = SQL1 & "       SEX,"
      SQL1 = SQL1 & "       AGE,"
      SQL1 = SQL1 & "       HOME_NO1,"
      SQL1 = SQL1 & "       MOO1,"
      SQL1 = SQL1 & "       SOI1,"
      SQL1 = SQL1 & "       ROAD1,"
      SQL1 = SQL1 & "       KWANG1,"
      SQL1 = SQL1 & "       KHATE1,"
      SQL1 = SQL1 & "       ZIPCODE1,"
      SQL1 = SQL1 & "       PROVINCE,"
      SQL1 = SQL1 & "       CARD_NO,"
      SQL1 = SQL1 & "       PHONE1,"
      SQL1 = SQL1 & "       REFERENTOR1,"
      SQL1 = SQL1 & "       PHONE2,"
      SQL1 = SQL1 & "       DOCUMENT_ALLOW1,"
      SQL1 = SQL1 & "       NOTE1,"
      SQL1 = SQL1 & "       YEAR_KNOW,"
      SQL1 = SQL1 & "       REASON_KNOW,"
      SQL1 = SQL1 & "       FAMILY_STATUS,"
      SQL1 = SQL1 & "       FAMILY_STATUS_OTH,"
      SQL1 = SQL1 & "       INHERIT,"
      SQL1 = SQL1 & "       INHERIT_OTH,"
      SQL1 = SQL1 & "       JOB_FLAG,"
      SQL1 = SQL1 & "       JOB_NAME,"
      SQL1 = SQL1 & "       JOB_SALARY,"
      SQL1 = SQL1 & "       JOB_OTH,"
      SQL1 = SQL1 & "       MEDICAL_EXPENSE,"
      SQL1 = SQL1 & "       EXPENSE_CODE,"
      SQL1 = SQL1 & "       EXPENSE_OTH,"
      SQL1 = SQL1 & "       CD4_VALUE,"
      SQL1 = SQL1 & "       CD4_DATE,"
      SQL1 = SQL1 & "       VIRAL_ROAD,"
      SQL1 = SQL1 & "       EQUIVALENCE,"
      SQL1 = SQL1 & "       EQUIVALENCE_DATE,"
      SQL1 = SQL1 & "       WEIGHT,"
      SQL1 = SQL1 & "       PRE_WEIGHT,"
      SQL1 = SQL1 & "       HEIGHT,"
      SQL1 = SQL1 & "       HEALT_OTH,"
      SQL1 = SQL1 & "       RELATIVE_CODE1,"
      SQL1 = SQL1 & "       RELATIVE_OTH1,"
      SQL1 = SQL1 & "       RELATIVE_CODE2,"
      SQL1 = SQL1 & "       RELATIVE_OTH2,"
      SQL1 = SQL1 & "       RELATIVE_CODE3,"
      SQL1 = SQL1 & "       RELATIVE_OTH3,"
      SQL1 = SQL1 & "       RELATIVE_CODE4,"
      SQL1 = SQL1 & "       RELATIVE_OTH4,"
      SQL1 = SQL1 & "       RELATIVE_CODE5,"
      SQL1 = SQL1 & "       RELATIVE_OTH5,"
      SQL1 = SQL1 & "       RELATIVE_CODE6,"
      SQL1 = SQL1 & "       RELATIVE_OTH6,"
      SQL1 = SQL1 & "       RELATIVE_CODE7,"
      SQL1 = SQL1 & "       RELATIVE_OTH7,"
      SQL1 = SQL1 & "       RELATIVE_CODE8,"
      SQL1 = SQL1 & "       RELATIVE_OTH8,"
      SQL1 = SQL1 & "       RELATIVE_CODE9,"
      SQL1 = SQL1 & "       RELATIVE_OTH9,"
      SQL1 = SQL1 & "       RELATIVE_CODE10,"
      SQL1 = SQL1 & "       RELATIVE_OTH10,"
      SQL1 = SQL1 & "       PSYCO_CODE1,"
      SQL1 = SQL1 & "       PSYCO_OTH1,"
      SQL1 = SQL1 & "       PSYCO_CODE2,"
      SQL1 = SQL1 & "       PSYCO_OTH2,"
      SQL1 = SQL1 & "       REPAIR_PLACE,"
      SQL1 = SQL1 & "       REPAIR_PERSON,"
      SQL1 = SQL1 & "       MEDICINE1,"
      SQL1 = SQL1 & "       MEDICINE2,"
      SQL1 = SQL1 & "       MEDICINE3,"
      SQL1 = SQL1 & "       MEDICINE4,"
      SQL1 = SQL1 & "       MEDICINE5,"
      SQL1 = SQL1 & "       CREATE_DATE,"
      SQL1 = SQL1 & "       CREATE_BY,"
      SQL1 = SQL1 & "       MODIFY_DATE,"
      SQL1 = SQL1 & "       MODIFY_BY"
      SQL1 = SQL1 & ") VALUES "
      SQL1 = SQL1 & "( "
      SQL1 = SQL1 & "       '" & ID & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(PATIENT_CODE)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(DateToStringInt(REGISTER_DATE))) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(CHANNEL)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(NAME)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(LAST_NAME)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(SEX)) & "', "
      SQL1 = SQL1 & "       '" & AGE & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(HOME_NO1)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(MOO1)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(SOI1)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(ROAD1)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(KWANG1)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(KHATE1)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(ZIPCODE1)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(PROVINCE)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(CARD_NO)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(PHONE1)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(REFERENTOR1)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(PHONE2)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(DOCUMENT_ALLOW1)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(NOTE1)) & "', "
      SQL1 = SQL1 & "       '" & YEAR_KNOW & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(REASON_KNOW)) & "', "
      SQL1 = SQL1 & "       '" & FAMILY_STATUS & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(FAMILY_STATUS_OTH)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(INHERIT)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(INHERIT_OTH)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(JOB_FLAG)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(JOB_NAME)) & "', "
      SQL1 = SQL1 & "       '" & JOB_SALARY & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(JOB_OTH)) & "', "
      SQL1 = SQL1 & "       '" & MEDICAL_EXPENSE & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(EXPENSE_CODE)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(EXPENSE_OTH)) & "', "
      SQL1 = SQL1 & "       '" & CD4_VALUE & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(DateToStringInt(CD4_DATE))) & "', "
      SQL1 = SQL1 & "       '" & VIRAL_ROAD & "', "
      SQL1 = SQL1 & "       '" & EQUIVALENCE & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(DateToStringInt(EQUIVALENCE_DATE))) & "', "
      SQL1 = SQL1 & "       '" & WEIGHT & "', "
      SQL1 = SQL1 & "       '" & PRE_WEIGHT & "', "
      SQL1 = SQL1 & "       '" & HEIGHT & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(HEALT_OTH)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(RELATIVE_CODE1)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(RELATIVE_OTH1)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(RELATIVE_CODE2)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(RELATIVE_OTH2)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(RELATIVE_CODE3)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(RELATIVE_OTH3)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(RELATIVE_CODE4)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(RELATIVE_OTH4)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(RELATIVE_CODE5)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(RELATIVE_OTH5)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(RELATIVE_CODE6)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(RELATIVE_OTH6)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(RELATIVE_CODE7)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(RELATIVE_OTH7)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(RELATIVE_CODE8)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(RELATIVE_OTH8)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(RELATIVE_CODE9)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(RELATIVE_OTH9)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(RELATIVE_CODE10)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(RELATIVE_OTH10)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(PSYCO_CODE1)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(PSYCO_OTH1)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(PSYCO_CODE2)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(PSYCO_OTH2)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(REPAIR_PLACE)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(REPAIR_PERSON)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(MEDICINE1)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(MEDICINE2)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(MEDICINE3)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(MEDICINE4)) & "', "
      SQL1 = SQL1 & "       '" & Trim(ChangeQuote(MEDICINE5)) & "', "
      SQL1 = SQL1 & "       '" & ServerDateTime & "', "
      SQL1 = SQL1 & "       '" & glbUser.USER_ID & "', "
      SQL1 = SQL1 & "       '" & ServerDateTime & "', "
      SQL1 = SQL1 & "       '" & glbUser.USER_ID & "'"
      SQL1 = SQL1 & ")"
      PATIENT_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE PATIENT SET "
      SQL1 = SQL1 & "    PATIENT_ID = '" & PATIENT_ID & "', "
      SQL1 = SQL1 & "    PATIENT_CODE = '" & Trim(ChangeQuote(PATIENT_CODE)) & "', "
      SQL1 = SQL1 & "    REGISTER_DATE = '" & Trim(ChangeQuote(DateToStringInt(REGISTER_DATE))) & "', "
      SQL1 = SQL1 & "    CHANNEL = '" & Trim(ChangeQuote(CHANNEL)) & "', "
      SQL1 = SQL1 & "    NAME = '" & Trim(ChangeQuote(NAME)) & "', "
      SQL1 = SQL1 & "    LAST_NAME = '" & Trim(ChangeQuote(LAST_NAME)) & "', "
      SQL1 = SQL1 & "    SEX = '" & Trim(ChangeQuote(SEX)) & "', "
      SQL1 = SQL1 & "    AGE = '" & AGE & "', "
      SQL1 = SQL1 & "    HOME_NO1 = '" & Trim(ChangeQuote(HOME_NO1)) & "', "
      SQL1 = SQL1 & "    MOO1 = '" & Trim(ChangeQuote(MOO1)) & "', "
      SQL1 = SQL1 & "    SOI1 = '" & Trim(ChangeQuote(SOI1)) & "', "
      SQL1 = SQL1 & "    ROAD1 = '" & Trim(ChangeQuote(ROAD1)) & "', "
      SQL1 = SQL1 & "    KWANG1 = '" & Trim(ChangeQuote(KWANG1)) & "', "
      SQL1 = SQL1 & "    KHATE1 = '" & Trim(ChangeQuote(KHATE1)) & "', "
      SQL1 = SQL1 & "    ZIPCODE1 = '" & Trim(ChangeQuote(ZIPCODE1)) & "', "
      SQL1 = SQL1 & "    PROVINCE = '" & Trim(ChangeQuote(PROVINCE)) & "', "
      SQL1 = SQL1 & "    CARD_NO = '" & Trim(ChangeQuote(CARD_NO)) & "', "
      SQL1 = SQL1 & "    PHONE1 = '" & Trim(ChangeQuote(PHONE1)) & "', "
      SQL1 = SQL1 & "    REFERENTOR1 = '" & Trim(ChangeQuote(REFERENTOR1)) & "', "
      SQL1 = SQL1 & "    PHONE2 = '" & Trim(ChangeQuote(PHONE2)) & "', "
      SQL1 = SQL1 & "    DOCUMENT_ALLOW1 = '" & Trim(ChangeQuote(DOCUMENT_ALLOW1)) & "', "
      SQL1 = SQL1 & "    NOTE1 = '" & Trim(ChangeQuote(NOTE1)) & "', "
      SQL1 = SQL1 & "    YEAR_KNOW = '" & YEAR_KNOW & "', "
      SQL1 = SQL1 & "    REASON_KNOW = '" & Trim(ChangeQuote(REASON_KNOW)) & "', "
      SQL1 = SQL1 & "    FAMILY_STATUS = '" & FAMILY_STATUS & "', "
      SQL1 = SQL1 & "    FAMILY_STATUS_OTH = '" & Trim(ChangeQuote(FAMILY_STATUS_OTH)) & "', "
      SQL1 = SQL1 & "    INHERIT = '" & Trim(ChangeQuote(INHERIT)) & "', "
      SQL1 = SQL1 & "    INHERIT_OTH = '" & Trim(ChangeQuote(INHERIT_OTH)) & "', "
      SQL1 = SQL1 & "    JOB_FLAG = '" & Trim(ChangeQuote(JOB_FLAG)) & "', "
      SQL1 = SQL1 & "    JOB_NAME = '" & Trim(ChangeQuote(JOB_NAME)) & "', "
      SQL1 = SQL1 & "    JOB_SALARY = '" & JOB_SALARY & "', "
      SQL1 = SQL1 & "    JOB_OTH = '" & Trim(ChangeQuote(JOB_OTH)) & "', "
      SQL1 = SQL1 & "    MEDICAL_EXPENSE = '" & MEDICAL_EXPENSE & "', "
      SQL1 = SQL1 & "    EXPENSE_CODE = '" & Trim(ChangeQuote(EXPENSE_CODE)) & "', "
      SQL1 = SQL1 & "    EXPENSE_OTH = '" & Trim(ChangeQuote(EXPENSE_OTH)) & "', "
      SQL1 = SQL1 & "    CD4_VALUE = '" & CD4_VALUE & "', "
      SQL1 = SQL1 & "    CD4_DATE = '" & Trim(ChangeQuote(DateToStringInt(CD4_DATE))) & "', "
      SQL1 = SQL1 & "    VIRAL_ROAD = '" & VIRAL_ROAD & "', "
      SQL1 = SQL1 & "    EQUIVALENCE = '" & EQUIVALENCE & "', "
      SQL1 = SQL1 & "    EQUIVALENCE_DATE = '" & Trim(ChangeQuote(DateToStringInt(EQUIVALENCE_DATE))) & "', "
      SQL1 = SQL1 & "    WEIGHT = '" & WEIGHT & "', "
      SQL1 = SQL1 & "    PRE_WEIGHT = '" & PRE_WEIGHT & "', "
      SQL1 = SQL1 & "    HEIGHT = '" & HEIGHT & "', "
      SQL1 = SQL1 & "    HEALT_OTH = '" & Trim(ChangeQuote(HEALT_OTH)) & "', "
      SQL1 = SQL1 & "    RELATIVE_CODE1 = '" & Trim(ChangeQuote(RELATIVE_CODE1)) & "', "
      SQL1 = SQL1 & "    RELATIVE_OTH1 = '" & Trim(ChangeQuote(RELATIVE_OTH1)) & "', "
      SQL1 = SQL1 & "    RELATIVE_CODE2 = '" & Trim(ChangeQuote(RELATIVE_CODE2)) & "', "
      SQL1 = SQL1 & "    RELATIVE_OTH2 = '" & Trim(ChangeQuote(RELATIVE_OTH2)) & "', "
      SQL1 = SQL1 & "    RELATIVE_CODE3 = '" & Trim(ChangeQuote(RELATIVE_CODE3)) & "', "
      SQL1 = SQL1 & "    RELATIVE_OTH3 = '" & Trim(ChangeQuote(RELATIVE_OTH3)) & "', "
      SQL1 = SQL1 & "    RELATIVE_CODE4 = '" & Trim(ChangeQuote(RELATIVE_CODE4)) & "', "
      SQL1 = SQL1 & "    RELATIVE_OTH4 = '" & Trim(ChangeQuote(RELATIVE_OTH4)) & "', "
      SQL1 = SQL1 & "    RELATIVE_CODE5 = '" & Trim(ChangeQuote(RELATIVE_CODE5)) & "', "
      SQL1 = SQL1 & "    RELATIVE_OTH5 = '" & Trim(ChangeQuote(RELATIVE_OTH5)) & "', "
      SQL1 = SQL1 & "    RELATIVE_CODE6 = '" & Trim(ChangeQuote(RELATIVE_CODE6)) & "', "
      SQL1 = SQL1 & "    RELATIVE_OTH6 = '" & Trim(ChangeQuote(RELATIVE_OTH6)) & "', "
      SQL1 = SQL1 & "    RELATIVE_CODE7 = '" & Trim(ChangeQuote(RELATIVE_CODE7)) & "', "
      SQL1 = SQL1 & "    RELATIVE_OTH7 = '" & Trim(ChangeQuote(RELATIVE_OTH7)) & "', "
      SQL1 = SQL1 & "    RELATIVE_CODE8 = '" & Trim(ChangeQuote(RELATIVE_CODE8)) & "', "
      SQL1 = SQL1 & "    RELATIVE_OTH8 = '" & Trim(ChangeQuote(RELATIVE_OTH8)) & "', "
      SQL1 = SQL1 & "    RELATIVE_CODE9 = '" & Trim(ChangeQuote(RELATIVE_CODE9)) & "', "
      SQL1 = SQL1 & "    RELATIVE_OTH9 = '" & Trim(ChangeQuote(RELATIVE_OTH9)) & "', "
      SQL1 = SQL1 & "    RELATIVE_CODE10 = '" & Trim(ChangeQuote(RELATIVE_CODE10)) & "', "
      SQL1 = SQL1 & "    RELATIVE_OTH10 = '" & Trim(ChangeQuote(RELATIVE_OTH10)) & "', "
      SQL1 = SQL1 & "    PSYCO_CODE1 = '" & Trim(ChangeQuote(PSYCO_CODE1)) & "', "
      SQL1 = SQL1 & "    PSYCO_OTH1 = '" & Trim(ChangeQuote(PSYCO_OTH1)) & "', "
      SQL1 = SQL1 & "    PSYCO_CODE2 = '" & Trim(ChangeQuote(PSYCO_CODE2)) & "', "
      SQL1 = SQL1 & "    PSYCO_OTH2 = '" & Trim(ChangeQuote(PSYCO_OTH2)) & "', "
      SQL1 = SQL1 & "    REPAIR_PLACE = '" & Trim(ChangeQuote(REPAIR_PLACE)) & "', "
      SQL1 = SQL1 & "    REPAIR_PERSON = '" & Trim(ChangeQuote(REPAIR_PERSON)) & "', "
      SQL1 = SQL1 & "    MEDICINE1 = '" & Trim(ChangeQuote(MEDICINE1)) & "', "
      SQL1 = SQL1 & "    MEDICINE2 = '" & Trim(ChangeQuote(MEDICINE2)) & "', "
      SQL1 = SQL1 & "    MEDICINE3 = '" & Trim(ChangeQuote(MEDICINE3)) & "', "
      SQL1 = SQL1 & "    MEDICINE4 = '" & Trim(ChangeQuote(MEDICINE4)) & "', "
      SQL1 = SQL1 & "    MEDICINE5 = '" & Trim(ChangeQuote(MEDICINE5)) & "', "
      SQL1 = SQL1 & "    MODIFY_DATE = '" & ServerDateTime & "', "
      SQL1 = SQL1 & "    MODIFY_BY = " & glbUser.USER_ID & " "
      SQL1 = SQL1 & "WHERE PATIENT_ID = " & PATIENT_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM PATIENT WHERE PATIENT_ID = " & PATIENT_ID
      m_Conn.Execute (SQL1)
End Function

