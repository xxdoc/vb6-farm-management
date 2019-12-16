VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Conn As ADODB.Connection

Public OLD_EMP_ID As Long

Public EMP_ID As Long
Public EMP_CODE As String
Public Name As String
Public LASTNAME As String
Public SEX_ID As Long
Public WEIGHT As Double
Public HEIGHT As Double
Public RACE As String
Public NATIONALITY As String
Public RELIGIOUS_ID As Long
Public MARITAL_ID As Long
Public MILITALY_ID As Long
Public BIRTH_DATE As Date
Public ENTRY_DATE As Date
Public DISCEASE As String
Public MOBILE_PHONE As String
Public EMAIL_ADDRESS As String
Public PASS_DATE As Date
Public RESIGN_DATE As Date
Public WORK_STATUS_ID As Long
Public CREMATION_DATE As Date
Public ALLOWANCE_DATE As Date
Public CURRENT_POSITION As Long
Public CURRENT_SECTION As Long
Public CURRENT_ORGANIZE As Long
Public CURRENT_COMPANY As Long
Public CURRENT_SALARY As Double
Public BLOOD_GROUP As Long
Public RESIGN_REASON As Long
Public DUE_DATE As Date
Public APPEND_DATE As Date
Public PASS_STATUS As String
Public CREDIT_AMOUNT As Double
Public BANK_BRANCH As String
Public BANK_ACCOUNT As String
Public BANK_ID As Long
Public EXTERNAL_FLAG As String
Public ASSURE_AMOUNT As Double
Public OT_RATE As Double
Public CARD_NO As String

Public POSITION_ID As Long
Public ORGANIZE_ID As Long
Public SECTION_ID As Long
Public POSITION_NAME As String
Public ORGANIZE_NAME As String
Public SECTION_NAME As String

Public KEY_ID As Long
Public KEY_LOOKUP As String

Public OrderBy As Long
Public OrderType As Long
Public QueryFlag As Long
Public AddEditMode As SHOW_MODE_TYPE
Public Flag As String

Public Contacts As Collection
Public Cards As Collection
Public EmpWorked As Collection
Public EmpEducation As Collection
Public EmpDependency As Collection
Public EmpChild As Collection
Public EmpHealty As Collection

Public EName As cName
Public EmpName As CEmployeeName

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   
   QueryFlag = -1
   Set Contacts = New Collection
   Set Cards = New Collection
   Set EName = New cName
   Set EmpName = New CEmployeeName
   Set EmpWorked = New Collection
   Set EmpEducation = New Collection
   Set EmpDependency = New Collection
   Set EmpChild = New Collection
   Set EmpHealty = New Collection
End Sub

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim Temp1 As String
Dim Temp2 As String
Dim Temp3 As String
Dim Temp4 As String
Dim Temp5 As String
Dim Temp6 As String
Dim Temp7 As String
Dim Temp8 As String

   Temp1 = RESIGN_REASON
   If RESIGN_REASON <= 0 Then
      Temp1 = "NULL "
   End If
   
   Temp2 = RELIGIOUS_ID
   If RELIGIOUS_ID <= 0 Then
      Temp2 = "NULL "
   End If
   
   Temp3 = MARITAL_ID
   If MARITAL_ID <= 0 Then
      Temp3 = "NULL "
   End If
   
   Temp4 = MILITALY_ID
   If MILITALY_ID <= 0 Then
      Temp4 = "NULL "
   End If
   
   Temp5 = BANK_ID
   If BANK_ID <= 0 Then
      Temp5 = "NULL "
   End If
   
   Temp6 = SEX_ID
   If SEX_ID <= 0 Then
      Temp6 = "NULL "
   End If
   
   Temp7 = BLOOD_GROUP
   If BLOOD_GROUP <= 0 Then
      Temp7 = "NULL "
   End If
   
   Temp8 = CURRENT_POSITION
   If CURRENT_POSITION <= 0 Then
      Temp8 = "NULL "
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("EMPLOYEE_SEQ", ID, glbErrorLog)
      If OLD_EMP_ID > 0 Then
         ID = OLD_EMP_ID
      End If
      SQL1 = "INSERT INTO EMPLOYEE " & _
                     "( " & _
                     "   EMP_ID, " & _
                     "   EMP_CODE, " & _
                     "   SEX_ID, " & _
                     "   WEIGHT, " & _
                     "   HEIGHT, " & _
                     "   RACE, " & _
                     "   NATIONALITY, " & _
                     "   RELIGIOUS_ID, " & _
                     "   MARITAL_ID, " & _
                     "   MILITALY_ID, " & _
                     "   BIRTH_DATE, " & _
                     "   ENTRY_DATE, " & _
                     "   DISCEASE, " & _
                     "   MOBILE_PHONE, " & _
                     "   EMAIL_ADDRESS, " & _
                     "   RESIGN_DATE, " & _
                     "   CREMATION_DATE, " & _
                     "   DUE_DATE, " & _
                     "   APPEND_DATE, " & _
                     "   ALLOWANCE_DATE, " & _
                     "   PASS_DATE, " & _
                     "   PASS_STATUS, "
               SQL1 = SQL1 & _
                     "   WORK_STATUS_ID, " & _
                     "   BLOOD_GROUP, " & _
                     "   RESIGN_REASON, " & _
                     "   CURRENT_POSITION, " & _
                     "   CURRENT_SECTION, " & _
                     "   CURRENT_ORGANIZE, " & _
                     "   CURRENT_SALARY, " & _
                     "   CREDIT_AMOUNT, " & _
                     "   BANK_BRANCH, " & _
                     "   BANK_ACCOUNT, " & _
                     "   BANK_ID, " & _
                     "   EXTERNAL_FLAG, " & _
                     "   ASSURE_AMOUNT, " & _
                     "   OT_RATE, " & _
                     "   CARD_NO, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
               SQL1 = SQL1 & "( " & _
                         ID & "," & _
                        "'" & ChangeQuote(Trim(EMP_CODE)) & "', " & _
                        Temp6 & ", " & _
                        WEIGHT & ", " & _
                        HEIGHT & ", " & _
                        "'" & ChangeQuote(Trim(RACE)) & "', " & _
                        "'" & ChangeQuote(Trim(NATIONALITY)) & "', " & _
                        Temp2 & ", " & _
                        Temp3 & ", " & _
                        Temp4 & ", " & _
                        "'" & ChangeQuote(DateToStringInt(BIRTH_DATE)) & "', " & _
                        "'" & ChangeQuote(DateToStringInt(ENTRY_DATE)) & "', " & _
                        "'" & ChangeQuote(Trim(DISCEASE)) & "', " & _
                        "'" & ChangeQuote(Trim(MOBILE_PHONE)) & "', " & _
                        "'" & ChangeQuote(Trim(EMAIL_ADDRESS)) & "', " & _
                        "'" & ChangeQuote(DateToStringInt(RESIGN_DATE)) & "', " & _
                        "'" & ChangeQuote(DateToStringInt(CREMATION_DATE)) & "', " & _
                        "'" & ChangeQuote(DateToStringInt(DUE_DATE)) & "', " & _
                        "'" & ChangeQuote(DateToStringInt(APPEND_DATE)) & "', " & _
                        "'" & ChangeQuote(DateToStringInt(ALLOWANCE_DATE)) & "', " & _
                        "'" & ChangeQuote(DateToStringInt(PASS_DATE)) & "', " & _
                        "'" & ChangeQuote(PASS_STATUS) & "', "
               SQL1 = SQL1 & _
                        WORK_STATUS_ID & ", " & _
                        Temp7 & ", " & _
                        Temp1 & ", " & _
                        Temp8 & ", " & _
                        CURRENT_SECTION & ", " & _
                        CURRENT_ORGANIZE & ", " & _
                        CURRENT_SALARY & ", " & _
                        CREDIT_AMOUNT & ", " & _
                        "'" & ChangeQuote(Trim(BANK_BRANCH)) & "', " & _
                        "'" & ChangeQuote(Trim(BANK_ACCOUNT)) & "', " & _
                        Temp5 & ", " & _
                        "'" & ChangeQuote(EXTERNAL_FLAG) & "', " & _
                        ASSURE_AMOUNT & ", " & _
                        OT_RATE & ", " & _
                        "'" & ChangeQuote(Trim(CARD_NO)) & "', " & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
                  EMP_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE EMPLOYEE SET " & _
                     "   EMP_CODE = '" & ChangeQuote(Trim(EMP_CODE)) & "', " & _
                     "   SEX_ID =  " & Temp6 & ", " & _
                     "   WEIGHT =  " & WEIGHT & ", " & _
                     "   HEIGHT =  " & HEIGHT & ", " & _
                     "   RACE = '" & ChangeQuote(Trim(RACE)) & "', " & _
                     "   NATIONALITY = '" & ChangeQuote(Trim(NATIONALITY)) & "', " & _
                     "   RELIGIOUS_ID =  " & Temp2 & ", " & _
                     "   MARITAL_ID =  " & Temp3 & ", " & _
                     "   MILITALY_ID =  " & Temp4 & ", " & _
                     "   BIRTH_DATE = '" & ChangeQuote(Trim(DateToStringInt(BIRTH_DATE))) & "', " & _
                     "   ENTRY_DATE = '" & ChangeQuote(Trim(DateToStringInt(ENTRY_DATE))) & "', " & _
                     "   DISCEASE = '" & ChangeQuote(Trim(DISCEASE)) & "', " & _
                     "   MOBILE_PHONE = '" & ChangeQuote(Trim(MOBILE_PHONE)) & "', " & _
                     "   EMAIL_ADDRESS = '" & ChangeQuote(Trim(EMAIL_ADDRESS)) & "', " & _
                     "   RESIGN_DATE = '" & ChangeQuote(Trim(DateToStringInt(RESIGN_DATE))) & "', " & _
                     "   CREMATION_DATE = '" & ChangeQuote(Trim(DateToStringInt(CREMATION_DATE))) & "', " & _
                     "   DUE_DATE = '" & ChangeQuote(Trim(DateToStringInt(DUE_DATE))) & "', " & _
                     "   APPEND_DATE = '" & ChangeQuote(Trim(DateToStringInt(APPEND_DATE))) & "', " & _
                     "   ALLOWANCE_DATE = '" & ChangeQuote(Trim(DateToStringInt(ALLOWANCE_DATE))) & "', " & _
                     "   PASS_DATE = '" & ChangeQuote(Trim(DateToStringInt(PASS_DATE))) & "', " & _
                     "   PASS_STATUS = '" & ChangeQuote(Trim(PASS_STATUS)) & "', "
         SQL1 = SQL1 & _
                     "   WORK_STATUS_ID =  " & WORK_STATUS_ID & ", " & _
                     "   BLOOD_GROUP =  " & Temp7 & ", " & _
                     "   RESIGN_REASON =  " & Temp1 & ", " & _
                     "   CURRENT_POSITION =  " & Temp8 & ", " & _
                     "   CURRENT_SECTION =  " & CURRENT_SECTION & ", " & _
                     "   CURRENT_ORGANIZE =  " & CURRENT_SECTION & ", " & _
                     "   CURRENT_SALARY =  " & CURRENT_SALARY & ", " & _
                     "   CREDIT_AMOUNT =  " & CREDIT_AMOUNT & ", " & _
                     "   BANK_BRANCH = '" & ChangeQuote(Trim(BANK_BRANCH)) & "', " & _
                     "   BANK_ACCOUNT = '" & ChangeQuote(Trim(BANK_ACCOUNT)) & "', " & _
                     "   BANK_ID =  " & Temp5 & ", " & _
                     "   EXTERNAL_FLAG = '" & ChangeQuote(Trim(EXTERNAL_FLAG)) & "', " & _
                     "   ASSURE_AMOUNT =  " & ASSURE_AMOUNT & ", " & _
                     "   OT_RATE =  " & OT_RATE & ", " & _
                     "   CARD_NO = '" & ChangeQuote(Trim(CARD_NO)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE EMP_ID = " & EMP_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM EMPLOYEE_NAME WHERE EMP_ID = " & EMP_ID
      m_Conn.Execute (SQL1)

'      SQL1 = "DELETE FROM EMPLOYEE_ADDRESS WHERE EMP_ID = " & EMP_ID
'      m_Conn.Execute (SQL1)
'
'      SQL1 = "DELETE FROM EMPLOYEE_PROOF WHERE EMP_ID = " & EMP_ID
'      m_Conn.Execute (SQL1)
'
'      SQL1 = "DELETE FROM EMP_WORKED WHERE EMP_ID = " & EMP_ID
'      m_Conn.Execute (SQL1)
'
'      SQL1 = "DELETE FROM EMP_EDUCATION WHERE EMP_ID = " & EMP_ID
'      m_Conn.Execute (SQL1)
'
'      SQL1 = "DELETE FROM EMP_DEPENDENCY WHERE EMP_ID = " & EMP_ID
'      m_Conn.Execute (SQL1)
'
'      SQL1 = "DELETE FROM EMP_CHILD WHERE EMP_ID = " & EMP_ID
'      m_Conn.Execute (SQL1)
'
'      SQL1 = "DELETE FROM EMP_HEALTY WHERE EMP_ID = " & EMP_ID
'      m_Conn.Execute (SQL1)

      SQL1 = "DELETE FROM EMPLOYEE WHERE EMP_ID = " & EMP_ID
      m_Conn.Execute (SQL1)
End Function

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

   SelectStr1 = "SELECT UG.*, RG.RELIGIOUS_NAME, SX.SEX_NAME, NM.*, EP.POSITION_NAME, EP.POSITION_DESC, WS.WORK_NAME, BG.BLOODGRP_NAME  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM EMPLOYEE UG "
   FromStr = FromStr & "LEFT OUTER JOIN RELIGIOUS_DATA RG ON (UG.RELIGIOUS_ID = RG.RELIGIOUS_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN SEX_DATA SX ON (UG.SEX_ID = SX.SEX_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (UG.EMP_ID = EN.EMP_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (EN.NAME_ID = NM.NAME_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN EMP_POSITION EP ON (UG.CURRENT_POSITION = EP.POSITION_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN WORK_STATUS WS ON (UG.WORK_STATUS_ID = WORK_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN BLOOD_GROUP BG ON (UG.BLOOD_GROUP = BG.BLOODGRP_ID) "
   
   WhereStr = ""
   If EMP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UG.EMP_ID = " & EMP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (UG.EMP_ID = " & EMP_ID & ") "
      End If
   End If
   
   If WORK_STATUS_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (WORK_STATUS_ID = " & WORK_STATUS_ID & ") "
      Else
         WhereStr = WhereStr & "AND (WORK_STATUS_ID = " & WORK_STATUS_ID & ") "
      End If
   End If
   
   If SEX_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SEX_ID = " & SEX_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SEX_ID = " & SEX_ID & ") "
      End If
   End If
   
   If CURRENT_POSITION > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CURRENT_POSITION = " & CURRENT_POSITION & ") "
      Else
         WhereStr = WhereStr & "AND (CURRENT_POSITION = " & CURRENT_POSITION & ") "
      End If
   End If
   
   If Trim(EMP_CODE) <> "" Then
      If WildCard(EMP_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(EMP_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(EMP_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (EMP_CODE = '" & ChangeQuote(Trim(EMP_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (EMP_CODE = '" & ChangeQuote(Trim(EMP_CODE)) & "')"
         End If
      End If
   End If
   
   If Trim(Name) <> "" Then
      If WildCard(Name, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(LONG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (LONG_NAME = '" & ChangeQuote(Trim(Name)) & "')"
         Else
            WhereStr = WhereStr & "AND (LONG_NAME = '" & ChangeQuote(Trim(Name)) & "')"
         End If
      End If
   End If
   
   If Trim(LASTNAME) <> "" Then
      If WildCard(LASTNAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(LAST_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(LAST_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (LAST_NAME = '" & ChangeQuote(Trim(LASTNAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (LAST_NAME = '" & ChangeQuote(Trim(LASTNAME)) & "')"
         End If
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY EMP_CODE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY LONG_NAME "
   ElseIf OrderBy = 3 Then
      OrderByStr = "ORDER BY LAST_NAME "
   ElseIf OrderBy = 4 Then
      OrderByStr = "ORDER BY POSITION_NAME "
   Else
      OrderByStr = "ORDER BY UG.EMP_ID "
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

Public Function QueryLegacyData(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT UG.*, 'N/A' EMP_CODE, UG.EMP_NAME LONG_NAME, UG.EMP_SURNAME LAST_NAME, RG.POSITION_NAME  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM EMPLOYEE_DATA UG "
   FromStr = FromStr & "LEFT OUTER JOIN POSITION_DATA RG ON (UG.EMP_POSITION_ID = RG.POSITION_ID) "
   
   WhereStr = ""
   If EMP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EMP_ID = " & EMP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (EMP_ID = " & EMP_ID & ") "
      End If
   End If
   
'   If WORK_STATUS_ID > 0 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (WORK_STATUS_ID = " & WORK_STATUS_ID & ") "
'      Else
'         WhereStr = WhereStr & "AND (WORK_STATUS_ID = " & WORK_STATUS_ID & ") "
'      End If
'   End If
'
'   If SEX_ID > 0 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (SEX_ID = " & SEX_ID & ") "
'      Else
'         WhereStr = WhereStr & "AND (SEX_ID = " & SEX_ID & ") "
'      End If
'   End If
'
   If CURRENT_POSITION > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (EMP_POSITION_ID = " & CURRENT_POSITION & ") "
      Else
         WhereStr = WhereStr & "AND (EMP_POSITION_ID = " & CURRENT_POSITION & ") "
      End If
   End If
   
   If Trim(EMP_CODE) <> "" Then
      If WildCard(EMP_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(EMP_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(EMP_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (EMP_CODE = '" & ChangeQuote(Trim(EMP_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (EMP_CODE = '" & ChangeQuote(Trim(EMP_CODE)) & "')"
         End If
      End If
   End If
   
   If Trim(Name) <> "" Then
      If WildCard(Name, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(EMP_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(EMP_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (EMP_NAME = '" & ChangeQuote(Trim(Name)) & "')"
         Else
            WhereStr = WhereStr & "AND (EMP_NAME = '" & ChangeQuote(Trim(Name)) & "')"
         End If
      End If
   End If
   
   If Trim(LASTNAME) <> "" Then
      If WildCard(LASTNAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(EMP_SURNAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(EMP_SURNAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (EMP_SURNAME = '" & ChangeQuote(Trim(LASTNAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (EMP_SURNAME = '" & ChangeQuote(Trim(LASTNAME)) & "')"
         End If
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY EMP_CODE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY EMP_NAME "
   ElseIf OrderBy = 3 Then
      OrderByStr = "ORDER BY EMP_SURNAME "
   Else
      OrderByStr = "ORDER BY EMP_ID "
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

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   EMP_ID = NVLI(Rs("EMP_ID"), -1)
   EMP_CODE = NVLS(Rs("EMP_CODE"), "")
   SEX_ID = NVLI(Rs("SEX_ID"), -1)
   WEIGHT = NVLD(Rs("WEIGHT"), 0)
   HEIGHT = NVLD(Rs("HEIGHT"), 0)
   RACE = NVLS(Rs("RACE"), "")
   NATIONALITY = NVLS(Rs("NATIONALITY"), "")
   RELIGIOUS_ID = NVLI(Rs("RELIGIOUS_ID"), -1)
   MARITAL_ID = NVLI(Rs("MARITAL_ID"), -1)
   MILITALY_ID = NVLI(Rs("MILITALY_ID"), -1)
   BIRTH_DATE = InternalDateToDate(NVLS(Rs("BIRTH_DATE"), ""))
   ENTRY_DATE = InternalDateToDate(NVLS(Rs("ENTRY_DATE"), ""))
   DISCEASE = NVLS(Rs("DISCEASE"), "")
   MOBILE_PHONE = NVLS(Rs("MOBILE_PHONE"), "")
   EMAIL_ADDRESS = NVLS(Rs("EMAIL_ADDRESS"), "")
   RESIGN_DATE = InternalDateToDate(NVLS(Rs("RESIGN_DATE"), ""))
   CREMATION_DATE = InternalDateToDate(NVLS(Rs("CREMATION_DATE"), ""))
   DUE_DATE = InternalDateToDate(NVLS(Rs("DUE_DATE"), ""))
   APPEND_DATE = InternalDateToDate(NVLS(Rs("APPEND_DATE"), ""))
   ALLOWANCE_DATE = InternalDateToDate(NVLS(Rs("ALLOWANCE_DATE"), ""))
   PASS_DATE = InternalDateToDate(NVLS(Rs("PASS_DATE"), ""))
   PASS_STATUS = NVLS(Rs("PASS_STATUS"), "")
   WORK_STATUS_ID = NVLI(Rs("WORK_STATUS_ID"), -1)
   BLOOD_GROUP = NVLI(Rs("BLOOD_GROUP"), -1)
   RESIGN_REASON = NVLI(Rs("RESIGN_REASON"), -1)
   CURRENT_POSITION = NVLI(Rs("CURRENT_POSITION"), -1)
   CURRENT_SECTION = NVLI(Rs("CURRENT_SECTION"), -1)
   CURRENT_ORGANIZE = NVLI(Rs("CURRENT_ORGANIZE"), -1)
   CURRENT_SALARY = NVLD(Rs("CURRENT_SALARY"), -1)
   BANK_BRANCH = NVLS(Rs("BANK_BRANCH"), "")
   BANK_ACCOUNT = NVLS(Rs("BANK_ACCOUNT"), "")
   BANK_ID = NVLI(Rs("BANK_ID"), -1)
   ASSURE_AMOUNT = NVLD(Rs("ASSURE_AMOUNT"), -1)
   OT_RATE = NVLD(Rs("OT_RATE"), 0)
   CARD_NO = NVLS(Rs("CARD_NO"), "")
   
   Name = NVLS(Rs("LONG_NAME"), "")
   LASTNAME = NVLS(Rs("LAST_NAME"), "")
   POSITION_NAME = NVLS(Rs("POSITION_DESC"), "")
   
   KEY_ID = EMP_ID
   KEY_LOOKUP = EMP_CODE
End Function

Private Sub Class_Terminate()
   Set EmpWorked = Nothing
   Set EmpEducation = Nothing
   Set EmpDependency = Nothing
   Set EmpChild = Nothing
   Set EmpHealty = Nothing
End Sub


