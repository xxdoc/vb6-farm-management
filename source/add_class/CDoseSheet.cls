VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDoseSheet"
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
Public ExtraFlag As Long
' End of generic data member

Public DOSE_SHEET_ID As Long
Public RECORD_DATE As Date
Public PATIENT_ID As Long
Public WEIGHT As Double
Public KS_PERCENT As Double
Public CD4_VALUE As Double
Public VL_VALUE As Double
Public LOG10_VALUE As Double
Public NOTE1 As String
Public NOTE2 As String

Public PATIENT_CODE As String
Public REGISTER_DATE As Date
Public CHANNEL As String
Public NAME As String
Public LAST_NAME As String
Public HOME_NO1 As String
Public MOO1 As String
Public SOI1 As String
Public ROAD1 As String
Public KWANG1 As String
Public KHATE1 As String
Public ZIPCODE1 As String
Public PHONE1 As String
Public CARD_NO As String
Public PROVINCE As String

Public DoseSheetItems As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   REGISTER_DATE = -1
   Set DoseSheetItems = New Collection
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

   SelectStr1 = "SELECT BL.*, PT.*, BL.WEIGHT RECORD_WEIGHT, BL.NOTE1 RECORD_NOTE1, BL.NOTE2 RECORD_NOTE2, BL.CD4_VALUE RECORD_CD4 "
   
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM DOSE_SHEET BL, PATIENT PT "
   
   WhereStr = "WHERE (BL.PATIENT_ID = PT.PATIENT_ID) "
   
   If DOSE_SHEET_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DOSE_SHEET_ID = " & DOSE_SHEET_ID & ") "
      Else
         WhereStr = WhereStr & "AND (DOSE_SHEET_ID = " & DOSE_SHEET_ID & ") "
      End If
   End If
         
   '===
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
         WhereStr = "WHERE (RECORD_DATE = '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (RECORD_DATE = '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   '===
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY RECORD_DATE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY RECORD_DATE "
   Else
      OrderByStr = "ORDER BY DOSE_SHEET_ID "
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
   DOSE_SHEET_ID = NVLI(Rs("DOSE_SHEET_ID"), -1)
   NOTE1 = NVLS(Rs("NOTE1"), "")
   CD4_VALUE = NVLD(Rs("CD4_VALUE"), 0)
   PATIENT_ID = NVLI(Rs("PATIENT_ID"), 0)
   VL_VALUE = NVLD(Rs("VL_VALUE"), -1)
   RECORD_DATE = InternalDateToDate(NVLS(Rs("RECORD_DATE"), ""))
   WEIGHT = NVLD(Rs("WEIGHT"), -1)
   KS_PERCENT = NVLD(Rs("KS_PERCENT"), -1)
   NOTE2 = NVLS(Rs("NOTE2"), "")
   LOG10_VALUE = NVLD(Rs("LOG10_VALUE"), -1)
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

   TempID1 = PATIENT_ID
   If PATIENT_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = VL_VALUE
   If VL_VALUE <= 0 Then
      TempID2 = "NULL"
   End If
         
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("DOSE_SHEET_SEQ", ID, glbErrorLog)
               
      SQL1 = "INSERT INTO DOSE_SHEET " & _
                     "( " & _
                     "   DOSE_SHEET_ID, " & _
                     "   NOTE1, " & _
                     "   PATIENT_ID, " & _
                     "   VL_VALUE, " & _
                     "   RECORD_DATE, " & _
                     "   WEIGHT, " & _
                     "   KS_PERCENT, " & _
                     "   LOG10_VALUE, " & _
                     "   CD4_VALUE, " & _
                     "   NOTE2, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
                  SQL1 = SQL1 & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(NOTE1)) & "'," & _
                         TempID1 & "," & _
                         TempID2 & "," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(RECORD_DATE))) & "'," & _
                         WEIGHT & "," & _
                         KS_PERCENT & "," & _
                         LOG10_VALUE & "," & _
                         CD4_VALUE & "," & _
                         "'" & ChangeQuote(Trim(NOTE2)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         DOSE_SHEET_ID = ID
         
   Else 'Show edit
      SQL1 = "UPDATE DOSE_SHEET SET " & _
                     "   NOTE1 = '" & ChangeQuote(Trim(NOTE1)) & "', " & _
                     "    PATIENT_ID = " & TempID1 & ", " & _
                     "    VL_VALUE = " & TempID2 & ", " & _
                     "   RECORD_DATE = '" & ChangeQuote(Trim(DateToStringInt(RECORD_DATE))) & "', " & _
                     "    WEIGHT = " & WEIGHT & ", " & _
                     "    KS_PERCENT = " & KS_PERCENT & ", " & _
                     "    LOG10_VALUE = " & LOG10_VALUE & ", " & _
                     "    CD4_VALUE = " & CD4_VALUE & ", " & _
                     "   NOTE2 = '" & ChangeQuote(Trim(NOTE2)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE DOSE_SHEET_ID = " & DOSE_SHEET_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
            
      SQL1 = "DELETE FROM DSHEET_ITEM WHERE DOSE_SHEET_ID = " & DOSE_SHEET_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM DOSE_SHEET WHERE DOSE_SHEET_ID = " & DOSE_SHEET_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set DoseSheetItems = Nothing
End Sub
