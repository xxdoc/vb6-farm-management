VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDSheetItem"
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

Public DSHEET_ITEM_ID As Long
Public DOSE_SHEET_ID As Long
Public DRUG_ID As Long
Public TOTAL_AMOUNT As Double
Public DOSE_AMOUNT As Double
Public PERIOD_AMOUNT As Long
Public PERIOD_DESC As String
Public EXTRA_DRUG_NAME As String

Public DRUG_NAME As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
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

   SelectStr1 = "SELECT BL.*, DR.* "
   
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM DSHEET_ITEM BL "
   FromStr = FromStr & "LEFT OUTER JOIN DRUG DR ON (BL.DRUG_ID = DR.DRUG_ID) "
   
   WhereStr = ""
   
   If DSHEET_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DSHEET_ITEM_ID = " & DSHEET_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (DSHEET_ITEM_ID = " & DSHEET_ITEM_ID & ") "
      End If
   End If
      
   If DOSE_SHEET_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DOSE_SHEET_ID = " & DOSE_SHEET_ID & ") "
      Else
         WhereStr = WhereStr & "AND (DOSE_SHEET_ID = " & DOSE_SHEET_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY DRUG_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY DRUG_ID "
   Else
      OrderByStr = "ORDER BY DRUG_GROUP_ID, DRUG_ID "
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

Public Function QueryData2(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT DISTINCT(PERIOD_DESC) "
   
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM DSHEET_ITEM BL "
   
   WhereStr = ""
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PERIOD_DESC "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PERIOD_DESC "
   Else
      OrderByStr = "ORDER BY PERIOD_DESC "
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Rs As ADODB.Recordset)
   DSHEET_ITEM_ID = NVLI(Rs("DSHEET_ITEM_ID"), -1)
   PERIOD_DESC = NVLS(Rs("PERIOD_DESC"), "")
   PERIOD_AMOUNT = NVLD(Rs("PERIOD_AMOUNT"), 0)
   DOSE_SHEET_ID = NVLI(Rs("DOSE_SHEET_ID"), 0)
   DRUG_ID = NVLI(Rs("DRUG_ID"), -1)
   TOTAL_AMOUNT = NVLD(Rs("TOTAL_AMOUNT"), -1)
   DOSE_AMOUNT = NVLD(Rs("DOSE_AMOUNT"), -1)
   EXTRA_DRUG_NAME = NVLS(Rs("EXTRA_DRUG_NAME"), "")
   DRUG_NAME = NVLS(Rs("DRUG_NAME"), "")
End Function

Public Function PopulateFromRS2(Rs As ADODB.Recordset)
   PERIOD_DESC = NVLS(Rs("PERIOD_DESC"), "")
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

   TempID1 = DOSE_SHEET_ID
   If DOSE_SHEET_ID <= 0 Then
      TempID1 = "NULL"
   End If
            
   TempID2 = DRUG_ID
   If DRUG_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("DSHEET_ITEM_SEQ", ID, glbErrorLog)
               
      SQL1 = "INSERT INTO DSHEET_ITEM " & _
                     "( " & _
                     "   DSHEET_ITEM_ID, " & _
                     "   PERIOD_DESC, " & _
                     "   DOSE_SHEET_ID, " & _
                     "   DRUG_ID, " & _
                     "   TOTAL_AMOUNT, " & _
                     "   DOSE_AMOUNT, " & _
                     "   PERIOD_AMOUNT, " & _
                     "   EXTRA_DRUG_NAME, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
                  SQL1 = SQL1 & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(PERIOD_DESC)) & "'," & _
                         TempID1 & "," & _
                         TempID2 & "," & _
                         TOTAL_AMOUNT & "," & _
                         DOSE_AMOUNT & "," & _
                         PERIOD_AMOUNT & "," & _
                         "'" & ChangeQuote(Trim(EXTRA_DRUG_NAME)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         DSHEET_ITEM_ID = ID
         
   Else 'Show edit
      SQL1 = "UPDATE DSHEET_ITEM SET " & _
                     "   PERIOD_DESC = '" & ChangeQuote(Trim(PERIOD_DESC)) & "', " & _
                     "    DOSE_SHEET_ID = " & TempID1 & ", " & _
                     "   DRUG_ID = " & TempID2 & ", " & _
                     "    TOTAL_AMOUNT = " & TOTAL_AMOUNT & ", " & _
                     "    DOSE_AMOUNT = " & DOSE_AMOUNT & ", " & _
                     "    PERIOD_AMOUNT = " & PERIOD_AMOUNT & ", " & _
                     "   EXTRA_DRUG_NAME = '" & ChangeQuote(Trim(EXTRA_DRUG_NAME)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE DSHEET_ITEM_ID = " & DSHEET_ITEM_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
            
      SQL1 = "DELETE FROM DSHEET_ITEM WHERE DSHEET_ITEM_ID = " & DSHEET_ITEM_ID
      m_Conn.Execute (SQL1)
End Function

