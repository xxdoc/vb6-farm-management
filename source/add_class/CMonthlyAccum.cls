VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMonthlyAccum"
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

Public MONTHLY_ACCUM_ID As Long
Public DOCUMENT_DATE As Date
Public PART_ITEM_ID As Long
Public LOCATION_ID As Long
Public IMPORT_AMOUNT As Double
Public EXPORT_AMOUNT As Double
Public BALANCE_AMOUNT1 As Double
Public BALANCE_AMOUNT2 As Double
Public SIMULATE_FLAG As String
Public BATCH_ID As Long
Public YYYYMM As String
Public BIRTH_AMOUNT As Double
Public SELL_AMOUNT As Double
Public STATUS_IN_AMOUNT As Double
Public STATUS_OUT_AMOUNT As Double

Public ADJUST_IN_AMOUNT As Double
Public ADJUST_OUT_AMOUNT As Double

Public GL_IN As Double
Public GL_OUT As Double

Public BUY_AMOUNT As Double

Public FROM_DATE As Date
Public TO_DATE As Date
Public TO_DATE1 As Date
Public MASTER_FLAG As String
Public PIG_FLAG As String
Public LOCATION_TYPE As Long
Public PART_TYPE As Long
 Public PIG_TYPE As String
 Public PART_NO As String
 Public PART_DESC As String
 
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

   SIMULATE_FLAG = glbUser.SIMULATE_FLAG
   
   If Ind = 1 Then
      SelectStr1 = "SELECT CM.*  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MONTHLY_ACCUM CM "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT DISTINCT PI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, PI.PIG_TYPE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MONTHLY_ACCUM CM "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CM.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = ""
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT CM.* "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MONTHLY_ACCUM CM "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CM.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = ""
   ElseIf Ind = 4 Then
      SelectStr1 = "SELECT CM.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, CM.YYYYMM, SUM(CM.IMPORT_AMOUNT) IMPORT_AMOUNT, SUM(CM.EXPORT_AMOUNT) EXPORT_AMOUNT "
      SelectStr1 = SelectStr1 & ", SUM(CM.BALANCE_AMOUNT1) BALANCE_AMOUNT1, SUM(CM.BALANCE_AMOUNT2) BALANCE_AMOUNT2 "
      SelectStr1 = SelectStr1 & ", SUM(CM.BIRTH_AMOUNT) BIRTH_AMOUNT, SUM(CM.BUY_AMOUNT) BUY_AMOUNT "
      SelectStr1 = SelectStr1 & ", SUM(CM.SELL_AMOUNT) SELL_AMOUNT, SUM(CM.STATUS_IN_AMOUNT) STATUS_IN_AMOUNT "
      SelectStr1 = SelectStr1 & ", SUM(CM.STATUS_OUT_AMOUNT) STATUS_OUT_AMOUNT "
      SelectStr1 = SelectStr1 & ", SUM(CM.ADJUST_IN_AMOUNT) ADJUST_IN_AMOUNT "
      SelectStr1 = SelectStr1 & ", SUM(CM.ADJUST_OUT_AMOUNT) ADJUST_OUT_AMOUNT "
      SelectStr1 = SelectStr1 & ", SUM(CM.GL_IN) GL_IN "
      SelectStr1 = SelectStr1 & ", SUM(CM.GL_OUT) GL_OUT "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM MONTHLY_ACCUM CM "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (CM.PART_ITEM_ID = PI.PART_ITEM_ID) "
      
      WhereStr = ""
   End If
   
   If MONTHLY_ACCUM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MONTHLY_ACCUM_ID = " & MONTHLY_ACCUM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (MONTHLY_ACCUM_ID = " & MONTHLY_ACCUM_ID & ") "
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
         WhereStr = "WHERE (LOCATION_ID = " & LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LOCATION_ID = " & LOCATION_ID & ") "
      End If
   End If
   
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
      
   If Len(PIG_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PIG_FLAG = '" & PIG_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (PIG_FLAG = '" & PIG_FLAG & "') "
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
   If Ind = 4 Then
      GroupByStr = "GROUP BY CM.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC, CM.YYYYMM "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY DOCUMENT_DATE "
      Else
         OrderByStr = "ORDER BY MONTHLY_ACCUM_ID "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY PI.PART_NO, PI.PIG_TYPE "
   ElseIf Ind = 3 Then
      OrderByStr = "ORDER BY PI.PART_NO, PI.PIG_TYPE "
   ElseIf Ind = 4 Then
      OrderByStr = "ORDER BY CM.PART_ITEM_ID ASC, CM.YYYYMM "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If

   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr

   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      MONTHLY_ACCUM_ID = NVLI(Rs("MONTHLY_ACCUM_ID"), 0)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      IMPORT_AMOUNT = NVLD(Rs("IMPORT_AMOUNT"), 0)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      BALANCE_AMOUNT1 = NVLD(Rs("BALANCE_AMOUNT1"), 0)
      BALANCE_AMOUNT2 = NVLD(Rs("BALANCE_AMOUNT2"), 0)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      SIMULATE_FLAG = NVLS(Rs("SIMULATE_FLAG"), "")
      BATCH_ID = NVLI(Rs("BATCH_ID"), 0)
      BIRTH_AMOUNT = NVLD(Rs("BIRTH_AMOUNT"), 0)
      SELL_AMOUNT = NVLD(Rs("SELL_AMOUNT"), 0)
      STATUS_IN_AMOUNT = NVLD(Rs("STATUS_IN_AMOUNT"), 0)
      STATUS_OUT_AMOUNT = NVLD(Rs("STATUS_OUT_AMOUNT"), 0)
      BUY_AMOUNT = NVLD(Rs("BUY_AMOUNT"), 0)
   ElseIf Ind = 2 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
      PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
   ElseIf Ind = 3 Then
      MONTHLY_ACCUM_ID = NVLI(Rs("MONTHLY_ACCUM_ID"), 0)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      IMPORT_AMOUNT = NVLD(Rs("IMPORT_AMOUNT"), 0)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      BALANCE_AMOUNT1 = NVLD(Rs("BALANCE_AMOUNT1"), 0)
      BALANCE_AMOUNT2 = NVLD(Rs("BALANCE_AMOUNT2"), 0)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      SIMULATE_FLAG = NVLS(Rs("SIMULATE_FLAG"), "")
      BATCH_ID = NVLI(Rs("BATCH_ID"), 0)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      BIRTH_AMOUNT = NVLD(Rs("BIRTH_AMOUNT"), 0)
      SELL_AMOUNT = NVLD(Rs("SELL_AMOUNT"), 0)
      STATUS_IN_AMOUNT = NVLD(Rs("STATUS_IN_AMOUNT"), 0)
      STATUS_OUT_AMOUNT = NVLD(Rs("STATUS_OUT_AMOUNT"), 0)
      BUY_AMOUNT = NVLD(Rs("BUY_AMOUNT"), 0)
   ElseIf Ind = 4 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      IMPORT_AMOUNT = NVLD(Rs("IMPORT_AMOUNT"), 0)
      EXPORT_AMOUNT = NVLD(Rs("EXPORT_AMOUNT"), 0)
      BALANCE_AMOUNT1 = NVLD(Rs("BALANCE_AMOUNT1"), 0)
      BALANCE_AMOUNT2 = NVLD(Rs("BALANCE_AMOUNT2"), 0)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      BIRTH_AMOUNT = NVLD(Rs("BIRTH_AMOUNT"), 0)
      SELL_AMOUNT = NVLD(Rs("SELL_AMOUNT"), 0)
      STATUS_IN_AMOUNT = NVLD(Rs("STATUS_IN_AMOUNT"), 0)
      STATUS_OUT_AMOUNT = NVLD(Rs("STATUS_OUT_AMOUNT"), 0)
      ADJUST_IN_AMOUNT = NVLD(Rs("ADJUST_IN_AMOUNT"), 0)
      ADJUST_OUT_AMOUNT = NVLD(Rs("ADJUST_OUT_AMOUNT"), 0)
      GL_IN = NVLD(Rs("GL_IN"), 0)
      GL_OUT = NVLD(Rs("GL_OUT"), 0)
      BUY_AMOUNT = NVLD(Rs("BUY_AMOUNT"), 0)
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
      Call glbDatabaseMngr.GetSeqID("MONTHLY_ACCUM_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO MONTHLY_ACCUM " & _
                     "( " & _
                     "   MONTHLY_ACCUM_ID, " & _
                     "   DOCUMENT_DATE, " & _
                     "   PART_ITEM_ID, " & _
                     "   IMPORT_AMOUNT, " & _
                     "   EXPORT_AMOUNT, " & _
                     "   BIRTH_AMOUNT, " & _
                     "   SELL_AMOUNT, " & _
                     "   STATUS_IN_AMOUNT, " & _
                     "   STATUS_OUT_AMOUNT, " & _
                     "   BUY_AMOUNT, " & _
                     "   BALANCE_AMOUNT1, " & _
                     "   BALANCE_AMOUNT2, " & _
                     "   LOCATION_ID, " & _
                     "   SIMULATE_FLAG, " & _
                     "   BATCH_ID, " & _
                     "   YYYYMM, " & _
                     "   ADJUST_IN_AMOUNT, " & _
                     "   ADJUST_OUT_AMOUNT, "
            SQL1 = SQL1 & _
                     "   GL_IN, " & _
                     "   GL_OUT, " & _
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
                        BIRTH_AMOUNT & "," & _
                        SELL_AMOUNT & "," & _
                        STATUS_IN_AMOUNT & "," & _
                        STATUS_OUT_AMOUNT & "," & _
                        BUY_AMOUNT & "," & _
                        BALANCE_AMOUNT1 & "," & _
                        BALANCE_AMOUNT2 & "," & _
                        LOCATION_ID & "," & _
                        "'" & ChangeQuote(Trim(SIMULATE_FLAG)) & "'," & _
                        TempID2 & "," & _
                        "'" & ChangeQuote(Trim(YYYYMM)) & "'," & _
                        ADJUST_IN_AMOUNT & ","
                     SQL1 = SQL1 & _
                        ADJUST_OUT_AMOUNT & "," & _
                        GL_IN & "," & _
                        GL_OUT & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         MONTHLY_ACCUM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE MONTHLY_ACCUM SET " & _
                     "    DOCUMENT_DATE = '" & ChangeQuote(DateToStringInt(Trim(DOCUMENT_DATE))) & "', " & _
                     "    PART_ITEM_ID = " & PART_ITEM_ID & ", " & _
                     "    IMPORT_AMOUNT = " & IMPORT_AMOUNT & ", " & _
                     "    EXPORT_AMOUNT = " & EXPORT_AMOUNT & ", " & _
                     "    BIRTH_AMOUNT = " & BIRTH_AMOUNT & ", " & _
                     "    SELL_AMOUNT = " & SELL_AMOUNT & ", " & _
                     "    STATUS_IN_AMOUNT = " & STATUS_IN_AMOUNT & ", " & _
                     "    STATUS_OUT_AMOUNT = " & STATUS_OUT_AMOUNT & ", " & _
                     "    BUY_AMOUNT = " & BUY_AMOUNT & ", " & _
                     "    BALANCE_AMOUNT1 = " & BALANCE_AMOUNT1 & ", " & _
                     "    BALANCE_AMOUNT2 = " & BALANCE_AMOUNT2 & ", " & _
                     "    LOCATION_ID = " & LOCATION_ID & ", " & _
                     "    SIMULATE_FLAG = '" & ChangeQuote(Trim(SIMULATE_FLAG)) & "', " & _
                     "    BATCH_ID = " & TempID2 & ", " & _
                     "    YYYYMM = '" & ChangeQuote(Trim(YYYYMM)) & "', " & _
                     "    ADJUST_IN_AMOUNT = " & ADJUST_IN_AMOUNT & ", " & _
                     "    ADJUST_OUT_AMOUNT = " & ADJUST_OUT_AMOUNT & ", " & _
                     "    GL_IN = " & GL_IN & ", " & _
                     "    GL_OUT = " & GL_OUT & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE MONTHLY_ACCUM_ID = " & MONTHLY_ACCUM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM MONTHLY_ACCUM WHERE MONTHLY_ACCUM_ID = " & MONTHLY_ACCUM_ID
      m_Conn.Execute (SQL1)
End Function

Public Function ClearData()
Dim SQL1 As String
Dim TempDate As String
Dim WhereStr As String

   SIMULATE_FLAG = glbUser.SIMULATE_FLAG
   
   SQL1 = "DELETE FROM MONTHLY_ACCUM "
   
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



