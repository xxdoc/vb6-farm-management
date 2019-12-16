VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIntake"
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

Public INTAKE_ID As Long
Public BATCH_ID As Long
Public DOCUMENT_DATE As Date
Public PIG_ID As Long
Public PART_ITEM_ID As Long
Public INTAKE_COST As Double                             'มูลค่าอาหารที่หมูกิน
Public INTAKE_AMOUNT As Double                     'น้ำหนักอาหารที่หมูกิน
Public CURRENT_PIG_AMOUNT As Double                 'ปริมาณหมูที่กินอาหารนั้นๆเข้าไป

Public SIMULATE_FLAG As String
Public YYYYMM As String
Public FROM_DATE As Date
Public TO_DATE As Date
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_DATE = -1
   TO_DATE = -1
End Sub
'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long, Optional MustCount As Boolean = False)
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
      SelectStr1 = "SELECT IT.* "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INTAKE IT "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT IT.PIG_ID,IT.YYYYMM,SUM(IT.INTAKE_COST) INTAKE_COST,SUM(IT.INTAKE_AMOUNT) INTAKE_AMOUNT,SUM(IT.CURRENT_PIG_AMOUNT) CURRENT_PIG_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INTAKE IT "
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT IT.PIG_ID,IT.PART_ITEM_ID,IT.YYYYMM,SUM(IT.INTAKE_COST) INTAKE_COST,SUM(IT.INTAKE_AMOUNT) INTAKE_AMOUNT,SUM(IT.CURRENT_PIG_AMOUNT) CURRENT_PIG_AMOUNT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM INTAKE IT "
    End If
     
   If INTAKE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (INTAKE_ID = " & INTAKE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (INTAKE_ID = " & INTAKE_ID & ") "
      End If
   End If
      
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   If PIG_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PIG_ID = " & PIG_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PIG_ID = " & PIG_ID & ") "
      End If
   End If
   
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
   
   If Ind = 2 Then
      GroupByStr = "GROUP BY IT.PIG_ID,IT.YYYYMM "
   ElseIf Ind = 3 Then
      GroupByStr = "GROUP BY IT.PIG_ID,IT.YYYYMM,IT.PART_ITEM_ID "
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY DOCUMENT_DATE "
      Else
         OrderByStr = "ORDER BY INTAKE_ID "
      End If
   End If
   
   If Not (Ind = 2 Or Ind = 3) Then
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      INTAKE_ID = NVLI(Rs("INTAKE_ID"), 0)
      BATCH_ID = NVLI(Rs("BATCH_ID"), -1)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      PIG_ID = NVLI(Rs("PIG_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      
      INTAKE_COST = NVLD(Rs("INTAKE_COST"), 0)
      INTAKE_AMOUNT = NVLD(Rs("INTAKE_AMOUNT"), 0)
      CURRENT_PIG_AMOUNT = NVLD(Rs("CURRENT_PIG_AMOUNT"), 0)
   ElseIf Ind = 2 Then
      PIG_ID = NVLI(Rs("PIG_ID"), -1)
      INTAKE_COST = NVLD(Rs("INTAKE_COST"), 0)
      INTAKE_AMOUNT = NVLD(Rs("INTAKE_AMOUNT"), 0)
      CURRENT_PIG_AMOUNT = NVLD(Rs("CURRENT_PIG_AMOUNT"), 0)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
   ElseIf Ind = 3 Then
      PIG_ID = NVLI(Rs("PIG_ID"), -1)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      INTAKE_COST = NVLD(Rs("INTAKE_COST"), 0)
      INTAKE_AMOUNT = NVLD(Rs("INTAKE_AMOUNT"), 0)
      CURRENT_PIG_AMOUNT = NVLD(Rs("CURRENT_PIG_AMOUNT"), 0)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
   
   TempID1 = PIG_ID
   If PIG_ID <= 0 Then
      TempID1 = "NULL"
   End If
      
   TempID2 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   SIMULATE_FLAG = glbUser.SIMULATE_FLAG
   YYYYMM = Mid(DateToStringInt(DOCUMENT_DATE), 1, 7)
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("INTAKE_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO INTAKE " & _
                     "( " & _
                     "   INTAKE_ID, " & _
                     "   BATCH_ID, " & _
                     "   DOCUMENT_DATE, " & _
                     "   PIG_ID, " & _
                     "   PART_ITEM_ID, " & _
                     "   INTAKE_COST, " & _
                     "   INTAKE_AMOUNT, " & _
                     "   CURRENT_PIG_AMOUNT, " & _
                     "   SIMULATE_FLAG, " & _
                     "   YYYYMM, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
            SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         BATCH_ID & "," & _
                        "'" & ChangeQuote(DateToStringInt(Trim(DOCUMENT_DATE))) & "'," & _
                        TempID1 & "," & _
                        TempID2 & "," & _
                        INTAKE_COST & "," & _
                        INTAKE_AMOUNT & "," & _
                        CURRENT_PIG_AMOUNT & "," & _
                        "'" & ChangeQuote(Trim(SIMULATE_FLAG)) & "'," & _
                        "'" & ChangeQuote(Trim(YYYYMM)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         INTAKE_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE INTAKE SET " & _
                     "    BATCH_ID = " & BATCH_ID & ", " & _
                     "    DOCUMENT_DATE = '" & ChangeQuote(DateToStringInt(Trim(DOCUMENT_DATE))) & "', " & _
                     "    PIG_ID = " & TempID1 & ", " & _
                     "    PART_ITEM_ID = " & TempID2 & ", " & _
                     "    INTAKE_COST = " & INTAKE_COST & ", " & _
                     "    INTAKE_AMOUNT = " & INTAKE_AMOUNT & ", " & _
                     "    CURRENT_PIG_AMOUNT = " & CURRENT_PIG_AMOUNT & ", " & _
                     "    SIMULATE_FLAG = '" & ChangeQuote(Trim(SIMULATE_FLAG)) & "', " & _
                     "    YYYYMM = '" & ChangeQuote(Trim(YYYYMM)) & "', " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE INTAKE_ID = " & INTAKE_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM INTAKE WHERE INTAKE_ID = " & INTAKE_ID
      m_Conn.Execute (SQL1)
End Function
