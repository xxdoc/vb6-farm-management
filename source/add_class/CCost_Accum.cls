VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCost_Accum"
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

Public COST_ACCUM_ID As Long
Public DOCUMENT_DATE As Date
Public DOCUMENT_TYPE As Long
Public DOCUMENT_CATEGORY As Long
Public PART_ITEM_ID As Long
Public LOCATION_ID As Long
Public BATCH_ID As Long
Public COST_RAW As Double
Public COST_EXP As Double
Public COST_PB As Double
Public COST_MOVEMENT As Double
Public YYYYMM As String
Public CUS_ID As Long


Public ITEM_AMOUNT As Double

Public FROM_DATE As Date
Public TO_DATE As Date

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
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
Dim GroupByStr As String

   YYYYMM = Mid(DateToStringInt(DOCUMENT_DATE), 1, 7)
   
   If Ind = 1 Then
      SelectStr1 = "SELECT CM.*  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COST_ACCUM CM "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT CM.YYYYMM "
      SelectStr1 = SelectStr1 & ",SUM(CM.COST_RAW) COST_RAW,SUM(CM.COST_EXP) COST_EXP,SUM(CM.COST_PB) COST_PB,SUM(CM.COST_MOVEMENT) COST_MOVEMENT "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM COST_ACCUM CM "
   End If

   If COST_ACCUM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.COST_ACCUM_ID = " & COST_ACCUM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CM.COST_ACCUM_ID = " & COST_ACCUM_ID & ") "
      End If
   End If
      
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (CM.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      End If
   End If
      
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CM.PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   If LOCATION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.LOCATION_ID = " & LOCATION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CM.LOCATION_ID = " & LOCATION_ID & ") "
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
                             
   If BATCH_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.BATCH_ID = " & BATCH_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CM.BATCH_ID = " & BATCH_ID & ") "
      End If
   Else
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE ((CM.BATCH_ID <= 0) OR (BATCH_ID IS NULL)) "
      Else
         WhereStr = WhereStr & "AND ((CM.BATCH_ID <= 0) OR (BATCH_ID IS NULL)) "
      End If
   End If
   
   GroupByStr = ""
   
   If Ind = 2 Then
      GroupByStr = "GROUP BY CM.YYYYMM "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY CM.DOCUMENT_DATE "
      Else
         OrderByStr = "ORDER BY CM.COST_ACCUM_ID "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY CM.YYYYMM "
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

   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      COST_ACCUM_ID = NVLI(Rs("COST_ACCUM_ID"), 0)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), 0)
      DOCUMENT_CATEGORY = NVLI(Rs("DOCUMENT_CATEGORY"), 0)
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), 0)
      LOCATION_ID = NVLI(Rs("LOCATION_ID"), 0)
      BATCH_ID = NVLI(Rs("BATCH_ID"), 0)
      COST_RAW = NVLD(Rs("COST_RAW"), 0)
      COST_EXP = NVLD(Rs("COST_EXP"), 0)
      COST_PB = NVLD(Rs("COST_PB"), 0)
      COST_MOVEMENT = NVLD(Rs("COST_MOVEMENT"), 0)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
      CUS_ID = NVLI(Rs("CUS_ID"), 0)
      ITEM_AMOUNT = NVLD(Rs("ITEM_AMOUNT"), 0)
   ElseIf Ind = 2 Then
      COST_RAW = NVLD(Rs("COST_RAW"), 0)
      COST_EXP = NVLD(Rs("COST_EXP"), 0)
      COST_PB = NVLD(Rs("COST_PB"), 0)
      COST_MOVEMENT = NVLD(Rs("COST_MOVEMENT"), 0)
      YYYYMM = NVLS(Rs("YYYYMM"), "")
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String

   TempID1 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID1 = "NULL"
   End If
      
   TempID2 = BATCH_ID
   If BATCH_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID3 = CUS_ID
   If CUS_ID <= 0 Then
      TempID3 = "NULL"
   End If
   
   'SIMULATE_FLAG = glbUser.SIMULATE_FLAG
   YYYYMM = Mid(DateToStringInt(DOCUMENT_DATE), 1, 7)
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("COST_ACCUM_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO COST_ACCUM " & _
                     "( " & _
                     "   COST_ACCUM_ID, " & _
                     "   DOCUMENT_DATE, " & _
                     "   DOCUMENT_TYPE, " & _
                     "   DOCUMENT_CATEGORY, " & _
                     "   PART_ITEM_ID, " & _
                     "   LOCATION_ID, " & _
                     "   BATCH_ID, " & _
                     "   COST_RAW, " & _
                     "   COST_EXP, " & _
                     "   YYYYMM, " & _
                     "   CUS_ID, " & _
                     "   ITEM_AMOUNT, " & _
                     "   COST_PB, " & _
                     "   COST_MOVEMENT, " & _
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
                        DOCUMENT_TYPE & "," & _
                        DOCUMENT_CATEGORY & "," & _
                        PART_ITEM_ID & "," & _
                        LOCATION_ID & "," & _
                        TempID2 & "," & _
                        COST_RAW & "," & _
                        COST_EXP & "," & _
                        "'" & ChangeQuote(Trim(YYYYMM)) & "'," & _
                        TempID3 & "," & _
                        ITEM_AMOUNT & "," & _
                        COST_PB & "," & _
                        COST_MOVEMENT & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         COST_ACCUM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE COST_ACCUM SET " & _
                     "    DOCUMENT_DATE = '" & ChangeQuote(DateToStringInt(Trim(DOCUMENT_DATE))) & "', " & _
                     "    DOCUMENT_TYPE = " & DOCUMENT_TYPE & ", " & _
                     "    DOCUMENT_CATEGORY = " & DOCUMENT_CATEGORY & ", " & _
                     "    PART_ITEM_ID = " & PART_ITEM_ID & ", " & _
                     "    LOCATION_ID = " & LOCATION_ID & ", " & _
                     "    BATCH_ID = " & TempID2 & ", " & _
                     "    COST_RAW = " & COST_RAW & ", " & _
                     "    COST_EXP = " & COST_EXP & ", " & _
                     "    YYYYMM = '" & ChangeQuote(Trim(YYYYMM)) & "', " & _
                     "    CUS_ID = " & TempID3 & ", " & _
                     "    ITEM_AMOUNT = " & ITEM_AMOUNT & ", " & _
                     "    COST_PB = " & COST_PB & ", " & _
                     "    COST_MOVEMENT = " & COST_MOVEMENT & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE COST_ACCUM_ID = " & COST_ACCUM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM COST_ACCUM WHERE COST_ACCUM_ID = " & COST_ACCUM_ID
      m_Conn.Execute (SQL1)
End Function
Public Function ClearData()
Dim SQL1 As String
Dim TempDate As String
 Dim WhereStr As String
 
   WhereStr = ""
      
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
   
   SQL1 = "DELETE FROM COST_ACCUM CM "
   SQL1 = SQL1 & WhereStr
   
   m_Conn.Execute (SQL1)

End Function
Public Function GetKey1() As String
   GetKey1 = DOCUMENT_DATE & "-" & DOCUMENT_CATEGORY & "-" & DOCUMENT_TYPE & "-" & LOCATION_ID & "-" & PART_ITEM_ID & "-" & CUS_ID
End Function
Public Function GetKey2() As String
   GetKey2 = LOCATION_ID & "-" & PART_ITEM_ID
End Function
