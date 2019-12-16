VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CYearWeek"
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

Public YEAR_WEEK_ID As Long
Public YEAR_SEQ_ID As Long
Public FROM_DATE As Date
Public TO_DATE As Date
Public WEEK_NO As Long
Public PART_ITEM_ID1 As Long
Public PART_ITEM_ID2 As Long
Public PART_ITEM_ID3 As Long
Public PART_ITEM_ID4 As Long
Public PART_ITEM_ID5 As Long
Public PART_ITEM_ID6 As Long

Public YEAR_NO As Long

Public PartItem1 As CPartItem
Public PartItem2 As CPartItem
Public PartItem3 As CPartItem
Public PartItem4 As CPartItem
Public PartItem5 As CPartItem
Public PartItem6 As CPartItem

Public PartItemMaps As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   
   Set PartItemMaps = New Collection
   
   Set PartItem1 = New CPartItem
   Set PartItem2 = New CPartItem
   Set PartItem3 = New CPartItem
   Set PartItem4 = New CPartItem
   Set PartItem5 = New CPartItem
   Set PartItem6 = New CPartItem
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

   If Ind = 1 Then
      SelectStr1 = "SELECT LT.*, YS.YEAR_NO  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM YEAR_WEEK LT, YEAR_SEQ YS "
      
      WhereStr = "WHERE (LT.YEAR_SEQ_ID = YS.YEAR_SEQ_ID) "
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT DISTINCT YW.WEEK_NO  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM YEAR_WEEK YW "
      
      WhereStr = ""
   End If
   
   If YEAR_WEEK_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (YEAR_WEEK_ID = " & YEAR_WEEK_ID & ") "
      Else
         WhereStr = WhereStr & "AND (YEAR_WEEK_ID = " & YEAR_WEEK_ID & ") "
      End If
   End If
   
   If YEAR_SEQ_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LT.YEAR_SEQ_ID = " & YEAR_SEQ_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LT.YEAR_SEQ_ID = " & YEAR_SEQ_ID & ") "
      End If
   End If

   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY LT.YEAR_WEEK_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY FROM_DATE "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY IP_ADDRESS "
      Else
         OrderByStr = "ORDER BY YEAR_SEQ_ID "
      End If
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY WEEK_NO "
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
   If Ind = 1 Then
      YEAR_WEEK_ID = NVLI(Rs("YEAR_WEEK_ID"), -1)
      YEAR_SEQ_ID = NVLI(Rs("YEAR_SEQ_ID"), 0)
      FROM_DATE = InternalDateToDate(NVLS(Rs("FROM_DATE"), ""))
      TO_DATE = InternalDateToDate(NVLS(Rs("TO_DATE"), ""))
      WEEK_NO = NVLI(Rs("WEEK_NO"), 0)
      YEAR_NO = NVLI(Rs("YEAR_NO"), 0)
      
      PART_ITEM_ID1 = NVLI(Rs("PART_ITEM_ID1"), -1)
      PART_ITEM_ID2 = NVLI(Rs("PART_ITEM_ID2"), -1)
      PART_ITEM_ID3 = NVLI(Rs("PART_ITEM_ID3"), -1)
      PART_ITEM_ID4 = NVLI(Rs("PART_ITEM_ID4"), -1)
      PART_ITEM_ID5 = NVLI(Rs("PART_ITEM_ID5"), -1)
      PART_ITEM_ID6 = NVLI(Rs("PART_ITEM_ID6"), -1)
   ElseIf Ind = 2 Then
      WEEK_NO = NVLI(Rs("WEEK_NO"), -1)
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

   TempID1 = PART_ITEM_ID1
   If PART_ITEM_ID1 <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = PART_ITEM_ID2
   If PART_ITEM_ID2 <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID3 = PART_ITEM_ID3
   If PART_ITEM_ID3 <= 0 Then
      TempID3 = "NULL"
   End If
   
   TempID4 = PART_ITEM_ID4
   If PART_ITEM_ID4 <= 0 Then
      TempID4 = "NULL"
   End If
   
   TempID5 = PART_ITEM_ID5
   If PART_ITEM_ID5 <= 0 Then
      TempID5 = "NULL"
   End If
   
   TempID6 = PART_ITEM_ID6
   If PART_ITEM_ID6 <= 0 Then
      TempID6 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("YEAR_WEEK_SEQ", ID, glbErrorLog)
      SQL1 = "INSERT INTO YEAR_WEEK " & _
                     "( " & _
                     "   YEAR_WEEK_ID, " & _
                     "   FROM_DATE, " & _
                     "   TO_DATE, " & _
                     "   WEEK_NO, " & _
                     "   YEAR_SEQ_ID, " & _
                     "   PART_ITEM_ID1, " & _
                     "   PART_ITEM_ID2, " & _
                     "   PART_ITEM_ID3, " & _
                     "   PART_ITEM_ID4, " & _
                     "   PART_ITEM_ID5, " & _
                     "   PART_ITEM_ID6, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
               SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(DateToStringIntLow(FROM_DATE))) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringIntHi(TO_DATE))) & "'," & _
                         "" & WEEK_NO & "," & _
                         "'" & ChangeQuote(Trim(YEAR_SEQ_ID)) & "'," & _
                         "" & TempID1 & "," & _
                         "" & TempID2 & "," & _
                         "" & TempID3 & "," & _
                         "" & TempID4 & "," & _
                         "" & TempID5 & "," & _
                         "" & TempID6 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
               YEAR_WEEK_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE YEAR_WEEK SET " & _
                     "   FROM_DATE = '" & ChangeQuote(Trim(DateToStringInt(FROM_DATE))) & "', " & _
                     "   TO_DATE = '" & ChangeQuote(Trim(DateToStringInt(TO_DATE))) & "', " & _
                     "   WEEK_NO = " & WEEK_NO & ", " & _
                     "   YEAR_SEQ_ID = '" & ChangeQuote(Trim(YEAR_SEQ_ID)) & "', " & _
                     "   PART_ITEM_ID1 = " & TempID1 & ", " & _
                     "   PART_ITEM_ID2 = " & TempID2 & ", " & _
                     "   PART_ITEM_ID3 = " & TempID3 & ", " & _
                     "   PART_ITEM_ID4 = " & TempID4 & ", " & _
                     "   PART_ITEM_ID5 = " & TempID5 & ", " & _
                     "   PART_ITEM_ID6 = " & TempID6 & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE YEAR_WEEK_ID = " & YEAR_WEEK_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM PRTITEM_MAP WHERE YEAR_WEEK_ID = " & YEAR_WEEK_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM YEAR_WEEK WHERE YEAR_WEEK_ID = " & YEAR_WEEK_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set PartItemMaps = Nothing
   
   Set PartItem1 = Nothing
   Set PartItem2 = Nothing
   Set PartItem3 = Nothing
   Set PartItem4 = Nothing
   Set PartItem5 = Nothing
   Set PartItem6 = Nothing
End Sub
