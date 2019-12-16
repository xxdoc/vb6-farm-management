VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPrtItemMap"
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

Public PRTITEM_MAP_ID As Long
Public YEAR_WEEK_ID As Long
Public PRODUCT_TYPE_ID As Long
Public PART_ITEM_ID As String

Public PartItem As CPartItem

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set PartItem = New CPartItem
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

   SelectStr1 = "SELECT UG.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PRTITEM_MAP UG "
   
   WhereStr = ""
   If PRTITEM_MAP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PRTITEM_MAP_ID = " & PRTITEM_MAP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PRTITEM_MAP_ID = " & PRTITEM_MAP_ID & ") "
      End If
   End If
   
   If YEAR_WEEK_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (YEAR_WEEK_ID = " & YEAR_WEEK_ID & ") "
      Else
         WhereStr = WhereStr & "AND (YEAR_WEEK_ID = " & YEAR_WEEK_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PRTITEM_MAP_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PRTITEM_MAP_ID "
   Else
      OrderByStr = "ORDER BY PRTITEM_MAP_ID "
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
   PRTITEM_MAP_ID = NVLI(Rs("PRTITEM_MAP_ID"), -1)
   YEAR_WEEK_ID = NVLI(Rs("YEAR_WEEK_ID"), -1)
   PRODUCT_TYPE_ID = NVLI(Rs("PRODUCT_TYPE_ID"), -1)
   PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("PRTITEM_MAP_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO PRTITEM_MAP " & _
                     "( " & _
                     "   PRTITEM_MAP_ID, " & _
                     "   YEAR_WEEK_ID, " & _
                     "   PRODUCT_TYPE_ID, " & _
                     "   PART_ITEM_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(YEAR_WEEK_ID)) & "'," & _
                         "'" & ChangeQuote(Trim(PRODUCT_TYPE_ID)) & "'," & _
                         "'" & ChangeQuote(Trim(PART_ITEM_ID)) & "',"
               SQL1 = SQL1 & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         YEAR_WEEK_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE PRTITEM_MAP SET " & _
                     "   YEAR_WEEK_ID = '" & ChangeQuote(Trim(YEAR_WEEK_ID)) & "', " & _
                     "   PRODUCT_TYPE_ID = '" & ChangeQuote(Trim(PRODUCT_TYPE_ID)) & "', " & _
                     "   PART_ITEM_ID = '" & ChangeQuote(Trim(PART_ITEM_ID)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PRTITEM_MAP_ID = " & PRTITEM_MAP_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM PRTITEM_MAP WHERE PRTITEM_MAP_ID = " & PRTITEM_MAP_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set PartItem = Nothing
End Sub
