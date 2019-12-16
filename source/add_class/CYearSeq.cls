VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CYearSeq"
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
'
Public YEAR_SEQ_ID As Long
Public YEAR_DESC As String
Public YEAR_NO As Long

Public YearWeeks As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set YearWeeks = New Collection
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

   SelectStr1 = "SELECT YS.* "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM YEAR_SEQ YS  "
   
   WhereStr = ""
   
   If YEAR_SEQ_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (YEAR_SEQ_ID = " & YEAR_SEQ_ID & ") "
      Else
         WhereStr = WhereStr & "AND (YEAR_SEQ_ID = " & YEAR_SEQ_ID & ") "
      End If
   End If
         
   If YEAR_NO > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (YEAR_NO = " & YEAR_NO & ") "
      Else
         WhereStr = WhereStr & "AND (YEAR_NO = " & YEAR_NO & ") "
      End If
   End If
            
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY YEAR_NO "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY YEAR_DESC "
   Else
      OrderByStr = "ORDER BY YEAR_NO "
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

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   YEAR_SEQ_ID = NVLI(Rs("YEAR_SEQ_ID"), -1)
   YEAR_NO = NVLI(Rs("YEAR_NO"), -1)
   YEAR_DESC = NVLS(Rs("YEAR_DESC"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
      
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("YEAR_SEQ_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO YEAR_SEQ " & _
                     "( " & _
                     "   YEAR_SEQ_ID, " & _
                     "   YEAR_NO, " & _
                     "   YEAR_DESC, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         YEAR_NO & "," & _
                         "'" & ChangeQuote(Trim(YEAR_DESC)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         YEAR_SEQ_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE YEAR_SEQ SET " & _
                     "   YEAR_NO = " & YEAR_NO & ", " & _
                     "   YEAR_DESC = '" & ChangeQuote(Trim(YEAR_DESC)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE YEAR_SEQ_ID = " & YEAR_SEQ_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String

   SQL1 = "DELETE FROM PRTITEM_MAP WHERE YEAR_WEEK_ID IN (SELECT YEAR_WEEK_ID FROM YEAR_WEEK WHERE YEAR_SEQ_ID =  " & YEAR_SEQ_ID & " )"
   m_Conn.Execute (SQL1)
   
   SQL1 = "DELETE FROM YEAR_WEEK WHERE YEAR_SEQ_ID = " & YEAR_SEQ_ID
   m_Conn.Execute (SQL1)
               
   SQL1 = "DELETE FROM YEAR_SEQ WHERE YEAR_SEQ_ID = " & YEAR_SEQ_ID
   m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set YearWeeks = Nothing
End Sub
