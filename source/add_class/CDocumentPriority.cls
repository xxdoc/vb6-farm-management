VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDocumentPriority"
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

Public DOCUMENT_PRIORITY_ID As Long
Public DOCUMENT_TYPE As Long
Public Area As Long
Public PRIORITY1 As Long
Public PRIORITY2 As Long

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

   SelectStr1 = "SELECT UG.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM DOCUMENT_PRIORITY UG "
   
   WhereStr = ""
   If DOCUMENT_PRIORITY_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DOCUMENT_PRIORITY_ID = " & DOCUMENT_PRIORITY_ID & ") "
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_PRIORITY_ID = " & DOCUMENT_PRIORITY_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY DOCUMENT_PRIORITY_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY AREA "
   Else
      OrderByStr = "ORDER BY DOCUMENT_PRIORITY_ID "
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
   DOCUMENT_PRIORITY_ID = NVLI(Rs("DOCUMENT_PRIORITY_ID"), -1)
   Area = NVLI(Rs("AREA"), -1)
   DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
   PRIORITY1 = NVLI(Rs("PRIORITY1"), -1)
   PRIORITY2 = NVLI(Rs("PRIORITY2"), -1)
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("DOCUMENT_PRIORITY_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO DOCUMENT_PRIORITY " & _
                     "( " & _
                     "   DOCUMENT_PRIORITY_ID, " & _
                     "   AREA, " & _
                     "   DOCUMENT_TYPE, " & _
                     "   PRIORITY1, " & _
                     "   PRIORITY2, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         Area & "," & _
                         DOCUMENT_TYPE & "," & _
                        PRIORITY1 & "," & _
                        PRIORITY2 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         DOCUMENT_PRIORITY_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE DOCUMENT_PRIORITY SET " & _
                     "   DOCUMENT_TYPE = " & DOCUMENT_TYPE & ", " & _
                     "   AREA = " & Area & ", " & _
                     "   PRIORITY1 = " & PRIORITY1 & ", " & _
                     "   PRIORITY2 = " & PRIORITY2 & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE DOCUMENT_PRIORITY_ID = " & DOCUMENT_PRIORITY_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM DOCUMENT_PRIORITY WHERE DOCUMENT_PRIORITY_ID = " & DOCUMENT_PRIORITY_ID
      m_Conn.Execute (SQL1)
End Function

Public Function ClearData()
Dim SQL1 As String
      SQL1 = "DELETE FROM DOCUMENT_PRIORITY "
      m_Conn.Execute (SQL1)
End Function



