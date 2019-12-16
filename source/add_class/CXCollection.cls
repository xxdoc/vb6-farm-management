VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXCollection"
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
Public ItemOrderType As Long
Public Flag As String
' End of generic data member

Public X_COLLECTION_ID As Long
Public X_COLLECTION_NAME As String
Public X_COLLECTION_DESC As String

Public XItems As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set XItems = New Collection
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
   
   FromStr = "FROM X_COLLECTION UG "
   
   WhereStr = ""
   If X_COLLECTION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (X_COLLECTION_ID = " & X_COLLECTION_ID & ") "
      Else
         WhereStr = "AND (X_COLLECTION_ID = " & X_COLLECTION_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY X_COLLECTION_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY X_COLLECTION_NAME "
   Else
      OrderByStr = "ORDER BY X_COLLECTION_NAME, X_COLLECTION_ID "
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
   X_COLLECTION_ID = NVLI(Rs("X_COLLECTION_ID"), -1)
   X_COLLECTION_NAME = NVLS(Rs("X_COLLECTION_NAME"), "")
   X_COLLECTION_DESC = NVLS(Rs("X_COLLECTION_DESC"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("X_COLLECTION_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO X_COLLECTION " & _
                     "( " & _
                     "   X_COLLECTION_ID, " & _
                     "   X_COLLECTION_NAME, " & _
                     "   X_COLLECTION_DESC, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(X_COLLECTION_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(X_COLLECTION_DESC)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         X_COLLECTION_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE X_COLLECTION SET " & _
                     "   X_COLLECTION_NAME = '" & ChangeQuote(Trim(X_COLLECTION_NAME)) & "', " & _
                     "   X_COLLECTION_DESC = '" & ChangeQuote(Trim(X_COLLECTION_DESC)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE X_COLLECTION_ID = " & X_COLLECTION_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM X_ITEM WHERE X_COLLECTION_ID = " & X_COLLECTION_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM X_COLLECTION WHERE X_COLLECTION_ID = " & X_COLLECTION_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set XItems = Nothing
End Sub
