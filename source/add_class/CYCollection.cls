VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CYCollection"
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

Public Y_COLLECTION_ID As Long
Public Y_COLLECTION_NAME As String
Public Y_COLLECTION_DESC As String
Public MASK1 As String
Public MASK2 As String
Public MASK3 As String
Public MASK4 As String
Public MASK5 As String
Public MASK6 As String
Public MASK7 As String
Public MASK8 As String
Public MASK9 As String
Public MASK10 As String

Public X_Collection As CXCollection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set X_Collection = New CXCollection
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
   
   FromStr = "FROM Y_COLLECTION UG "
   
   WhereStr = ""
   If Y_COLLECTION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (Y_COLLECTION_ID = " & Y_COLLECTION_ID & ") "
      Else
         WhereStr = "AND (Y_COLLECTION_ID = " & Y_COLLECTION_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY Y_COLLECTION_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY Y_COLLECTION_NAME "
   Else
      OrderByStr = "ORDER BY Y_COLLECTION_NAME, Y_COLLECTION_ID "
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
   Y_COLLECTION_ID = NVLI(Rs("Y_COLLECTION_ID"), -1)
   Y_COLLECTION_NAME = NVLS(Rs("Y_COLLECTION_NAME"), "")
   Y_COLLECTION_DESC = NVLS(Rs("Y_COLLECTION_DESC"), "")
   MASK1 = NVLS(Rs("MASK1"), "")
   MASK2 = NVLS(Rs("MASK2"), "")
   MASK3 = NVLS(Rs("MASK3"), "")
   MASK4 = NVLS(Rs("MASK4"), "")
   MASK5 = NVLS(Rs("MASK5"), "")
   MASK6 = NVLS(Rs("MASK6"), "")
   MASK7 = NVLS(Rs("MASK7"), "")
   MASK8 = NVLS(Rs("MASK8"), "")
   MASK9 = NVLS(Rs("MASK9"), "")
   MASK10 = NVLS(Rs("MASK10"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("Y_COLLECTION_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO Y_COLLECTION " & _
                     "( " & _
                     "   Y_COLLECTION_ID, " & _
                     "   Y_COLLECTION_NAME, " & _
                     "   Y_COLLECTION_DESC, " & _
                     "   MASK1, " & _
                     "   MASK2, " & _
                     "   MASK3, " & _
                     "   MASK4, " & _
                     "   MASK5, " & _
                     "   MASK6, "
            SQL1 = SQL1 & _
                     "   MASK7, " & _
                     "   MASK8, " & _
                     "   MASK9, " & _
                     "   MASK10, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(Y_COLLECTION_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(Y_COLLECTION_DESC)) & "'," & _
                         "'" & ChangeQuote(Trim(MASK1)) & "'," & _
                         "'" & ChangeQuote(Trim(MASK2)) & "'," & _
                         "'" & ChangeQuote(Trim(MASK3)) & "'," & _
                         "'" & ChangeQuote(Trim(MASK4)) & "'," & _
                         "'" & ChangeQuote(Trim(MASK5)) & "'," & _
                         "'" & ChangeQuote(Trim(MASK6)) & "',"
            SQL1 = SQL1 & _
                        "'" & ChangeQuote(Trim(MASK7)) & "'," & _
                        "'" & ChangeQuote(Trim(MASK8)) & "'," & _
                        "'" & ChangeQuote(Trim(MASK9)) & "'," & _
                        "'" & ChangeQuote(Trim(MASK10)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         Y_COLLECTION_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE Y_COLLECTION SET " & _
                     "   Y_COLLECTION_NAME = '" & ChangeQuote(Trim(Y_COLLECTION_NAME)) & "', " & _
                     "   Y_COLLECTION_DESC = '" & ChangeQuote(Trim(Y_COLLECTION_DESC)) & "', " & _
                     "   MASK1 = '" & ChangeQuote(Trim(MASK1)) & "', " & _
                     "   MASK2 = '" & ChangeQuote(Trim(MASK2)) & "', " & _
                     "   MASK3 = '" & ChangeQuote(Trim(MASK3)) & "', " & _
                     "   MASK4 = '" & ChangeQuote(Trim(MASK4)) & "', " & _
                     "   MASK5 = '" & ChangeQuote(Trim(MASK5)) & "', " & _
                     "   MASK6 = '" & ChangeQuote(Trim(MASK6)) & "', " & _
                     "   MASK7= '" & ChangeQuote(Trim(MASK7)) & "', " & _
                     "   MASK8 = '" & ChangeQuote(Trim(MASK8)) & "', " & _
                     "   MASK9 = '" & ChangeQuote(Trim(MASK9)) & "', " & _
                     "   MASK10 = '" & ChangeQuote(Trim(MASK10)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE Y_COLLECTION_ID = " & Y_COLLECTION_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM Y_COLLECTION WHERE Y_COLLECTION_ID = " & Y_COLLECTION_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set X_Collection = Nothing
End Sub
