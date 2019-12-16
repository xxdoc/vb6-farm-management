VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CXItem"
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

Public X_ITEM_ID As Long
Public X_COLLECTION_ID As Long
Public ITEM_DATE As Date
Public ITEM_SEQ As Long
Public ITEM_VALUE As Double

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   ITEM_DATE = -1
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

   SelectStr1 = "SELECT BL.* "
   
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM X_ITEM BL "
   
   WhereStr = ""
   
   If X_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (X_ITEM_ID = " & X_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (X_ITEM_ID = " & X_ITEM_ID & ") "
      End If
   End If
         
   If X_COLLECTION_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (X_COLLECTION_ID = " & X_COLLECTION_ID & ") "
      Else
         WhereStr = WhereStr & "AND (X_COLLECTION_ID = " & X_COLLECTION_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY ITEM_DATE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY ITEM_DATE "
   Else
      OrderByStr = "ORDER BY X_ITEM_ID "
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

Public Function PopulateFromRS(Rs As ADODB.Recordset)
   X_ITEM_ID = NVLI(Rs("X_ITEM_ID"), -1)
   X_COLLECTION_ID = NVLI(Rs("X_COLLECTION_ID"), -1)
   ITEM_SEQ = NVLI(Rs("ITEM_SEQ"), 0)
   ITEM_DATE = InternalDateToDate(NVLS(Rs("ITEM_DATE"), ""))
   ITEM_VALUE = NVLD(Rs("ITEM_VALUE"), -1)
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String

   TempID1 = ITEM_SEQ
   If ITEM_SEQ <= 0 Then
      TempID1 = "NULL"
   End If
         
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("X_ITEM_SEQ", ID, glbErrorLog)
               
      SQL1 = "INSERT INTO X_ITEM " & _
                     "( " & _
                     "   X_ITEM_ID, " & _
                     "   X_COLLECTION_ID, " & _
                     "   ITEM_SEQ, " & _
                     "   ITEM_DATE, " & _
                     "   ITEM_VALUE, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
                  SQL1 = SQL1 & _
                     "( " & _
                         ID & "," & _
                         X_COLLECTION_ID & "," & _
                         TempID1 & "," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(ITEM_DATE))) & "'," & _
                         ITEM_VALUE & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         X_ITEM_ID = ID
         
   Else 'Show edit
      SQL1 = "UPDATE X_ITEM SET " & _
                     "    ITEM_SEQ = " & TempID1 & ", " & _
                     "    X_COLLECTION_ID = " & X_COLLECTION_ID & ", " & _
                     "   ITEM_DATE = '" & ChangeQuote(Trim(DateToStringInt(ITEM_DATE))) & "', " & _
                     "    ITEM_VALUE = " & ITEM_VALUE & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE X_ITEM_ID = " & X_ITEM_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
            
      SQL1 = "DELETE FROM X_ITEM WHERE X_ITEM_ID = " & X_ITEM_ID
      m_Conn.Execute (SQL1)
End Function

