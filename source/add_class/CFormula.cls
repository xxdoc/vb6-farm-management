VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFormula"
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

Public FORMULA_ID As Long
Public FORMULA_NAME As String
Public FORMULA_DESC As String
Public X_COLLECTION_ID As Long
Public Y_COLLECTION_ID As Long
Public X_START As Long
Public X_END As Long

Public XCollection As CXCollection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set XCollection = New CXCollection
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
   
   FromStr = "FROM FORMULA BL "
   
   WhereStr = ""
   
   If FORMULA_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FORMULA_ID = " & FORMULA_ID & ") "
      Else
         WhereStr = WhereStr & "AND (FORMULA_ID = " & FORMULA_ID & ") "
      End If
   End If
         
   '===
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY FORMULA_NAME "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY FORMULA_NAME "
   Else
      OrderByStr = "ORDER BY FORMULA_ID "
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
   FORMULA_ID = NVLI(Rs("FORMULA_ID"), -1)
   FORMULA_NAME = NVLS(Rs("FORMULA_NAME"), "")
   X_END = NVLI(Rs("X_END"), 0)
   X_COLLECTION_ID = NVLI(Rs("X_COLLECTION_ID"), 0)
   Y_COLLECTION_ID = NVLI(Rs("Y_COLLECTION_ID"), -1)
   X_START = NVLD(Rs("X_START"), -1)
   FORMULA_DESC = NVLS(Rs("FORMULA_DESC"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String

   TempID1 = X_COLLECTION_ID
   If X_COLLECTION_ID <= 0 Then
      TempID1 = "NULL"
   End If
         
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("FORMULA_SEQ", ID, glbErrorLog)
               
      SQL1 = "INSERT INTO FORMULA " & _
                     "( " & _
                     "   FORMULA_ID, " & _
                     "   FORMULA_NAME, " & _
                     "   X_COLLECTION_ID, " & _
                     "   Y_COLLECTION_ID, " & _
                     "   X_START, " & _
                     "   X_END, " & _
                     "   FORMULA_DESC, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
                  SQL1 = SQL1 & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(FORMULA_NAME)) & "'," & _
                         TempID1 & "," & _
                         Y_COLLECTION_ID & "," & _
                         X_START & "," & _
                         X_END & "," & _
                         "'" & ChangeQuote(Trim(FORMULA_DESC)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         FORMULA_ID = ID
         
   Else 'Show edit
      SQL1 = "UPDATE FORMULA SET " & _
                     "   FORMULA_NAME = '" & ChangeQuote(Trim(FORMULA_NAME)) & "', " & _
                     "    X_COLLECTION_ID = " & TempID1 & ", " & _
                     "    Y_COLLECTION_ID = " & Y_COLLECTION_ID & ", " & _
                     "    X_START = " & X_START & ", " & _
                     "    X_END = " & X_END & ", " & _
                     "   FORMULA_DESC = '" & ChangeQuote(Trim(FORMULA_DESC)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE FORMULA_ID = " & FORMULA_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM FORMULA WHERE FORMULA_ID = " & FORMULA_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set XCollection = Nothing
End Sub
