VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProductType"
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

Public PRODUCT_TYPE_ID As Long
Public PRODUCT_TYPE_NAME As String
Public PRODUCT_TYPE_NO As String
Public CAPITAL_FLAG As String

Public LOCK_WEIGHT_FLAG As String


Public KEY_ID As Long
Public KEY_LOOKUP As String

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
   
   FromStr = "FROM PRODUCT_TYPE UG "
   
   WhereStr = ""
   If PRODUCT_TYPE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PRODUCT_TYPE_ID = " & PRODUCT_TYPE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PRODUCT_TYPE_ID = " & PRODUCT_TYPE_ID & ") "
      End If
   End If
   
   If Len(CAPITAL_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CAPITAL_FLAG = '" & CAPITAL_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (CAPITAL_FLAG = '" & CAPITAL_FLAG & "') "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PRODUCT_TYPE_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PRODUCT_TYPE_NAME "
   Else
      OrderByStr = "ORDER BY PRODUCT_TYPE_NO "
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
   PRODUCT_TYPE_ID = NVLI(Rs("PRODUCT_TYPE_ID"), -1)
   PRODUCT_TYPE_NAME = NVLS(Rs("PRODUCT_TYPE_NAME"), "")
   PRODUCT_TYPE_NO = NVLS(Rs("PRODUCT_TYPE_NO"), "")
   CAPITAL_FLAG = NVLS(Rs("CAPITAL_FLAG"), "")
      
   LOCK_WEIGHT_FLAG = NVLS(Rs("LOCK_WEIGHT_FLAG"), "")
   
   KEY_ID = PRODUCT_TYPE_ID
   KEY_LOOKUP = PRODUCT_TYPE_NO
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("PRODUCT_TYPE_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO PRODUCT_TYPE " & _
                     "( " & _
                     "   PRODUCT_TYPE_ID, " & _
                     "   PRODUCT_TYPE_NAME, " & _
                     "   PRODUCT_TYPE_NO, " & _
                     "   CAPITAL_FLAG, " & _
                     "   LOCK_WEIGHT_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(PRODUCT_TYPE_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(PRODUCT_TYPE_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(CAPITAL_FLAG)) & "'," & _
                         "'" & ChangeQuote(Trim(LOCK_WEIGHT_FLAG)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         PRODUCT_TYPE_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE PRODUCT_TYPE SET " & _
                     "   PRODUCT_TYPE_NAME = '" & ChangeQuote(Trim(PRODUCT_TYPE_NAME)) & "', " & _
                     "   PRODUCT_TYPE_NO = '" & ChangeQuote(Trim(PRODUCT_TYPE_NO)) & "', " & _
                     "   CAPITAL_FLAG = '" & ChangeQuote(Trim(CAPITAL_FLAG)) & "', " & _
                    "   LOCK_WEIGHT_FLAG = '" & ChangeQuote(Trim(LOCK_WEIGHT_FLAG)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PRODUCT_TYPE_ID = " & PRODUCT_TYPE_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM PRODUCT_TYPE WHERE PRODUCT_TYPE_ID = " & PRODUCT_TYPE_ID
      m_Conn.Execute (SQL1)
End Function

Public Function DeleteAllData()
Dim SQL1 As String
      SQL1 = "DELETE FROM PRODUCT_TYPE  "
      m_Conn.Execute (SQL1)
End Function




