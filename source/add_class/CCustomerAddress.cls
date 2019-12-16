VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCustomerAddress"
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

Public CUSTOMER_ADDRESS_ID As Long
Public CUSTOMER_ID As Long
Public ADDRESS_ID As Long
Public ADDRESS_TYPE As Long

Public Addresses As CAddress

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set Addresses = New CAddress
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
   
   FromStr = "FROM CUSTOMER_ADDRESS UG "

   WhereStr = ""
   If CUSTOMER_ADDRESS_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CUSTOMER_ADDRESS_ID = " & CUSTOMER_ADDRESS_ID & ") "
      Else
         WhereStr = "AND (CUSTOMER_ADDRESS_ID = " & CUSTOMER_ADDRESS_ID & ") "
      End If
   End If
   
   If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CUSTOMER_ID = " & CUSTOMER_ID & ") "
      Else
         WhereStr = "AND (CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   If OrderBy > 0 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY CUSTOMER_ADDRESS_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY CUSTOMER_ADDRESS_ID "
      Else
         OrderByStr = "ORDER BY CUSTOMER_ADDRESS_ID "
      End If
      
      If OrderType = 1 Then
         OrderTypeStr = "ASC "
      ElseIf OrderType = 2 Then
         OrderTypeStr = "DESC "
      Else
         OrderTypeStr = "DESC "
      End If
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Rs As ADODB.Recordset)
   CUSTOMER_ADDRESS_ID = NVLI(Rs("CUSTOMER_ADDRESS_ID"), -1)
   CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
   ADDRESS_ID = NVLI(Rs("ADDRESS_ID"), -1)
   ADDRESS_TYPE = NVLI(Rs("ADDRESS_TYPE"), -1)
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String

   TempID1 = ADDRESS_TYPE
   If ADDRESS_TYPE <= 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("CUSTOMER_ADDRESS_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO CUSTOMER_ADDRESS " & _
                     "( " & _
                     "   CUSTOMER_ADDRESS_ID, " & _
                     "   CUSTOMER_ID, " & _
                     "   ADDRESS_ID, " & _
                     "   ADDRESS_TYPE, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "" & CUSTOMER_ID & "," & _
                         "" & ADDRESS_ID & "," & _
                         TempID1 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         CUSTOMER_ADDRESS_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE CUSTOMER_ADDRESS SET " & _
                     "   CUSTOMER_ID = " & CUSTOMER_ID & ", " & _
                     "   ADDRESS_ID = " & ADDRESS_ID & ", " & _
                     "   ADDRESS_TYPE = " & TempID1 & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE CUSTOMER_ADDRESS_ID = " & CUSTOMER_ADDRESS_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM CUSTOMER_ADDRESS WHERE CUSTOMER_ADDRESS_ID = " & CUSTOMER_ADDRESS_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set Addresses = Nothing
End Sub


