VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSupplierName"
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

Public SUPPLIER_NAME_ID As Long
Public SUPPLIER_ID As Long
Public NAME_ID As Long

Public Name As cName

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set Name = New cName
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
   
   FromStr = "FROM SUPPLIER_NAME UG "

   WhereStr = ""
   If SUPPLIER_NAME_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUPPLIER_NAME_ID = " & SUPPLIER_NAME_ID & ") "
      Else
         WhereStr = "AND (SUPPLIER_NAME_ID = " & SUPPLIER_NAME_ID & ") "
      End If
   End If
   
   If SUPPLIER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUPPLIER_ID = " & SUPPLIER_ID & ") "
      Else
         WhereStr = "AND (SUPPLIER_ID = " & SUPPLIER_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   If OrderBy > 0 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY SUPPLIER_NAME_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY SUPPLIER_NAME_ID "
      Else
         OrderByStr = "ORDER BY SUPPLIER_NAME_ID "
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
   SUPPLIER_NAME_ID = NVLI(Rs("SUPPLIER_NAME_ID"), -1)
   SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), -1)
   NAME_ID = NVLI(Rs("NAME_ID"), -1)
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("SUPPLIER_NAME_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO SUPPLIER_NAME " & _
                     "( " & _
                     "   SUPPLIER_NAME_ID, " & _
                     "   SUPPLIER_ID, " & _
                     "   NAME_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "" & SUPPLIER_ID & "," & _
                         "" & NAME_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         SUPPLIER_NAME_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE SUPPLIER_NAME SET " & _
                     "   SUPPLIER_ID = " & SUPPLIER_ID & ", " & _
                     "   NAME_ID = " & NAME_ID & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE SUPPLIER_NAME_ID = " & SUPPLIER_NAME_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM SUPPLIER_NAME WHERE SUPPLIER_NAME_ID = " & SUPPLIER_NAME_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set Name = Nothing
End Sub


