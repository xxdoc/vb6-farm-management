VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDrug"
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

Public DRUG_ID As Long
Public DRUG_NAME As String
Public DRUG_GROUP_ID As Long

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
   
   FromStr = "FROM DRUG UG "
   
   WhereStr = ""
   If DRUG_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DRUG_ID = " & DRUG_ID & ") "
      Else
         WhereStr = "AND (DRUG_ID = " & DRUG_ID & ") "
      End If
   End If
   
   If DRUG_GROUP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DRUG_GROUP_ID = " & DRUG_GROUP_ID & ") "
      Else
         WhereStr = "AND (DRUG_GROUP_ID = " & DRUG_GROUP_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY DRUG_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY DRUG_NAME "
   Else
      OrderByStr = "ORDER BY DRUG_GROUP_ID, DRUG_ID "
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
   DRUG_ID = NVLI(Rs("DRUG_ID"), -1)
   DRUG_NAME = NVLS(Rs("DRUG_NAME"), "")
   DRUG_GROUP_ID = NVLI(Rs("DRUG_GROUP_ID"), -1)
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("DRUG_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO DRUG " & _
                     "( " & _
                     "   DRUG_ID, " & _
                     "   DRUG_NAME, " & _
                     "   DRUG_GROUP_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(DRUG_NAME)) & "'," & _
                         DRUG_GROUP_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         DRUG_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE DRUG SET " & _
                     " DRUG_GROUP_ID = " & DRUG_GROUP_ID & ", " & _
                     "   DRUG_NAME = '" & ChangeQuote(Trim(DRUG_NAME)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE DRUG_ID = " & DRUG_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM DRUG WHERE DRUG_ID = " & DRUG_ID
      m_Conn.Execute (SQL1)
End Function


