VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLegacy_d"
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

Public LEGACY_D_ID As Long
Public LEGACY_H_ID As Long
Public D_FIELD1 As String
Public D_FIELD2 As String
Public D_FIELD3 As String
Public D_FIELD4 As String
Public D_FIELD5 As String
Public D_FIELD6 As String
Public D_FIELD7 As String
Public D_FIELD8 As String
Public D_FIELD9 As String
Public D_FIELD10 As String
Public D_FIELD11 As String

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
   
   FromStr = "FROM LEGACY_D UG "
   
   WhereStr = ""
   If LEGACY_D_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LEGACY_D_ID = " & LEGACY_D_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LEGACY_D_ID = " & LEGACY_D_ID & ") "
      End If
   End If
   
   If LEGACY_H_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LEGACY_H_ID = " & LEGACY_H_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LEGACY_H_ID = " & LEGACY_H_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY LEGACY_D_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY D_FIELD2 "
   Else
      OrderByStr = "ORDER BY LEGACY_D_ID "
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
   LEGACY_D_ID = NVLI(Rs("LEGACY_D_ID"), -1)
   LEGACY_H_ID = NVLI(Rs("LEGACY_H_ID"), -1)
   D_FIELD1 = NVLS(Rs("D_FIELD1"), "")
   D_FIELD2 = NVLS(Rs("D_FIELD2"), "")
   D_FIELD3 = NVLS(Rs("D_FIELD3"), "")
   D_FIELD4 = NVLS(Rs("D_FIELD4"), "")
   D_FIELD5 = NVLS(Rs("D_FIELD5"), "")
   D_FIELD6 = NVLS(Rs("D_FIELD6"), "")
   D_FIELD7 = NVLS(Rs("D_FIELD7"), "")
   D_FIELD8 = NVLS(Rs("D_FIELD8"), "")
   D_FIELD9 = NVLS(Rs("D_FIELD9"), "")
   D_FIELD10 = NVLS(Rs("D_FIELD10"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("LEGACY_D_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO LEGACY_D " & _
                     "( " & _
                     "   LEGACY_D_ID, " & _
                     "   LEGACY_H_ID, " & _
                     "   D_FIELD1, " & _
                     "   D_FIELD2, " & _
                     "   D_FIELD3, " & _
                     "   D_FIELD4, " & _
                     "   D_FIELD5, " & _
                     "   D_FIELD6, " & _
                     "   D_FIELD7, " & _
                     "   D_FIELD8, " & _
                     "   D_FIELD9, " & _
                     "   D_FIELD10, "
            SQL1 = SQL1 & _
                     "   D_FIELD11, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         LEGACY_H_ID & "," & _
                         "'" & ChangeQuote(Trim(D_FIELD1)) & "'," & _
                         "'" & ChangeQuote(Trim(D_FIELD2)) & "'," & _
                         "'" & ChangeQuote(Trim(D_FIELD3)) & "'," & _
                         "'" & ChangeQuote(Trim(D_FIELD4)) & "'," & _
                         "'" & ChangeQuote(Trim(D_FIELD5)) & "'," & _
                         "'" & ChangeQuote(Trim(D_FIELD6)) & "'," & _
                         "'" & ChangeQuote(Trim(D_FIELD7)) & "'," & _
                         "'" & ChangeQuote(Trim(D_FIELD8)) & "'," & _
                         "'" & ChangeQuote(Trim(D_FIELD9)) & "'," & _
                         "'" & ChangeQuote(Trim(D_FIELD10)) & "',"
               SQL1 = SQL1 & _
                        "'" & ChangeQuote(Trim(D_FIELD11)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         LEGACY_D_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE LEGACY_D SET " & _
                     "   D_FIELD1 = '" & ChangeQuote(Trim(D_FIELD1)) & "', " & _
                     "   D_FIELD2 = '" & ChangeQuote(Trim(D_FIELD2)) & "', " & _
                     "   D_FIELD3 = '" & ChangeQuote(Trim(D_FIELD3)) & "', " & _
                     "   D_FIELD4 = '" & ChangeQuote(Trim(D_FIELD4)) & "', " & _
                     "   D_FIELD5 = '" & ChangeQuote(Trim(D_FIELD5)) & "', " & _
                     "   D_FIELD6 = '" & ChangeQuote(Trim(D_FIELD6)) & "', " & _
                     "   D_FIELD7 = '" & ChangeQuote(Trim(D_FIELD7)) & "', " & _
                     "   D_FIELD8 = '" & ChangeQuote(Trim(D_FIELD8)) & "', " & _
                     "   D_FIELD9 = '" & ChangeQuote(Trim(D_FIELD9)) & "', " & _
                     "   D_FIELD10 = '" & ChangeQuote(Trim(D_FIELD10)) & "', " & _
                     "   D_FIELD11 = '" & ChangeQuote(Trim(D_FIELD11)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE LEGACY_D_ID = " & LEGACY_D_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM LEGACY_D WHERE LEGACY_D_ID = " & LEGACY_D_ID
      m_Conn.Execute (SQL1)
End Function



