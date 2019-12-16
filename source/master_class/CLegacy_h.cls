VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLegacy_h"
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

Public LEGACY_H_ID As Long
Public DOCUMENT_ID As Long
Public DOCUMENT_SEQ As Long
Public DOCUMENT_DATE As String
Public H_FIELD1 As String
Public H_FIELD2 As String
Public H_FIELD3 As String
Public H_FIELD4 As String
Public H_FIELD5 As String
Public H_FIELD6 As String
Public H_FIELD7 As String
Public H_FIELD8 As String
Public H_FIELD9 As String
Public H_FIELD10 As String

Public LegacyDetails As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set LegacyDetails = New Collection
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
   
   FromStr = "FROM LEGACY_H UG "
   
   WhereStr = ""
   If LEGACY_H_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LEGACY_H_ID = " & LEGACY_H_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LEGACY_H_ID = " & LEGACY_H_ID & ") "
      End If
   End If
   
   If DOCUMENT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (DOCUMENT_ID = " & DOCUMENT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_ID = " & DOCUMENT_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY LEGACY_H_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY DOCUMENT_DATE "
   Else
      OrderByStr = "ORDER BY LEGACY_H_ID "
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
   LEGACY_H_ID = NVLI(Rs("LEGACY_H_ID"), -1)
   H_FIELD1 = NVLS(Rs("H_FIELD1"), "")
   H_FIELD2 = NVLS(Rs("H_FIELD2"), "")
   H_FIELD3 = NVLS(Rs("H_FIELD3"), "")
   H_FIELD4 = NVLS(Rs("H_FIELD4"), "")
   H_FIELD5 = NVLS(Rs("H_FIELD5"), "")
   H_FIELD6 = NVLS(Rs("H_FIELD6"), "")
   H_FIELD7 = NVLS(Rs("H_FIELD7"), "")
   H_FIELD8 = NVLS(Rs("H_FIELD8"), "")
   H_FIELD9 = NVLS(Rs("H_FIELD9"), "")
   H_FIELD10 = NVLS(Rs("H_FIELD10"), "")
   
   DOCUMENT_ID = NVLI(Rs("DOCUMENT_ID"), -1)
   DOCUMENT_SEQ = NVLI(Rs("DOCUMENT_SEQ"), -1)
   DOCUMENT_DATE = NVLS(Rs("DOCUMENT_SEQ"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("LEGACY_H_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO LEGACY_H " & _
                     "( " & _
                     "   LEGACY_H_ID, " & _
                     "   DOCUMENT_ID, " & _
                     "   DOCUMENT_SEQ, " & _
                     "   H_FIELD1, " & _
                     "   H_FIELD2, " & _
                     "   H_FIELD3, " & _
                     "   H_FIELD4, " & _
                     "   H_FIELD5, " & _
                     "   H_FIELD6, " & _
                     "   H_FIELD7, " & _
                     "   H_FIELD8, " & _
                     "   H_FIELD9, " & _
                     "   H_FIELD10, "
            SQL1 = SQL1 & _
                     "   DOCUMENT_DATE, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         DOCUMENT_ID & "," & _
                         DOCUMENT_SEQ & "," & _
                         "'" & ChangeQuote(Trim(H_FIELD1)) & "'," & _
                         "'" & ChangeQuote(Trim(H_FIELD2)) & "'," & _
                         "'" & ChangeQuote(Trim(H_FIELD3)) & "'," & _
                         "'" & ChangeQuote(Trim(H_FIELD4)) & "'," & _
                         "'" & ChangeQuote(Trim(H_FIELD5)) & "'," & _
                         "'" & ChangeQuote(Trim(H_FIELD6)) & "'," & _
                         "'" & ChangeQuote(Trim(H_FIELD7)) & "'," & _
                         "'" & ChangeQuote(Trim(H_FIELD8)) & "'," & _
                         "'" & ChangeQuote(Trim(H_FIELD9)) & "'," & _
                         "'" & ChangeQuote(Trim(H_FIELD10)) & "',"
               SQL1 = SQL1 & _
                        "'" & ChangeQuote(Trim(DOCUMENT_DATE)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         LEGACY_H_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE LEGACY_H SET " & _
                     "   H_FIELD1 = '" & ChangeQuote(Trim(H_FIELD1)) & "', " & _
                     "   H_FIELD2 = '" & ChangeQuote(Trim(H_FIELD2)) & "', " & _
                     "   H_FIELD3 = '" & ChangeQuote(Trim(H_FIELD3)) & "', " & _
                     "   H_FIELD4 = '" & ChangeQuote(Trim(H_FIELD4)) & "', " & _
                     "   H_FIELD5 = '" & ChangeQuote(Trim(H_FIELD5)) & "', " & _
                     "   H_FIELD6 = '" & ChangeQuote(Trim(H_FIELD6)) & "', " & _
                     "   H_FIELD7 = '" & ChangeQuote(Trim(H_FIELD7)) & "', " & _
                     "   H_FIELD8 = '" & ChangeQuote(Trim(H_FIELD8)) & "', " & _
                     "   H_FIELD9 = '" & ChangeQuote(Trim(H_FIELD9)) & "', " & _
                     "   H_FIELD10 = '" & ChangeQuote(Trim(H_FIELD10)) & "', " & _
                     "   DOCUMENT_DATE = '" & ChangeQuote(Trim(DOCUMENT_DATE)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE LEGACY_H_ID = " & LEGACY_H_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function SetDocType()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE LEGACY_H LGH SET LGH.DOCUMENT_ID = 6 WHERE LGH.LEGACY_H_ID IN "
   SQL1 = SQL1 & "( "
   SQL1 = SQL1 & "   SELECT DISTINCT (LGD.LEGACY_H_ID) "
   SQL1 = SQL1 & "   FROM LEGACY_D LGD "
   SQL1 = SQL1 & "   Where (LGD.D_FIELD2 <> LGD.D_FIELD3) "
   SQL1 = SQL1 & ") "
   SQL1 = SQL1 & "AND (LGH.DOCUMENT_ID = 0) "
   m_Conn.Execute (SQL1)

   SQL1 = "UPDATE LEGACY_H LGH SET LGH.DOCUMENT_ID = 7 WHERE LGH.LEGACY_H_ID IN "
   SQL1 = SQL1 & "( "
   SQL1 = SQL1 & "   SELECT DISTINCT (LGD.LEGACY_H_ID) "
   SQL1 = SQL1 & "   FROM LEGACY_D LGD "
   SQL1 = SQL1 & "   Where (LGD.D_FIELD2 = LGD.D_FIELD3) "
   SQL1 = SQL1 & ") "
   SQL1 = SQL1 & "AND (LGH.DOCUMENT_ID = 0) "
   m_Conn.Execute (SQL1)
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM LEGACY_H WHERE LEGACY_H_ID = " & LEGACY_H_ID
      m_Conn.Execute (SQL1)
End Function

Public Function DeleteAllData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM LEGACY_D "
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM LEGACY_H "
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set LegacyDetails = Nothing
End Sub
