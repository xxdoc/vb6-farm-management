VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CName"
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

Public NAME_ID As Long
Public PREFIX_ID As Long
Public LONG_NAME As String
Public SHORT_NAME As String
Public NICK_NAME As String
Public MIDDLE_NAME As String
Public LAST_NAME As String
Public LANGUAGE_ID As Long
Public EMAIL As String

Public PREFIX_NAME As String
Public ENTERPRISE_ID As Long
Public MASTER_FLAG As String
Public ADDRESS_ID As Long

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
Dim SubLen As Long
Dim NewStr As String

   SelectStr1 = "SELECT NM.*, NP.PREFIX_NAME  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM NAME NM "
   FromStr = FromStr & "LEFT OUTER JOIN NAME_PREFIX NP ON (NP.PREFIX_ID = NM.PREFIX_ID) "
   
   WhereStr = ""
   
   If NAME_ID >= 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (NAME_ID = " & NAME_ID & ") "
      Else
         WhereStr = WhereStr & "AND (NAME_ID = " & NAME_ID & ") "
      End If
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""
   If OrderBy > 0 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY NAME_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY NAME_ID "
      Else
         OrderByStr = "ORDER BY NAME_ID "
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

'In this class we will not handle any error or any transaction
Public Function QueryData2(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT NM.*, NP.PREFIX_NAME, AD.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM NAME NM "
   FromStr = FromStr & "LEFT OUTER JOIN NAME_PREFIX NP ON (NP.PREFIX_ID = NM.PREFIX_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN ENTERPRISE_PERSON EP ON (EP.NAME_ID = NM.NAME_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN ADDRESS AD ON (EP.ADDRESS_ID = AD.ADDRESS_ID) "
   
   WhereStr = ""
   
   If NAME_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (NAME_ID = " & NAME_ID & ") "
      Else
         WhereStr = WhereStr & "AND (NAME_ID = " & NAME_ID & ") "
      End If
   End If
      
   If ENTERPRISE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ENTERPRISE_ID = " & ENTERPRISE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (ENTERPRISE_ID = " & ENTERPRISE_ID & ") "
      End If
   End If
      
   If Len(MASTER_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MASTER_FLAG = '" & MASTER_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (MASTER_FLAG = '" & MASTER_FLAG & "') "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY NAME_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY NAME_ID "
   Else
      OrderByStr = "ORDER BY NAME_ID "
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
   NAME_ID = NVLI(Rs("NAME_ID"), -1)
   PREFIX_ID = NVLI(Rs("PREFIX_ID"), -1)
   LONG_NAME = NVLS(Rs("LONG_NAME"), "")
   SHORT_NAME = NVLS(Rs("SHORT_NAME"), "")
   NICK_NAME = NVLS(Rs("NICK_NAME"), "")
   MIDDLE_NAME = NVLS(Rs("MIDDLE_NAME"), "")
   LAST_NAME = NVLS(Rs("LAST_NAME"), "")
   EMAIL = NVLS(Rs("EMAIL"), "")
   PREFIX_NAME = NVLS(Rs("PREFIX_NAME"), "")
End Function

Public Function PopulateFromRS2(Rs As ADODB.Recordset)
   NAME_ID = NVLI(Rs("NAME_ID"), -1)
   PREFIX_ID = NVLI(Rs("PREFIX_ID"), -1)
   LONG_NAME = NVLS(Rs("LONG_NAME"), "")
   SHORT_NAME = NVLS(Rs("SHORT_NAME"), "")
   NICK_NAME = NVLS(Rs("NICK_NAME"), "")
   MIDDLE_NAME = NVLS(Rs("MIDDLE_NAME"), "")
   LAST_NAME = NVLS(Rs("LAST_NAME"), "")
   EMAIL = NVLS(Rs("EMAIL"), "")
   PREFIX_NAME = NVLS(Rs("PREFIX_NAME"), "")
   ADDRESS_ID = NVLI(Rs("ADDRESS_ID"), -1)
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("NAME_SEQ", ID, glbErrorLog)
   
      SQL1 = "INSERT INTO NAME " & _
                     "( " & _
                     "   NAME_ID, " & _
                     "   PREFIX_ID, " & _
                     "   LONG_NAME, " & _
                     "   SHORT_NAME, " & _
                     "   NICK_NAME, " & _
                     "   MIDDLE_NAME, " & _
                     "   LAST_NAME, " & _
                     "   EMAIL, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "" & PREFIX_ID & "," & _
                         "'" & ChangeQuote(Trim(LONG_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(SHORT_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(NICK_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(MIDDLE_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(LAST_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(EMAIL)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         NAME_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE NAME SET " & _
                     "   PREFIX_ID = " & PREFIX_ID & ", " & _
                     "   LONG_NAME = '" & ChangeQuote(Trim(LONG_NAME)) & "', " & _
                     "   SHORT_NAME = '" & ChangeQuote(Trim(SHORT_NAME)) & "', " & _
                     "   NICK_NAME = '" & ChangeQuote(Trim(NICK_NAME)) & "', " & _
                     "   MIDDLE_NAME = '" & ChangeQuote(Trim(MIDDLE_NAME)) & "', " & _
                     "   LAST_NAME = '" & ChangeQuote(Trim(LAST_NAME)) & "', " & _
                     "   EMAIL = '" & ChangeQuote(Trim(EMAIL)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE NAME_ID = " & NAME_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM NAME WHERE NAME_ID = " & NAME_ID
      m_Conn.Execute (SQL1)
End Function


