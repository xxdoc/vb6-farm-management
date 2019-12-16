VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUserAccount"
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

Public USER_ID As Long
Public USER_NAME As String
Public USER_PASSWORD As String
Public USER_DESC As String
Public USER_STATUS As String
Public GROUP_ID As Long
Public CHECK_EXPIRE As String
Public EXPIRE_DATE As Date
Public FROM_TIME As Date
Public TO_TIME As Date
Public EXCEPTION_FLAG As String

Public GROUP_NAME As String

Public Logins As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set Logins = New Collection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, ItemCount As Long)
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

   If Ind = 1 Then
      SelectStr1 = "SELECT UA.*, UG.GROUP_NAME  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM USER_ACCOUNT UA, USER_GROUP UG "
      
      WhereStr = "WHERE (UA.GROUP_ID = UG.GROUP_ID) "
   End If
   
   If GROUP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (GROUP_ID = " & GROUP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (GROUP_ID = " & GROUP_ID & ") "
      End If
   End If
   
   If USER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (USER_ID = " & USER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (USER_ID = " & USER_ID & ") "
      End If
   End If
   
   If Len(USER_STATUS) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (USER_STATUS = '" & USER_STATUS & "') "
      Else
         WhereStr = WhereStr & "AND (USER_STATUS = '" & USER_STATUS & "') "
      End If
   End If
   
   If Trim(USER_NAME) <> "" Then
      If WildCard(USER_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(USER_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(USER_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (USER_NAME = '" & ChangeQuote(Trim(USER_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (USER_NAME = '" & ChangeQuote(Trim(USER_NAME)) & "')"
         End If
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY USER_NAME "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY USER_NAME "
   Else
      OrderByStr = "ORDER BY GROUP_ID "
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
   If Ind = 1 Then
      USER_ID = NVLI(Rs("USER_ID"), -1)
      USER_NAME = NVLS(Rs("USER_NAME"), "")
      USER_PASSWORD = NVLS(Rs("USER_PASSWORD"), "")
      USER_DESC = NVLS(Rs("USER_DESC"), "")
      USER_STATUS = NVLS(Rs("USER_STATUS"), "")
      GROUP_ID = NVLI(Rs("GROUP_ID"), -1)
      CHECK_EXPIRE = NVLS(Rs("CHECK_EXPIRE"), "")
      EXPIRE_DATE = InternalDateToDate(NVLS(Rs("EXPIRE_DATE"), ""))
      FROM_TIME = InternalDateToDate(NVLS(Rs("FROM_TIME"), ""))
      TO_TIME = InternalDateToDate(NVLS(Rs("TO_TIME"), ""))
      GROUP_NAME = NVLS(Rs("GROUP_NAME"), "")
      EXCEPTION_FLAG = NVLS(Rs("EXCEPTION_FLAG"), "")
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("USER_ACCOUNT_SEQ", ID, glbErrorLog)
   
      SQL1 = "INSERT INTO USER_ACCOUNT " & _
                     "( " & _
                     "   USER_ID, " & _
                     "   USER_NAME, " & _
                     "   USER_PASSWORD, " & _
                     "   USER_DESC, " & _
                     "   USER_STATUS, " & _
                     "   GROUP_ID, " & _
                     "   CHECK_EXPIRE, " & _
                     "   EXPIRE_DATE, " & _
                     "   FROM_TIME, " & _
                     "   TO_TIME, " & _
                     "   EXCEPTION_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(USER_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(EncryptText(USER_PASSWORD))) & "'," & _
                         "'" & ChangeQuote(Trim(USER_DESC)) & "'," & _
                         "'" & ChangeQuote(Trim(USER_STATUS)) & "'," & _
                         "" & GROUP_ID & "," & _
                         "'" & ChangeQuote(Trim(CHECK_EXPIRE)) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(EXPIRE_DATE))) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(FROM_TIME))) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(TO_TIME))) & "'," & _
                         "'" & ChangeQuote(Trim(EXCEPTION_FLAG)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         USER_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE USER_ACCOUNT SET " & _
                     "   USER_NAME = '" & ChangeQuote(Trim(USER_NAME)) & "', " & _
                     "   USER_PASSWORD = '" & ChangeQuote(Trim(EncryptText(USER_PASSWORD))) & "', " & _
                     "   USER_DESC = '" & ChangeQuote(Trim(USER_DESC)) & "', " & _
                     "   USER_STATUS = '" & ChangeQuote(Trim(USER_STATUS)) & "', " & _
                     "   GROUP_ID = " & GROUP_ID & ", " & _
                     "   CHECK_EXPIRE = '" & ChangeQuote(Trim(CHECK_EXPIRE)) & "', " & _
                     "   EXPIRE_DATE = '" & ChangeQuote(Trim(DateToStringInt(EXPIRE_DATE))) & "', " & _
                     "   FROM_TIME = '" & ChangeQuote(Trim(DateToStringInt(FROM_TIME))) & "', " & _
                     "   TO_TIME = '" & ChangeQuote(Trim(DateToStringInt(TO_TIME))) & "', " & _
                     "   EXCEPTION_FLAG = '" & ChangeQuote(Trim(EXCEPTION_FLAG)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE USER_ID = " & USER_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData(Ind As Long)
Dim SQL1 As String
      SQL1 = "DELETE FROM USER_ACCOUNT WHERE USER_ID = " & USER_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set Logins = Nothing
End Sub
