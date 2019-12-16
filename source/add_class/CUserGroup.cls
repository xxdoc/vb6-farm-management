VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUserGroup"
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

Public GROUP_ID As Long
Public GROUP_NAME As String
Public GROUP_DESC As String
Public MAX_USER As Long
Public GROUP_STATUS As String

Public USER_COUNT As Long

Public RightItems As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set RightItems = New Collection
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

   SelectStr1 = "SELECT UG.*, (SELECT COUNT (*) FROM USER_ACCOUNT UA WHERE UA.GROUP_ID = UG.GROUP_ID)  USER_COUNT "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM USER_GROUP UG "
   
   WhereStr = ""
   If GROUP_ID >= 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (GROUP_ID = " & GROUP_ID & ") "
      Else
         WhereStr = "AND (GROUP_ID = " & GROUP_ID & ") "
      End If
   End If
   
   If Trim(GROUP_NAME) <> "" Then
      If WildCard(GROUP_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(GROUP_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(GROUP_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (GROUP_NAME = '" & ChangeQuote(Trim(GROUP_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (GROUP_NAME = '" & ChangeQuote(Trim(GROUP_NAME)) & "')"
         End If
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY GROUP_NAME "
   Else
      OrderByStr = "ORDER BY GROUP_ID "
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

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   GROUP_ID = NVLI(Rs("GROUP_ID"), 0)
   GROUP_NAME = NVLS(Rs("GROUP_NAME"), "")
   GROUP_DESC = NVLS(Rs("GROUP_DESC"), "")
   MAX_USER = NVLI(Rs("MAX_USER"), 0)
   GROUP_STATUS = NVLS(Rs("GROUP_STATUS"), "")
   
   USER_COUNT = NVLI(Rs("USER_COUNT"), 0)
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("USER_GROUP_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO USER_GROUP " & _
                     "( " & _
                     "   GROUP_ID, " & _
                     "   GROUP_NAME, " & _
                     "   GROUP_DESC, " & _
                     "   MAX_USER, " & _
                     "   GROUP_STATUS, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(GROUP_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(GROUP_DESC)) & "'," & _
                         MAX_USER & "," & _
                         "'" & ChangeQuote(Trim(GROUP_STATUS)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         GROUP_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE USER_GROUP SET " & _
                     "   GROUP_NAME = '" & ChangeQuote(Trim(GROUP_NAME)) & "', " & _
                     "   GROUP_DESC = '" & ChangeQuote(Trim(GROUP_DESC)) & "', " & _
                     "   MAX_USER = " & MAX_USER & ", " & _
                     "   GROUP_STATUS = '" & ChangeQuote(Trim(GROUP_STATUS)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE GROUP_ID = " & GROUP_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData(Ind As Long)
Dim SQL1 As String

      SQL1 = "DELETE FROM USER_GROUP WHERE GROUP_ID = " & GROUP_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set RightItems = Nothing
End Sub
