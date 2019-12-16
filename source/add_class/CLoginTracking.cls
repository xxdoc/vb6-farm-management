VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLoginTracking"
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

Public LOGIN_TRACKING_ID As Long
Public LOGIN_DATE As Date
Public LOGOUT_DATE As Date
Public USER_ID As Long
Public USER_NAME As String
Public IP_ADDRESS As String

Public LOGIN_FROM_DATE As Date
Public LOGIN_TO_DATE As Date

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
Dim TempDate As String

   SelectStr1 = "SELECT LT.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM LOGIN_TRACKING LT "
   
   WhereStr = ""
   If LOGIN_TRACKING_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LOGIN_TRACKING_ID = " & LOGIN_TRACKING_ID & ") "
      Else
         WhereStr = "AND (LOGIN_TRACKING_ID = " & LOGIN_TRACKING_ID & ") "
      End If
   End If
   
   If USER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (USER_ID = " & USER_ID & ") "
      Else
         WhereStr = "AND (USER_ID = " & USER_ID & ") "
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
   
   If LOGIN_FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(LOGIN_FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (LOGIN_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (LOGIN_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If LOGIN_TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(LOGIN_TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (LOGIN_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (LOGIN_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY LOGIN_DATE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY USER_NAME "
   ElseIf OrderBy = 3 Then
      OrderByStr = "ORDER BY IP_ADDRESS "
   Else
      OrderByStr = "ORDER BY USER_NAME "
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
   LOGIN_TRACKING_ID = NVLI(Rs("LOGIN_TRACKING_ID"), -1)
   LOGIN_DATE = InternalDateToDate(NVLS(Rs("LOGIN_DATE"), ""))
   LOGOUT_DATE = InternalDateToDate(NVLS(Rs("LOGOUT_DATE"), ""))
   USER_ID = NVLI(Rs("USER_ID"), -1)
   USER_NAME = NVLS(Rs("USER_NAME"), "")
   IP_ADDRESS = NVLS(Rs("IP_ADDRESS"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("LOGIN_TRACKING_SEQ", ID, glbErrorLog)
      SQL1 = "INSERT INTO LOGIN_TRACKING " & _
                     "( " & _
                     "   LOGIN_TRACKING_ID, " & _
                     "   LOGIN_DATE, " & _
                     "   LOGOUT_DATE, " & _
                     "   USER_ID, " & _
                     "   USER_NAME, " & _
                     "   IP_ADDRESS, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
               SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(LOGIN_DATE))) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(LOGOUT_DATE))) & "'," & _
                         "" & USER_ID & "," & _
                         "'" & ChangeQuote(Trim(USER_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(IP_ADDRESS)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
               glbUser.USER_ID = ID
               LOGIN_TRACKING_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE LOGIN_TRACKING SET " & _
                     "   LOGIN_DATE = '" & ChangeQuote(Trim(DateToStringInt(LOGIN_DATE))) & "', " & _
                     "   LOGOUT_DATE = '" & ChangeQuote(Trim(DateToStringInt(LOGOUT_DATE))) & "', " & _
                     "   USER_ID = " & USER_ID & ", " & _
                     "   USER_NAME = '" & ChangeQuote(Trim(USER_NAME)) & "', " & _
                     "   IP_ADDRESS = '" & ChangeQuote(Trim(IP_ADDRESS)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE LOGIN_TRACKING_ID = " & LOGIN_TRACKING_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM LOGIN_TRACKING WHERE LOGIN_TRACKING_ID = " & LOGIN_TRACKING_ID
      m_Conn.Execute (SQL1)
End Function



