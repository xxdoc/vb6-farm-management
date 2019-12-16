VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCountry"
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

Public COUNTRY_ID As Long
Public COUNTRY_NO As String
Public COUNTRY_NAME As String
Public CONTINENT_ID As Long

Public RightItems As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set RightItems = New Collection
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
   
   FromStr = "FROM COUNTRY UG "
   
   WhereStr = ""
   If COUNTRY_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (COUNTRY_ID = " & COUNTRY_ID & ") "
      Else
         WhereStr = "AND (COUNTRY_ID = " & COUNTRY_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   If OrderBy > 0 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY COUNTRY_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY COUNTRY_NAME "
      Else
         OrderByStr = "ORDER BY COUNTRY_ID "
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

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   COUNTRY_ID = NVLI(Rs("COUNTRY_ID"), -1)
   COUNTRY_NAME = NVLS(Rs("COUNTRY_NAME"), "")
   COUNTRY_NO = NVLS(Rs("COUNTRY_NO"), "")
   CONTINENT_ID = NVLI(Rs("CONTINENT_ID"), -1)
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String

   TempID1 = CONTINENT_ID
   If CONTINENT_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("COUNTRY_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO COUNTRY " & _
                     "( " & _
                     "   COUNTRY_ID, " & _
                     "   COUNTRY_NAME, " & _
                     "   COUNTRY_NO, " & _
                     "   CONTINENT_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(COUNTRY_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(COUNTRY_NO)) & "'," & _
                         TempID1 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         COUNTRY_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE COUNTRY SET " & _
                     "   CONTINENT_ID = " & CONTINENT_ID & ", " & _
                     "   COUNTRY_NAME = '" & ChangeQuote(Trim(COUNTRY_NAME)) & "', " & _
                     "   COUNTRY_NO = '" & ChangeQuote(Trim(COUNTRY_NO)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE COUNTRY_ID = " & COUNTRY_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM COUNTRY WHERE COUNTRY_ID = " & COUNTRY_ID
      m_Conn.Execute (SQL1)
End Function

