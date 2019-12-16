VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBloodSpec"
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

Public BLOOD_SPEC_ID As Long
Public SPEC_NAME As String
Public MALE_STD As String
Public FEMALE_STD As String

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
   
   FromStr = "FROM BLOOD_SPEC UG "
   
   WhereStr = ""
   If BLOOD_SPEC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BLOOD_SPEC_ID = " & BLOOD_SPEC_ID & ") "
      Else
         WhereStr = "AND (BLOOD_SPEC_ID = " & BLOOD_SPEC_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY BLOOD_SPEC_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY SPEC_NAME "
   Else
      OrderByStr = "ORDER BY SPEC_NAME, BLOOD_SPEC_ID "
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
   BLOOD_SPEC_ID = NVLI(Rs("BLOOD_SPEC_ID"), -1)
   SPEC_NAME = NVLS(Rs("SPEC_NAME"), "")
   MALE_STD = NVLS(Rs("MALE_STD"), "")
   FEMALE_STD = NVLS(Rs("FEMALE_STD"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("BLOOD_SPEC_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO BLOOD_SPEC " & _
                     "( " & _
                     "   BLOOD_SPEC_ID, " & _
                     "   SPEC_NAME, " & _
                     "   MALE_STD, " & _
                     "   FEMALE_STD, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(SPEC_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(MALE_STD)) & "'," & _
                         "'" & ChangeQuote(Trim(FEMALE_STD)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         BLOOD_SPEC_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE BLOOD_SPEC SET " & _
                     "   SPEC_NAME = '" & ChangeQuote(Trim(SPEC_NAME)) & "', " & _
                     "   MALE_STD = '" & ChangeQuote(Trim(MALE_STD)) & "', " & _
                     "   FEMALE_STD = '" & ChangeQuote(Trim(FEMALE_STD)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE BLOOD_SPEC_ID = " & BLOOD_SPEC_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM BLOOD_SPEC WHERE BLOOD_SPEC_ID = " & BLOOD_SPEC_ID
      m_Conn.Execute (SQL1)
End Function




