VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPackage"
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
'
Public PKG_ID As Long
Public PKG_CODE As String
Public PKG_NAME As String
Public PKG_BASIC_FLAG As String
Public PKG_TYPE As Long

Public PACKAGE_TYPE_NAME As String

Public KEY_ID As Long
Public KEY_LOOKUP As String

Public PackageDetail As Collection


Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
    
    
    
    Set PackageDetail = New Collection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, itemcount As Long)
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

   SelectStr1 = "SELECT PKG.*,PKT.PACKAGE_TYPE_NAME "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PACKAGE PKG "
   FromStr = FromStr & "LEFT OUTER JOIN PACKAGE_TYPE PKT ON (PKG.PKG_TYPE = PKT.PACKAGE_TYPE_ID) "
   
   WhereStr = ""
   
   If PKG_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PKG_ID = " & PKG_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PKG_ID = " & PKG_ID & ") "
      End If
   End If
   
   If PKG_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PKG_TYPE = " & PKG_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (PKG_TYPE = " & PKG_TYPE & ") "
      End If
   End If
      
   If Len(PKG_BASIC_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PKG_BASIC_FLAG = '" & PKG_BASIC_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (PKG_BASIC_FLAG = '" & PKG_BASIC_FLAG & "') "
      End If
   End If
   
   If Len(Trim(PKG_CODE)) > 0 Then
      If WildCard(PKG_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(PKG_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(PKG_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PKG_CODE = '" & ChangeQuote(Trim(PKG_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (PKG_CODE = '" & ChangeQuote(Trim(PKG_CODE)) & "')"
         End If
      End If
   End If

   If Len(Trim(PKG_NAME)) > 0 Then
      If WildCard(PKG_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(PKG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(PKG_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PKG_NAME = '" & ChangeQuote(Trim(PKG_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (PKG_NAME = '" & ChangeQuote(Trim(PKG_NAME)) & "')"
         End If
      End If
   End If

   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PKG_CODE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PKG_NAME "
    ElseIf OrderBy = 3 Then
      OrderByStr = "ORDER BY PKG_TYPE "
   Else
      OrderByStr = "ORDER BY PKG_ID "
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   PKG_ID = NVLI(Rs("PKG_ID"), -1)
   PKG_CODE = NVLS(Rs("PKG_CODE"), "")
   PKG_NAME = NVLS(Rs("PKG_NAME"), "")
   PKG_TYPE = NVLI(Rs("PKG_TYPE"), -1)
   PKG_BASIC_FLAG = NVLS(Rs("PKG_BASIC_FLAG"), "")
   
   PACKAGE_TYPE_NAME = NVLS(Rs("PACKAGE_TYPE_NAME"), "")
   
   KEY_ID = PKG_ID
   KEY_LOOKUP = PKG_CODE
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String

   TempID1 = PKG_TYPE
   If PKG_TYPE <= 0 Then
      TempID1 = "NULL "
   End If
      
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("PACKAGE_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO PACKAGE " & _
                     "( " & _
                     "   PKG_ID, " & _
                     "   PKG_CODE, " & _
                     "   PKG_NAME, " & _
                     "   PKG_TYPE, " & _
                     "   PKG_BASIC_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(PKG_CODE)) & "'," & _
                         "'" & ChangeQuote(Trim(PKG_NAME)) & "'," & _
                         "" & TempID1 & "," & _
                         "'" & ChangeQuote(Trim(PKG_BASIC_FLAG)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         PKG_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE PACKAGE SET " & _
                     "   PKG_CODE = '" & ChangeQuote(Trim(PKG_CODE)) & "', " & _
                     "   PKG_NAME = '" & ChangeQuote(Trim(PKG_NAME)) & "', " & _
                     "   PKG_TYPE = " & TempID1 & ", " & _
                     "   PKG_BASIC_FLAG = '" & ChangeQuote(Trim(PKG_BASIC_FLAG)) & "', " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PKG_ID = " & PKG_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM PACKAGE_DETAIL WHERE PKG_ID = " & PKG_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM PACKAGE WHERE PKG_ID = " & PKG_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
    Set PackageDetail = Nothing
End Sub
