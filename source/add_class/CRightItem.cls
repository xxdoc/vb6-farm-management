VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRightItem"
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

Public RIGHT_ID As Long
Public RIGHT_ITEM_DESC As String
Public DEFAULT_VALUE As String
Public RIGHT_ITEM_NAME As String
Public PARENT_ID As Long

Public KEY_ID As Long
Public KEY_LOOKUP As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
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

   SelectStr1 = "SELECT ADR.* "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM RIGHT_ITEM ADR  "
   
   WhereStr = ""
   
   If RIGHT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (RIGHT_ID = " & RIGHT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (RIGHT_ID = " & RIGHT_ID & ") "
      End If
   End If
                  
   If Len(RIGHT_ITEM_NAME) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (RIGHT_ITEM_NAME = '" & RIGHT_ITEM_NAME & "') "
      Else
         WhereStr = WhereStr & "AND (RIGHT_ITEM_NAME = '" & RIGHT_ITEM_NAME & "') "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY RIGHT_ID "
   Else
      OrderByStr = "ORDER BY RIGHT_ID "
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
   RIGHT_ID = NVLI(Rs("RIGHT_ID"), -1)
   RIGHT_ITEM_DESC = NVLS(Rs("RIGHT_ITEM_DESC"), "")
   RIGHT_ITEM_NAME = NVLS(Rs("RIGHT_ITEM_NAME"), "")
   DEFAULT_VALUE = NVLS(Rs("DEFAULT_VALUE"), "")
   PARENT_ID = NVLI(Rs("PARENT_ID"), 0)
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String

   TempID1 = PARENT_ID
   If PARENT_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("RIGHT_ITEM_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO RIGHT_ITEM " & _
                     "( " & _
                     "   RIGHT_ID, " & _
                     "   RIGHT_ITEM_DESC, " & _
                     "   RIGHT_ITEM_NAME, " & _
                     "   DEFAULT_VALUE, " & _
                     "   PARENT_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(RIGHT_ITEM_DESC)) & "'," & _
                         "'" & ChangeQuote(Trim(RIGHT_ITEM_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(DEFAULT_VALUE)) & "'," & _
                         TempID1 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         RIGHT_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE RIGHT_ITEM SET " & _
                     "   RIGHT_ITEM_DESC = '" & Trim(ChangeQuote(RIGHT_ITEM_DESC)) & "', " & _
                     "   RIGHT_ITEM_NAME = '" & Trim(ChangeQuote(RIGHT_ITEM_NAME)) & "', " & _
                     "   DEFAULT_VALUE = " & DEFAULT_VALUE & ", " & _
                     "   PARENT_ID = " & TempID1 & " , " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE RIGHT_ID = " & RIGHT_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM RIGHT_ITEM WHERE RIGHT_ID = " & RIGHT_ID
      m_Conn.Execute (SQL1)
End Function




