VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGroupRight"
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

Public RTYPE_ID    As Long
Public RTYPE_NAME As String

Public GROUP_RIGHT_ID As Long
Public GROUP_ID As Long
Public RIGHT_ID  As Long
Public RIGHT_STATUS As String

Public RIGHT_ITEM_DESC As String
Public PARENT_ID As Long
Public RIGHT_ITEM_NAME As String

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

   If Ind = 1 Then
      SelectStr1 = "SELECT GR.*  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM GROUP_RIGHT GR "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT RI.*, GR.*, GR.RIGHT_ID RIGHT_ID2, GR.RIGHT_STATUS RIGHT_STATUS2, RI.RIGHT_ID RIGHT_ID1, RI.DEFAULT_VALUE RIGHT_STATUS1 "
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = "FROM RIGHT_ITEM RI "
      FromStr = FromStr & "LEFT OUTER JOIN GROUP_RIGHT GR ON ((RI.RIGHT_ID = GR.RIGHT_ID) AND (GR.GROUP_ID = " & GROUP_ID & "))"
   
      WhereStr = ""
   End If
   
   If GROUP_RIGHT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (GROUP_RIGHT_ID = " & GROUP_RIGHT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (GROUP_RIGHT_ID = " & GROUP_RIGHT_ID & ") "
      End If
   End If
   
'   If GROUP_ID > 0 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (GROUP_ID = " & GROUP_ID & ") "
'      Else
'         WhereStr = WhereStr & "AND (GROUP_ID = " & GROUP_ID & ") "
'      End If
'   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY GR.GROUP_RIGHT_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY GR.GROUP_ID "
   Else
      OrderByStr = "ORDER BY RI.RIGHT_ID "
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
'
''In this class we will not handle any error or any transaction
'Public Function QueryData2(Rs As ADODB.Recordset, ItemCount As Long)
'Dim SQL1 As String
'Dim SQL2 As String
'Dim SelectStr1 As String
'Dim SelectStr2 As String
'Dim FromStr As String
'Dim WhereStr As String
'Dim OrderByStr As String
'Dim OrderTypeStr As String
'
'   SelectStr1 = "SELECT RI.*, GR.RIGHT_ID RIGHT_ID2  "
'   SelectStr2 = "SELECT COUNT(*)  "
'
'   FromStr = "FROM RIGHT_ITEM RI "
'   FromStr = "LEFT OUTER JOIN GROUP_RIGHT GR ON (RI.RIGHT_ID = GR.RIGHT_ID) "
'
'   WhereStr = ""
'   If GROUP_ID > 0 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (GROUP_ID = " & GROUP_ID & ") "
'      Else
'         WhereStr = "AND (GROUP_ID = " & GROUP_ID & ") "
'      End If
'   End If
'
'   OrderByStr = ""
'   OrderTypeStr = ""
'
'   If OrderBy = 1 Then
'      OrderByStr = "ORDER BY GROUP_RIGHT_ID "
'   ElseIf OrderBy = 2 Then
'      OrderByStr = "ORDER BY GROUP_ID "
'   Else
'      OrderByStr = "ORDER BY RIGHT_ID "
'   End If
'
'   If OrderType = 1 Then
'      OrderTypeStr = "ASC "
'   ElseIf OrderType = 2 Then
'      OrderTypeStr = "DESC "
'   Else
'      OrderTypeStr = "DESC "
'   End If
'
'   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
'   SQL2 = SelectStr2 & FromStr & WhereStr
'
'   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
'End Function
'
Public Function QueryData3(Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String

   SelectStr1 = "SELECT GR.*, GR.RIGHT_ID RIGHT_ID_ORG, RI.RIGHT_ITEM_NAME, RI.DEFAULT_VALUE  "
   SelectStr2 = "SELECT COUNT(*)  "

   FromStr = "FROM GROUP_RIGHT GR "
   FromStr = FromStr & "LEFT OUTER JOIN RIGHT_ITEM RI ON (GR.RIGHT_ID = RI.RIGHT_ID) "

   WhereStr = ""
   If GROUP_RIGHT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (GROUP_RIGHT_ID = " & GROUP_RIGHT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (GROUP_RIGHT_ID = " & GROUP_RIGHT_ID & ") "
      End If
   End If

   If GROUP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (GROUP_ID = " & GROUP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (GROUP_ID = " & GROUP_ID & ") "
      End If
   End If

   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY GROUP_RIGHT_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY GROUP_ID "
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
   If Ind = 1 Then
      GROUP_RIGHT_ID = NVLI(Rs("RIGHT_ID"), -1)
      GROUP_ID = NVLI(Rs("RIGHT_ID"), -1)
      RIGHT_ID = NVLI(Rs("RIGHT_ID"), -1)
      RIGHT_STATUS = NVLS(Rs("RIGHT_ITEM_NAME"), "")
   ElseIf Ind = 2 Then
         GROUP_ID = NVLI(Rs("GROUP_ID"), -1)
         GROUP_RIGHT_ID = NVLI(Rs("GROUP_RIGHT_ID"), -1)
         PARENT_ID = NVLI(Rs("PARENT_ID"), 0)
         If NVLI(Rs("RIGHT_ID1"), 0) <> NVLI(Rs("RIGHT_ID2"), 0) Then
            RIGHT_ID = NVLI(Rs("RIGHT_ID1"), 0)
            RIGHT_STATUS = NVLS(Rs("RIGHT_STATUS1"), "")
            Flag = "A"
         Else
            RIGHT_ID = NVLI(Rs("RIGHT_ID2"), 0)
            RIGHT_STATUS = NVLS(Rs("RIGHT_STATUS2"), "")
            Flag = "I"
         End If
         RIGHT_ITEM_DESC = NVLS(Rs("RIGHT_ITEM_DESC"), "")
         RIGHT_ITEM_NAME = NVLS(Rs("RIGHT_ITEM_NAME"), "")
   End If
End Function

Public Function PopulateFromRS3(Rs As ADODB.Recordset)
   GROUP_RIGHT_ID = NVLI(Rs("GROUP_RIGHT_ID"), -1)
   GROUP_ID = NVLI(Rs("GROUP_ID"), -1)
   RIGHT_ID = NVLI(Rs("RIGHT_ID_ORG"), -1)
   RIGHT_STATUS = NVLS(Rs("RIGHT_STATUS"), "")
   If RIGHT_ID <= 0 Then
      RIGHT_STATUS = "Y"
   End If
   RIGHT_ITEM_NAME = NVLS(Rs("RIGHT_ITEM_NAME"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("GROUP_RIGHT_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO GROUP_RIGHT " & _
                     "( " & _
                     "   GROUP_RIGHT_ID, " & _
                     "   GROUP_ID, " & _
                     "   RIGHT_ID, " & _
                     "   RIGHT_STATUS, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "" & GROUP_ID & "," & _
                         "" & RIGHT_ID & "," & _
                         "'" & ChangeQuote(Trim(RIGHT_STATUS)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
   Else 'Show edit
      SQL1 = "UPDATE GROUP_RIGHT SET " & _
                     "   GROUP_ID = " & GROUP_ID & ", " & _
                     "   RIGHT_ID = " & RIGHT_ID & ", " & _
                     "   RIGHT_STATUS = '" & ChangeQuote(Trim(RIGHT_STATUS)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE GROUP_RIGHT_ID = " & GROUP_RIGHT_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM GROUP_RIGHT WHERE GROUP_RIGHT_ID = " & GROUP_RIGHT_ID
      m_Conn.Execute (SQL1)
End Function


