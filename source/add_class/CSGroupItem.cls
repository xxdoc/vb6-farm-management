VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSGroupItem"
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

Public SGROUP_ITEM_ID As Long
Public STATUS_GROUP_ID As Long
Public STATUS_ID As Long
Public SELECT_FLAG As String
Public GUI_ID As Long

Public ST_STATUS_ID As Long
Public STATUS_NO As String
Public STATUS_NAME As String

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

   SelectStr1 = "SELECT PS.PRODUCT_STATUS_NO STATUS_NO, PS.PRODUCT_STATUS_NAME STATUS_NAME, PS.PRODUCT_STATUS_ID ST_STATUS_ID, HGI.STATUS_ID HGI_STATUS_ID, HGI.* "
   
   FromStr = "FROM SGROUP_ITEM HGI "
   FromStr = FromStr & "LEFT OUTER JOIN STATUS_GROUP HG ON (HGI.STATUS_GROUP_ID = HG.STATUS_GROUP_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PS ON (HGI.STATUS_ID = PS.PRODUCT_STATUS_ID) "

   SelectStr2 = "SELECT COUNT(*)  "
      
   WhereStr = ""
   
   If SGROUP_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (HGI.SGROUP_ITEM_ID = " & SGROUP_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (HGI.SGROUP_ITEM_ID = " & SGROUP_ITEM_ID & ") "
      End If
   End If
   
   If STATUS_GROUP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (HGI.STATUS_GROUP_ID = " & STATUS_GROUP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (HGI.STATUS_GROUP_ID = " & STATUS_GROUP_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PRODUCT_STATUS_NO, GUI_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY STATUS_GROUP_ID, GUI_ID "
   Else
      OrderByStr = "ORDER BY HGI.SGROUP_ITEM_ID, GUI_ID "
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
   SGROUP_ITEM_ID = NVLI(Rs("SGROUP_ITEM_ID"), -1)
   STATUS_GROUP_ID = NVLI(Rs("STATUS_GROUP_ID"), -1)
   STATUS_ID = NVLI(Rs("HGI_STATUS_ID"), -1)
   GUI_ID = NVLI(Rs("GUI_ID"), -1)
   
   ST_STATUS_ID = NVLI(Rs("ST_STATUS_ID"), -1)
   STATUS_NO = NVLS(Rs("STATUS_NO"), "")
   STATUS_NAME = NVLS(Rs("STATUS_NAME"), "")
   
   If STATUS_ID <> ST_STATUS_ID Then
      SELECT_FLAG = "N"
   Else
      SELECT_FLAG = NVLS(Rs("SELECT_FLAG"), "")
   End If
End Function

Public Function CopyField(Ind As Long, D As CSGroupItem)
   SGROUP_ITEM_ID = D.SGROUP_ITEM_ID
   STATUS_GROUP_ID = D.STATUS_GROUP_ID
   STATUS_ID = D.STATUS_ID
   GUI_ID = D.GUI_ID
   
   ST_STATUS_ID = D.ST_STATUS_ID
   STATUS_NO = D.STATUS_NO
   STATUS_NAME = D.STATUS_NAME
   SELECT_FLAG = D.SELECT_FLAG
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String

   TempID1 = SGROUP_ITEM_ID
   If SGROUP_ITEM_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = GUI_ID
   If GUI_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("SGROUP_ITEM_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO SGROUP_ITEM " & _
                     "( " & _
                     "   SGROUP_ITEM_ID, " & _
                     "   STATUS_GROUP_ID, " & _
                     "   STATUS_ID, " & _
                     "   SELECT_FLAG, " & _
                     "   GUI_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(STATUS_GROUP_ID)) & "'," & _
                         "'" & ChangeQuote(Trim(STATUS_ID)) & "'," & _
                         "'" & ChangeQuote(Trim(SELECT_FLAG)) & "'," & _
                        TempID2 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         SGROUP_ITEM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE SGROUP_ITEM SET " & _
                     "   STATUS_GROUP_ID = '" & ChangeQuote(Trim(STATUS_GROUP_ID)) & "', " & _
                     "   STATUS_ID = '" & ChangeQuote(Trim(STATUS_ID)) & "', " & _
                     "   SELECT_FLAG = '" & ChangeQuote(Trim(SELECT_FLAG)) & "', " & _
                     "    GUI_ID = " & TempID2 & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE SGROUP_ITEM_ID = " & SGROUP_ITEM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM SGROUP_ITEM WHERE SGROUP_ITEM_ID = " & SGROUP_ITEM_ID
      m_Conn.Execute (SQL1)
End Function


