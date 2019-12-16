VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHGroupItem"
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

Public HGROUP_ITEM_ID As Long
Public HOUSE_GROUP_ID As Long
Public LOCATION_ID As Long
Public SELECT_FLAG As String

Public LC_LOCATION_ID As Long
Public LOCATION_NO As String
Public LOCATION_NAME As String
Public EXTRA_FLAG As String
Public HOUSE_GROUP_NAME As String
Public EXIST_FLAG As String

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

If Ind = 1 Then
   SelectStr1 = "SELECT LC.LOCATION_NO, LC.LOCATION_NAME, LC.LOCATION_ID LC_LOCATION_ID, HGI.LOCATION_ID HGI_LOCATION_ID, HGI.* "
  
   FromStr = "FROM HGROUP_ITEM HGI "
   FromStr = FromStr & "LEFT OUTER JOIN HOUSE_GROUP HG ON (HGI.HOUSE_GROUP_ID = HG.HOUSE_GROUP_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN LOCATION LC ON (HGI.LOCATION_ID = LC.LOCATION_ID) "

   SelectStr2 = "SELECT COUNT(*)  "
      
   WhereStr = "WHERE (LC.LOCATION_TYPE = 1) "
ElseIf Ind = 2 Then
    SelectStr1 = "SELECT HGI.* , HG.HOUSE_GROUP_NAME , HG.EXTRA_FLAG "
  
   FromStr = "FROM HGROUP_ITEM HGI "
   FromStr = FromStr & "LEFT OUTER JOIN HOUSE_GROUP HG ON (HGI.HOUSE_GROUP_ID = HG.HOUSE_GROUP_ID) "
   SelectStr2 = "SELECT COUNT(*)  "
   
   WhereStr = "WHERE (HG.EXTRA_FLAG = 'N') AND (HGI.SELECT_FLAG='Y') "
ElseIf Ind = 3 Then
    SelectStr1 = "SELECT DISTINCT HG.HOUSE_GROUP_ID , HG.HOUSE_GROUP_NAME "
  
   FromStr = "FROM HGROUP_ITEM HGI "
   FromStr = FromStr & "LEFT OUTER JOIN HOUSE_GROUP HG ON (HGI.HOUSE_GROUP_ID = HG.HOUSE_GROUP_ID) "
   SelectStr2 = "SELECT COUNT(*)  "
   
   WhereStr = "WHERE (HG.EXTRA_FLAG = 'N') AND (HGI.SELECT_FLAG='Y') "
End If

   If HGROUP_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (HGI.HGROUPHGI_ITEM_ID = " & HGROUP_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (HGI.HGROUP_ITEM_ID = " & HGROUP_ITEM_ID & ") "
      End If
   End If
   
   If HOUSE_GROUP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (HGI.HOUSE_GROUP_ID = " & HOUSE_GROUP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (HGI.HOUSE_GROUP_ID = " & HOUSE_GROUP_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY HGI.LOCATION_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY HGI.HOUSE_GROUP_ID "
   Else
      OrderByStr = "ORDER BY HGI.HGROUP_ITEM_ID "
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
 If Ind = 1 Then
   HGROUP_ITEM_ID = NVLI(Rs("HGROUP_ITEM_ID"), -1)
   HOUSE_GROUP_ID = NVLI(Rs("HOUSE_GROUP_ID"), -1)
   LOCATION_ID = NVLI(Rs("HGI_LOCATION_ID"), -1)
   
   LC_LOCATION_ID = NVLI(Rs("LC_LOCATION_ID"), -1)
   LOCATION_NO = NVLS(Rs("LOCATION_NO"), "")
   LOCATION_NAME = NVLS(Rs("LOCATION_NAME"), "")
   
   If LOCATION_ID <> LC_LOCATION_ID Then
      SELECT_FLAG = "N"
   Else
      SELECT_FLAG = NVLS(Rs("SELECT_FLAG"), "")
   End If
ElseIf Ind = 2 Then
   HGROUP_ITEM_ID = NVLI(Rs("HGROUP_ITEM_ID"), -1)
   HOUSE_GROUP_ID = NVLI(Rs("HOUSE_GROUP_ID"), -1)
   HOUSE_GROUP_NAME = NVLS(Rs("HOUSE_GROUP_NAME"), "")
   LOCATION_ID = NVLI(Rs("LOCATION_ID"), -1)
   EXTRA_FLAG = NVLS(Rs("EXTRA_FLAG"), "")
ElseIf Ind = 3 Then
   HOUSE_GROUP_ID = NVLI(Rs("HOUSE_GROUP_ID"), -1)
   HOUSE_GROUP_NAME = NVLS(Rs("HOUSE_GROUP_NAME"), "")
End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String

   TempID1 = HGROUP_ITEM_ID
   If HGROUP_ITEM_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("HGROUP_ITEM_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO HGROUP_ITEM " & _
                     "( " & _
                     "   HGROUP_ITEM_ID, " & _
                     "   HOUSE_GROUP_ID, " & _
                     "   LOCATION_ID, " & _
                     "   SELECT_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(HOUSE_GROUP_ID)) & "'," & _
                         "'" & ChangeQuote(Trim(LOCATION_ID)) & "'," & _
                         "'" & ChangeQuote(Trim(SELECT_FLAG)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         HGROUP_ITEM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE HGROUP_ITEM SET " & _
                     "   HOUSE_GROUP_ID = '" & ChangeQuote(Trim(HOUSE_GROUP_ID)) & "', " & _
                     "   LOCATION_ID = '" & ChangeQuote(Trim(LOCATION_ID)) & "', " & _
                     "   SELECT_FLAG = '" & ChangeQuote(Trim(SELECT_FLAG)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE HGROUP_ITEM_ID = " & HGROUP_ITEM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM HGROUP_ITEM WHERE HGROUP_ITEM_ID = " & HGROUP_ITEM_ID
      m_Conn.Execute (SQL1)
End Function




