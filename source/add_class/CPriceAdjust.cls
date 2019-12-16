VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPriceAdjust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Conn As ADODB.Connection

Public PRICE_ADJUST_ID As Long
Public PART_ITEM_ID As Long
Public AVG_PRICE As Double

Public PART_NO As String
Public PART_DESC As String
Public UNIT_NAME As String

Public OrderBy As Long
Public OrderType As Long
Public QueryFlag As Long
Public AddEditMode As SHOW_MODE_TYPE
Public Flag As String
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   
   QueryFlag = -1
End Sub

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim Temp1 As String

   Temp1 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      Temp1 = "NULL "
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("PRICE_ADJUST_SEQ", ID, glbErrorLog)
      SQL1 = "INSERT INTO PRICE_ADJUST " & _
                     "( " & _
                     "   PRICE_ADJUST_ID, " & _
                     "   PART_ITEM_ID, " & _
                     "   AVG_PRICE, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
               SQL1 = SQL1 & "( " & _
                         ID & "," & _
                        Temp1 & ", " & _
                        AVG_PRICE & ", " & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
                  PRICE_ADJUST_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE PRICE_ADJUST SET " & _
                     "   PART_ITEM_ID =  " & Temp1 & ", " & _
                     "   AVG_PRICE =  " & AVG_PRICE & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PRICE_ADJUST_ID = " & PRICE_ADJUST_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM PRICE_ADJUST WHERE PRICE_ADJUST_ID = " & PRICE_ADJUST_ID
      m_Conn.Execute (SQL1)
End Function
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
   
   SelectStr1 = "SELECT PA.*,PI.PART_NO,PI.PART_DESC "
   SelectStr1 = SelectStr1 & ",UI.UNIT_NAME "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PRICE_ADJUST PA "
   FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PI.PART_ITEM_ID = PA.PART_ITEM_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN UNIT UI ON (UI.UNIT_ID = PI.UNIT_COUNT) "
   
   WhereStr = ""
   
   If PRICE_ADJUST_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PRICE_ADJUST_ID = " & PRICE_ADJUST_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PRICE_ADJUST_ID = " & PRICE_ADJUST_ID & ") "
      End If
   End If
   
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   If Trim(PART_NO) <> "" Then
      If WildCard(PART_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "')"
         End If
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PART_NO "
   Else
      OrderByStr = "ORDER BY PART_NO "
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
   PRICE_ADJUST_ID = NVLI(Rs("PRICE_ADJUST_ID"), -1)
   PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
   AVG_PRICE = NVLD(Rs("AVG_PRICE"), 0)
   
   PART_NO = NVLS(Rs("PART_NO"), "")
   PART_DESC = NVLS(Rs("PART_DESC"), "")
   UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
End Function


