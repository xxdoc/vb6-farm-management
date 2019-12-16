VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPlanningItem"
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

Public PLANNING_ITEM_ID As Long
Public PLANNING_ID As Long
Public PLANNING_SUB_TYPE As Long
Public PART_ITEM_ID As Long
Public PLAN_AMOUNT As Double
Public NOTE As String

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public ID As Long

Public FROM_DATE As Date
Public TO_DATE As Date
Public FROM_DATE2  As Date
Public PART_NO As String
Public PART_DESC As String

Public PLANNING_AREA As Long
Public PLANNING_DATE As Date
Public PLAN_VERSION As Long
Public UNIT_NAME As String
Public STR_SQL_DATE As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   
   FROM_DATE = -1
   TO_DATE = -1
End Sub

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
Dim TempDate As String
Dim GroupByStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT PNI.* "
      SelectStr1 = SelectStr1 & ",PI.PART_NO,PI.PART_DESC "
      SelectStr1 = SelectStr1 & ",UN.UNIT_NAME "
      
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM PLANNING_ITEM PNI "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PNI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (PI.UNIT_COUNT = UN.UNIT_ID) "
   
   ElseIf Ind = 2 Then  'LoadPlanningItemDateAmount
      SelectStr1 = "SELECT PNI.PLANNING_ITEM_ID,PNI.PART_ITEM_ID,PN.PLANNING_DATE,PNI.PLAN_AMOUNT,PN.PLAN_VERSION,PN.PLANNING_ID "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM PLANNING_ITEM PNI "
      FromStr = FromStr & "LEFT OUTER JOIN PLANNING PN ON (PNI.PLANNING_ID = PN.PLANNING_ID) "
   ElseIf Ind = 3 Then  'CReportPlanning001
      SelectStr1 = "SELECT DISTINCT PNI.PART_ITEM_ID "
      SelectStr1 = SelectStr1 & ",PI.PART_NO,PI.PART_DESC "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM PLANNING_ITEM PNI "
      FromStr = FromStr & "LEFT OUTER JOIN PART_ITEM PI ON (PNI.PART_ITEM_ID = PI.PART_ITEM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PLANNING PN ON (PNI.PLANNING_ID = PN.PLANNING_ID) "
      
      WhereStr = " WHERE (PNI.PLANNING_SUB_TYPE = 1 OR PNI.PLANNING_SUB_TYPE = 3) "
   End If
   
   If PLANNING_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PNI.PLANNING_ITEM_ID = " & PLANNING_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PNI.PLANNING_ITEM_ID = " & PLANNING_ITEM_ID & ") "
      End If
   End If
   
   If PLANNING_SUB_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PNI.PLANNING_SUB_TYPE = " & PLANNING_SUB_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (PNI.PLANNING_SUB_TYPE = " & PLANNING_SUB_TYPE & ") "
      End If
   End If
   
   If PLANNING_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PNI.PLANNING_ID = " & PLANNING_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PNI.PLANNING_ID = " & PLANNING_ID & ") "
      End If
   End If
   
   If PLANNING_AREA > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PN.PLANNING_AREA = " & PLANNING_AREA & ") "
      Else
         WhereStr = WhereStr & "AND (PN.PLANNING_AREA = " & PLANNING_AREA & ") "
      End If
   End If
   
   If FROM_DATE > -1 Then
      TempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (PN.PLANNING_DATE >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (PN.PLANNING_DATE >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      TempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (PN.PLANNING_DATE <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (PN.PLANNING_DATE <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
   
    If Len(STR_SQL_DATE) > 0 Then 'for LoadPlanningItemDateAmount only
      If WhereStr = "" Then
         WhereStr = "WHERE " & STR_SQL_DATE
      Else
         WhereStr = WhereStr & "AND " & STR_SQL_DATE
      End If
   End If
   
   If PLAN_VERSION > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PN.PLAN_VERSION = " & PLAN_VERSION & ") "
      Else
         WhereStr = WhereStr & "AND (PN.PLAN_VERSION = " & PLAN_VERSION & ") "
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
   ' GroupByStr = "GROUP BY PNI.PART_ITEM_ID,PN.PLANNING_DATE,Pni.PLAN_AMOUNT "
   ElseIf Ind = 3 Then
   '   GroupByStr = "GROUP BY PNI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC,PI.PARCEL_TYPE "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 4 Then
        OrderByStr = "ORDER BY PI.PART_NO "
      ElseIf OrderBy = 5 Then
         OrderByStr = "ORDER BY PI.PART_DESC "
      ElseIf OrderBy = 6 Then
         OrderByStr = "ORDER BY PNI.PLAN_AMOUNT "
      Else
         OrderByStr = "ORDER BY PNI.PLANNING_ITEM_ID "
      End If
   ElseIf Ind = 3 Or Ind = 4 Then
      OrderByStr = "ORDER BY PI.PART_NO "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   If Len(OrderByStr) <= 0 Then
      OrderTypeStr = ""
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr

   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      PLANNING_ITEM_ID = NVLI(Rs("PLANNING_ITEM_ID"), -1)
      PLANNING_ID = NVLI(Rs("PLANNING_ID"), -1)
      PLANNING_SUB_TYPE = NVLI(Rs("PLANNING_SUB_TYPE"), -1)
      
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PLAN_AMOUNT = NVLD(Rs("PLAN_AMOUNT"), 0)
      
      NOTE = NVLS(Rs("NOTE"), "")
      
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
            
      UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
            
   ElseIf Ind = 2 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PLANNING_DATE = InternalDateToDate(NVLS(Rs("PLANNING_DATE"), ""))
      PLAN_AMOUNT = NVLD(Rs("PLAN_AMOUNT"), 0)
      PLAN_VERSION = NVLD(Rs("PLAN_VERSION"), 0)
      PLANNING_ITEM_ID = NVLI(Rs("PLANNING_ITEM_ID"), -1)
      PLANNING_ID = NVLI(Rs("PLANNING_ID"), -1)
   ElseIf Ind = 3 Then
      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
      PART_NO = NVLS(Rs("PART_NO"), "")
      PART_DESC = NVLS(Rs("PART_DESC"), "")
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String
Dim TempID5 As String
Dim TempID6 As String

   TempID1 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID1 = "NULL "
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO PLANNING_ITEM " & _
                     "( " & _
                     "   PLANNING_ID, " & _
                     "   PLANNING_SUB_TYPE, " & _
                     "   PART_ITEM_ID, " & _
                     "   PLAN_AMOUNT, " & _
                     "   NOTE, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                        PLANNING_ID & "," & _
                        PLANNING_SUB_TYPE & "," & _
                        TempID1 & "," & _
                        PLAN_AMOUNT & "," & _
                        "'" & ChangeQuote(Trim(NOTE)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ")  RETURNING PLANNING_ITEM_ID "
   Else 'Show edit
      SQL1 = "UPDATE PLANNING_ITEM SET " & _
                     "  PART_ITEM_ID = " & TempID1 & ", " & _
                     "  PLAN_AMOUNT = " & PLAN_AMOUNT & ", " & _
                      "  NOTE = '" & ChangeQuote(Trim(NOTE)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PLANNING_ITEM_ID = " & PLANNING_ITEM_ID
   End If
   
   If AddEditMode = SHOW_ADD Then
      PLANNING_ITEM_ID = m_Conn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf AddEditMode = SHOW_EDIT Then
      Call m_Conn.Execute(SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM PLANNING_ITEM WHERE PLANNING_ITEM_ID = " & PLANNING_ITEM_ID
      m_Conn.Execute (SQL1)
End Function

