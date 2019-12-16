VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAuthenPOItem"
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

Public AUTHEN_DETAIL_ID As Long
Public AUTHEN_PO_ID As Long
Public AUTHEN_AREA As Long
Public AUTHEN_USER As String
Public AUTHEN_PO_REALNAME As String
Public AUTHEN_USERNAME As String

Public AUTHEN_PO_AMOUNT As Double
'Public AUTHEN_PO_SUB_TYPE As Long
'Public PART_ITEM_ID As Long
'Public AUTHEN_PO_AMOUNT As Double
'Public NOTE As String
'
'Public KEY_LOOKUP As String
'Public KEY_ID As Long
'Public ID As Long
'
'Public FROM_DATE As Date
'Public TO_DATE As Date
'
'Public PART_NO As String
'Public PART_DESC As String
'
'Public PLANNING_AREA As Long
'Public PLANNING_DATE As Date
'Public UNIT_NAME As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   
'   FROM_DATE = -1
'   TO_DATE = -1
End Sub

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
Dim TempDate As String
Dim GroupByStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT ATD.*,UA.REAL_NAME,UA.USER_NAME "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM AUTHEN_DETAIL ATD "
      FromStr = FromStr & "LEFT JOIN USER_ACCOUNT UA ON (ATD.AUTHEN_USER=UA.USER_ID)  "
   
   ElseIf Ind = 2 Then  'LoadPlanningItemDateAmount
      SelectStr1 = "SELECT PNI.PART_ITEM_ID,PN.PLANNING_DATE,PNI.PLAN_AMOUNT "
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
   
   
   If AUTHEN_PO_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ATD.AUTHEN_PO_ID = " & AUTHEN_PO_ID & ") "
      Else
         WhereStr = WhereStr & "AND (ATD.AUTHEN_PO_ID = " & AUTHEN_PO_ID & ") "
      End If
   End If
   
   If AUTHEN_AREA > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ATD.AUTHEN_AREA = " & AUTHEN_AREA & ") "
      Else
         WhereStr = WhereStr & "AND (ATD.AUTHEN_AREA = " & AUTHEN_AREA & ") "
      End If
   End If
   
   GroupByStr = ""
   If Ind = 2 Then
   '   GroupByStr = "GROUP BY PNI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
   ElseIf Ind = 3 Then
   '   GroupByStr = "GROUP BY PNI.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC,PI.PARCEL_TYPE "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY ATD.AUTHEN_DETAIL_ID "
   ElseIf Ind = 3 Then
'      OrderByStr = "ORDER BY PI.PART_NO "
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

   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      AUTHEN_DETAIL_ID = NVLI(Rs("AUTHEN_DETAIL_ID"), -1)
      AUTHEN_PO_ID = NVLI(Rs("AUTHEN_PO_ID"), -1)
      AUTHEN_AREA = NVLI(Rs("AUTHEN_AREA"), -1)
      AUTHEN_USER = NVLI(Rs("AUTHEN_USER"), -1)
      AUTHEN_PO_REALNAME = NVLS(Rs("REAL_NAME"), "")
      AUTHEN_USERNAME = NVLS(Rs("USER_NAME"), "")
            
   ElseIf Ind = 2 Then
'      AUTHEN_PO_ID = NVLI(Rs("AUTHEN_PO_ID"), -1)
'      AUTHEN_PO_GROUP = NVLI(Rs("AUTHEN_PO_GROUP"), -1)
'      AUTHEN_PO_FROM = NVLI(Rs("AUTHEN_PO_FROM"), -1)
'      AUTHEN_PO_TO = NVLI(Rs("AUTHEN_PO_TO"), -1)
'      AUTHEN_PO_DESC = NVLS(Rs("AUTHEN_PO_DESC"), -1)
'   ElseIf Ind = 3 Then
'      PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
'      PART_NO = NVLS(Rs("PART_NO"), "")
'      PART_DESC = NVLS(Rs("PART_DESC"), "")
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

'   TempID1 = PART_ITEM_ID
'   If PART_ITEM_ID <= 0 Then
'      TempID1 = "NULL "
'   End If

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO AUTHEN_DETAIL " & _
                     "( " & _
                     "   AUTHEN_PO_ID, " & _
                     "   AUTHEN_AREA, " & _
                     "   AUTHEN_USER, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                        AUTHEN_PO_ID & "," & _
                        AUTHEN_AREA & ",'" & _
                        AUTHEN_USER & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ")  RETURNING AUTHEN_DETAIL_ID "
   Else 'Show edit
      SQL1 = "UPDATE AUTHEN_DETAIL SET " & _
                     "  AUTHEN_PO_ID = " & AUTHEN_PO_ID & ", " & _
                     "  AUTHEN_AREA = " & AUTHEN_AREA & ", " & _
                     "  AUTHEN_USER = " & AUTHEN_USER & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE AUTHEN_DETAIL_ID = " & AUTHEN_DETAIL_ID
   End If

   If AddEditMode = SHOW_ADD Then
      AUTHEN_DETAIL_ID = m_Conn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf AddEditMode = SHOW_EDIT Then
      Call m_Conn.Execute(SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM AUTHEN_DETAIL WHERE AUTHEN_DETAIL_ID = " & AUTHEN_DETAIL_ID
      m_Conn.Execute (SQL1)
End Function

