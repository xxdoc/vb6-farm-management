VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLegacy"
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
Public FIELD1 As String
Public FIELD2 As String
Public FIELD3 As String
Public FIELD4 As String
Public FIELD5 As String
Public FIELD6 As String
Public FIELD7 As String
Public FIELD8 As String
Public FIELD9 As String
Public FIELD10 As String


Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData1(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT DISTINCT(PRO_UNIT) AS UNIT_NAME "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PRODUCT "
   
   WhereStr = ""
               
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PRO_UNIT "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PRO_UNIT "
   Else
      OrderByStr = "ORDER BY PRO_UNIT "
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
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

'Product type
Public Function QueryData2(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT *  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PRODUCTTYPE "
   
   WhereStr = ""
               
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PRO_TYPECODE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PRO_TYPECODE "
   Else
      OrderByStr = "ORDER BY PRO_TYPECODE "
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
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

'House
Public Function QueryData3(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT *  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM HOUSE "
   
   WhereStr = ""
               
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY H_CODE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY H_CODE "
   Else
      OrderByStr = "ORDER BY H_CODE "
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
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

'Pigtype
Public Function QueryData4(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT *  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PIGTYPE "
   
   WhereStr = ""
               
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PIG_TYPECODE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PIG_TYPECODE "
   Else
      OrderByStr = "ORDER BY PIG_TYPECODE "
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
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

'Pig status
Public Function QueryData5(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT *  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PIGSTATUS "
   
   WhereStr = ""
               
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PIG_STATUSCODE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PIG_STATUSCODE "
   Else
      OrderByStr = "ORDER BY PIG_STATUSCODE "
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
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function QueryData6(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT *  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PRODUCT "
   
   WhereStr = ""
               
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PRO_CODE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PRO_CODE "
   Else
      OrderByStr = "ORDER BY PRO_CODE "
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
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function QueryData7(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT *  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM SUPPILER "
   
   WhereStr = ""
               
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY SUP_CODE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY SUP_CODE "
   Else
      OrderByStr = "ORDER BY SUP_CODE "
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
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function QueryData8(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT *  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM MIS "
   
   WhereStr = ""
               
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY MIS_CODE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY MIS_CODE "
   Else
      OrderByStr = "ORDER BY MIS_CODE "
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
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function QueryData9(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT PW.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PASSWORD PW "
   
   WhereStr = ""
               
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PW.USERNAME "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PW.USERNAME "
   Else
      OrderByStr = "ORDER BY PW.USERNAME "
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
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function QueryData10(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT *  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM ARMAS "
   
   WhereStr = ""
               
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY CUSCOD "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY CUSCOD "
   Else
      OrderByStr = "ORDER BY CUSCOD "
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
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

'Feed stock
Public Function QueryData11(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT *  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM FEEDSTOCK_H "
   
   WhereStr = ""
               
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY FEED_NO "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY FEED_NO "
   Else
      OrderByStr = "ORDER BY FEED_NO "
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
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

'Feed stock detail
Public Function QueryData12(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT *  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM FEEDSTOCK "
   
   WhereStr = ""
               
   If Len(FIELD1) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (FEED_NO = '" & FIELD1 & "') "
      Else
         WhereStr = WhereStr & "AND (FEED_NO = '" & FIELD1 & "') "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY FEED_NO "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY FEED_NO "
   Else
      OrderByStr = "ORDER BY FEED_NO "
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
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

'Bill Header
Public Function QueryData13(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT *  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM BILL_H "
   
   WhereStr = ""
               
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY BH_NO "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY BH_NO "
   Else
      OrderByStr = "ORDER BY BH_NO "
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
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function QueryData14(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT *  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM BILL_D "
   
   WhereStr = ""
               
   If Len(FIELD1) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BD_NO = '" & FIELD1 & "') "
      Else
         WhereStr = WhereStr & "AND (BD_NO = '" & FIELD1 & "') "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY BD_NO "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY BD_NO "
   Else
      OrderByStr = "ORDER BY BD_NO "
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
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

'Pig move header
Public Function QueryData15(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT *  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM MOVESTOCK_H "
   
   WhereStr = ""
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY MOVE_NO "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY MOVE_NO "
   Else
      OrderByStr = "ORDER BY MOVE_NO "
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
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

'Move stock detail
Public Function QueryData16(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT *  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM MOVESTOCK "
   
   WhereStr = ""
               
   If Len(FIELD1) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MOVE_NO = '" & FIELD1 & "') "
      Else
         WhereStr = WhereStr & "AND (MOVE_NO = '" & FIELD1 & "') "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY MOVE_NO "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY MOVE_NO "
   Else
      OrderByStr = "ORDER BY MOVE_NO "
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
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

'Move pig birth
Public Function QueryData17(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT *  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PIGBIRTH "
   
   WhereStr = ""
                  
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PB_CODE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PB_CODE "
   Else
      OrderByStr = "ORDER BY PB_CODE "
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
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

'Move pig sell
Public Function QueryData18(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT *  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PIGSELL_H "
   
   WhereStr = ""
                  
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PS_NO "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PS_NO "
   Else
      OrderByStr = "ORDER BY PS_NO "
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
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

'Move pig sell detail
Public Function QueryData19(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT *  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PIGSELL_D "
   
   WhereStr = ""
                  
   If Len(FIELD1) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PS_NO = '" & FIELD1 & "') "
      Else
         WhereStr = WhereStr & "AND (PS_NO = '" & FIELD1 & "') "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PS_NO "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PS_NO "
   Else
      OrderByStr = "ORDER BY PS_NO "
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
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function QueryData20(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT DISTINCT MPH.MP_NO, MPH.MP_DATE, MPH.MP_MISCODE   "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM MOVEPIG_D MPD, MOVEPIG_H MPH "
   
   WhereStr = "WHERE (MPH.MP_NO = MPD.MP_NO) "
'AND (""MPD.MP_Y/W"" = ""MPD.MP_Y/WW"")
'   If Len(FIELD1) > 0 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (MP_NO = '" & FIELD1 & "') "
'      Else
'         WhereStr = WhereStr & "AND (MP_NO = '" & FIELD1 & "') "
'      End If
'   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY MPH.MP_NO "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY MPH.MP_NO "
   Else
      OrderByStr = "ORDER BY MPH.MP_NO "
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
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, False, ItemCount, Rs, glbErrorLog)
End Function

Public Function QueryData21(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT * "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM MOVEPIG_D "
   
   WhereStr = ""
                  
   If Len(FIELD1) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MP_NO = '" & FIELD1 & "') "
      Else
         WhereStr = WhereStr & "AND (MP_NO = '" & FIELD1 & "') "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY MP_NO "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY MP_NO "
   Else
      OrderByStr = "ORDER BY MP_NO "
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
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      FIELD1 = NVLS(Rs("UNIT_NAME"), "")
   ElseIf Ind = 2 Then
      FIELD1 = NVLS(Rs("PRO_TYPECODE"), "")
      FIELD2 = NVLS(Rs("PRO_TYPENAME"), "")
   ElseIf Ind = 3 Then
      FIELD1 = NVLS(Rs("H_CODE"), "")
      FIELD2 = NVLS(Rs("H_NAME"), "")
   ElseIf Ind = 4 Then
      FIELD1 = NVLS(Rs("PIG_TYPECODE"), "")
      FIELD2 = NVLS(Rs("PIG_TYPENAME"), "")
   ElseIf Ind = 5 Then
      FIELD1 = NVLS(Rs("PIG_STATUSCODE"), "")
      FIELD2 = NVLS(Rs("PIG_STATUSNAME"), "")
   ElseIf Ind = 6 Then
      FIELD1 = NVLS(Rs("PRO_TYPECODE"), "")
      FIELD2 = NVLS(Rs("PRO_CODE"), "")
      FIELD3 = NVLS(Rs("PRO_NAME"), "")
      FIELD4 = NVLS(Rs("PRO_UNIT"), "")
   ElseIf Ind = 7 Then
      FIELD1 = NVLS(Rs("SUP_CODE"), "")
      FIELD2 = NVLS(Rs("SUP_NAME"), "")
   ElseIf Ind = 8 Then
      FIELD1 = NVLS(Rs("MIS_CODE"), "")
      FIELD2 = NVLS(Rs("MIS_NAME"), "")
   ElseIf Ind = 9 Then
      FIELD1 = NVLS(Rs("USERNAME"), "")
      FIELD2 = NVLS(Rs("PASSWORD"), "")
   ElseIf Ind = 10 Then
      FIELD1 = NVLS(Rs("CUSCOD"), "")
      FIELD2 = NVLS(Rs("CUSTYP"), "")
      FIELD3 = NVLS(Rs("PRENAM"), "")
      FIELD4 = NVLS(Rs("CUSNAM"), "")
   ElseIf Ind = 11 Then
      FIELD1 = NVLS(Rs("FEED_NO"), "")
      FIELD2 = ReFormatDate(NVLS(Rs("FEED_DATE"), ""))
      FIELD3 = NVLS(Rs("FEED_MISCODE"), "")
   ElseIf Ind = 12 Then
      FIELD1 = NVLS(Rs("FEED_NO"), "")
      FIELD2 = NVLS(Rs("FEED_CODE"), "")
      FIELD3 = NVLS(Rs("FEED_Y/W"), "")
      FIELD4 = NVLS(Rs("FEED_FROMHOUSE"), "")
      FIELD5 = NVLS(Rs("FEED_TOHOUSE"), "")
      FIELD6 = NVLS(Rs("FEED_NUM"), "")
      FIELD7 = NVLS(Rs("FEED_PRICEUNIT"), "")
      FIELD8 = NVLS(Rs("FEED_TYPECODE"), "")
   ElseIf Ind = 13 Then
      FIELD1 = NVLS(Rs("BH_NO"), "")
      FIELD2 = NVLS(Rs("BH_NOR"), "")
      FIELD3 = NVLS(Rs("BH_CODE1"), "")
      FIELD4 = ReFormatDate(NVLS(Rs("BH_DATE"), ""))
      FIELD5 = NVLS(Rs("BH_LCAR"), "")
      FIELD6 = NVLS(Rs("BH_CODE2"), "")
      FIELD7 = NVLS(Rs("BH_TRANNO"), "")
      FIELD8 = NVLS(Rs("BH_NAMEG"), "")
      FIELD9 = NVLS(Rs("BH_NAMER"), "")
      FIELD10 = NVLS(Rs("BH_PRICE"), "")
   ElseIf Ind = 14 Then
      FIELD1 = NVLS(Rs("BD_NO"), "")
      FIELD2 = NVLS(Rs("BD_CODE"), "")
      FIELD3 = NVLS(Rs("BD_NUM"), "")
      FIELD4 = NVLS(Rs("BD_PRICEUNIT"), "")
      FIELD5 = NVLS(Rs("BD_PRICEUNIT_KG"), "")
      FIELD6 = NVLS(Rs("BD_TOTALPRICE"), "")
      FIELD7 = NVLS(Rs("BD_TYPECODE"), "")
   ElseIf Ind = 15 Then
      FIELD1 = NVLS(Rs("MOVE_NO"), "")
      FIELD2 = ReFormatDate(NVLS(Rs("MOVE_DATE"), ""))
      FIELD3 = NVLS(Rs("MOVE_MISCODE"), "")
   ElseIf Ind = 16 Then
      FIELD1 = NVLS(Rs("MOVE_NO"), "")
      FIELD2 = NVLS(Rs("MOVE_CODE"), "")
      FIELD3 = NVLS(Rs("MOVE_TOSTOCK"), "")
      FIELD4 = NVLS(Rs("MOVE_FROMSTOCK"), "")
      FIELD5 = NVLS(Rs("MOVE_NUM"), "")
      FIELD6 = NVLS(Rs("MOVE_PRICEUNIT"), "")
      FIELD7 = NVLS(Rs("MOVE_TYPECODE"), "")
   ElseIf Ind = 17 Then
      FIELD1 = NVLS(Rs("PB_CODE"), "")
      FIELD2 = NVLS(Rs("PB_WEEK"), "")
      FIELD3 = NVLS(Rs("PB_HOUSE"), "")
      FIELD4 = ReFormatDate(NVLS(Rs("PB_DATE"), ""))
      FIELD5 = NVLS(Rs("PB_MOTHER"), "")
      FIELD6 = NVLS(Rs("PB_FATHER"), "")
      FIELD7 = NVLS(Rs("PB_NUM"), "")
      FIELD8 = NVLS(Rs("PB_WEIGTH"), "")
      FIELD9 = NVLS(Rs("PB_NUMWEIGTH"), "")
      FIELD10 = NVLS(Rs("ID"), "")
   ElseIf Ind = 18 Then
      FIELD1 = NVLS(Rs("PS_NO"), "")
      FIELD2 = ReFormatDate(NVLS(Rs("PS_DATE"), ""))
   ElseIf Ind = 19 Then
      FIELD1 = NVLS(Rs("PS_NO"), "")
      FIELD2 = NVLS(Rs("PS_FROMSTOCK"), "")
      FIELD3 = NVLS(Rs("PS_WEEK"), "")
      FIELD4 = NVLS(Rs("PS_STATUS"), "")
      FIELD5 = NVLS(Rs("PS_NUM"), "")
      FIELD6 = NVLS(Rs("PS_NUMWEIGTH"), "")
   ElseIf Ind = 20 Then
      FIELD1 = NVLS(Rs("MP_NO"), "")
      FIELD2 = ReFormatDate(NVLS(Rs("MP_DATE"), ""))
      FIELD3 = NVLS(Rs("MP_MISCODE"), "")
   ElseIf Ind = 21 Then
      FIELD1 = NVLS(Rs("MP_NO"), "")
      FIELD2 = NVLS(Rs("MP_Y/W"), "")
      FIELD3 = NVLS(Rs("MP_Y/WW"), "")
      FIELD4 = NVLS(Rs("MP_TOSTOCK"), "")
      FIELD5 = NVLS(Rs("MP_FROMSTOCK"), "")
      FIELD6 = NVLS(Rs("MP_NUM"), "")
      FIELD7 = NVLS(Rs("MP_WEIGTH"), "")
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
      
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
'   If AddEditMode = SHOW_ADD Then
'      Call glbDatabaseMngr.GetSeqID("YEAR_SEQ_SEQ", ID, glbErrorLog)
'
'      SQL1 = "INSERT INTO YEAR_SEQ " & _
'                     "( " & _
'                     "   YEAR_SEQ_ID, " & _
'                     "   YEAR_NO, " & _
'                     "   YEAR_DESC, " & _
'                     "   CREATE_DATE, " & _
'                    "    CREATE_BY, " & _
'                    "    MODIFY_DATE, " & _
'                    "    MODIFY_BY " & _
'                     ") " & _
'                     "VALUES "
'         SQL1 = SQL1 & "( " & _
'                         ID & "," & _
'                         YEAR_NO & "," & _
'                         "'" & ChangeQuote(Trim(YEAR_DESC)) & "'," & _
'                         "'" & ServerDateTime & "'," & _
'                         glbUser.USER_ID & "," & _
'                         "'" & ServerDateTime & "'," & _
'                         glbUser.USER_ID & _
'                     ") "
'         YEAR_SEQ_ID = ID
'   Else 'Show edit
'      SQL1 = "UPDATE YEAR_SEQ SET " & _
'                     "   YEAR_NO = " & YEAR_NO & ", " & _
'                     "   YEAR_DESC = '" & ChangeQuote(Trim(YEAR_DESC)) & "', " & _
'                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
'                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
'                     "WHERE YEAR_SEQ_ID = " & YEAR_SEQ_ID
'   End If
'
'   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
'      m_Conn.Execute (SQL1)
'   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
            
'      SQL1 = "DELETE FROM YEAR_WEEK WHERE YEAR_SEQ_ID = " & YEAR_SEQ_ID
'      m_Conn.Execute (SQL1)
'
'      SQL1 = "DELETE FROM PART_ITEM WHERE PART_ITEM_ID IN (SELECT YW.PART_ITEM_ID1 FROM YEAR_WEEK YW WHERE YW.YEAR_SEQ_ID = " & YEAR_SEQ_ID & ") "
'      m_Conn.Execute (SQL1)
'
'      SQL1 = "DELETE FROM PART_ITEM WHERE PART_ITEM_ID IN (SELECT YW.PART_ITEM_ID2 FROM YEAR_WEEK YW WHERE YW.YEAR_SEQ_ID = " & YEAR_SEQ_ID & ") "
'      m_Conn.Execute (SQL1)
'
'      SQL1 = "DELETE FROM PART_ITEM WHERE PART_ITEM_ID IN (SELECT YW.PART_ITEM_ID3 FROM YEAR_WEEK YW WHERE YW.YEAR_SEQ_ID = " & YEAR_SEQ_ID & ") "
'      m_Conn.Execute (SQL1)
'
'      SQL1 = "DELETE FROM YEAR_SEQ WHERE YEAR_SEQ_ID = " & YEAR_SEQ_ID
'      m_Conn.Execute (SQL1)
End Function

