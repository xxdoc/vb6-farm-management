VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAuthenPO"
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

'Public AUTHEN_PO_ID As Long
Public AUTHEN_PO_FROM As Double
Public AUTHEN_PO_TO As Double
Public AUTHEN_PO_DESC As String
Public AUTHEN_PO_GROUP As Double
Public AUTHEN_PO_POTYPE As String

Public AUTHEN_DETAIL_ID As Double
Public AUTHEN_PO_ID As Double
Public AUTHEN_AREA As Double
Public AUTHEN_USER As String
Public AUTHEN_USER_NAME As String

Public TOTAL_PRICE As Double

Public KEY_LOOKUP As String
Public KEY_ID As Long
Public ID As Long


Public FROM_DATE As Date
Public TO_DATE As Date

Public PART_NO As String
Public PRODUCT_NO As String

Public CollAuthenPOCheck As Collection
Public CollAuthenPOApprove As Collection
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   
   FROM_DATE = -1
   TO_DATE = -1
   
   Set CollAuthenPOCheck = New Collection
   Set CollAuthenPOApprove = New Collection
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
      SelectStr1 = "SELECT ATP.*  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
   
      FromStr = "FROM AUTHEN_PO ATP "
   ElseIf Ind = 2 Then        'LoadAuthenPO_Approve ,LoadAuthenPO_Verify
      SelectStr1 = "SELECT ATP.*,ATD.*,UA.USER_NAME  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      WhereStr = ""
      
      FromStr = "FROM AUTHEN_PO ATP "
      FromStr = FromStr & "LEFT OUTER JOIN AUTHEN_DETAIL ATD ON (ATP.AUTHEN_PO_ID=ATD.AUTHEN_PO_ID) "
      FromStr = FromStr & " LEFT OUTER JOIN USER_ACCOUNT UA ON (ATD.AUTHEN_USER=UA.USER_ID) "
   End If
   
   If AUTHEN_PO_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ATP.AUTHEN_PO_ID = " & AUTHEN_PO_ID & ") "
      Else
         WhereStr = WhereStr & "AND (ATP.AUTHEN_PO_ID = " & AUTHEN_PO_ID & ") "
      End If
   End If
 
   If AUTHEN_PO_GROUP > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ATP.AUTHEN_PO_GROUP = " & AUTHEN_PO_GROUP & ") "
      Else
         WhereStr = WhereStr & "AND (ATP.AUTHEN_PO_GROUP = " & AUTHEN_PO_GROUP & ") "
      End If
   End If
   
   If TOTAL_PRICE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ATP.AUTHEN_PO_FROM < " & TOTAL_PRICE & ")   AND (ATP.AUTHEN_PO_TO >= " & TOTAL_PRICE & ") "
      Else
         WhereStr = WhereStr & "AND (ATP.AUTHEN_PO_FROM < " & TOTAL_PRICE & ")   AND (ATP.AUTHEN_PO_TO >= " & TOTAL_PRICE & ") "
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
   '   GroupByStr = "GROUP BY PN.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC "
   ElseIf Ind = 3 Then
   '   GroupByStr = "GROUP BY PN.PART_ITEM_ID, PI.PART_NO, PI.PART_DESC,PI.PARCEL_TYPE "
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
    OrderByStr = "ORDER BY ATP.CREATE_DATE "
'      If OrderBy = 1 Then
'         OrderByStr = "ORDER BY ATP.CREATE_DATE "
'      ElseIf OrderBy = 2 Then
'         OrderByStr = "ORDER BY ATP.AUTHEN_PO_FROM "
'      ElseIf OrderBy = 3 Then
'         OrderByStr = "ORDER BY ATP.AUTHEN_PO_TO "
'      Else
'         OrderByStr = "ORDER BY ATP.PLANNING_DATE "
'      End If
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
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
      AUTHEN_PO_ID = NVLI(Rs("AUTHEN_PO_ID"), -1)
      AUTHEN_PO_GROUP = NVLI(Rs("AUTHEN_PO_GROUP"), -1)
      AUTHEN_PO_FROM = NVLI(Rs("AUTHEN_PO_FROM"), -1)
      AUTHEN_PO_TO = NVLI(Rs("AUTHEN_PO_TO"), -1)
      AUTHEN_PO_DESC = NVLS(Rs("AUTHEN_PO_DESC"), "")
   ElseIf Ind = 2 Then
      AUTHEN_PO_ID = NVLI(Rs("AUTHEN_PO_ID"), -1)
      AUTHEN_PO_GROUP = NVLI(Rs("AUTHEN_PO_GROUP"), -1)
      AUTHEN_PO_FROM = NVLI(Rs("AUTHEN_PO_FROM"), -1)
      AUTHEN_PO_TO = NVLI(Rs("AUTHEN_PO_TO"), -1)
      AUTHEN_PO_DESC = NVLS(Rs("AUTHEN_PO_DESC"), "")
      
      AUTHEN_DETAIL_ID = NVLI(Rs("AUTHEN_DETAIL_ID"), -1)
      AUTHEN_PO_ID = NVLI(Rs("AUTHEN_PO_ID"), -1)
      AUTHEN_AREA = NVLI(Rs("AUTHEN_AREA"), -1)
      AUTHEN_USER = NVLS(Rs("AUTHEN_USER"), "")
      AUTHEN_USER_NAME = NVLS(Rs("USER_NAME"), "")

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

'   TempID1 = PROCESS_ID
'   If PROCESS_ID <= 0 Then
'      TempID1 = "NULL "
'   End If

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   If AddEditMode = SHOW_ADD Then
      SQL1 = "INSERT INTO AUTHEN_PO " & _
                     "( " & _
                     "   AUTHEN_PO_ID, " & _
                     "   AUTHEN_PO_GROUP, " & _
                     "   AUTHEN_PO_FROM, " & _
                     "   AUTHEN_PO_TO, " & _
                     "   AUTHEN_PO_DESC, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                        AUTHEN_PO_ID & "," & _
                        "" & AUTHEN_PO_GROUP & "," & _
                         "" & AUTHEN_PO_FROM & "," & _
                         "" & AUTHEN_PO_TO & "," & _
                        "'" & AUTHEN_PO_DESC & "'," & _
                         "'" & ServerDateTime & "'," & _
                         "" & glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ")  RETURNING AUTHEN_PO_ID "
   Else 'Show edit
      SQL1 = "UPDATE AUTHEN_PO SET " & _
                     "   AUTHEN_PO_GROUP= " & AUTHEN_PO_GROUP & ", " & _
                     "   AUTHEN_PO_FROM= '" & AUTHEN_PO_FROM & "', " & _
                     "   AUTHEN_PO_TO= '" & AUTHEN_PO_TO & "', " & _
                      "  AUTHEN_PO_DESC = '" & AUTHEN_PO_DESC & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE AUTHEN_PO_ID = " & AUTHEN_PO_ID
   End If

   If AddEditMode = SHOW_ADD Then
      AUTHEN_PO_ID = m_Conn.Execute(SQL1)(0) 'Support FireBird Version 2.1 Or Higher
   ElseIf AddEditMode = SHOW_EDIT Then
      Call m_Conn.Execute(SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM AUTHEN_DETAIL WHERE AUTHEN_PO_ID = " & AUTHEN_PO_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM AUTHEN_PO WHERE AUTHEN_PO_ID = " & AUTHEN_PO_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set CollAuthenPOCheck = New Collection
   Set CollAuthenPOApprove = New Collection
End Sub
