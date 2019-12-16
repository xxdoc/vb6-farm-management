VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPartItem"
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
Public PART_ITEM_ID As Long
Public PART_NO As String
Public PART_DESC As String
Public UNIT_COUNT As Long
Public PART_TYPE As Long
Public MINIMUM_ALLOW As Double
Public PIG_FLAG As String
Public PIG_TYPE As String
Public PARENT_ID As Long
Public SPECIFIC_FLAG As String
Public FEED_GROUP As Long

Public PIG_AGE As Double
Public AGE_CODE As String
Public UNIT_NAME As String
Public PART_TYPE_NAME As String
Public PART_TYPE_NO As String
Public INTAKE_FLAG As String
Public PART_GROUP_ID As Long

Public KEY_ID As Long
Public KEY_LOOKUP As String

Public SUPPLIER_ID As Long
Public REGIST_NUM As String
Public DRUG_SIZE  As String
Public DRUG_NAME  As String

Public PartLocations As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set PartLocations = New Collection
   INTAKE_FLAG = "Y"
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

   If Ind = 1 Then
      SelectStr1 = "SELECT ADR.*, PT.*, UN.* "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM PART_ITEM ADR  "
      FromStr = FromStr & "LEFT OUTER JOIN PART_TYPE PT ON (ADR.PART_TYPE = PT.PART_TYPE_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN UNIT UN ON (ADR.UNIT_COUNT = UN.UNIT_ID) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
   End If
   
   If PART_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PART_ITEM_ID = " & PART_ITEM_ID & ") "
      End If
   End If
   
   If PART_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PART_TYPE = " & PART_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (PART_TYPE = " & PART_TYPE & ") "
      End If
   End If
      
   If UNIT_COUNT > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UNIT_COUNT = " & UNIT_COUNT & ") "
      Else
         WhereStr = WhereStr & "AND (UNIT_COUNT = " & UNIT_COUNT & ") "
      End If
   End If
      
   If PARENT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PARENT_ID = " & PARENT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PARENT_ID = " & PARENT_ID & ") "
      End If
   End If
      
   If Len(PIG_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PIG_FLAG = '" & PIG_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (PIG_FLAG = '" & PIG_FLAG & "') "
      End If
   End If
   
   If Len(SPECIFIC_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SPECIFIC_FLAG = '" & SPECIFIC_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (SPECIFIC_FLAG = '" & SPECIFIC_FLAG & "') "
      End If
   End If
   
   If Len(PIG_TYPE) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PIG_TYPE = '" & PIG_TYPE & "') "
      Else
         WhereStr = WhereStr & "AND (PIG_TYPE = '" & PIG_TYPE & "') "
      End If
   End If
   
   If Len(Trim(PART_NO)) > 0 Then
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

   If Len(Trim(PART_DESC)) > 0 Then
      If WildCard(PART_DESC, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(PART_DESC,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(PART_DESC,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (PART_DESC = '" & ChangeQuote(Trim(PART_DESC)) & "')"
         Else
            WhereStr = WhereStr & "AND (PART_DESC = '" & ChangeQuote(Trim(PART_DESC)) & "')"
         End If
      End If
   End If
   
   If PART_GROUP_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PT.PART_GROUP_ID = " & PART_GROUP_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PT.PART_GROUP_ID = " & PART_GROUP_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PART_NO "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY PART_DESC "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY PART_TYPE_NO "
      ElseIf OrderBy = 4 Then
         OrderByStr = "ORDER BY PART_TYPE_NAME "
      ElseIf OrderBy = 5 Then
         OrderByStr = "ORDER BY PART_NO "
      ElseIf OrderBy = 6 Then
         OrderByStr = "ORDER BY PART_DESC "
      Else
         OrderByStr = "ORDER BY PART_ITEM_ID "
      End If
   ElseIf Ind = 2 Then
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
   PART_ITEM_ID = NVLI(Rs("PART_ITEM_ID"), -1)
   PART_NO = NVLS(Rs("PART_NO"), "")
   UNIT_COUNT = NVLI(Rs("UNIT_COUNT"), -1)
   MINIMUM_ALLOW = NVLS(Rs("MINIMUM_ALLOW"), "")
   PART_TYPE = NVLI(Rs("PART_TYPE"), -1)
   PIG_FLAG = NVLS(Rs("PIG_FLAG"), "")
   PART_DESC = NVLS(Rs("PART_DESC"), "")
   PIG_TYPE = NVLS(Rs("PIG_TYPE"), "")
   INTAKE_FLAG = NVLS(Rs("INTAKE_FLAG"), "")
   PARENT_ID = NVLI(Rs("PARENT_ID"), -1)
   SPECIFIC_FLAG = NVLS(Rs("SPECIFIC_FLAG"), "")
   FEED_GROUP = NVLI(Rs("FEED_GROUP"), -1)
   
   UNIT_NAME = NVLS(Rs("UNIT_NAME"), "")
   PART_TYPE_NAME = NVLS(Rs("PART_TYPE_NAME"), "")
   PART_TYPE_NO = NVLS(Rs("PART_TYPE_NO"), "")
   
   REGIST_NUM = NVLS(Rs("REGIST_NUM"), "")
   DRUG_NAME = NVLS(Rs("DRUG_NAME"), "")
   DRUG_SIZE = NVLS(Rs("DRUG_SIZE"), "")
   SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), -1)
   
   KEY_ID = PART_ITEM_ID
   KEY_LOOKUP = PART_NO
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

   TempID1 = PART_TYPE
   If PART_TYPE <= 0 Then
      TempID1 = "NULL "
   End If
      
   TempID2 = UNIT_COUNT
   If UNIT_COUNT <= 0 Then
      TempID2 = "NULL "
   End If
      
PARENT_ID = 0
   TempID3 = PARENT_ID
   If PARENT_ID <= 0 Then
      TempID3 = "NULL "
   End If
   
   TempID5 = SUPPLIER_ID
   If SUPPLIER_ID <= 0 Then
      TempID5 = "NULL"
   End If
   
SPECIFIC_FLAG = "N"
   If Len(SPECIFIC_FLAG) <= 0 Then
      SPECIFIC_FLAG = "N"
   End If
   
   TempID4 = FEED_GROUP
   If FEED_GROUP <= 0 Then
      TempID4 = "NULL "
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("PART_ITEM_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO PART_ITEM " & _
                     "( " & _
                     "   PART_ITEM_ID, " & _
                     "   PART_NO, " & _
                     "   UNIT_COUNT, " & _
                     "   MINIMUM_ALLOW, " & _
                     "   PART_TYPE, " & _
                     "   PIG_FLAG, " & _
                     "   PART_DESC, " & _
                     "   PIG_TYPE, " & _
                     "   INTAKE_FLAG, " & _
                     "   PARENT_ID, " & _
                     "   SPECIFIC_FLAG, " & _
                     "   FEED_GROUP, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                     "   REGIST_NUM, " & _
                     "   DRUG_SIZE, " & _
                     "   DRUG_NAME, " & _
                     "   SUPPLIER_ID, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(PART_NO)) & "'," & _
                         "" & TempID2 & "," & _
                         "'" & ChangeQuote(Trim(MINIMUM_ALLOW)) & "'," & _
                         "" & TempID1 & "," & _
                         "'" & ChangeQuote(Trim(PIG_FLAG)) & "'," & _
                         "'" & ChangeQuote(Trim(PART_DESC)) & "'," & _
                         "'" & ChangeQuote(Trim(PIG_TYPE)) & "'," & _
                         "'" & ChangeQuote(Trim(INTAKE_FLAG)) & "'," & _
                         "" & TempID3 & "," & _
                         "'" & ChangeQuote(Trim(SPECIFIC_FLAG)) & "'," & _
                        "" & TempID4 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ChangeQuote(Trim(REGIST_NUM)) & "'," & _
                         "'" & ChangeQuote(Trim(DRUG_SIZE)) & "'," & _
                         "'" & ChangeQuote(Trim(DRUG_NAME)) & "'," & _
                          "" & TempID5 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         PART_ITEM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE PART_ITEM SET " & _
                     "   PART_NO = '" & ChangeQuote(Trim(PART_NO)) & "', " & _
                     "   UNIT_COUNT = " & TempID2 & ", " & _
                     "   MINIMUM_ALLOW = '" & MINIMUM_ALLOW & "', " & _
                     "   PART_TYPE = " & TempID1 & ", " & _
                     "   PIG_FLAG = '" & ChangeQuote(Trim(PIG_FLAG)) & "', " & _
                     "   PART_DESC = '" & ChangeQuote(Trim(PART_DESC)) & "', " & _
                     "   PIG_TYPE = '" & ChangeQuote(Trim(PIG_TYPE)) & "', " & _
                     "   INTAKE_FLAG = '" & ChangeQuote(Trim(INTAKE_FLAG)) & "', " & _
                     "   PARENT_ID = " & TempID3 & ", " & _
                     "   SPECIFIC_FLAG = '" & ChangeQuote(Trim(SPECIFIC_FLAG)) & "', " & _
                     "   FEED_GROUP = " & TempID4 & ", " & _
                     "   REGIST_NUM = '" & ChangeQuote(Trim(REGIST_NUM)) & "', " & _
                     "   DRUG_SIZE = '" & ChangeQuote(Trim(DRUG_SIZE)) & "', " & _
                     "   DRUG_NAME = '" & ChangeQuote(Trim(DRUG_NAME)) & "', " & _
                    "   SUPPLIER_ID = " & TempID5 & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PART_ITEM_ID = " & PART_ITEM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM PART_ITEM WHERE PART_ITEM_ID = " & PART_ITEM_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set PartLocations = Nothing
End Sub
