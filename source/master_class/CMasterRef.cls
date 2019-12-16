VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMasterRef"
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

Public KEY_ID As Long
Public KEY_CODE As String
Public KEY_NAME As String
Public MASTER_AREA As Long
Public CBO_ID As Long
Public CASE_NO As Double
Public TEMP_ID1 As Long
Public TEMP_ID2 As Long

Public KEY_LOOKUP As String

Public RightItems As Collection
Public PrtItemSets As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set RightItems = New Collection
   Set PrtItemSets = New Collection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Rs As ADODB.Recordset, ItemCount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String

   SelectStr1 = "SELECT UG.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM MASTER_REF UG "
   
   WhereStr = ""
   
   If KEY_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (KEY_ID = " & KEY_ID & ") "
      Else
         WhereStr = "AND (KEY_ID = " & KEY_ID & ") "
      End If
   End If
   
   If MASTER_AREA > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MASTER_AREA = " & MASTER_AREA & ") "
      Else
         WhereStr = WhereStr & "AND (MASTER_AREA = " & MASTER_AREA & ") "
      End If
   End If
   
   If TEMP_ID1 > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (TEMP_ID1 = " & TEMP_ID1 & ") "
      Else
         WhereStr = WhereStr & "AND (TEMP_ID1 = " & TEMP_ID1 & ") "
      End If
   End If
   
   If TEMP_ID2 > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (TEMP_ID2 = " & TEMP_ID2 & ") "
      Else
         WhereStr = WhereStr & "AND (TEMP_ID2 = " & TEMP_ID2 & ") "
      End If
   End If
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY KEY_CODE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY KEY_NAME "
   ElseIf OrderBy = 3 Then
      OrderByStr = "ORDER BY CBO_ID "
   Else
      OrderByStr = "ORDER BY KEY_CODE "
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
   KEY_ID = NVLI(Rs("KEY_ID"), -1)
   KEY_NAME = NVLS(Rs("KEY_NAME"), "")
   KEY_CODE = NVLS(Rs("KEY_CODE"), "")
   MASTER_AREA = NVLI(Rs("MASTER_AREA"), -1)
   TEMP_ID1 = NVLI(Rs("TEMP_ID1"), -1)
   TEMP_ID2 = NVLI(Rs("TEMP_ID2"), -1)
   
   KEY_LOOKUP = KEY_CODE
End Function

Public Function AddEditData(Optional AutoGenerateFlag As Boolean = True)
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String
Dim TempID4 As String

   TempID1 = MASTER_AREA
   If MASTER_AREA <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = CBO_ID
   If CBO_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID3 = TEMP_ID1
   If TEMP_ID1 <= 0 Then
      TempID3 = "NULL"
   End If
   
   TempID4 = TEMP_ID2
   If TEMP_ID2 <= 0 Then
      TempID4 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      If AutoGenerateFlag Then
         Call glbDatabaseMngr.GetSeqID("MASTER_REF_SEQ", ID, glbErrorLog)
      Else
         ID = KEY_ID
      End If
      
      SQL1 = "INSERT INTO MASTER_REF " & _
                     "( " & _
                     "   KEY_ID, " & _
                     "   KEY_NAME, " & _
                     "   KEY_CODE, " & _
                     "   MASTER_AREA, " & _
                     "   TEMP_ID1, " & _
                     "   TEMP_ID2, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(KEY_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(KEY_CODE)) & "'," & _
                         TempID1 & ","
           SQL1 = SQL1 & _
                        TempID3 & "," & _
                        TempID4 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         KEY_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE MASTER_REF SET " & _
                     "   MASTER_AREA = " & MASTER_AREA & ", " & _
                     "   KEY_NAME = '" & ChangeQuote(Trim(KEY_NAME)) & "', " & _
                     "   KEY_CODE = '" & ChangeQuote(Trim(KEY_CODE)) & "', " & _
                     "   TEMP_ID1 = " & TempID3 & ", " & _
                     "   TEMP_ID2 = " & TempID4 & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE KEY_ID = " & KEY_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      
'      SQL1 = "DELETE FROM PRTITEM_SET WHERE MASTER_REF_ID = " & KEY_ID
'      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM MASTER_REF WHERE KEY_ID = " & KEY_ID
      m_Conn.Execute (SQL1)
End Function
Private Sub Class_Terminate()
   Set PrtItemSets = Nothing
End Sub
