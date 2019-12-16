VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBSheetItem"
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
Public ExtraFlag As Long
' End of generic data member

Public BSHEET_ITEM_ID As Long
Public BLOOD_SHEET_ID As Long
Public BLOOD_SPEC_ID As Long
Public SPEC_VALUE As Double

Public SPEC_NAME As String
Public MALE_STD As String
Public FEMALE_STD As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
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
Dim SubLen As Long
Dim NewStr As String
Dim TempDate As String

   SelectStr1 = "SELECT BL.*, DR.* "
   
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM BSHEET_ITEM BL "
   FromStr = FromStr & "LEFT OUTER JOIN BLOOD_SPEC DR ON (BL.BLOOD_SPEC_ID = DR.BLOOD_SPEC_ID) "
   
   WhereStr = ""
   
   If BSHEET_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BSHEET_ITEM_ID = " & BSHEET_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BSHEET_ITEM_ID = " & BSHEET_ITEM_ID & ") "
      End If
   End If
      
   If BLOOD_SHEET_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BLOOD_SHEET_ID = " & BLOOD_SHEET_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BLOOD_SHEET_ID = " & BLOOD_SHEET_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY BLOOD_SPEC_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY BLOOD_SPEC_ID "
   Else
      OrderByStr = "ORDER BY SPEC_NAME, BLOOD_SPEC_ID "
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

Public Function PopulateFromRS(Rs As ADODB.Recordset)
   BSHEET_ITEM_ID = NVLI(Rs("BSHEET_ITEM_ID"), -1)
   BLOOD_SHEET_ID = NVLI(Rs("BLOOD_SHEET_ID"), 0)
   BLOOD_SPEC_ID = NVLI(Rs("BLOOD_SPEC_ID"), -1)
   SPEC_VALUE = NVLD(Rs("SPEC_VALUE"), -1)
   
   SPEC_NAME = NVLS(Rs("SPEC_NAME"), "")
   MALE_STD = NVLS(Rs("MALE_STD"), "")
   FEMALE_STD = NVLS(Rs("FEMALE_STD"), "")
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
Dim TempID7 As String
Dim TempID8 As String
Dim TempID9 As String

   TempID1 = BLOOD_SHEET_ID
   If BLOOD_SHEET_ID <= 0 Then
      TempID1 = "NULL"
   End If
            
   TempID2 = BLOOD_SPEC_ID
   If BLOOD_SPEC_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("BSHEET_ITEM_SEQ", ID, glbErrorLog)
               
      SQL1 = "INSERT INTO BSHEET_ITEM " & _
                     "( " & _
                     "   BSHEET_ITEM_ID, " & _
                     "   BLOOD_SHEET_ID, " & _
                     "   BLOOD_SPEC_ID, " & _
                     "   SPEC_VALUE, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
                  SQL1 = SQL1 & _
                     "( " & _
                         ID & "," & _
                         TempID1 & "," & _
                         TempID2 & "," & _
                         SPEC_VALUE & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         BSHEET_ITEM_ID = ID
         
   Else 'Show edit
      SQL1 = "UPDATE BSHEET_ITEM SET " & _
                     "    BLOOD_SHEET_ID = " & TempID1 & ", " & _
                     "   BLOOD_SPEC_ID = " & TempID2 & ", " & _
                     "    SPEC_VALUE = " & SPEC_VALUE & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE BSHEET_ITEM_ID = " & BSHEET_ITEM_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
                  
      SQL1 = "DELETE FROM BSHEET_ITEM WHERE BSHEET_ITEM_ID = " & BSHEET_ITEM_ID
      m_Conn.Execute (SQL1)
End Function



