VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAccount"
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

Public ACCOUNT_ID As Long
Public ACCOUNT_NO As String
Public ACCOUNT_TYPE As Long
Public ACCOUNT_STATUS As Long
Public CUSTOMER_ID As Long

Public CUSTOMER_NAME As String
Public CUSTOMER_CODE As String

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
Dim SubLen As Long
Dim NewStr As String

   SelectStr1 = "SELECT AC.*, NM.LONG_NAME, CM.CUSTOMER_CODE  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM ACCOUNT AC, NAME NM, CUSTOMER_NAME CN, CUSTOMER CM "
   
   WhereStr = "WHERE (AC.CUSTOMER_ID = CM.CUSTOMER_ID) AND (CM.CUSTOMER_ID = CN.CUSTOMER_ID) AND (CN.NAME_ID = NM.NAME_ID) "
   If ACCOUNT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (AC.ACCOUNT_ID = " & ACCOUNT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (AC.ACCOUNT_ID = " & ACCOUNT_ID & ") "
      End If
   End If
      
   If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CM.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CM.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If
         
   OrderByStr = ""
   OrderTypeStr = ""
   
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY AC.ACCOUNT_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY AC.ACCOUNT_ID "
   Else
      OrderByStr = "ORDER BY AC.ACCOUNT_ID "
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
   ACCOUNT_ID = NVLI(Rs("ACCOUNT_ID"), -1)
   ACCOUNT_NO = NVLS(Rs("ACCOUNT_NO"), "")
   ACCOUNT_STATUS = NVLI(Rs("ACCOUNT_STATUS"), -1)
   CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
   ACCOUNT_TYPE = NVLI(Rs("ACCOUNT_TYPE"), -1)
   
   CUSTOMER_NAME = NVLS(Rs("LONG_NAME"), "")
   CUSTOMER_CODE = NVLS(Rs("CUSTOMER_CODE"), "")
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String

   TempID1 = ACCOUNT_TYPE
   If ACCOUNT_TYPE <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = ACCOUNT_STATUS
   If ACCOUNT_STATUS <= 0 Then
      TempID2 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("ACCOUNT_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO ACCOUNT " & _
                     "( " & _
                     "   ACCOUNT_ID, " & _
                     "   ACCOUNT_NO, " & _
                     "   ACCOUNT_STATUS, " & _
                     "   CUSTOMER_ID, " & _
                     "   ACCOUNT_TYPE, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(ACCOUNT_NO)) & "'," & _
                         "" & TempID2 & "," & _
                         " " & CUSTOMER_ID & "," & _
                         "" & TempID1 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         ACCOUNT_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE ACCOUNT SET " & _
                     "   ACCOUNT_NO = '" & ChangeQuote(Trim(ACCOUNT_NO)) & "', " & _
                     "   ACCOUNT_STATUS = " & TempID2 & ", " & _
                     "   CUSTOMER_ID = '" & CUSTOMER_ID & "', " & _
                     "   ACCOUNT_TYPE = " & TempID1 & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE ACCOUNT_ID = " & ACCOUNT_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
            
      SQL1 = "DELETE FROM ACCOUNT WHERE ACCOUNT_ID = " & ACCOUNT_ID
      m_Conn.Execute (SQL1)
End Function

