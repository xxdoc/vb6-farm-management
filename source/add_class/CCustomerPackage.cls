VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCustomerPackage"
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
Public CUSTOMER_PACKAGE_ID As Long
Public CUSTOMER_ID As Long
Public PKG_TYPE As Long
Public PKG_ID As Long

Public PACKAGE_TYPE_NAME As String
Public PKG_NAME As String

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
Dim SubLen As Long
Dim NewStr As String

   SelectStr1 = "SELECT CPK.*,PKG.PKG_NAME,PKT.PACKAGE_TYPE_NAME "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM CUSTOMER_PACKAGE CPK  "
   FromStr = FromStr & "LEFT OUTER JOIN PACKAGE PKG ON (CPK.PKG_ID = PKG.PKG_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN PACKAGE_TYPE PKT ON (CPK.PKG_TYPE = PKT.PACKAGE_TYPE_ID) "
   
   WhereStr = ""
   
   If CUSTOMER_PACKAGE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CUSTOMER_PACKAGE_ID = " & CUSTOMER_PACKAGE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CUSTOMER_PACKAGE_ID = " & CUSTOMER_PACKAGE_ID & ") "
      End If
   End If
   
   
   If PKG_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PKG_ID = " & PKG_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PKG_ID = " & PKG_ID & ") "
      End If
   End If
      
   If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (CUSTOMER_ID = " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If
   
   If PKG_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PKG_TYPE = " & PKG_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (PKG_TYPE = " & PKG_TYPE & ") "
      End If
   End If
      
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY CUSTOMER_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PKG_TYPE "
   ElseIf OrderBy = 3 Then
      OrderByStr = "ORDER BY PKG_TPYE "
   Else
      OrderByStr = "ORDER BY CUSTOMER_PACKAGE_ID "
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
   CUSTOMER_PACKAGE_ID = NVLI(Rs("CUSTOMER_PACKAGE_ID"), -1)
   CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
   PKG_TYPE = NVLI(Rs("PKG_TYPE"), -1)
   PKG_ID = NVLI(Rs("PKG_ID"), -1)
   
   PKG_NAME = NVLS(Rs("PKG_NAME"), "")
   PACKAGE_TYPE_NAME = NVLS(Rs("PACKAGE_TYPE_NAME"), "")
   
'   KEY_ID = PART_ITEM_ID
'   KEY_LOOKUP = PART_NO
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String

   TempID1 = CUSTOMER_ID
   If CUSTOMER_ID <= 0 Then
      TempID1 = "NULL "
   End If
      
   TempID2 = PKG_TYPE
   If PKG_TYPE <= 0 Then
      TempID2 = "NULL "
   End If
    
    TempID3 = PKG_ID
   If PKG_ID <= 0 Then
      TempID3 = "NULL "
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("CUSTOMER_PACKAGE_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO CUSTOMER_PACKAGE " & _
                     "( " & _
                     "   CUSTOMER_PACKAGE_ID, " & _
                     "   CUSTOMER_ID, " & _
                     "   PKG_TYPE, " & _
                     "   PKG_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "" & TempID1 & "," & _
                         "" & TempID2 & "," & _
                         "" & TempID3 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         CUSTOMER_PACKAGE_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE CUSTOMER_PACKAGE SET " & _
                     "   CUSTOMER_ID = " & TempID1 & ", " & _
                     "   PKG_TYPE = " & TempID2 & ", " & _
                     "   PKG_ID = " & TempID3 & ", " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE CUSTOMER_PACKAGE_ID = " & CUSTOMER_PACKAGE_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM CUSTOMER_PACKAGE WHERE CUSTOMER_PACKAGE_ID = " & CUSTOMER_PACKAGE_ID
      m_Conn.Execute (SQL1)
End Function

