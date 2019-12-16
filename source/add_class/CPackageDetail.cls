VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPackageDetail"
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
Public PKG_DETAIL_ID As Long
Public PKG_ID As Long
Public STATUS_BUY_ID As Long
Public FROM_WEIGHT As Double
Public TO_WEIGHT As Double
Public CUT_WEIGHT As Double
Public COST_PER_EXCEED As Double
Public PEDIGREE_COST As Double
Public COST_PER_WEIGHT As Double
Public COST_PER_UNIT As Double
Public PKG_BASIC_FLAG As String

Public PRODUCT_STATUS_NAME As String
Public PKG_TYPE As Double

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

   SelectStr1 = "SELECT PKD.*,PDS.PRODUCT_STATUS_NAME,PKG.PKG_BASIC_FLAG,PKG_TYPE "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM PACKAGE_DETAIL PKD  "
   FromStr = FromStr & "LEFT OUTER JOIN PACKAGE PKG ON (PKD.PKG_ID = PKG.PKG_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_STATUS PDS ON (PKD.STATUS_BUY_ID = PDS.PRODUCT_STATUS_ID) "
   
   WhereStr = ""
   
   If PKG_DETAIL_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PKG_DETAIL_ID = " & PKG_DETAIL_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PKG_DETAIL_ID = " & PKG_DETAIL_ID & ") "
      End If
   End If
   
   
   If PKG_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (PKG.PKG_ID = " & PKG_ID & ") "
      Else
         WhereStr = WhereStr & "AND (PKG.PKG_ID = " & PKG_ID & ") "
      End If
   End If
      
   If STATUS_BUY_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (STATUS_BUY_ID = " & STATUS_BUY_ID & ") "
      Else
         WhereStr = WhereStr & "AND (STATUS_BUY_ID = " & STATUS_BUY_ID & ") "
      End If
   End If
      
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY PKG_DETAIL_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY PKG_ID "
   ElseIf OrderBy = 3 Then
      OrderByStr = "ORDER BY STATUS_BY_ID "
   Else
      OrderByStr = "ORDER BY PKG_DETAIL_ID "
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
   PKG_DETAIL_ID = NVLI(Rs("PKG_DETAIL_ID"), -1)
   PKG_ID = NVLI(Rs("PKG_ID"), -1)
   STATUS_BUY_ID = NVLI(Rs("STATUS_BUY_ID"), -1)
   FROM_WEIGHT = NVLD(Rs("FROM_WEIGHT"), 0)
   TO_WEIGHT = NVLD(Rs("TO_WEIGHT"), 0)
   CUT_WEIGHT = NVLD(Rs("CUT_WEIGHT"), 0)
   COST_PER_EXCEED = NVLD(Rs("COST_PER_EXCEED"), 0)
   PEDIGREE_COST = NVLD(Rs("PEDIGREE_COST"), 0)
   COST_PER_WEIGHT = NVLD(Rs("COST_PER_WEIGHT"), 0)
   COST_PER_UNIT = NVLD(Rs("COST_PER_UNIT"), 0)
   
   PKG_TYPE = NVLI(Rs("PKG_TYPE"), -1)
   
   PRODUCT_STATUS_NAME = NVLS(Rs("PRODUCT_STATUS_NAME"), "")
   PKG_BASIC_FLAG = NVLS(Rs("PKG_BASIC_FLAG"), "")
   
'   KEY_ID = PART_ITEM_ID
'   KEY_LOOKUP = PART_NO
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String

   TempID1 = PKG_ID
   If PKG_ID <= 0 Then
      TempID1 = "NULL "
   End If
      
   TempID2 = STATUS_BUY_ID
   If STATUS_BUY_ID <= 0 Then
      TempID2 = "NULL "
   End If
      
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("PACKAGE_DETAIL_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO PACKAGE_DETAIL " & _
                     "( " & _
                     "   PKG_DETAIL_ID, " & _
                     "   PKG_ID, " & _
                     "   STATUS_BUY_ID, " & _
                     "   FROM_WEIGHT, " & _
                     "   TO_WEIGHT, " & _
                     "   CUT_WEIGHT, " & _
                     "   COST_PER_EXCEED, " & _
                     "   PEDIGREE_COST, " & _
                     "   COST_PER_WEIGHT, " & _
                     "   COST_PER_UNIT, " & _
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
                         FROM_WEIGHT & "," & _
                         TO_WEIGHT & "," & _
                         CUT_WEIGHT & "," & _
                         COST_PER_EXCEED & "," & _
                         PEDIGREE_COST & "," & _
                         COST_PER_WEIGHT & "," & _
                         COST_PER_UNIT & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         PKG_DETAIL_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE PACKAGE_DETAIL SET " & _
                     "   PKG_ID = " & TempID1 & ", " & _
                     "   STATUS_BUY_ID = " & TempID2 & ", " & _
                     "    FROM_WEIGHT = " & FROM_WEIGHT & ", " & _
                     "    TO_WEIGHT = " & TO_WEIGHT & ", " & _
                     "    CUT_WEIGHT = " & CUT_WEIGHT & ", " & _
                     "    COST_PER_EXCEED = " & COST_PER_EXCEED & ", " & _
                     "    PEDIGREE_COST = " & PEDIGREE_COST & ", " & _
                     "    COST_PER_WEIGHT = " & COST_PER_WEIGHT & ", " & _
                     "    COST_PER_UNIT = " & COST_PER_UNIT & ", " & _
                     "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE PKG_DETAIL_ID = " & PKG_DETAIL_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM PACKAGE_DETAIL WHERE PKG_DETAIL_ID = " & PKG_DETAIL_ID
      m_Conn.Execute (SQL1)
End Function

