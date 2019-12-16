VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CManagementConfig"
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

Public MANAGEMENT_CONFIG_ID As Long
Public TARGET As Double
Public ACTUAL_BIRTH As Double
Public DIFF As Double
Public AVERAGE As Double
Public BIRTH_DATE As Date
Public MONTH1 As Double
Public LEFT1 As Double
Public MIX1 As Double
Public MONTH2 As Double
Public LEFT2 As Double
Public MIX2 As Double
Public MONTH3 As Double
Public LEFT3 As Double
Public MIX3 As Double
Public MONTH4 As Double
Public LEFT4 As Double
Public MIX4 As Double

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

   SelectStr1 = "SELECT ADR.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM MANAGEMENT_CONFIG ADR "
   
   If MANAGEMENT_CONFIG_ID >= 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (MANAGEMENT_CONFIG_ID = " & MANAGEMENT_CONFIG_ID & ") "
      Else
         WhereStr = WhereStr & "AND (MANAGEMENT_CONFIG_ID = " & MANAGEMENT_CONFIG_ID & ") "
      End If
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""
   If OrderBy > 0 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY MANAGEMENT_CONFIG_ID "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY MANAGEMENT_CONFIG_ID "
      Else
         OrderByStr = "ORDER BY MANAGEMENT_CONFIG_ID "
      End If
      
      If OrderType = 1 Then
         OrderTypeStr = "ASC "
      ElseIf OrderType = 2 Then
         OrderTypeStr = "DESC "
      Else
         OrderTypeStr = "DESC "
      End If
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function
 
Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   MANAGEMENT_CONFIG_ID = NVLI(Rs("MANAGEMENT_CONFIG_ID"), -1)
   TARGET = NVLD(Rs("TARGET"), 0)
   ACTUAL_BIRTH = NVLD(Rs("ACTUAL_BIRTH"), 0)
   DIFF = NVLD(Rs("DIFF"), 0)
   AVERAGE = NVLD(Rs("AVERAGE"), 0)
   BIRTH_DATE = InternalDateToDate(NVLS(Rs("BIRTH_DATE"), ""))
   MONTH1 = NVLD(Rs("MONTH1"), 0)
   LEFT1 = NVLD(Rs("LEFT1"), 0)
   MIX1 = NVLD(Rs("MIX1"), 0)
   MONTH2 = NVLD(Rs("MONTH2"), 0)
   LEFT2 = NVLD(Rs("LEFT2"), 0)
   MIX2 = NVLD(Rs("MIX2"), 0)
   MONTH3 = NVLD(Rs("MONTH3"), 0)
   LEFT3 = NVLD(Rs("LEFT3"), 0)
   MIX3 = NVLD(Rs("MIX3"), 0)
   MONTH4 = NVLD(Rs("MONTH4"), 0)
   LEFT4 = NVLD(Rs("LEFT4"), 0)
   MIX4 = NVLD(Rs("MIX4"), 0)
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("MANAGEMENT_CONFIG_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO MANAGEMENT_CONFIG " & _
                     "( " & _
                     "   MANAGEMENT_CONFIG_ID, " & _
                     "   TARGET, " & _
                     "   ACTUAL_BIRTH, " & _
                     "   DIFF, " & _
                     "   AVERAGE, " & _
                     "   BIRTH_DATE, " & _
                     "   MONTH1, " & _
                     "   LEFT1, " & _
                     "   MIX1, " & _
                     "   MONTH2, " & _
                     "   LEFT2, " & _
                     "   MIX2, " & _
                     "   MONTH3, " & _
                     "   LEFT3, " & _
                     "   MIX3, "
               SQL1 = SQL1 & _
                     "   MONTH4, " & _
                     "   LEFT4, " & _
                     "   MIX4, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         TARGET & "," & _
                         ACTUAL_BIRTH & "," & _
                         DIFF & "," & _
                         AVERAGE & "," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(BIRTH_DATE))) & "'," & _
                         MONTH1 & "," & _
                         LEFT1 & "," & _
                         MIX1 & "," & _
                         MONTH2 & "," & _
                         LEFT2 & "," & _
                         MIX2 & "," & _
                         MONTH3 & "," & _
                         LEFT3 & "," & _
                         MIX3 & "," & _
                         MONTH4 & "," & _
                         LEFT4 & "," & _
                         MIX4 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         MANAGEMENT_CONFIG_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE MANAGEMENT_CONFIG SET " & _
                     "   TARGET = " & TARGET & ", " & _
                     "   ACTUAL_BIRTH = '" & ACTUAL_BIRTH & "', " & _
                     "   DIFF = " & DIFF & ", " & _
                     "   AVERAGE = " & AVERAGE & ", " & _
                     "   BIRTH_DATE = '" & ChangeQuote(Trim(DateToStringInt(BIRTH_DATE))) & "', " & _
                     "   MONTH1 = " & MONTH1 & ", " & _
                     "   LEFT1 = '" & LEFT1 & "', " & _
                     "   MIX1 = " & MIX1 & ", " & _
                     "   MONTH2 = " & MONTH2 & ", " & _
                     "   LEFT2 = '" & LEFT2 & "', " & _
                     "   MIX2 = " & MIX2 & ", " & _
                     "   MONTH3 = " & MONTH3 & ", " & _
                     "   LEFT3 = '" & LEFT3 & "', " & _
                     "   MIX3 = " & MIX3 & ", " & _
                     "   MONTH4 = " & MONTH4 & ", " & _
                     "   LEFT4 = '" & LEFT4 & "', " & _
                     "   MIX4 = " & MIX4 & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE MANAGEMENT_CONFIG_ID = " & MANAGEMENT_CONFIG_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM MANAGEMENT_CONFIG WHERE MANAGEMENT_CONFIG_ID = " & MANAGEMENT_CONFIG_ID
      m_Conn.Execute (SQL1)
End Function

