VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGLAgeAmount"
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

Public GL_AGE_AMOUNT_ID As Long
Public GL_AGE As Long
Public GL_AMOUNT As Double
Public PARAM_ID As Long

Public PIG_TYPE As Long
Public PIG_TYPE_NAME As String
Public FROM_DATE As Date
Public TO_DATE As Date
 
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_DATE = -1
   TO_DATE = -1
End Sub
'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, itemcount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim TempDate As String
Dim GroupByStr As String
Dim MustCount As Boolean
   
   MustCount = True
   
   If Ind = 1 Then
      SelectStr1 = "SELECT GLA.*,PM.PIG_TYPE,PT.PRODUCT_TYPE_NO  PIG_TYPE_NAME  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM GL_AGE_AMOUNT GLA "
      FromStr = FromStr & "LEFT OUTER JOIN PARAMETERS PM ON (PM.PARAM_ID = GLA.PARAM_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN PRODUCT_TYPE PT ON (PT.PRODUCT_TYPE_ID = PM.PIG_TYPE) "
      
      WhereStr = ""
   ElseIf Ind = 2 Then
      
   End If
   
   If GL_AGE_AMOUNT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (GLA.GL_AGE_AMOUNT_ID = " & GL_AGE_AMOUNT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (GLA.GL_AGE_AMOUNT_ID = " & GL_AGE_AMOUNT_ID & ") "
      End If
   End If
      
   If PARAM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (GLA.PARAM_ID = " & PARAM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (GLA.PARAM_ID = " & PARAM_ID & ") "
      End If
   Else
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE ((GLA.PARAM_ID <= 0) OR (PARAM_ID IS NULL)) "
      Else
         WhereStr = WhereStr & "AND ((GLA.PARAM_ID <= 0) OR (PARAM_ID IS NULL)) "
      End If
   End If
   
'   If Ind = 17 Then
'      GroupByStr = "GROUP BY PI.FEED_GROUP "
'   End If
      
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY GL_AGE_AMOUNT_ID "
      Else
         OrderByStr = "ORDER BY GL_AGE "
      End If
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, MustCount, itemcount, Rs, glbErrorLog)
End Function
Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      GL_AGE_AMOUNT_ID = NVLI(Rs("GL_AGE_AMOUNT_ID"), 0)
      GL_AGE = NVLD(Rs("GL_AGE"), 0)
      GL_AMOUNT = NVLD(Rs("GL_AMOUNT"), 0)
      PARAM_ID = NVLI(Rs("PARAM_ID"), 0)
      
      PIG_TYPE = NVLI(Rs("PIG_TYPE"), 0)
      PIG_TYPE_NAME = NVLS(Rs("PIG_TYPE_NAME"), "")
   ElseIf Ind = 2 Then
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("GL_AGE_AMOUNT_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO GL_AGE_AMOUNT " & _
                     "( " & _
                     "   GL_AGE_AMOUNT_ID, " & _
                     "   PARAM_ID, " & _
                     "   GL_AGE, " & _
                     "   GL_AMOUNT, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
            SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                        PARAM_ID & "," & _
                        GL_AGE & "," & _
                        GL_AMOUNT & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         GL_AGE_AMOUNT_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE GL_AGE_AMOUNT SET " & _
                     "    PARAM_ID = " & PARAM_ID & ", " & _
                     "    GL_AGE = " & GL_AGE & ", " & _
                     "    GL_AMOUNT = " & GL_AMOUNT & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE GL_AGE_AMOUNT_ID = " & GL_AGE_AMOUNT_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function
Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM GL_AGE_AMOUNT WHERE GL_AGE_AMOUNT_ID = " & GL_AGE_AMOUNT_ID
      m_Conn.Execute (SQL1)
End Function
