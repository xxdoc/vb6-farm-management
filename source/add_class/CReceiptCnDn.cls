VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReceiptCnDn"
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

Public RECEIPT_CNDN_ID As Long
Public BILLING_DOC_ID As Long
Public DO_ID As Long
Public CN_AMOUNT As Double
Public DN_AMOUNT As Double
Public CNDN_ID As Long

Public CNDN_NO As String
Public CNDN_DATE As Date
Public DO_NO As String

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
Dim TempDate As String
Dim SubLen As Long
Dim NewStr As String
Dim GroupByStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT UG.*, BD2.DOCUMENT_NO DO_NO, BD3.DOCUMENT_NO CNDN_NO, BD3.DOCUMENT_DATE CNDN_DATE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM RECEIPT_CNDN UG "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD1 ON (UG.BILLING_DOC_ID = BD1.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD2 ON (UG.DO_ID = BD2.BILLING_DOC_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN BILLING_DOC BD3 ON (UG.CNDN_ID = BD3.BILLING_DOC_ID) "
      
      WhereStr = ""
   End If

   If RECEIPT_CNDN_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (RECEIPT_CNDN_ID = " & RECEIPT_CNDN_ID & ") "
      Else
         WhereStr = WhereStr & "AND (RECEIPT_CNDN_ID = " & RECEIPT_CNDN_ID & ") "
      End If
   End If
   
   If BILLING_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BD1.BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BD1.BILLING_DOC_ID = " & BILLING_DOC_ID & ") "
      End If
   End If
         
   GroupByStr = ""
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      OrderByStr = "ORDER BY RECEIPT_CNDN_ID "
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      RECEIPT_CNDN_ID = NVLI(Rs("RECEIPT_CNDN_ID"), -1)
      BILLING_DOC_ID = NVLI(Rs("BILLING_DOC_ID"), -1)
      DO_ID = NVLI(Rs("DO_ID"), -1)
      CN_AMOUNT = NVLD(Rs("CN_AMOUNT"), 0)
      DN_AMOUNT = NVLD(Rs("DN_AMOUNT"), 0)
      CNDN_ID = NVLI(Rs("CNDN_ID"), -1)
      
      CNDN_NO = NVLS(Rs("CNDN_NO"), "")
      CNDN_DATE = InternalDateToDate(NVLS(Rs("CNDN_DATE"), ""))
      DO_NO = NVLS(Rs("DO_NO"), "")
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim ID2 As Long
Dim Bd As CBillingDoc
Dim TempID1 As String
Dim TempID2 As String

   TempID1 = CNDN_ID
   If CNDN_ID <= 0 Then
      TempID1 = "NULL"
   End If
      
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("RECEIPT_CNDN_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO RECEIPT_CNDN " & _
                     "( " & _
                     "   RECEIPT_CNDN_ID, " & _
                     "   BILLING_DOC_ID, " & _
                     "   DO_ID, " & _
                     "   CN_AMOUNT, "
            SQL1 = SQL1 & _
                     "   DN_AMOUNT, " & _
                     "   CNDN_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "" & BILLING_DOC_ID & "," & _
                         "" & DO_ID & "," & _
                         "" & CN_AMOUNT & ","
            SQL1 = SQL1 & _
                        "" & DN_AMOUNT & "," & _
                        TempID1 & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         RECEIPT_CNDN_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE RECEIPT_CNDN SET " & _
                     "   BILLING_DOC_ID = " & BILLING_DOC_ID & ", " & _
                     "   DO_ID = " & DO_ID & ", " & _
                     "   CN_AMOUNT = " & CN_AMOUNT & ", " & _
                     "   DN_AMOUNT = " & DN_AMOUNT & ", " & _
                     "   CNDN_ID = " & TempID1 & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE RECEIPT_CNDN_ID = " & RECEIPT_CNDN_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
Dim Bd As CBillingDoc

   SQL1 = "DELETE FROM RECEIPT_CNDN WHERE RECEIPT_CNDN_ID = " & RECEIPT_CNDN_ID
   m_Conn.Execute (SQL1)
End Function



