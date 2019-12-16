VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBalanceAccumBack"
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

Public BALANCE_ACCUM_ID As Long
Public DOCUMENT_DATE As Date
Public PART_ITEM_ID As Long
Public IMPORT_AMOUNT As Double
Public EXPORT_AMOUNT As Double
Public BALANCE_AMOUNT As Double
Public TOTAL_INCLUDE_PRICE As Double
Public LOCATION_ID As Long
Public AVG_PRICE As Double
Public YYYYMM As String
Public PIG_AGE As Long
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub
Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String
Dim TempID3 As String

   TempID1 = PART_ITEM_ID
   If PART_ITEM_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = LOCATION_ID
   If LOCATION_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   If AddEditMode = SHOW_ADD Then
      
      SQL1 = "INSERT INTO BALANCE_ACCUM_BACK " & _
                     "( " & _
                     "   DOCUMENT_DATE, " & _
                     "   PART_ITEM_ID, " & _
                     "   IMPORT_AMOUNT, " & _
                     "   EXPORT_AMOUNT, " & _
                     "   BALANCE_AMOUNT, " & _
                     "   TOTAL_INCLUDE_PRICE, " & _
                     "   LOCATION_ID, " & _
                     "   AVG_PRICE, " & _
                     "   YYYYMM, " & _
                     "   PIG_AGE, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") "
            SQL1 = SQL1 & _
                     "VALUES " & _
                     "( " & _
                         "'" & ChangeQuote(DateToStringInt(Trim(DOCUMENT_DATE))) & "'," & _
                        PART_ITEM_ID & "," & _
                        IMPORT_AMOUNT & "," & _
                        EXPORT_AMOUNT & "," & _
                        BALANCE_AMOUNT & "," & _
                        TOTAL_INCLUDE_PRICE & "," & _
                        TempID2 & "," & _
                        AVG_PRICE & "," & _
                        "'" & ChangeQuote(Trim(YYYYMM)) & "'," & _
                        PIG_AGE & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
   End If
   
   If (AddEditMode = SHOW_ADD) Then
      m_Conn.Execute (SQL1)
   End If
End Function
