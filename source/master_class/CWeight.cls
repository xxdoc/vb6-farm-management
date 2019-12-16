VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWeight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_LegacyADOConn As ADODB.Connection
Private m_Conn As ADODB.Connection

Private Const MODULE_NAME = "CWeight"
Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long
Public Flag As String
' End of generic data member
Public WEIGHT_ACTUAL_ID As Long
Public WEIGHT_ID As String
Public TRUCK_ID As String
Public CUST_ID As String
Public PRODUCT_ID As String
Public Date1 As Date
Public Date2 As Date
Public Time1  As String
Public Time2  As String
Public WEIGHT1 As Double
Public WEIGHT2 As Double
Public NetWeight As Double
Public REMARK As String
Public DOCUMENT_NO As String
Public RightItems As Collection
Public TempCollection As Collection

Private Sub Class_Initialize()
   Set m_LegacyADOConn = glbDatabaseMngr.DBLegacyConnection
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set RightItems = New Collection
   Set TempCollection = New Collection
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
If Ind = 1 Then
   SelectStr1 = "SELECT sale.*  "
   SelectStr2 = "SELECT COUNT(*) as RC  "
   FromStr = "FROM sale "
   WhereStr = ""
ElseIf Ind = 2 Then
    SelectStr1 = "SELECT tmpsale.*  "
   SelectStr2 = "SELECT COUNT(*) as RC  "
   FromStr = "FROM tmpsale "
   WhereStr = ""
End If
   If Len(CUST_ID) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (cust_id LIKE '" & CUST_ID & "') "
      Else
         WhereStr = WhereStr & "AND (cust_id LIKE '" & CUST_ID & "') "
      End If
   End If
   
    If Len(Date1) > 0 Then
      If Ind = 1 Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (Format(date1,'dd/MM/YYYY') LIKE '" & Date1 & "') "
         Else
            WhereStr = WhereStr & "AND (Format(date1,'dd/MM/YYYY') LIKE '" & Date1 & "') "
         End If
      ElseIf Ind = 2 Then
         If Len(WhereStr) <= 0 Then
            WhereStr = "WHERE (Format(date,'dd/MM/YYYY') LIKE '" & Date1 & "') "
         Else
            WhereStr = WhereStr & "AND (Format(date,'dd/MM/YYYY') LIKE '" & Date1 & "') "
         End If
      End If
   End If
   
   If Ind = 1 Then
      OrderByStr = "ORDER BY date1,time1 "
   ElseIf Ind = 2 Then
      OrderByStr = "ORDER BY date,time "
   End If
   
   OrderTypeStr = "DESC "
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr2.GetLegacyRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      WEIGHT_ID = NVLS(Rs("WEIGHT_ID"), "")
      TRUCK_ID = NVLS(Rs("TRUCK_ID"), "")
      CUST_ID = NVLS(Rs("CUST_ID"), "")
      PRODUCT_ID = NVLS(Rs("PRODUCT_ID"), "")
      Date1 = InternalDateToDate(NVLS(Rs("DATE1"), ""))
      Date2 = InternalDateToDate(NVLS(Rs("DATE1"), ""))
      Time1 = NVLS(Rs("Time1"), "")
      Time2 = NVLS(Rs("Time2"), "")
      WEIGHT1 = NVLD(Rs("WEIGHT1"), 0)
      WEIGHT2 = NVLD(Rs("WEIGHT2"), 0)
      NetWeight = NVLD(Rs("NETWEIGHT"), 0)
      REMARK = NVLS(Rs("REMARK"), "")
   ElseIf Ind = 2 Then
      WEIGHT_ID = NVLS(Rs("WEIGHT_ID"), "")
      TRUCK_ID = NVLS(Rs("TRUCK_ID"), "")
      CUST_ID = NVLS(Rs("CUST_ID"), "")
      PRODUCT_ID = NVLS(Rs("PRODUCT_ID"), "")
      Date1 = InternalDateToDate(NVLS(Rs("DATE"), ""))
      Date2 = -1
      Time1 = NVLS(Rs("Time"), "")
      WEIGHT1 = NVLD(Rs("WEIGHT"), 0)
      REMARK = NVLS(Rs("REMARK"), "")
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("WEIGHT_ACTUAL_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO WEIGHT_ACTUAL " & _
                     "( " & _
                     "   WEIGHT_ACTUAL_ID, " & _
                     "   DOCUMENT_NO, " & _
                     "   WEIGHT_IN, " & _
                     "   WEIGHT_OUT, " & _
                     "   WEIGHT_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "'" & DOCUMENT_NO & "'," & _
                         "" & WEIGHT1 & "," & _
                         "" & WEIGHT2 & "," & _
                         "'" & WEIGHT_ID & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         WEIGHT_ACTUAL_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE WEIGHT_ACTUAL SET " & _
                     "   DOCUMENT_NO = '" & DOCUMENT_NO & "', " & _
                     "   WEIGHT_IN = " & WEIGHT1 & ", " & _
                     "   WEIGHT_OUT = " & WEIGHT2 & ", " & _
                     "   WEIGHT_ID = '" & WEIGHT_ID & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE WEIGHT_ACTUAL_ID = " & WEIGHT_ACTUAL_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM WEIGHT_ACTUAL WHERE WEIGHT_ACTUAL_ID = " & WEIGHT_ACTUAL_ID
      m_Conn.Execute (SQL1)
End Function

Public Function CopyObject(Ind As Long, TempObj As CWeight)
   WEIGHT_ID = TempObj.WEIGHT_ID
   TRUCK_ID = TempObj.TRUCK_ID
   CUST_ID = TempObj.CUST_ID
   PRODUCT_ID = TempObj.PRODUCT_ID
   Date1 = TempObj.Date1
   Date2 = TempObj.Date2
   WEIGHT1 = TempObj.WEIGHT1
   WEIGHT2 = TempObj.WEIGHT2
   Time1 = TempObj.Time1
   Time2 = TempObj.Time2
   REMARK = TempObj.REMARK
   NetWeight = TempObj.NetWeight
End Function



