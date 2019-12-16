VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInventoryDoc"
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

Public INVENTORY_DOC_ID As Long
Public DOCUMENT_NO As String
Public DOCUMENT_DATE As Date
Public DOCUMENT_DESC As String
Public BILL_NO As String
Public DO_NO As String
Public TRUCK_NO As String
Public SUPPLIER_ID As Long
Public DELIVERY_ID As Long
Public DELIVERY_FEE As Double
Public SENDER_NAME As String
Public RECEIVE_NAME As String
Public DOCUMENT_TYPE As Long
Public EMP_ID As Long
Public SURETY_ID As Long
Public COMMIT_FLAG As String
Public SALE_FLAG As String
Public EXCEPTION_FLAG As String
Public DOCUMENT_SUBTYPE As Long
Public BATCH_ID As Long
Public SIMULATE_FLAG As String
Public DUE_DATE As Date

Public OLD_COMMIT_FLAG As String
Public PART_NO As String
Public SUPPLIER_CODE As String
Public SUPPLIER_NAME As String
Public RESPONSE_NAME As String
Public RESPONSE_LNAME As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public DELETE_BALANCE_FLAG As String
Public YYYYMM As String
Public YYYYMM2 As String
Public CUS_ID As Long
Public PO_NO As String

Public PART_NO_EXPORT As String
Public PART_NO_IMPORT As String
Public PIG_NO_EXPORT As String

Public PARTGROUP_ID_EXPORT As Long
Public PARTGROUP_ID_IMPORT As Long

Public EXPORTER_NAME As String
Public EXPORTER_LNAME As String

Public RUNNING_NO As Long
Public CONFIG_DOC_TYPE As Long


Public ImportItems As Collection
Public ExportItems As Collection
Public ImportExports As Collection
Public TransferItems As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_DATE = -1
   TO_DATE = -1
   DOCUMENT_DATE = -1
   COMMIT_FLAG = "N"
   SALE_FLAG = "N"
   
   Set ImportItems = New Collection
   Set ExportItems = New Collection
   Set ImportExports = New Collection
   Set TransferItems = New Collection
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

   SIMULATE_FLAG = glbUser.SIMULATE_FLAG
   
   SelectStr1 = "SELECT ADR.*, NM1.LONG_NAME SUPPLIER_NAME, NM3.LONG_NAME RESPONSE_NAME, NM3.LAST_NAME RESPONSE_LNAME, PR.PRIORITY1  "
   SelectStr1 = SelectStr1 & ", NM4.LONG_NAME EXPORTER_NAME, NM4.LAST_NAME EXPORTER_LNAME "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM INVENTORY_DOC ADR "
   FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SN.SUPPLIER_ID = ADR.SUPPLIER_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM1 ON (SN.NAME_ID = NM1.NAME_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN ON (EN.EMP_ID = ADR.EMP_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM3 ON (EN.NAME_ID = NM3.NAME_ID) "
   
   FromStr = FromStr & "LEFT OUTER JOIN EMPLOYEE_NAME EN1 ON (EN1.EMP_ID = ADR.SURETY_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM4 ON (EN1.NAME_ID = NM4.NAME_ID) "
   
   FromStr = FromStr & "LEFT OUTER JOIN DOCUMENT_PRIORITY PR ON ((ADR.DOCUMENT_TYPE = PR.DOCUMENT_TYPE) AND (PR.AREA = 1)) "
   
   If INVENTORY_DOC_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") "
      Else
         WhereStr = WhereStr & "AND (INVENTORY_DOC_ID = " & INVENTORY_DOC_ID & ") "
      End If
   End If
         
   If DOCUMENT_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ADR.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (ADR.DOCUMENT_TYPE = " & DOCUMENT_TYPE & ") "
      End If
   End If
   
   
   
   If Len(COMMIT_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ADR.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (ADR.COMMIT_FLAG = '" & COMMIT_FLAG & "') "
      End If
   End If
   
   If SIMULATE_FLAG = "Y" Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SIMULATE_FLAG = '" & SIMULATE_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (SIMULATE_FLAG = '" & SIMULATE_FLAG & "') "
      End If
   Else
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE ((SIMULATE_FLAG <> 'Y') OR (SIMULATE_FLAG IS NULL)) "
      Else
         WhereStr = WhereStr & "AND ((SIMULATE_FLAG <> 'Y') OR (SIMULATE_FLAG IS NULL)) "
      End If
   End If
   
   If Trim(DOCUMENT_NO) <> "" Then
      If WildCard(DOCUMENT_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(DOCUMENT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "')"
         End If
      End If
   End If
         
   If Trim(SUPPLIER_CODE) <> "" Then
      Call SearchSupplierCode(FromStr)
      If WildCard(SUPPLIER_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(SUPPLIER_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(SUPPLIER_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (SUPPLIER_CODE = '" & ChangeQuote(Trim(SUPPLIER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUPPLIER_CODE = '" & ChangeQuote(Trim(SUPPLIER_CODE)) & "')"
         End If
      End If
   End If
   
   If DOCUMENT_DATE > -1 Then
      TempDate = Trim(DateToStringInt(DOCUMENT_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DOCUMENT_DATE = '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_DATE = '" & ChangeQuote(TempDate) & "')"
      End If
   End If
      
 If FROM_DATE > -1 Then
   TempDate = DateToStringIntLow(Trim(FROM_DATE))
   If WhereStr = "" Then
      WhereStr = "WHERE (DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
   Else
      WhereStr = WhereStr & "AND (DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
   End If
End If

If TO_DATE > -1 Then
   TempDate = DateToStringIntHi(Trim(TO_DATE))
   If WhereStr = "" Then
      WhereStr = "WHERE (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
   Else
      WhereStr = WhereStr & "AND (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
   End If
End If
      
   If BATCH_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (BATCH_ID = " & BATCH_ID & ") "
      Else
         WhereStr = WhereStr & "AND (BATCH_ID = " & BATCH_ID & ") "
      End If
   Else
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE ((BATCH_ID <= 0) OR (BATCH_ID IS NULL)) "
      Else
         WhereStr = WhereStr & "AND ((BATCH_ID <= 0) OR (BATCH_ID IS NULL)) "
      End If
   End If
     
   If Len(PART_NO_EXPORT) > 0 Then
      If Len(WhereStr) <= 0 Then
         If WildCard(PART_NO_EXPORT, SubLen, NewStr) Then
            WhereStr = "WHERE ((SELECT COUNT(*) FROM EXPORT_ITEM EI LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (EI.INVENTORY_DOC_ID = ADR.INVENTORY_DOC_ID) AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = "WHERE ((SELECT COUNT(*) FROM EXPORT_ITEM EI LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (EI.INVENTORY_DOC_ID = ADR.INVENTORY_DOC_ID) AND (PI.PART_NO = '" & ChangeQuote(Trim(PART_NO_EXPORT)) & "')) > 0) "
         End If
      Else
         If WildCard(PART_NO_EXPORT, SubLen, NewStr) Then
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM EXPORT_ITEM EI LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (EI.INVENTORY_DOC_ID = ADR.INVENTORY_DOC_ID) AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM EXPORT_ITEM EI LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (EI.INVENTORY_DOC_ID = ADR.INVENTORY_DOC_ID) AND (PI.PART_NO = '" & ChangeQuote(Trim(PART_NO_EXPORT)) & "')) > 0) "
         End If
      End If
   End If
   
   If Len(PART_NO_IMPORT) > 0 Then
      If Len(WhereStr) <= 0 Then
         If WildCard(PART_NO_IMPORT, SubLen, NewStr) Then
            WhereStr = "WHERE ((SELECT COUNT(*) FROM IMPORT_ITEM EI LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (EI.INVENTORY_DOC_ID = ADR.INVENTORY_DOC_ID) AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = "WHERE ((SELECT COUNT(*) FROM IMPORT_ITEM EI LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (EI.INVENTORY_DOC_ID = ADR.INVENTORY_DOC_ID) AND (PI.PART_NO = '" & ChangeQuote(Trim(PART_NO_IMPORT)) & "')) > 0) "
         End If
      Else
         If WildCard(PART_NO_IMPORT, SubLen, NewStr) Then
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM IMPORT_ITEM EI LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (EI.INVENTORY_DOC_ID = ADR.INVENTORY_DOC_ID) AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM IMPORT_ITEM EI LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) WHERE (EI.INVENTORY_DOC_ID = ADR.INVENTORY_DOC_ID) AND (PI.PART_NO = '" & ChangeQuote(Trim(PART_NO_IMPORT)) & "')) > 0) "
         End If
      End If
   End If
   
   If Len(PIG_NO_EXPORT) > 0 Then
      If Len(WhereStr) <= 0 Then
         If WildCard(PIG_NO_EXPORT, SubLen, NewStr) Then
            WhereStr = "WHERE ((SELECT COUNT(*) FROM EXPORT_ITEM EI LEFT OUTER JOIN PART_ITEM PI ON (EI.PIG_ID = PI.PART_ITEM_ID) WHERE (EI.INVENTORY_DOC_ID = ADR.INVENTORY_DOC_ID) AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = "WHERE ((SELECT COUNT(*) FROM EXPORT_ITEM EI LEFT OUTER JOIN PART_ITEM PI ON (EI.PIG_ID = PI.PART_ITEM_ID) WHERE (EI.INVENTORY_DOC_ID = ADR.INVENTORY_DOC_ID) AND (PI.PART_NO = '" & ChangeQuote(Trim(PIG_NO_EXPORT)) & "')) > 0) "
         End If
      Else
         If WildCard(PIG_NO_EXPORT, SubLen, NewStr) Then
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM EXPORT_ITEM EI LEFT OUTER JOIN PART_ITEM PI ON (EI.PIG_ID = PI.PART_ITEM_ID) WHERE (EI.INVENTORY_DOC_ID = ADR.INVENTORY_DOC_ID) AND (SUBSTR(PI.PART_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')) > 0) "
         Else
            WhereStr = WhereStr & "AND ((SELECT COUNT(*) FROM EXPORT_ITEM EI LEFT OUTER JOIN PART_ITEM PI ON (EI.PIG_ID = PI.PART_ITEM_ID) WHERE (EI.INVENTORY_DOC_ID = ADR.INVENTORY_DOC_ID) AND (PI.PART_NO = '" & ChangeQuote(Trim(PIG_NO_EXPORT)) & "')) > 0) "
         End If
      End If
   End If
   
   
   If SUPPLIER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ADR.SUPPLIER_ID = " & SUPPLIER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (ADR.SUPPLIER_ID = " & SUPPLIER_ID & ") "
      End If
   End If
   
   
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      If OrderType = 1 Then
         OrderByStr = "ORDER BY DOCUMENT_NO ASC, INVENTORY_DOC_ID "
      Else
         OrderByStr = "ORDER BY DOCUMENT_NO DESC, INVENTORY_DOC_ID "
      End If
   ElseIf OrderBy = 2 Then
      If OrderType = 1 Then
         OrderByStr = "ORDER BY DOCUMENT_DATE ASC, INVENTORY_DOC_ID "
      Else
         OrderByStr = "ORDER BY DOCUMENT_DATE DESC, INVENTORY_DOC_ID "
      End If
   ElseIf OrderBy = 3 Then
      OrderByStr = "ORDER BY ADR.DOCUMENT_DATE ASC, PR.PRIORITY1 ASC, ADR.DOCUMENT_NO "
   Else
      OrderByStr = "ORDER BY INVENTORY_DOC_ID "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      If Len(OrderByStr) <= 0 Then
         OrderTypeStr = ""
      Else
         OrderTypeStr = "DESC "
      End If
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PatchDocDate() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE INVENTORY_DOC SET " & _
                  "   DOCUMENT_DATE = '" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "', " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE INVENTORY_DOC_ID = " & INVENTORY_DOC_ID
                  
   m_Conn.Execute (SQL1)
End Function

Public Function UpdateCommitFlag() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE INVENTORY_DOC SET " & _
                  "   COMMIT_FLAG = '" & COMMIT_FLAG & "', " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE INVENTORY_DOC_ID = " & INVENTORY_DOC_ID
                  
   m_Conn.Execute (SQL1)
End Function

Public Function DeleteData()
Dim SQL1 As String
Dim TempRs As ADODB.Recordset
Dim II As CImportItem
Dim EI As CExportItem
Dim iCount As Long
Dim I As Long
   Set TempRs = New ADODB.Recordset
   
   I = 0
   Set II = New CImportItem
   II.IMPORT_ITEM_ID = -1
   II.INVENTORY_DOC_ID = INVENTORY_DOC_ID
   Call II.QueryData(1, TempRs, iCount)
   While Not TempRs.EOF
      I = I + 1
      '''debug.print (I)
      Call II.PopulateFromRS(1, TempRs)
      Call II.DeleteData
      Call TempRs.MoveNext
   Wend
   Set II = Nothing
   
   I = 0
   Set EI = New CExportItem
   EI.EXPORT_ITEM_ID = -1
   EI.INVENTORY_DOC_ID = INVENTORY_DOC_ID
   Call EI.QueryData(1, TempRs, iCount)
   While Not TempRs.EOF
      I = I + 1
      '''debug.print (I)
      Call EI.PopulateFromRS(1, TempRs)
      Call EI.DeleteData
      Call TempRs.MoveNext
   Wend
   Set EI = Nothing
   
   If TempRs.State = adStateOpen Then
      Call TempRs.Close
   End If
   Set TempRs = Nothing
                        
   SQL1 = "DELETE FROM INVENTORY_DOC WHERE INVENTORY_DOC_ID = " & INVENTORY_DOC_ID
   m_Conn.Execute (SQL1)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   INVENTORY_DOC_ID = NVLI(Rs("INVENTORY_DOC_ID"), -1)
   DOCUMENT_NO = NVLS(Rs("DOCUMENT_NO"), "")
   SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), -1)
   DELIVERY_ID = NVLI(Rs("DELIVERY_ID"), 0)
   DELIVERY_FEE = NVLD(Rs("DELIVERY_FEE"), 0)
   DOCUMENT_DESC = NVLS(Rs("DOCUMENT_DESC"), "")
   RECEIVE_NAME = NVLS(Rs("RECEIVE_NAME"), "")
   BILL_NO = NVLS(Rs("BILL_NO"), "")
   DO_NO = NVLS(Rs("DO_NO"), "")
   DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
   TRUCK_NO = DecryptText(NVLS(Rs("TRUCK_NO"), ""))
   DOCUMENT_TYPE = NVLI(Rs("DOCUMENT_TYPE"), -1)
   SENDER_NAME = NVLS(Rs("SENDER_NAME"), "")
   EMP_ID = NVLI(Rs("EMP_ID"), -1)
   SURETY_ID = NVLI(Rs("SURETY_ID"), -1)
   COMMIT_FLAG = NVLS(Rs("COMMIT_FLAG"), "")
   OLD_COMMIT_FLAG = COMMIT_FLAG
   SALE_FLAG = NVLS(Rs("SALE_FLAG"), "")
   EXCEPTION_FLAG = NVLS(Rs("EXCEPTION_FLAG"), "")
   DOCUMENT_SUBTYPE = NVLI(Rs("DOCUMENT_SUBTYPE"), -1)
   BATCH_ID = NVLI(Rs("BATCH_ID"), -1)
   SIMULATE_FLAG = NVLS(Rs("SIMULATE_FLAG"), "")
   YYYYMM = NVLS(Rs("YYYYMM"), "")
   DUE_DATE = InternalDateToDate(NVLS(Rs("DUE_DATE"), ""))
   YYYYMM2 = NVLS(Rs("YYYYMM2"), "")
   CUS_ID = NVLI(Rs("CUS_ID"), -1)
   PO_NO = NVLS(Rs("PO_NO"), "")
   
   SUPPLIER_NAME = NVLS(Rs("SUPPLIER_NAME"), "")
   RESPONSE_NAME = NVLS(Rs("RESPONSE_NAME"), "")
   RESPONSE_LNAME = NVLS(Rs("RESPONSE_LNAME"), "")
   
   EXPORTER_NAME = NVLS(Rs("EXPORTER_NAME"), "")
   EXPORTER_LNAME = NVLS(Rs("EXPORTER_LNAME"), "")
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

   TempID1 = SUPPLIER_ID
   If SUPPLIER_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = DELIVERY_ID
   If DELIVERY_ID <= 0 Then
      TempID2 = "NULL"
   End If
   
   TempID3 = EMP_ID
   If EMP_ID <= 0 Then
      TempID3 = "NULL"
   End If
   
   TempID4 = BATCH_ID
   If BATCH_ID <= 0 Then
      TempID4 = "NULL"
   End If
   
   TempID5 = CUS_ID
   If CUS_ID <= 0 Then
      TempID5 = "NULL"
   End If
   
   TempID6 = SURETY_ID
   If SURETY_ID <= 0 Then
      TempID6 = "NULL"
   End If
   
   If DUE_DATE <= 0 Then
      DUE_DATE = DOCUMENT_DATE
   End If
   
   SIMULATE_FLAG = glbUser.SIMULATE_FLAG
   YYYYMM = Mid(DateToStringInt(DOCUMENT_DATE), 1, 7)
   YYYYMM2 = Mid(DateToStringInt(DUE_DATE), 1, 7)
   
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("INVENTORY_DOC_SEQ", ID, glbErrorLog)

      SQL1 = "INSERT INTO INVENTORY_DOC " & _
                     "( " & _
                     "   INVENTORY_DOC_ID, " & _
                     "   DOCUMENT_NO, " & _
                     "   SUPPLIER_ID, " & _
                     "   DELIVERY_ID, " & _
                     "   DOCUMENT_DESC, " & _
                     "   RECEIVE_NAME, " & _
                     "   BILL_NO, " & _
                     "   DO_NO, " & _
                     "   TRUCK_NO, " & _
                     "   DOCUMENT_DATE, " & _
                     "   DOCUMENT_TYPE, " & _
                     "   SENDER_NAME, " & _
                     "   DELIVERY_FEE, " & _
                     "   EMP_ID, " & _
                     "   SURETY_ID, " & _
                     "   COMMIT_FLAG, " & _
                     "   SALE_FLAG, " & _
                     "   EXCEPTION_FLAG, "
         SQL1 = SQL1 & _
                     "   DOCUMENT_SUBTYPE, " & _
                     "   BATCH_ID, " & _
                     "   SIMULATE_FLAG, " & _
                     "   YYYYMM, " & _
                     "   YYYYMM2, " & _
                     "   DUE_DATE, " & _
                     "   CUS_ID, " & _
                     "   PO_NO, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(DOCUMENT_NO)) & "'," & _
                         TempID1 & "," & _
                         TempID2 & "," & _
                         "'" & ChangeQuote(Trim(DOCUMENT_DESC)) & "'," & _
                         "'" & ChangeQuote(Trim(RECEIVE_NAME)) & "'," & _
                         "'" & ChangeQuote(Trim(BILL_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(DO_NO)) & "'," & _
                         "'" & ChangeQuote(Trim(EncryptText(TRUCK_NO))) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "'," & _
                         DOCUMENT_TYPE & "," & _
                         "'" & ChangeQuote(Trim(SENDER_NAME)) & "'," & _
                         DELIVERY_FEE & "," & _
                         TempID3 & "," & _
                         TempID6 & "," & _
                         "'" & ChangeQuote(Trim(COMMIT_FLAG)) & "'," & _
                         "'" & ChangeQuote(Trim(SALE_FLAG)) & "'," & _
                         "'" & ChangeQuote(Trim(EXCEPTION_FLAG)) & "'," & _
                         DOCUMENT_SUBTYPE & "," & _
                         TempID4 & ","
               SQL1 = SQL1 & _
                         "'" & ChangeQuote(Trim(SIMULATE_FLAG)) & "'," & _
                         "'" & ChangeQuote(Trim(YYYYMM)) & "'," & _
                         "'" & ChangeQuote(Trim(YYYYMM2)) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(DUE_DATE))) & "'," & _
                         TempID5 & "," & _
                         "'" & ChangeQuote(Trim(PO_NO)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         INVENTORY_DOC_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE INVENTORY_DOC SET " & _
                     "   DOCUMENT_NO = '" & ChangeQuote(Trim(DOCUMENT_NO)) & "', " & _
                     "   SUPPLIER_ID = " & TempID1 & ", " & _
                     "   DELIVERY_ID = " & TempID2 & ", " & _
                     "   DOCUMENT_DESC = '" & ChangeQuote(Trim(DOCUMENT_DESC)) & "', " & _
                     "   RECEIVE_NAME = '" & ChangeQuote(Trim(RECEIVE_NAME)) & "', " & _
                     "   BILL_NO = '" & ChangeQuote(Trim(BILL_NO)) & "', " & _
                     "   DO_NO = '" & ChangeQuote(Trim(DO_NO)) & "', " & _
                     "   TRUCK_NO = '" & ChangeQuote(Trim(EncryptText(TRUCK_NO))) & "', " & _
                     "   DOCUMENT_DATE = '" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "', " & _
                     "   DOCUMENT_TYPE = " & DOCUMENT_TYPE & ", " & _
                     "   SENDER_NAME = '" & ChangeQuote(Trim(SENDER_NAME)) & "', " & _
                     "   DELIVERY_FEE = '" & DELIVERY_FEE & "', " & _
                     "   EMP_ID = " & TempID3 & ", " & _
                      "  SURETY_ID = " & TempID6 & ", " & _
                     "   COMMIT_FLAG = '" & ChangeQuote(Trim(COMMIT_FLAG)) & "', " & _
                     "   SALE_FLAG = '" & ChangeQuote(Trim(SALE_FLAG)) & "', " & _
                     "   EXCEPTION_FLAG = '" & ChangeQuote(Trim(EXCEPTION_FLAG)) & "', " & _
                     "   DOCUMENT_SUBTYPE = " & DOCUMENT_SUBTYPE & ", " & _
                     "   BATCH_ID = " & TempID4 & ", " & _
                     "   SIMULATE_FLAG = '" & ChangeQuote(Trim(SIMULATE_FLAG)) & "', " & _
                     "   YYYYMM = '" & ChangeQuote(Trim(YYYYMM)) & "', " & _
                     "   YYYYMM2 = '" & ChangeQuote(Trim(YYYYMM2)) & "', "
               SQL1 = SQL1 & _
                     "   DUE_DATE = '" & ChangeQuote(Trim(DateToStringInt(DUE_DATE))) & "', " & _
                     "   CUS_ID = " & TempID5 & ", " & _
                     "   PO_NO = '" & ChangeQuote(Trim(PO_NO)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE INVENTORY_DOC_ID = " & INVENTORY_DOC_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function

'Public Function DeleteData()
'Dim SQL1 As String
'
'      SQL1 = "DELETE FROM EXPORT_ITEM WHERE INVENTORY_DOC_ID = " & INVENTORY_DOC_ID
'      m_Conn.Execute (SQL1)
'
'      SQL1 = "DELETE FROM IMPORT_ITEM WHERE INVENTORY_DOC_ID = " & INVENTORY_DOC_ID
'      m_Conn.Execute (SQL1)
'
'      SQL1 = "DELETE FROM INVENTORY_DOC WHERE INVENTORY_DOC_ID = " & INVENTORY_DOC_ID
'      m_Conn.Execute (SQL1)
'End Function

Private Sub Class_Terminate()
   Set ImportItems = Nothing
   Set ExportItems = Nothing
   Set ImportExports = Nothing
   Set TransferItems = Nothing
End Sub
Public Sub SearchSupplierCode(FromStr As String)
   If InStr(1, FromStr, "SP") <= 0 Then
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (SP.SUPPLIER_ID = ADR.SUPPLIER_ID) "
   End If
End Sub
Public Function UndoCommit() As Boolean
Dim ServerDateTime As String
Dim SQL1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL1 = "UPDATE INVENTORY_DOC SET " & _
                  "   COMMIT_FLAG = '" & ChangeQuote(Trim(COMMIT_FLAG)) & "', " & _
                 "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                 "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                  "WHERE INVENTORY_DOC_ID = " & Val(INVENTORY_DOC_ID)
   
   Call m_Conn.Execute(SQL1)
   
End Function




