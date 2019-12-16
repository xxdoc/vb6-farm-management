VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long
Public Flag As String
' End of generic data member
Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub
Public Function IsPatch(VersionStr As String) As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim TempRs As ADODB.Recordset
Dim itemcount As Long
Dim Result As Boolean

   Set TempRs = New ADODB.Recordset
   Result = True
   
   SQL = "SELECT * FROM PATCH_DELIVERY WHERE PATCH_VERSION = '" & VersionStr & "' "
   Call glbDatabaseMngr.GetRs(SQL, "", False, itemcount, TempRs, glbErrorLog)

   If TempRs.EOF Then
      Result = fasle
   End If
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   
   IsPatch = Result
   Exit Function
   
ErrorHandler:
   IsPatch = False
End Function

Public Function Patch_1_0_12_2() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_2"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE IMPORT_ITEM ADD TOTAL_ACTUAL_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE IMPORT_ITEM ADD TOTAL_INCLUDE_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE IMPORT_ITEM SET TOTAL_ACTUAL_PRICE = TOTAL_PRICE "
   m_Conn.Execute (SQL)

   SQL = "UPDATE IMPORT_ITEM SET TOTAL_INCLUDE_PRICE = TOTAL_PRICE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE IMPORT_ITEM DROP TOTAL_PRICE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_2 = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_3() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "1_0_12_3"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_3 = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_9() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Yw As CYearWeek
Dim Pt As CProductType

   VersionStr = "1_0_12_9"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "CREATE TABLE PRTITEM_MAP "
   SQL = SQL & "( "
   SQL = SQL & "PRTITEM_MAP_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "YEAR_WEEK_ID           ID_TYPE NOT NULL, "
   SQL = SQL & "PRODUCT_TYPE_ID          ID_TYPE NOT NULL, "
   SQL = SQL & "PART_ITEM_ID          ID_TYPE NOT NULL, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR PRTITEM_MAP_SEQ "
   m_Conn.Execute (SQL)
      
   SQL = "ALTER TABLE PRODUCT_TYPE ADD CONSTRAINT PRODUCT_TYPE_ID_PK PRIMARY KEY (PRODUCT_TYPE_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PRTITEM_MAP ADD CONSTRAINT PRTITEM_MAP_ID_PK PRIMARY KEY (PRTITEM_MAP_ID) "
   m_Conn.Execute (SQL)
      
   SQL = "ALTER TABLE PRTITEM_MAP ADD CONSTRAINT PRTITEM_MAP_WEEK_ID_FK FOREIGN KEY (YEAR_WEEK_ID) REFERENCES YEAR_WEEK "
   m_Conn.Execute (SQL)
      
   SQL = "ALTER TABLE PRTITEM_MAP ADD CONSTRAINT PRTITEM_MAP_PRODUCT_TYPE_ID_FK FOREIGN KEY (PRODUCT_TYPE_ID) REFERENCES PRODUCT_TYPE "
   m_Conn.Execute (SQL)
      
   SQL = "ALTER TABLE PRTITEM_MAP ADD CONSTRAINT PRTITEM_MAP_PART_ITEM_ID_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
      
   Set Pt = New CProductType
   Call glbDaily.PopulateProductType(Pt)
   Set Pt = Nothing
   
   Set Yw = New CYearWeek
   Yw.YEAR_WEEK_ID = -1
   Call glbDaily.MigrateYearWeek(Yw, IsOK, False, glbErrorLog)
   Set Yw = Nothing
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_9 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_10() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Yw As CYearWeek
Dim Pt As CProductType

   VersionStr = "1_0_12_10"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
            
   m_Conn.BeginTrans
   HasBegin = True
         
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_10 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_11() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Yw As CYearWeek
Dim Pt As CProductType

   VersionStr = "1_0_12_11"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
            
   m_Conn.BeginTrans
   HasBegin = True
         
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_11 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_12() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Yw As CYearWeek
Dim Pt As CProductType

   VersionStr = "1_0_12_12"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
            
   m_Conn.BeginTrans
   HasBegin = True
         
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_12 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_13() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Yw As CYearWeek
Dim Pt As CProductType

   VersionStr = "1_0_12_13"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
            
   m_Conn.BeginTrans
   HasBegin = True
         
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_13 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_14() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_14"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
            
   SQL = "ALTER TABLE CAPITAL_MOVEMENT ADD COMMIT_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
            
   m_Conn.BeginTrans
   HasBegin = True
            
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_14 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_15() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Yw As CImportItem
Dim Ivd As CInventoryDoc
Dim Dp As CDocumentPriority

   Set Dp = New CDocumentPriority
   
   VersionStr = "1_0_12_15"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "ALTER TABLE EXPENSE_TYPE ADD BUY_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
            
   SQL = "CREATE TABLE DOCUMENT_PRIORITY"
   SQL = SQL & "("
   SQL = SQL & "   DOCUMENT_PRIORITY_ID  ID_TYPE NOT NULL,"
   SQL = SQL & "   DOCUMENT_TYPE         ID_TYPE NOT NULL,"
   SQL = SQL & "   AREA                  ID_TYPE NOT NULL,"
   SQL = SQL & "   PRIORITY1             ID_TYPE NOT NULL,"
   SQL = SQL & "   PRIORITY2             ID_TYPE NOT NULL,"
   SQL = SQL & "   CREATE_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   CREATE_BY             ID_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_BY ID_TYPE NOT NULL"
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR DOCUMENT_PRIORITY_SEQ "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DOCUMENT_PRIORITY ADD CONSTRAINT DOCUMENT_PRIORITY_ID_PK PRIMARY KEY (DOCUMENT_PRIORITY_ID) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   '=====
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 1 'ใบนำเข้าวัตถุดิบ
   Dp.Area = 1
   Dp.PRIORITY1 = 1
   Dp.PRIORITY2 = 0
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 3 'ใบโอนวัตถุดิบ
   Dp.Area = 1
   Dp.PRIORITY1 = 2
   Dp.PRIORITY2 = 0
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 2 'ใบเบิกวัตถุดิบ
   Dp.Area = 1
   Dp.PRIORITY1 = 3
   Dp.PRIORITY2 = 0
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 4 'ใบปรับยอดวัตถุดิบ
   Dp.Area = 1
   Dp.PRIORITY1 = 4
   Dp.PRIORITY2 = 0
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 11 'ใบนำเข้าสุกร
   Dp.Area = 1
   Dp.PRIORITY1 = 5
   Dp.PRIORITY2 = 0
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 5 'ใบสุกรคลอด
   Dp.Area = 1
   Dp.PRIORITY1 = 6
   Dp.PRIORITY2 = 0
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 6 'ใบโอนย้ายสุกร
   Dp.Area = 1
   Dp.PRIORITY1 = 7
   Dp.PRIORITY2 = 0
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 7 'ใบโอนสุกรเข้าเรือนขาย
   Dp.Area = 1
   Dp.PRIORITY1 = 8
   Dp.PRIORITY2 = 0
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 8 'ใบเปลี่ยนสถานะสุกร
   Dp.Area = 1
   Dp.PRIORITY1 = 9
   Dp.PRIORITY2 = 0
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 10 'ใบเบิกสุกรจากใบส่งสินค้า
   Dp.Area = 1
   Dp.PRIORITY1 = 10
   Dp.PRIORITY2 = 0
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 9 'ใบปรับยอดสุกร
   Dp.Area = 1
   Dp.PRIORITY1 = 11
   Dp.PRIORITY2 = 0
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 12 'ใบเปลี่ยนสถานะสุกรในเรือนขาย
   Dp.Area = 1
   Dp.PRIORITY1 = 12
   Dp.PRIORITY2 = 0
   Call Dp.AddEditData
   '=====
   
   SQL = "UPDATE EXPENSE_TYPE SET BUY_FLAG = 'N' "
   m_Conn.Execute (SQL)

   Set Ivd = New CInventoryDoc
   Call glbDaily.PatchIncludePrice(Ivd, IsOK, False, glbErrorLog)
   Set Ivd = Nothing

   Set Yw = New CImportItem
   Yw.IMPORT_ITEM_ID = -1
   Call glbDaily.PatchTxSeq(Yw, IsOK, False, glbErrorLog)
   Set Yw = Nothing

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   Set Dp = Nothing
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_15 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_16() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_16"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
            
   SQL = "ALTER TABLE PART_ITEM ADD INTAKE_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
            
   m_Conn.BeginTrans
   HasBegin = True
            
   SQL = "UPDATE PART_ITEM SET INTAKE_FLAG = 'Y' "
   m_Conn.Execute (SQL)
            
   SQL = "UPDATE PART_ITEM SET INTAKE_FLAG = 'N' WHERE PART_ITEM_ID = 7314 "
   m_Conn.Execute (SQL)
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_16 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_17() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_17"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
            
   SQL = "CREATE TABLE BALANCE_ACCUM"
   SQL = SQL & "("
   SQL = SQL & "   BALANCE_ACCUM_ID      ID_TYPE NOT NULL,"
   SQL = SQL & "   PART_ITEM_ID          ID_TYPE NOT NULL,"
   SQL = SQL & "   DOCUMENT_DATE         DATE_TYPE NOT NULL,"
   SQL = SQL & "   IMPORT_AMOUNT         MONEY_TYPE,"
   SQL = SQL & "   EXPORT_AMOUNT         MONEY_TYPE,"
   SQL = SQL & "   CREATE_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   CREATE_BY             ID_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_BY ID_TYPE NOT NULL"
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR BALANCE_ACCUM_SEQ "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE BALANCE_ACCUM ADD CONSTRAINT BALANCE_ACCUM_ID_PK PRIMARY KEY (BALANCE_ACCUM_ID) "
   m_Conn.Execute (SQL)
            
   SQL = "ALTER TABLE BALANCE_ACCUM ADD CONSTRAINT BALANCE_ACCUM_PART_ID_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM "
   m_Conn.Execute (SQL)
            
   '===
   SQL = "CREATE TABLE EXPTYPE_RATIO "
   SQL = SQL & "("
   SQL = SQL & "      EXPTYPE_RATIO_ID      ID_TYPE NOT NULL,"
   SQL = SQL & "      EXPENSE_TYPE_ID       ID_TYPE NOT NULL,"
   SQL = SQL & "      LOCATION_ID           ID_TYPE NOT NULL,"
   SQL = SQL & "      RATIO                 MONEY_TYPE,"
   SQL = SQL & "      RATIO_AMOUNT          MONEY_TYPE,"
   SQL = SQL & "   CREATE_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   CREATE_BY             ID_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_BY ID_TYPE NOT NULL"
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR EXPTYPE_RATIO_SEQ "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EXPTYPE_RATIO ADD CONSTRAINT EXPTYPE_RATIO_ID_PK PRIMARY KEY (EXPTYPE_RATIO_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EXPTYPE_RATIO ADD CONSTRAINT EXPTYPE_RATIO_HOUSE_ID_FK FOREIGN KEY (LOCATION_ID) REFERENCES LOCATION "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE EXPTYPE_RATIO ADD CONSTRAINT EXPTYPE_RATIO_EXPENSE_ID_FK FOREIGN KEY (EXPENSE_TYPE_ID) REFERENCES EXPENSE_TYPE "
   m_Conn.Execute (SQL)
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
               
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_17 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_18() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_18"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
            
   SQL = "ALTER TABLE BALANCE_ACCUM ADD BALANCE_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE BALANCE_ACCUM ADD TOTAL_INCLUDE_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)
            
   SQL = "ALTER TABLE BALANCE_ACCUM ADD LOCATION_ID ID_TYPE "
   m_Conn.Execute (SQL)
               
   SQL = "ALTER TABLE BALANCE_ACCUM ADD CONSTRAINT BALANCE_ACCUM_LOCATION_ID_FK FOREIGN KEY (LOCATION_ID) REFERENCES LOCATION "
   m_Conn.Execute (SQL)
               
   SQL = "ALTER TABLE EXPORT_ITEM ADD EXPORT_TOTAL_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)
               
   m_Conn.BeginTrans
   HasBegin = True
               
   SQL = "DELETE FROM BALANCE_ACCUM "
   m_Conn.Execute (SQL)
               
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_18 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_19() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Dp As CDocumentPriority

   VersionStr = "1_0_12_19"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
                           
   m_Conn.BeginTrans
   HasBegin = True
               
   '=====
   Set Dp = New CDocumentPriority
   Call Dp.ClearData
   
   '=== บัญชี
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 5 'ใบรับของ
   Dp.Area = 2
   Dp.PRIORITY1 = 9
   Dp.PRIORITY2 = 8
   Call Dp.AddEditData
   
   '=== หมู
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 5 'ใบสุกรคลอด
   Dp.Area = 1
   Dp.PRIORITY1 = 10
   Dp.PRIORITY2 = 9
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 11 'ใบนำเข้าสุกร
   Dp.Area = 1
   Dp.PRIORITY1 = 11
   Dp.PRIORITY2 = 14
   Call Dp.AddEditData
      
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 6 'ใบโอนย้ายสุกร
   Dp.Area = 1
   Dp.PRIORITY1 = 12
   Dp.PRIORITY2 = 15
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 7 'ใบโอนสุกรเข้าเรือนขาย
   Dp.Area = 1
   Dp.PRIORITY1 = 13
   Dp.PRIORITY2 = 16
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 8 'ใบเปลี่ยนสถานะสุกร
   Dp.Area = 1
   Dp.PRIORITY1 = 14
   Dp.PRIORITY2 = 17
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 10 'ใบเบิกสุกรจากใบส่งสินค้า
   Dp.Area = 1
   Dp.PRIORITY1 = 15
   Dp.PRIORITY2 = 18
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 9 'ใบปรับยอดสุกร
   Dp.Area = 1
   Dp.PRIORITY1 = 16
   Dp.PRIORITY2 = 19
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 12 'ใบเปลี่ยนสถานะสุกรในเรือนขาย
   Dp.Area = 1
   Dp.PRIORITY1 = 17
   Dp.PRIORITY2 = 20
   Call Dp.AddEditData
   
   '==== คลัง
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 1 'ใบนำเข้าวัตถุดิบ
   Dp.Area = 1
   Dp.PRIORITY1 = 18
   Dp.PRIORITY2 = 10
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 3 'ใบโอนวัตถุดิบ
   Dp.Area = 1
   Dp.PRIORITY1 = 19
   Dp.PRIORITY2 = 11
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 2 'ใบเบิกวัตถุดิบ
   Dp.Area = 1
   Dp.PRIORITY1 = 20
   Dp.PRIORITY2 = 12
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 4 'ใบปรับยอดวัตถุดิบ
   Dp.Area = 1
   Dp.PRIORITY1 = 21
   Dp.PRIORITY2 = 13
   Call Dp.AddEditData
         
   Set Dp = Nothing
               
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_19 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_20() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Dp As CDocumentPriority

   VersionStr = "1_0_12_20"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
                           
   SQL = "ALTER TABLE CAPITAL_MOVEMENT ADD TX_SEQ ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
               
   '=====
   Set Dp = New CDocumentPriority
   Call Dp.ClearData
   
   '=== บัญชี
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 5 'ใบรับของ
   Dp.Area = 2
   Dp.PRIORITY1 = 9
   Dp.PRIORITY2 = 8
   Call Dp.AddEditData
   
   '=== หมู
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 5 'ใบสุกรคลอด
   Dp.Area = 1
   Dp.PRIORITY1 = 10
   Dp.PRIORITY2 = 9
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 11 'ใบนำเข้าสุกร
   Dp.Area = 1
   Dp.PRIORITY1 = 11
   Dp.PRIORITY2 = 14
   Call Dp.AddEditData
      
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 6 'ใบโอนย้ายสุกร
   Dp.Area = 1
   Dp.PRIORITY1 = 12
   Dp.PRIORITY2 = 15
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 7 'ใบโอนสุกรเข้าเรือนขาย
   Dp.Area = 1
   Dp.PRIORITY1 = 13
   Dp.PRIORITY2 = 16
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 8 'ใบเปลี่ยนสถานะสุกร
   Dp.Area = 1
   Dp.PRIORITY1 = 14
   Dp.PRIORITY2 = 17
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 10 'ใบเบิกสุกรจากใบส่งสินค้า
   Dp.Area = 1
   Dp.PRIORITY1 = 15
   Dp.PRIORITY2 = 18
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 9 'ใบปรับยอดสุกร
   Dp.Area = 1
   Dp.PRIORITY1 = 16
   Dp.PRIORITY2 = 19
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 12 'ใบเปลี่ยนสถานะสุกรในเรือนขาย
   Dp.Area = 1
   Dp.PRIORITY1 = 17
   Dp.PRIORITY2 = 20
   Call Dp.AddEditData
   
   '==== คลัง
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 1 'ใบนำเข้าวัตถุดิบ
   Dp.Area = 1
   Dp.PRIORITY1 = 18
   Dp.PRIORITY2 = 10
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 3 'ใบโอนวัตถุดิบ
   Dp.Area = 1
   Dp.PRIORITY1 = 19
   Dp.PRIORITY2 = 11
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 2 'ใบเบิกวัตถุดิบ
   Dp.Area = 1
   Dp.PRIORITY1 = 20
   Dp.PRIORITY2 = 12
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 4 'ใบปรับยอดวัตถุดิบ
   Dp.Area = 1
   Dp.PRIORITY1 = 21
   Dp.PRIORITY2 = 13
   Call Dp.AddEditData
         
   Set Dp = Nothing
               
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_20 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_21() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Dp As CDocumentPriority

   VersionStr = "1_0_12_21"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
                              
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_21 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_22() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Ivd As CInventoryDoc

   VersionStr = "1_0_12_22"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
                              
   SQL = "ALTER TABLE PRODUCT_TYPE ADD CAPITAL_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
                              
                              
   m_Conn.BeginTrans
   HasBegin = True
   
'   Set Ivd = New CInventoryDoc
'   Call glbDaily.PatchIncludePrice(Ivd, IsOK, False, glbErrorLog)
'   Set Ivd = Nothing
   
   SQL = "UPDATE PRODUCT_TYPE SET CAPITAL_FLAG = 'N' WHERE PRODUCT_TYPE_NO IN ('N', 'R', 'TB') "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE PRODUCT_TYPE SET CAPITAL_FLAG = 'Y' WHERE PRODUCT_TYPE_NO IN ('G', 'L', 'B') "
   m_Conn.Execute (SQL)
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_22 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_23() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Ivd As CInventoryDoc

   VersionStr = "1_0_12_23"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
                              
   SQL = "ALTER TABLE CAPITAL_MOVEMENT ADD TO_PIG_COUNT MONEY_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE INVENTORY_DOC ADD EXCEPTION_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "UPDATE INVENTORY_DOC SET EXCEPTION_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
'   Set Ivd = New CInventoryDoc
'   Ivd.COMMIT_FLAG = ""
'   Ivd.DELETE_BALANCE_FLAG = "Y"
'   Ivd.FROM_DATE = -1
'   Ivd.TO_DATE = -1
'   Call glbDaily.ClearCapitalMovement(Ivd, IsOK, False, glbErrorLog)
'   Set Ivd = Nothing

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done., Don't forget to import capital movement again."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_23 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_24() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Ba As CBalanceAccum

   VersionStr = "1_0_12_24"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
                                 
   m_Conn.BeginTrans
   HasBegin = True
   
   Set Ba = New CBalanceAccum
   Call Ba.ClearData
   Set Ba = Nothing
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done., Don't forget to calculate avg price agin (1 April 2005 - 31 May 2005) "
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_24 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_25() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Ba As CBalanceAccum

   VersionStr = "1_0_12_25"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE IMPORT_ITEM ADD REPLACE_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
                                 
   SQL = "ALTER TABLE EXPORT_ITEM ADD REPLACE_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "UPDATE IMPORT_ITEM SET REPLACE_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE EXPORT_ITEM SET REPLACE_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done"
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_25 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_26() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Ba As CBalanceAccum

   VersionStr = "1_0_12_26"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE CAPITAL_MOVEMENT ADD REPLACE_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "UPDATE CAPITAL_MOVEMENT SET REPLACE_FLAG = 'N' "
   m_Conn.Execute (SQL)
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done"
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_26 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_27() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Ba As CBalanceAccum

   VersionStr = "1_0_12_27"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE CAPITAL_MOVEMENT ADD BILLING_DOC_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CAPITAL_MOVEMENT ADD CONSTRAINT CAPITAL_MOVEMENT_DOC_ID_FK1 FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done"
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_27 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_28() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Ba As CBalanceAccum

   VersionStr = "1_0_12_28"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done"
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_28 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_29() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Ba As CBalanceAccum

   VersionStr = "1_0_12_29"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "ALTER TABLE EXPENSE_TYPE ADD DEPLICATE_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "UPDATE EXPENSE_TYPE SET DEPLICATE_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE EXPENSE_TYPE SET DEPLICATE_FLAG = 'Y' WHERE EXPENSE_TYPE_NO = '06' "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE DOCUMENT_PRIORITY SET PRIORITY1 = 9 WHERE DOCUMENT_PRIORITY_ID = 13"
   m_Conn.Execute (SQL)
   SQL = "UPDATE DOCUMENT_PRIORITY SET PRIORITY1 = 10 WHERE DOCUMENT_PRIORITY_ID = 14"
   m_Conn.Execute (SQL)
   SQL = "UPDATE DOCUMENT_PRIORITY SET PRIORITY1 = 11 WHERE DOCUMENT_PRIORITY_ID = 15"
   m_Conn.Execute (SQL)
   SQL = "UPDATE DOCUMENT_PRIORITY SET PRIORITY1 = 15 WHERE DOCUMENT_PRIORITY_ID = 16"
   m_Conn.Execute (SQL)
   SQL = "UPDATE DOCUMENT_PRIORITY SET PRIORITY1 = 16 WHERE DOCUMENT_PRIORITY_ID = 17"
   m_Conn.Execute (SQL)
   SQL = "UPDATE DOCUMENT_PRIORITY SET PRIORITY1 = 18 WHERE DOCUMENT_PRIORITY_ID = 18"
   m_Conn.Execute (SQL)
   SQL = "UPDATE DOCUMENT_PRIORITY SET PRIORITY1 = 17 WHERE DOCUMENT_PRIORITY_ID = 19"
   m_Conn.Execute (SQL)
   SQL = "UPDATE DOCUMENT_PRIORITY SET PRIORITY1 = 19 WHERE DOCUMENT_PRIORITY_ID = 20"
   m_Conn.Execute (SQL)
   SQL = "UPDATE DOCUMENT_PRIORITY SET PRIORITY1 = 20 WHERE DOCUMENT_PRIORITY_ID = 21"
   m_Conn.Execute (SQL)
   SQL = "UPDATE DOCUMENT_PRIORITY SET PRIORITY1 = 12 WHERE DOCUMENT_PRIORITY_ID = 22"
   m_Conn.Execute (SQL)
   SQL = "UPDATE DOCUMENT_PRIORITY SET PRIORITY1 = 13 WHERE DOCUMENT_PRIORITY_ID = 23"
   m_Conn.Execute (SQL)
   SQL = "UPDATE DOCUMENT_PRIORITY SET PRIORITY1 = 14 WHERE DOCUMENT_PRIORITY_ID = 24"
   m_Conn.Execute (SQL)
   SQL = "UPDATE DOCUMENT_PRIORITY SET PRIORITY1 = 21 WHERE DOCUMENT_PRIORITY_ID = 25"
   m_Conn.Execute (SQL)

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done, Don't forget to update weighted average cost from 1 April 2005 - 30 June 2005 and create capital movement again."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_29 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_30() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_30"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "ALTER TABLE BALANCE_ACCUM ADD AVG_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_30 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_31() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
'Dim Ba As CBalanceAccum

   VersionStr = "1_0_12_31"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
            
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "DELETE FROM BALANCE_ACCUM "
   m_Conn.Execute (SQL)

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done, Don't forget to update weighted average cost from 1 April 2005 - 30 June 2005 and create capital movement again."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_31 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_32() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
'Dim Ba As CBalanceAccum

   VersionStr = "1_0_12_32"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
            
   SQL = "ALTER TABLE BILLING_DOC ADD DOCUMENT_SUBTYPE ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_DOC ADD DOCUMENT_SUBTYPE ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EXPORT_ITEM ALTER COLUMN EXPORT_AMOUNT TYPE MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "UPDATE BILLING_DOC SET DOCUMENT_SUBTYPE = 1 WHERE DOCUMENT_TYPE = 1 "
   m_Conn.Execute (SQL)

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done"
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_32 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_33() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
'Dim Ba As CBalanceAccum

   VersionStr = "1_0_12_33"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
               
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done"
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_33 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_34() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
'Dim Ba As CBalanceAccum

   VersionStr = "1_0_12_34"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
               
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done"
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_34 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_35() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
'Dim Ba As CBalanceAccum

   VersionStr = "1_0_12_35"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
               
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done"
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_35 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_36() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
'Dim Ba As CBalanceAccum

   VersionStr = "1_0_12_36"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
               
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done"
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_36 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_37() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
'Dim Ba As CBalanceAccum

   VersionStr = "1_0_12_37"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
               
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done"
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_37 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_38() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
'Dim Ba As CBalanceAccum

   VersionStr = "1_0_12_38"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
               
   SQL = "ALTER TABLE LOCATION ADD MASTER_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "UPDATE LOCATION SET MASTER_FLAG = 'N'  "
   m_Conn.Execute (SQL)

   SQL = "UPDATE LOCATION SET MASTER_FLAG = 'Y' WHERE (LOCATION_NO = '00') AND (LOCATION_TYPE = 2) "
   m_Conn.Execute (SQL)

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done"
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_38 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_39() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_39"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
            
   SQL = "CREATE TABLE BANK"
   SQL = SQL & "("
   SQL = SQL & "   BANK_ID      ID_TYPE NOT NULL,"
   SQL = SQL & "   BANK_NO          CODE_TYPE NOT NULL,"
   SQL = SQL & "   BANK_NAME         NAME_TYPE NOT NULL,"
   SQL = SQL & "   CREATE_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   CREATE_BY             ID_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_BY ID_TYPE NOT NULL"
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE BANK ADD CONSTRAINT BANK_ID_PK PRIMARY KEY (BANK_ID) "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE BALANCE_ACCUM ADD CONSTRAINT BALANCE_ACCUM_ID_PK PRIMARY KEY (BALANCE_ACCUM_ID) "
   m_Conn.Execute (SQL)
            
   SQL = "CREATE GENERATOR BANK_SEQ "
   m_Conn.Execute (SQL)
            
   
   SQL = "CREATE TABLE BANK_BRANCH"
   SQL = SQL & "("
   SQL = SQL & "   BBRANCH_ID      ID_TYPE NOT NULL,"
   SQL = SQL & "   BBRANCH_NO          CODE_TYPE NOT NULL,"
   SQL = SQL & "   BBRANCH_NAME         NAME_TYPE NOT NULL,"
   SQL = SQL & "   BANK_ID      ID_TYPE NOT NULL,"
   SQL = SQL & "   CREATE_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   CREATE_BY             ID_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_BY ID_TYPE NOT NULL"
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BANK_BRANCH ADD CONSTRAINT BANK_BRANCH_ID_PK PRIMARY KEY (BBRANCH_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BANK_BRANCH ADD CONSTRAINT BANK_BRANCH_BANK_ID_FK FOREIGN KEY (BANK_ID) REFERENCES BANK "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR BBRANCH_SEQ "
   m_Conn.Execute (SQL)
      
   
   SQL = "ALTER TABLE BILLING_DOC ADD BANK_BRANCH_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD CONSTRAINT BD_BRANCH_BANK_ID_FK FOREIGN KEY (BANK_BRANCH_ID) REFERENCES BANK_BRANCH "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD BANK_NOTE DESC_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD PAYMENT_TYPE ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD CHECK_NO CODE_TYPE "
   m_Conn.Execute (SQL)
      
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
               
   SQL = "UPDATE BALANCE_ACCUM SET TOTAL_INCLUDE_PRICE = 4357.1428572 WHERE (LOCATION_ID = 254) AND (PART_ITEM_ID = 7426) AND (DOCUMENT_DATE = '2005-05-30 00:00:00') "
   m_Conn.Execute (SQL)
               
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_39 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_40() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_40"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
            
   SQL = "ALTER TABLE BILLING_DOC ADD PAYMENT_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD CONSTRAINT BILLING_DOC_PMT_ID_FK FOREIGN KEY (PAYMENT_ID) REFERENCES PAYMENT "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE TABLE PAYMENT"
   SQL = SQL & "("
   SQL = SQL & "      PAYMENT_ID         ID_TYPE NOT NULL,"
   SQL = SQL & "      PAYMENT_NO         CODE_TYPE NOT NULL,"
   SQL = SQL & "      PAYMENT_DATE       NAME_TYPE NOT NULL,"
   SQL = SQL & "      TO_BANK_ID         ID_TYPE,"
   SQL = SQL & "      TO_BANK_BRANCH     ID_TYPE,"
   SQL = SQL & "      ACCOUNT_NO         CODE_TYPE,"
   SQL = SQL & "      COMMIT_FLAG        FLAG_TYPE,"
   SQL = SQL & "      TX_TYPE            CODE_TYPE,"
   SQL = SQL & "   CREATE_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   CREATE_BY             ID_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_BY ID_TYPE NOT NULL"
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE PAYMENT ADD CONSTRAINT PAYMENT_ID_PK PRIMARY KEY (PAYMENT_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PAYMENT ADD CONSTRAINT PAYMENT_BANK_ID_FK FOREIGN KEY (TO_BANK_ID) REFERENCES BANK "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PAYMENT ADD CONSTRAINT PAYMENT_BBRANCH_ID_FK FOREIGN KEY (TO_BANK_BRANCH) REFERENCES BANK_BRANCH "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR PAYMENT_SEQ "
   m_Conn.Execute (SQL)
               
   SQL = "CREATE TABLE PAYMENT_ITEM"
   SQL = SQL & "("
   SQL = SQL & "   PAYMENT_ITEM_ID    ID_TYPE NOT NULL,"
   SQL = SQL & "   PAYMENT_ID         ID_TYPE NOT NULL,"
   SQL = SQL & "   PAYMENT_TYPE       ID_TYPE,"
   SQL = SQL & "   CHECK_NO           CODE_TYPE,"
   SQL = SQL & "   BANK_ID            ID_TYPE,"
   SQL = SQL & "   BANK_BRANCH        ID_TYPE,"
   SQL = SQL & "   BANK_NOTE          DESC_TYPE,"
   SQL = SQL & "   CREATE_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   CREATE_BY             ID_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_BY ID_TYPE NOT NULL"
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PAYMENT_ITEM ADD CONSTRAINT PAYMENT_ITEM_ID_PK PRIMARY KEY (PAYMENT_ITEM_ID) "
   m_Conn.Execute (SQL)
   SQL = "ALTER TABLE PAYMENT_ITEM ADD CONSTRAINT PAYMENT_ITEM_PMT_ID_FK FOREIGN KEY (PAYMENT_ID) REFERENCES PAYMENT "
   m_Conn.Execute (SQL)
   SQL = "ALTER TABLE PAYMENT_ITEM ADD CONSTRAINT PAYMENT_ITEM_BANK_ID_FK FOREIGN KEY (BANK_ID) REFERENCES BANK "
   m_Conn.Execute (SQL)
   SQL = "ALTER TABLE PAYMENT_ITEM ADD CONSTRAINT PAYMENT_ITEM_BBANK_ID_FK FOREIGN KEY (BANK_BRANCH) REFERENCES BANK_BRANCH "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR PAYMENT_ITEM_SEQ "
   m_Conn.Execute (SQL)
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                              
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_40 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_41() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_41"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
            
   SQL = "ALTER TABLE PAYMENT ADD INTERNAL_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PAYMENT ADD ACCEPT_BY ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PAYMENT ADD CONSTRAINT PAYMENT_ACCEPT_BY_FK FOREIGN KEY (ACCEPT_BY) REFERENCES EMPLOYEE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PAYMENT ADD TOTAL_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PAYMENT_ITEM ADD PAY_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER ADD CREDIT_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE TABLE REGION"
   SQL = SQL & "("
   SQL = SQL & "      REGION_ID         ID_TYPE NOT NULL,"
   SQL = SQL & "      REGION_NO         CODE_TYPE NOT NULL,"
   SQL = SQL & "      REGION_NAME       NAME_TYPE NOT NULL,"
   SQL = SQL & "   CREATE_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   CREATE_BY             ID_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_BY ID_TYPE NOT NULL"
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE REGION ADD CONSTRAINT REGION_ID_PK PRIMARY KEY (REGION_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR REGION_SEQ "
   m_Conn.Execute (SQL)
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                              
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_41 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_42() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim EI As CExportItem

   VersionStr = "1_0_12_42"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
            
   SQL = "ALTER TABLE CUSTOMER ADD REGION_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER ADD CONSTRAINT CUSTOMER_REGION_ID_FK FOREIGN KEY (REGION_ID) REFERENCES REGION "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD REGION_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD CONSTRAINT BILLING_DOC_REGION_ID_FK FOREIGN KEY (REGION_ID) REFERENCES REGION "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE RECEIPT_ITEM ADD PAID_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD PAID_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EXPORT_ITEM ADD PIG_AGE MONEY_TYPE "
   m_Conn.Execute (SQL)
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                              
   Set EI = New CExportItem
   Call glbDaily.PatchPigAge(EI, IsOK, False, glbErrorLog)
   Set EI = Nothing
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_42 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_43() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim EI As CExportItem

   VersionStr = "1_0_12_43"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
            
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                              
   Set EI = New CExportItem
   Call glbDaily.PatchPigAge(EI, IsOK, False, glbErrorLog)
   Set EI = Nothing
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_43 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_44() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean

   VersionStr = "1_0_12_44"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
            
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                              
   SQL = "UPDATE BILLING_DOC SET DOCUMENT_SUBTYPE = 1 WHERE DOCUMENT_TYPE = 2 "
   m_Conn.Execute (SQL)
   
   Set II = New CImportItem
   Call glbDaily.PatchPigBirthPartID(II, IsOK, False, glbErrorLog)
   Set II = Nothing
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_44 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_45() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean

   VersionStr = "1_0_12_45"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
            
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                              
   Set II = New CImportItem
   Call glbDaily.PatchPigBirthPartID(II, IsOK, False, glbErrorLog)
   Set II = Nothing
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done., Don't forget to update average price again."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_45 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_46() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean

   VersionStr = "1_0_12_46"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
            
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                              
   Set II = New CImportItem
   Call glbDaily.PatchPigBirthPartID(II, IsOK, False, glbErrorLog)
   Set II = Nothing
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_46 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_47() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean

   VersionStr = "1_0_12_47"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
            
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                              
   SQL = "UPDATE BILLING_DOC SET COMMIT_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE INVENTORY_DOC SET COMMIT_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_47 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_48() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean

   VersionStr = "1_0_12_48"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
            
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                                 
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_48 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_49() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean
Dim Dp As CDocumentPriority

   VersionStr = "1_0_12_49"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE LOCATION ADD CAPITAL_BIRTH_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                                 
   '====
   Set Dp = New CDocumentPriority
   Call Dp.ClearData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 5
   Dp.Area = 2
   Dp.PRIORITY1 = 9
   Dp.PRIORITY2 = 8
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 5
   Dp.Area = 1
   Dp.PRIORITY1 = 10
   Dp.PRIORITY2 = 9
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 11
   Dp.Area = 1
   Dp.PRIORITY1 = 11
   Dp.PRIORITY2 = 14
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 6
   Dp.Area = 1
   Dp.PRIORITY1 = 15
   Dp.PRIORITY2 = 15
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 7
   Dp.Area = 1
   Dp.PRIORITY1 = 16
   Dp.PRIORITY2 = 16
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 8
   Dp.Area = 1
   Dp.PRIORITY1 = 18
   Dp.PRIORITY2 = 17
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 10
   Dp.Area = 1
   Dp.PRIORITY1 = 17
   Dp.PRIORITY2 = 18
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 9
   Dp.Area = 1
   Dp.PRIORITY1 = 19
   Dp.PRIORITY2 = 19
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 12
   Dp.Area = 1
   Dp.PRIORITY1 = 20
   Dp.PRIORITY2 = 20
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 1
   Dp.Area = 1
   Dp.PRIORITY1 = 12
   Dp.PRIORITY2 = 10
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 3
   Dp.Area = 1
   Dp.PRIORITY1 = 13
   Dp.PRIORITY2 = 11
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 2
   Dp.Area = 1
   Dp.PRIORITY1 = 14
   Dp.PRIORITY2 = 12
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 4
   Dp.Area = 1
   Dp.PRIORITY1 = 21
   Dp.PRIORITY2 = 13
   Call Dp.AddEditData
   
   Set Dp = Nothing
   '====
   
   SQL = "UPDATE LOCATION SET CAPITAL_BIRTH_FLAG = 'Y' "
   m_Conn.Execute (SQL)

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_49 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_50() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean
Dim Dp As CDocumentPriority

   VersionStr = "1_0_12_50"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                                 
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_50 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_51() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean

   VersionStr = "1_0_12_51"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE CAPITAL_MOVEMENT ADD TO_PIG_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CAPITAL_MOVEMENT ADD CONSTRAINT CAPITAL_MOVEMENT_TO_PIG_FK FOREIGN KEY (TO_PIG_ID) REFERENCES PART_ITEM "
   m_Conn.Execute (SQL)
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                                 
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_51 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_52() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean

   VersionStr = "1_0_12_52"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE DO_ITEM ADD CONSTRAINT DO_ITEM_PART_ITEM_ID_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD CONSTRAINT DO_ITEM_LOCATION_ID_FK FOREIGN KEY (LOCATION_ID) REFERENCES LOCATION "
   m_Conn.Execute (SQL)
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                                 
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_52 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_53() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean

   VersionStr = "1_0_12_53"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                                 
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_53 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_54() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean

   VersionStr = "1_0_12_54"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "CREATE TABLE REVENUE_TYPE"
   SQL = SQL & "("
   SQL = SQL & "      REVENUE_TYPE_ID         ID_TYPE NOT NULL,"
   SQL = SQL & "      REVENUE_NO         CODE_TYPE NOT NULL,"
   SQL = SQL & "      REVENUE_NAME       NAME_TYPE NOT NULL,"
   SQL = SQL & "   CREATE_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   CREATE_BY             ID_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_BY ID_TYPE NOT NULL"
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE REVENUE_TYPE ADD CONSTRAINT REVENUE_TYPE_ID_PK PRIMARY KEY (REVENUE_TYPE_ID) "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR REVENUE_TYPE_SEQ "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE DO_ITEM ADD REVENUE_ID ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE DO_ITEM ADD CONSTRAINT DO_ITEM_REVENUE_ID_FK FOREIGN KEY (REVENUE_ID) REFERENCES REVENUE_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE DO_ITEM ADD PART_ITEM_ID_EX ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE DO_ITEM ADD LOCATION_ID_EX ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "UPDATE DO_ITEM SET PART_ITEM_ID_EX = PART_ITEM_ID "
   m_Conn.Execute (SQL)

   SQL = "UPDATE DO_ITEM SET LOCATION_ID_EX = LOCATION_ID "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM DROP PART_ITEM_ID "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM DROP LOCATION_ID "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD PART_ITEM_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD LOCATION_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD CONSTRAINT DO_ITEM_PART_ITEM_ID_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD CONSTRAINT DO_ITEM_LOCATION_ITEM_ID_FK FOREIGN KEY (LOCATION_ID) REFERENCES LOCATION "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE DO_ITEM SET PART_ITEM_ID = PART_ITEM_ID_EX "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE DO_ITEM SET LOCATION_ID = LOCATION_ID_EX "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
                                 
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_54 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_55() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean

   VersionStr = "1_0_12_55"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "CREATE TABLE CAPITAL_LOSS"
   SQL = SQL & "("
   SQL = SQL & "         CAPITAL_LOSS_ID        ID_TYPE NOT NULL,"
   SQL = SQL & "         DOCUMENT_NO            CODE_TYPE,"
   SQL = SQL & "         DOCUMENT_DATE          DATE_TYPE,"
   SQL = SQL & "         IVD_ID                 ID_TYPE,"
   SQL = SQL & "         BL_ID                  ID_TYPE,"
   SQL = SQL & "         DOCUMENT_CATEGORY      ID_TYPE,"
   SQL = SQL & "         DOCUMENT_TYPE          ID_TYPE,"
   SQL = SQL & "         TX_AMOUNT              MONEY_TYPE,"
   SQL = SQL & "         TX_TYPE                CODE_TYPE,"
   SQL = SQL & "      FROM_HOUSE_ID          ID_TYPE,"
   SQL = SQL & "         TO_HOUSE_ID            ID_TYPE,"
   SQL = SQL & "         PIG_STATUS             ID_TYPE,"
   SQL = SQL & "         IMPORT_ITEM_ID         ID_TYPE,"
   SQL = SQL & "         EXPORT_ITEM_ID         ID_TYPE,"
   SQL = SQL & "         PIG_ID                 ID_TYPE,"
   SQL = SQL & "         COMMIT_FLAG            FLAG_TYPE,"
   SQL = SQL & "         TX_SEQ                 ID_TYPE,"
   SQL = SQL & "         TO_PIG_COUNT           MONEY_TYPE,"
   SQL = SQL & "      REPLACE_FLAG           FLAG_TYPE,"
   SQL = SQL & "         BILLING_DOC_ID         ID_TYPE,"
   SQL = SQL & "         TO_PIG_ID              ID_TYPE,"
   SQL = SQL & "         CREATE_DATE            DATE_TYPE NOT NULL,"
   SQL = SQL & "         CREATE_BY              ID_TYPE NOT NULL,"
   SQL = SQL & "         MODIFY_DATE            DATE_TYPE NOT NULL,"
   SQL = SQL & "         MODIFY_BY  ID_TYPE NOT NULL"
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CAPITAL_LOSS ADD CONSTRAINT CAPITAL_LOSS_ID_PK PRIMARY KEY (CAPITAL_LOSS_ID) "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE CAPITAL_LOSS ADD CONSTRAINT CAPITAL_LOSS_BLD_FK1 FOREIGN KEY (BL_ID) REFERENCES BILLING_DOC (BILLING_DOC_ID) "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE CAPITAL_LOSS ADD CONSTRAINT CAPITAL_LOSS_DOC_ID_FK1 FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC (BILLING_DOC_ID) "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE CAPITAL_LOSS ADD CONSTRAINT CAPITAL_LOSS_EXP_FK1 FOREIGN KEY (EXPORT_ITEM_ID) REFERENCES EXPORT_ITEM (EXPORT_ITEM_ID) "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE CAPITAL_LOSS ADD CONSTRAINT CAPITAL_LOSS_HOUSE_FK1 FOREIGN KEY (FROM_HOUSE_ID) REFERENCES LOCATION (LOCATION_ID) "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE CAPITAL_LOSS ADD CONSTRAINT CAPITAL_LOSS_HOUSE_FK2 FOREIGN KEY (TO_HOUSE_ID) REFERENCES LOCATION (LOCATION_ID) "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE CAPITAL_LOSS ADD CONSTRAINT CAPITAL_LOSS_IMP_FK1 FOREIGN KEY (IMPORT_ITEM_ID) REFERENCES IMPORT_ITEM (IMPORT_ITEM_ID) "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE CAPITAL_LOSS ADD CONSTRAINT CAPITAL_LOSS_IVD_FK1 FOREIGN KEY (IVD_ID) REFERENCES INVENTORY_DOC (INVENTORY_DOC_ID) "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE CAPITAL_LOSS ADD CONSTRAINT CAPITAL_LOSS_PIG_FK1 FOREIGN KEY (PIG_ID) REFERENCES PART_ITEM (PART_ITEM_ID) "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE CAPITAL_LOSS ADD CONSTRAINT CAPITAL_LOSS_STATUS_FK1 FOREIGN KEY (PIG_STATUS) REFERENCES PRODUCT_STATUS (PRODUCT_STATUS_ID) "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE CAPITAL_LOSS ADD CONSTRAINT CAPITAL_LOSS_TO_PIG_FK FOREIGN KEY (TO_PIG_ID) REFERENCES PART_ITEM (PART_ITEM_ID) "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR CAPITAL_LOSS_SEQ "
   m_Conn.Execute (SQL)

   SQL = "CREATE TABLE LOSS_ITEM"
   SQL = SQL & "("
   SQL = SQL & "            LOSS_ITEM_ID          ID_TYPE NOT NULL,"
   SQL = SQL & "            CAPITAL_LOSS_ID       ID_TYPE NOT NULL,"
   SQL = SQL & "            EXPENSE_TYPE          ID_TYPE,"
   SQL = SQL & "            PART_ITEM_ID          ID_TYPE,"
   SQL = SQL & "            CAPITAL_AMOUNT        MONEY_TYPE,"
   SQL = SQL & "            CREATE_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "            CREATE_BY             ID_TYPE NOT NULL,"
   SQL = SQL & "            MODIFY_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "            MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE LOSS_ITEM ADD CONSTRAINT LOSS_ITEM_ID_PK PRIMARY KEY (LOSS_ITEM_ID) "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE LOSS_ITEM ADD CONSTRAINT LOSS_ITEM_CPM_ID_FK FOREIGN KEY (CAPITAL_LOSS_ID) REFERENCES CAPITAL_LOSS (CAPITAL_LOSS_ID) "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE LOSS_ITEM ADD CONSTRAINT LOSS_ITEM_EXPENSE_ID_FK FOREIGN KEY (EXPENSE_TYPE) REFERENCES EXPENSE_TYPE (EXPENSE_TYPE_ID) "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE LOSS_ITEM ADD CONSTRAINT LOSS_ITEM_PRTITM_ID_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM (PART_ITEM_ID) "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR LOSS_ITEM_SEQ "
   m_Conn.Execute (SQL)

   SQL = "UPDATE DOCUMENT_PRIORITY SET PRIORITY1 = PRIORITY1 + 1 WHERE PRIORITY1 > 16 "
   m_Conn.Execute (SQL)

   SQL = "UPDATE DOCUMENT_PRIORITY SET PRIORITY1 = 17 WHERE DOCUMENT_TYPE = 12 "
   m_Conn.Execute (SQL)

   SQL = "UPDATE DOCUMENT_PRIORITY SET PRIORITY1 = 300 WHERE DOCUMENT_TYPE = 6 "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE PART_ITEM ADD PARENT_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PART_ITEM ADD CONSTRAINT PART_ITEM_PARENT_ID_FK1 FOREIGN KEY (PARENT_ID) REFERENCES PART_ITEM "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PART_ITEM ADD SPECIFIC_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE PART_ITEM SET SPECIFIC_FLAG = 'N' "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True
                                 
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_55 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_56() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean

   VersionStr = "1_0_12_56"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                                 
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_56 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_57() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean

   VersionStr = "1_0_12_57"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                                 
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_57 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_58() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean

   VersionStr = "1_0_12_58"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                                 
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_58 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_59() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean

   VersionStr = "1_0_12_59"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                                 
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_59 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_60() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean

   VersionStr = "1_0_12_60"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                                 
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_60 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_61() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean

   VersionStr = "1_0_12_61"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
      
   SQL = "CREATE GENERATOR DN_NUMBER_SEQ "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR CN_NUMBER_SEQ "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE RECEIPT_ITEM ADD DEBIT_CREDIT_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE TABLE CNDN_REASON "
   SQL = SQL & "("
   SQL = SQL & "         REASON_ID        ID_TYPE NOT NULL,"
   SQL = SQL & "         REASON_NO            CODE_TYPE,"
   SQL = SQL & "         REASON_NAME          NAME_TYPE,"
   SQL = SQL & "         CREATE_DATE            DATE_TYPE NOT NULL,"
   SQL = SQL & "         CREATE_BY              ID_TYPE NOT NULL,"
   SQL = SQL & "         MODIFY_DATE            DATE_TYPE NOT NULL,"
   SQL = SQL & "         MODIFY_BY  ID_TYPE NOT NULL"
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
      
   SQL = "ALTER TABLE CNDN_REASON ADD CONSTRAINT REASON_ID_PK PRIMARY KEY (REASON_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR CNDN_REASON_SEQ "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE RECEIPT_ITEM ADD STATUS_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE RECEIPT_ITEM ADD CONSTRAINT RECEIPT_ITEM_STATUS_ID_FK FOREIGN KEY (STATUS_ID) REFERENCES PRODUCT_STATUS "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE RECEIPT_ITEM ADD REASON_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE RECEIPT_ITEM ADD CONSTRAINT RECEIPT_ITEM_REASON_ID_FK FOREIGN KEY (REASON_ID) REFERENCES CNDN_REASON "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD DEBIT_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD CREDIT_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
                                 
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_61 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_62() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean

   VersionStr = "1_0_12_62"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                                 
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_62 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_63() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean

   VersionStr = "1_0_12_63"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                                 
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_63 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_64() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean
Dim Pm As CPayment
Dim Pmi As CPaymentItem

   VersionStr = "1_0_12_64"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===

   SQL = "ALTER TABLE BILLING_DOC ADD CONSTRAINT BILLING_DOC_PAYMENT_ID_FK FOREIGN KEY (PAYMENT_ID) REFERENCES PAYMENT "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
                                 
   SQL = "UPDATE PAYMENT SET TOTAL_AMOUNT = 0 WHERE PAYMENT_DATE < '2005-09-01 00:00:00' "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE PAYMENT_ITEM PI SET PAY_AMOUNT = 0 WHERE (PI.PAYMENT_ID IN (SELECT PAYMENT_ID FROM PAYMENT WHERE PAYMENT_DATE < '2005-09-01 00:00:00')) "
   m_Conn.Execute (SQL)
                                 
   Set Pm = New CPayment
   Pm.AddEditMode = SHOW_ADD
   Pm.PAYMENT_NO = "ยกมา"
   Pm.PAYMENT_DATE = InternalDateToDate("2005-08-31 00:00:00")
   Pm.COMMIT_FLAG = "N"
   Pm.INTERNAL_FLAG = "Y"
   Pm.TX_TYPE = "I"
   Pm.TOTAL_AMOUNT = 250292
   
   Set Pmi = New CPaymentItem
   Pmi.Flag = "A"
   Pmi.PAYMENT_TYPE = CASH_PMT
   Pmi.PAY_AMOUNT = 250292
   Call Pm.PaymentItems.Add(Pmi)
   Set Pmi = Nothing
   
   Call glbDaily.AddEditPayment(Pm, IsOK, False, glbErrorLog)
   Set Pm = Nothing
                                                                     
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_64 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_65() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Dp As CDocumentPriority

   VersionStr = "1_0_12_65"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
      
   m_Conn.BeginTrans
   HasBegin = True
                                                                                                      
   SQL = "DELETE FROM DOCUMENT_PRIORITY "
   m_Conn.Execute (SQL)
                                                                                                      
   '=====
   Set Dp = New CDocumentPriority
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 11
   Dp.Area = 1
   Dp.PRIORITY1 = 9
   Dp.PRIORITY2 = 9
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 1
   Dp.Area = 1
   Dp.PRIORITY1 = 10
   Dp.PRIORITY2 = 10
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 3
   Dp.Area = 1
   Dp.PRIORITY1 = 11
   Dp.PRIORITY2 = 11
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 2
   Dp.Area = 1
   Dp.PRIORITY1 = 12
   Dp.PRIORITY2 = 12
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 5
   Dp.Area = 2
   Dp.PRIORITY1 = 1
   Dp.PRIORITY2 = 1
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 5
   Dp.Area = 1
   Dp.PRIORITY1 = 13
   Dp.PRIORITY2 = 13
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 6
   Dp.Area = 1
   Dp.PRIORITY1 = 14
   Dp.PRIORITY2 = 14
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 17
   Dp.Area = 1
   Dp.PRIORITY1 = 15
   Dp.PRIORITY2 = 15
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 13
   Dp.Area = 1
   Dp.PRIORITY1 = 16
   Dp.PRIORITY2 = 16
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 10
   Dp.Area = 1
   Dp.PRIORITY1 = 16
   Dp.PRIORITY2 = 16
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 8
   Dp.Area = 1
   Dp.PRIORITY1 = 17
   Dp.PRIORITY2 = 17
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 5
   Dp.Area = 1
   Dp.PRIORITY1 = 18
   Dp.PRIORITY2 = 18
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 12
   Dp.Area = 1
   Dp.PRIORITY1 = 19
   Dp.PRIORITY2 = 19
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 9
   Dp.Area = 1
   Dp.PRIORITY1 = 20
   Dp.PRIORITY2 = 20
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 4
   Dp.Area = 1
   Dp.PRIORITY1 = 21
   Dp.PRIORITY2 = 21
   Call Dp.AddEditData
   
   Set Dp = Nothing
   '=====
                                                                                                      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_65 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_66() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Dp As CDocumentPriority

   VersionStr = "1_0_12_66"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
      
   m_Conn.BeginTrans
   HasBegin = True
                                                                                                      
   SQL = "DELETE FROM DOCUMENT_PRIORITY "
   m_Conn.Execute (SQL)
                                                                                                      
   '=====
   Set Dp = New CDocumentPriority
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 11
   Dp.Area = 1
   Dp.PRIORITY1 = 9
   Dp.PRIORITY2 = 9
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 1
   Dp.Area = 1
   Dp.PRIORITY1 = 10
   Dp.PRIORITY2 = 10
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 3
   Dp.Area = 1
   Dp.PRIORITY1 = 11
   Dp.PRIORITY2 = 11
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 2
   Dp.Area = 1
   Dp.PRIORITY1 = 12
   Dp.PRIORITY2 = 12
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 5
   Dp.Area = 2
   Dp.PRIORITY1 = 1
   Dp.PRIORITY2 = 1
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 5
   Dp.Area = 1
   Dp.PRIORITY1 = 13
   Dp.PRIORITY2 = 13
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 6
   Dp.Area = 1
   Dp.PRIORITY1 = 14
   Dp.PRIORITY2 = 14
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 7
   Dp.Area = 1
   Dp.PRIORITY1 = 15
   Dp.PRIORITY2 = 15
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 13
   Dp.Area = 1
   Dp.PRIORITY1 = 16
   Dp.PRIORITY2 = 16
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 10
   Dp.Area = 1
   Dp.PRIORITY1 = 16
   Dp.PRIORITY2 = 16
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 8
   Dp.Area = 1
   Dp.PRIORITY1 = 17
   Dp.PRIORITY2 = 17
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 12
   Dp.Area = 1
   Dp.PRIORITY1 = 19
   Dp.PRIORITY2 = 19
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 9
   Dp.Area = 1
   Dp.PRIORITY1 = 20
   Dp.PRIORITY2 = 20
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 4
   Dp.Area = 1
   Dp.PRIORITY1 = 21
   Dp.PRIORITY2 = 21
   Call Dp.AddEditData
   
   Set Dp = Nothing
   '=====
                                                                                                      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_66 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_67() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Dp As CDocumentPriority

   VersionStr = "1_0_12_67"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
      
   m_Conn.BeginTrans
   HasBegin = True
                                                                                                      
   SQL = "DELETE FROM DOCUMENT_PRIORITY "
   m_Conn.Execute (SQL)
                                                                                                      
   '=====
   Set Dp = New CDocumentPriority
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 11
   Dp.Area = 1
   Dp.PRIORITY1 = 9
   Dp.PRIORITY2 = 9
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 1
   Dp.Area = 1
   Dp.PRIORITY1 = 10
   Dp.PRIORITY2 = 10
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 3
   Dp.Area = 1
   Dp.PRIORITY1 = 11
   Dp.PRIORITY2 = 11
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 2
   Dp.Area = 1
   Dp.PRIORITY1 = 12
   Dp.PRIORITY2 = 12
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 5
   Dp.Area = 2
   Dp.PRIORITY1 = 1
   Dp.PRIORITY2 = 1
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 5
   Dp.Area = 1
   Dp.PRIORITY1 = 13
   Dp.PRIORITY2 = 13
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 6
   Dp.Area = 1
   Dp.PRIORITY1 = 14
   Dp.PRIORITY2 = 14
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 7
   Dp.Area = 1
   Dp.PRIORITY1 = 15
   Dp.PRIORITY2 = 15
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 13
   Dp.Area = 1
   Dp.PRIORITY1 = 16
   Dp.PRIORITY2 = 16
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 10
   Dp.Area = 1
   Dp.PRIORITY1 = 16
   Dp.PRIORITY2 = 16
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 8
   Dp.Area = 1
   Dp.PRIORITY1 = 17
   Dp.PRIORITY2 = 17
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 12
   Dp.Area = 1
   Dp.PRIORITY1 = 19
   Dp.PRIORITY2 = 19
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 9
   Dp.Area = 1
   Dp.PRIORITY1 = 20
   Dp.PRIORITY2 = 20
   Call Dp.AddEditData
   
   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 4
   Dp.Area = 1
   Dp.PRIORITY1 = 21
   Dp.PRIORITY2 = 21
   Call Dp.AddEditData
   
   Set Dp = Nothing
   '=====
                                                                                                      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_67 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_68() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean

   VersionStr = "1_0_12_68"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                                 
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_68 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_69() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Dp As CDocumentPriority

   VersionStr = "1_0_12_69"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
      
   m_Conn.BeginTrans
   HasBegin = True
                                                                                                                                                                                                            
   '=====
   Set Dp = New CDocumentPriority

   Dp.AddEditMode = SHOW_ADD
   Dp.DOCUMENT_TYPE = 14 'ใบโอนย้ายต้นทุนเป็นขาดทุน
   Dp.Area = 1
   Dp.PRIORITY1 = 300
   Dp.PRIORITY2 = 300
   Call Dp.AddEditData

   Set Dp = Nothing
   '=====
                                                                                                      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_69 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_70() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Dp As CDocumentPriority

   VersionStr = "1_0_12_70"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
      
   m_Conn.BeginTrans
   HasBegin = True
                                                  
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_70 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_71() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Dp As CDocumentPriority

   VersionStr = "1_0_12_71"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
      
   m_Conn.BeginTrans
   HasBegin = True
                                                  
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_71 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_72() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Dp As CDocumentPriority

   VersionStr = "1_0_12_72"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
      
   m_Conn.BeginTrans
   HasBegin = True
                                                  
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_72 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_73() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim Dp As CDocumentPriority

   VersionStr = "1_0_12_73"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
      
   m_Conn.BeginTrans
   HasBegin = True
                                                  
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError

   Patch_1_0_12_73 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_74() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean
Dim Pm As CPayment
Dim Pmi As CPaymentItem
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_74"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===

   UpdateFlag = False
   If glbEnterPrise.BRANCH_CODE = "MA2" Then
      PaidAmount = 903465
      UpdateFlag = True
   ElseIf glbEnterPrise.BRANCH_CODE = "MH" Then
      PaidAmount = 484175
      UpdateFlag = True
   ElseIf glbEnterPrise.BRANCH_CODE = "DTS" Then
      PaidAmount = 114916
      UpdateFlag = True
   End If
   
   If UpdateFlag Then
      SQL = "ALTER TABLE BILLING_DOC ADD CONSTRAINT BILLING_DOC_PAYMENT_ID_FK FOREIGN KEY (PAYMENT_ID) REFERENCES PAYMENT "
      m_Conn.Execute (SQL)
         
      m_Conn.BeginTrans
      HasBegin = True
                                    
      SQL = "UPDATE PAYMENT SET TOTAL_AMOUNT = 0 WHERE PAYMENT_DATE < '2005-10-01 00:00:00' "
      m_Conn.Execute (SQL)
      
      SQL = "UPDATE PAYMENT_ITEM PI SET PAY_AMOUNT = 0 WHERE (PI.PAYMENT_ID IN (SELECT PAYMENT_ID FROM PAYMENT WHERE PAYMENT_DATE < '2005-10-01 00:00:00')) "
      m_Conn.Execute (SQL)
                                    
      Set Pm = New CPayment
      Pm.AddEditMode = SHOW_ADD
      Pm.PAYMENT_NO = "ยกมา"
      Pm.PAYMENT_DATE = InternalDateToDate("2005-09-31 00:00:00")
      Pm.COMMIT_FLAG = "N"
      Pm.INTERNAL_FLAG = "Y"
      Pm.TX_TYPE = "I"
      Pm.TOTAL_AMOUNT = PaidAmoun
      
      Set Pmi = New CPaymentItem
      Pmi.Flag = "A"
      Pmi.PAYMENT_TYPE = CASH_PMT
      Pmi.PAY_AMOUNT = PaidAmount
      Call Pm.PaymentItems.Add(Pmi)
      Set Pmi = Nothing
      
      Call glbDaily.AddEditPayment(Pm, IsOK, False, glbErrorLog)
      Set Pm = Nothing
   End If
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_74 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_75() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean
Dim Pm As CPayment
Dim Pmi As CPaymentItem
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_75"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   glbErrorLog.LocalErrorMsg = "ใช่ฟาร์ม MP หรือไม่"
   If glbErrorLog.AskMessage = vbYes Then
      SQL = "UPDATE PART_ITEM SET PART_NO = '254825', PARENT_ID = NULL, SPECIFIC_FLAG = 'N' WHERE PART_ITEM_ID = 8879 "
      m_Conn.Execute (SQL)
   End If
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_75 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_76() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean
Dim Pm As CPayment
Dim Pmi As CPaymentItem
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_76"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
'   glbErrorLog.LocalErrorMsg = "ใช่ฟาร์ม MP หรือไม่"
'   If glbErrorLog.AskMessage = vbYes Then
'      SQL = "UPDATE PART_ITEM SET PART_NO = '254825', PARENT_ID = NULL, SPECIFIC_FLAG = 'N' WHERE PART_ITEM_ID = 8879 "
'      m_Conn.Execute (SQL)
'   End If
   
   UpdateFlag = False
   If glbEnterPrise.BRANCH_CODE = "MA2" Then
      PaidAmount = 903465
      UpdateFlag = True
   ElseIf glbEnterPrise.BRANCH_CODE = "MH" Then
      PaidAmount = 484175
      UpdateFlag = True
   ElseIf glbEnterPrise.BRANCH_CODE = "DTS" Then
      PaidAmount = 114916
      UpdateFlag = True
   End If
   
   If UpdateFlag Then
      SQL = "UPDATE PAYMENT SET TOTAL_AMOUNT = 0 WHERE PAYMENT_DATE < '2005-10-01 00:00:00' "
      m_Conn.Execute (SQL)
      
      SQL = "UPDATE PAYMENT_ITEM PI SET PAY_AMOUNT = 0 WHERE (PI.PAYMENT_ID IN (SELECT PAYMENT_ID FROM PAYMENT WHERE PAYMENT_DATE < '2005-10-01 00:00:00')) "
      m_Conn.Execute (SQL)
                                    
      Set Pm = New CPayment
      Pm.AddEditMode = SHOW_ADD
      Pm.PAYMENT_NO = "ยกมา"
      Pm.PAYMENT_DATE = InternalDateToDate("2005-09-30 00:00:00")
      Pm.COMMIT_FLAG = "N"
      Pm.INTERNAL_FLAG = "Y"
      Pm.TX_TYPE = "I"
      Pm.TOTAL_AMOUNT = PaidAmount
      
      Set Pmi = New CPaymentItem
      Pmi.Flag = "A"
      Pmi.PAYMENT_TYPE = CASH_PMT
      Pmi.PAY_AMOUNT = PaidAmount
      Call Pm.PaymentItems.Add(Pmi)
      Set Pmi = Nothing
      
      Call glbDaily.AddEditPayment(Pm, IsOK, False, glbErrorLog)
      Set Pm = Nothing
   End If
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_76 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_77() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean
Dim Pm As CPayment
Dim Pmi As CPaymentItem
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_77"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   glbErrorLog.LocalErrorMsg = "ใช่ฟาร์ม MP หรือไม่"
   If glbErrorLog.AskMessage = vbYes Then
      SQL = "UPDATE PART_ITEM SET PART_NO = '254826', PARENT_ID = NULL, SPECIFIC_FLAG = 'N' WHERE PART_ITEM_ID = 8885 "
      m_Conn.Execute (SQL)
   
      SQL = "UPDATE PART_ITEM SET PART_NO = '254827', PART_DESC = 'สุกรทั่วไป 254827', PARENT_ID = NULL, SPECIFIC_FLAG = 'N' WHERE PART_ITEM_ID = 8891 "
      m_Conn.Execute (SQL)
   End If
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_77 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_78() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean
Dim Pm As CPayment
Dim Pmi As CPaymentItem
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_78"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   glbErrorLog.LocalErrorMsg = "ใช่ฟาร์ม DTS หรือไม่"
   If glbErrorLog.AskMessage = vbYes Then
      SQL = "UPDATE PART_ITEM SET PART_NO = '254820', PART_DESC = 'สุกรทั่วไป 254820', PARENT_ID = NULL, SPECIFIC_FLAG = 'N' WHERE PART_ITEM_ID = 11631 "
      m_Conn.Execute (SQL)
   
      SQL = "UPDATE PART_ITEM SET PART_NO = '254821', PART_DESC = 'สุกรทั่วไป 254821', PARENT_ID = NULL, SPECIFIC_FLAG = 'N' WHERE PART_ITEM_ID = 11637 "
      m_Conn.Execute (SQL)
   
      SQL = "UPDATE PART_ITEM SET PART_NO = '254823', PART_DESC = 'สุกรทั่วไป 254823', PARENT_ID = NULL, SPECIFIC_FLAG = 'N' WHERE PART_ITEM_ID = 11649 "
      m_Conn.Execute (SQL)
   End If
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_78 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_79() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean
Dim Pm As CPayment
Dim Pmi As CPaymentItem
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_79"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===

   glbErrorLog.LocalErrorMsg = "ใช่ฟาร์ม MP หรือไม่"
   If glbErrorLog.AskMessage = vbYes Then
      SQL = "UPDATE PART_ITEM SET PART_NO = '254825', PART_DESC = 'สุกรทั่วไป 254825', PARENT_ID = NULL, SPECIFIC_FLAG = 'N' WHERE PART_ITEM_ID = 8879 "
      m_Conn.Execute (SQL)
   End If

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_79 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_80() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean
Dim Pm As CPayment
Dim Pmi As CPaymentItem
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_80"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_80 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_81() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean
Dim Pm As CPayment
Dim Pmi As CPaymentItem
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_81"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "CREATE TABLE BANK_ACCOUNT"
   SQL = SQL & "("
   SQL = SQL & "   BANK_ACCOUNT_ID      ID_TYPE NOT NULL,"
   SQL = SQL & "   ACCOUNT_NO          CODE_TYPE NOT NULL,"
   SQL = SQL & "   ACCOUNT_NAME         NAME_TYPE NOT NULL,"
   SQL = SQL & "   CREATE_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   CREATE_BY             ID_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_BY ID_TYPE NOT NULL"
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BANK_ACCOUNT ADD CONSTRAINT BANK_ACCOUNT_ID_PK PRIMARY KEY (BANK_ACCOUNT_ID)"
   m_Conn.Execute (SQL)
      
   SQL = "CREATE GENERATOR BANK_ACCOUNT_SEQ"
   m_Conn.Execute (SQL)
      
   SQL = "ALTER TABLE PAYMENT ADD BANK_ACCOUNT_ID ID_TYPE"
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PAYMENT ADD CONSTRAINT PAYMENT_BANK_ACCOUNT_ID_FK FOREIGN KEY (BANK_ACCOUNT_ID) REFERENCES BANK_ACCOUNT"
   m_Conn.Execute (SQL)
            
   Call m_Conn.BeginTrans
   HasBegin = True
      
   If glbEnterPrise.BRANCH_CODE = "MA2" Then
      SQL = "DELETE FROM PAYMENT WHERE PAYMENT_ID = 142"
      m_Conn.Execute (SQL)
   
      SQL = "DELETE FROM PAYMENT_ITEM WHERE PAYMENT_ID = 142"
      m_Conn.Execute (SQL)
   End If
   
   'Call glbDaily.CreateBankAccount(False)
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_81 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_82() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean
Dim Pm As CPayment
Dim Pmi As CPaymentItem
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_82"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
               
   Call m_Conn.BeginTrans
   HasBegin = True
         
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_82 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_83() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean
Dim Pm As CPayment
Dim Pmi As CPaymentItem
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_83"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
               
   Call m_Conn.BeginTrans
   HasBegin = True
         
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_83 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_84() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean
Dim Pm As CPayment
Dim Pmi As CPaymentItem
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_84"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "CREATE TABLE RECEIPT_CNDN"
   SQL = SQL & "("
   SQL = SQL & "   RECEIPT_CNDN_ID      ID_TYPE NOT NULL,"
   SQL = SQL & "   BILLING_DOC_ID      ID_TYPE NOT NULL,"
   SQL = SQL & "   DO_ID      ID_TYPE,"
   SQL = SQL & "   CNDN_ID      ID_TYPE,"
   SQL = SQL & "   CN_AMOUNT      MONEY_TYPE,"
   SQL = SQL & "   DN_AMOUNT      MONEY_TYPE,"
   SQL = SQL & "   CREATE_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   CREATE_BY             ID_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_DATE           DATE_TYPE NOT NULL,"
   SQL = SQL & "   MODIFY_BY ID_TYPE NOT NULL"
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
         
   SQL = "CREATE GENERATOR RECEIPT_CNDN_SEQ"
   m_Conn.Execute (SQL)
      
   SQL = "ALTER TABLE RECEIPT_CNDN ADD CONSTRAINT RECEIPT_CNDN_ID_PK PRIMARY KEY (RECEIPT_CNDN_ID) "
   m_Conn.Execute (SQL)
      
   SQL = "ALTER TABLE RECEIPT_CNDN ADD CONSTRAINT RECEIPT_CNDN_BLD_ID_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE RECEIPT_CNDN ADD CONSTRAINT RECEIPT_CNDN_DO_ID_FK FOREIGN KEY (DO_ID) REFERENCES BILLING_DOC "
   m_Conn.Execute (SQL)
            
   SQL = "ALTER TABLE RECEIPT_CNDN ADD CONSTRAINT RECEIPT_CNDN_CNDN_ID_FK FOREIGN KEY (CNDN_ID) REFERENCES BILLING_DOC "
   m_Conn.Execute (SQL)
            
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_84 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_85() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean
Dim Pm As CPayment
Dim Pmi As CPaymentItem
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_85"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
            
   SQL = "ALTER TABLE DO_ITEM ADD AGE_CODE CODE_TYPE "
   m_Conn.Execute (SQL)
      
   SQL = "ALTER TABLE DO_ITEM ADD PIG_AGE MONEY_TYPE "
   m_Conn.Execute (SQL)
      
   SQL = "ALTER TABLE EXPORT_ITEM ADD AGE_CODE CODE_TYPE "
   m_Conn.Execute (SQL)

   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_85 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_86() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean
Dim Pm As CPayment
Dim Pmi As CPaymentItem
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_86"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
            
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_86 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_87() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean
Dim Pm As CPayment
Dim Pmi As CPaymentItem
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_87"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
            
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_87 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_88() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean
Dim Pm As CPayment
Dim Pmi As CPaymentItem
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_88"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
            
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_88 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_89() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim II As CImportItem
Dim IsOK As Boolean
Dim Pm As CPayment
Dim Pmi As CPaymentItem
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_89"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
            
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_89 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_91() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_91"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
            
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_91 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_92() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_92"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "ALTER DOMAIN FLAG_TYPE DROP CONSTRAINT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER DOMAIN FLAG_TYPE ADD CONSTRAINT CHECK(VALUE IN ('Y', 'N', '')) "
   m_Conn.Execute (SQL)
   

   SQL = "ALTER TABLE BILLING_DOC ADD SIMULATE_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
            
   SQL = "ALTER TABLE BILLING_DOC ADD BATCH_ID ID_TYPE "
   m_Conn.Execute (SQL)
            
   SQL = "ALTER TABLE INVENTORY_DOC ADD SIMULATE_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
            
   SQL = "ALTER TABLE INVENTORY_DOC ADD BATCH_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "UPDATE BILLING_DOC SET SIMULATE_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE INVENTORY_DOC SET SIMULATE_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_92 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function



Public Function Patch_1_0_12_93() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_93"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===

   
   SQL = "CREATE TABLE PARAMETERS "
   SQL = SQL & "( "
   SQL = SQL & "PARAM_ID              ID_TYPE NOT NULL, "
   SQL = SQL & "PARAM_NO              CODE_TYPE, "
   SQL = SQL & "PARAM_DATE            DATE_TYPE, "
   SQL = SQL & "PARAM_AREA            ID_TYPE, "
   SQL = SQL & "FROM_SALE_DATE        DATE_TYPE, "
   SQL = SQL & "TO_SALE_DATE          DATE_TYPE, "
   SQL = SQL & "FROM_AGE              MONEY_TYPE, "
   SQL = SQL & "TO_AGE                MONEY_TYPE, "
      SQL = SQL & "PARAM_DESC            DESC_TYPE, "
      SQL = SQL & "COMMIT_FLAG           FLAG_TYPE, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PARAMETERS ADD CONSTRAINT PARAMETER_ID_PK PRIMARY KEY (PARAM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR PARAMETER_SEQ "
   m_Conn.Execute (SQL)
      
   SQL = "CREATE TABLE USEDPRM_ITEM "
   SQL = SQL & "( "
   SQL = SQL & "USEDPRM_ITEM_ID       ID_TYPE NOT NULL, "
   SQL = SQL & "PARAM_ID              ID_TYPE NOT NULL, "
   SQL = SQL & "FOOD_ID               ID_TYPE NOT NULL, "
   SQL = SQL & "USED_RATE             MONEY_TYPE, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE USEDPRM_ITEM ADD CONSTRAINT USEDPRM_ITEM_ID_PK PRIMARY KEY (USEDPRM_ITEM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE USEDPRM_ITEM ADD CONSTRAINT USEDPRM_ITEM_PARAM_ID_FK FOREIGN KEY (PARAM_ID) REFERENCES PARAMETERS "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE USEDPRM_ITEM ADD CONSTRAINT USEDPRM_ITEM_FOOD_ID_FK FOREIGN KEY (FOOD_ID) REFERENCES PART_ITEM "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR USEDPRM_ITEM_SEQ "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE TABLE SALEPRM_ITEM "
   SQL = SQL & "( "
   SQL = SQL & "SALEPRM_ITEM_ID       ID_TYPE NOT NULL, "
   SQL = SQL & "PARAM_ID              ID_TYPE NOT NULL, "
   SQL = SQL & "PIG_STATUS            ID_TYPE, "
   SQL = SQL & "SALE_RATE             MONEY_TYPE, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALEPRM_ITEM ADD CONSTRAINT SALEPRM_ITEM_ID_PK PRIMARY KEY (SALEPRM_ITEM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALEPRM_ITEM ADD CONSTRAINT SALEPRM_ITEM_PARAM_ID_FK FOREIGN KEY (PARAM_ID) REFERENCES PARAMETERS "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALEPRM_ITEM ADD CONSTRAINT SALEPRM_ITEM_PIG_STATUS_FK FOREIGN KEY (PIG_STATUS) REFERENCES PRODUCT_STATUS "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR SALEPRM_ITEM_SEQ "
   
   SQL = "CREATE TABLE WEIGHTPRM_ITEM "
   SQL = SQL & "( "
   SQL = SQL & "WEIGHTPRM_ITEM_ID     ID_TYPE NOT NULL, "
   SQL = SQL & "PARAM_ID              ID_TYPE NOT NULL, "
   SQL = SQL & "FROM_AGE              MONEY_TYPE, "
   SQL = SQL & "TO_AGE                MONEY_TYPE, "
   SQL = SQL & "UNIT_WEIGHT           MONEY_TYPE, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE WEIGHTPRM_ITEM ADD CONSTRAINT WEIGHTPRM_ITEM_ID_PK PRIMARY KEY (WEIGHTPRM_ITEM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE WEIGHTPRM_ITEM ADD CONSTRAINT WEIGHTPRM_ITEM_PARAM_ID_FK FOREIGN KEY (PARAM_ID) REFERENCES PARAMETERS "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR WEIGHTPRM_ITEM_SEQ "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE TABLE TRNPRM_ITEM "
   SQL = SQL & "( "
   SQL = SQL & "TRNPRM_ITEM_ID        ID_TYPE NOT NULL, "
   SQL = SQL & "PARAM_ID              ID_TYPE NOT NULL, "
   SQL = SQL & "PIG_STATUS            ID_TYPE, "
   SQL = SQL & "LOSS_RATE             MONEY_TYPE, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE TRNPRM_ITEM ADD CONSTRAINT TRNPRM_ITEM_ID_PK PRIMARY KEY (TRNPRM_ITEM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE TRNPRM_ITEM ADD CONSTRAINT TRNPRM_ITEM_PARAM_ID_FK FOREIGN KEY (PARAM_ID) REFERENCES PARAMETERS "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE TRNPRM_ITEM ADD CONSTRAINT TRNPRM_ITEM_PIG_STATUS_FK FOREIGN KEY (PIG_STATUS) REFERENCES PRODUCT_STATUS "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR TRNPRM_ITEM_SEQ "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE TABLE BATCH "
   SQL = SQL & "( "
   SQL = SQL & "BATCH_ID              ID_TYPE NOT NULL, "
   SQL = SQL & "BATCH_NO              CODE_TYPE, "
   SQL = SQL & "BATCH_DATE            DATE_TYPE, "
   SQL = SQL & "EXECUTE_FROM          DATE_TYPE, "
   SQL = SQL & "EXECUTE_TO            DATE_TYPE, "
   SQL = SQL & "BATCH_DESC            DESC_TYPE, "
   SQL = SQL & "COMMIT_FLAG           FLAG_TYPE, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BATCH ADD CONSTRAINT BATCH_ID_PK PRIMARY KEY (BATCH_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR BATCH_SEQ "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE TABLE BATCH_ITEM "
   SQL = SQL & "( "
   SQL = SQL & "BATCH_ITEM_ID         ID_TYPE NOT NULL, "
   SQL = SQL & "BATCH_ID              ID_TYPE NOT NULL, "
   SQL = SQL & "PARAM_ID              ID_TYPE NOT NULL, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BATCH_ITEM ADD CONSTRAINT BATCH_ITEM_ID_PK PRIMARY KEY (BATCH_ITEM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BATCH_ITEM ADD CONSTRAINT BATCH_ITEM_BATCH_ID_FK FOREIGN KEY (BATCH_ID) REFERENCES BATCH "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BATCH_ITEM ADD CONSTRAINT BATCH_ITEM_PARAM_ID_FK FOREIGN KEY (PARAM_ID) REFERENCES PARAMETERS "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR BATCH_ITEM_SEQ "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_93 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_94() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_94"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "ALTER DOMAIN FLAG_TYPE DROP CONSTRAINT "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER DOMAIN FLAG_TYPE ADD CONSTRAINT CHECK(VALUE IN ('Y', 'N', '')) "
   m_Conn.Execute (SQL)
      
   SQL = "CREATE TABLE COSTPRM_ITEM "
   SQL = SQL & "( "
   SQL = SQL & "COSTPRM_ITEM_ID       ID_TYPE NOT NULL, "
   SQL = SQL & "PARAM_ID              ID_TYPE NOT NULL, "
   SQL = SQL & "FOOD_ID               ID_TYPE NOT NULL, "
   SQL = SQL & "COST_RATE             MONEY_TYPE, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE COSTPRM_ITEM ADD CONSTRAINT COSTPRM_ITEM_ID_PK PRIMARY KEY (COSTPRM_ITEM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE COSTPRM_ITEM ADD CONSTRAINT COSTPRM_ITEM_PARAM_ID_FK FOREIGN KEY (PARAM_ID) REFERENCES PARAMETERS "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE COSTPRM_ITEM ADD CONSTRAINT COSTPRM_ITEM_FOOD_ID_FK FOREIGN KEY (FOOD_ID) REFERENCES PART_ITEM "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR COSTPRM_ITEM_SEQ "
   m_Conn.Execute (SQL)
         
   SQL = "CREATE TABLE ADJPRM_ITEM "
   SQL = SQL & "( "
   SQL = SQL & "ADJPRM_ITEM_ID        ID_TYPE NOT NULL, "
   SQL = SQL & "PARAM_ID              ID_TYPE NOT NULL, "
   SQL = SQL & "PIG_ID                ID_TYPE NOT NULL, "
   SQL = SQL & "PIG_AMOUNT            MONEY_TYPE, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
         
   SQL = "ALTER TABLE ADJPRM_ITEM ADD CONSTRAINT ADJPRM_ITEM_ID_PK PRIMARY KEY (ADJPRM_ITEM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE ADJPRM_ITEM ADD CONSTRAINT ADJPRM_ITEM_PARAM_ID_FK FOREIGN KEY (PARAM_ID) REFERENCES PARAMETERS "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE ADJPRM_ITEM ADD CONSTRAINT ADJPRM_ITEM_FOOD_ID_FK FOREIGN KEY (PIG_ID) REFERENCES PART_ITEM "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR ADJPRM_ITEM_SEQ "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE TABLE BRTPRM_ITEM "
   SQL = SQL & "( "
   SQL = SQL & "BRTPRM_ITEM_ID        ID_TYPE NOT NULL, "
   SQL = SQL & "PARAM_ID              ID_TYPE NOT NULL, "
   SQL = SQL & "FROM_BREED            DATE_TYPE, "
   SQL = SQL & "TO_BREED              DATE_TYPE, "
   SQL = SQL & "FROM_BIRTH            DATE_TYPE, "
   SQL = SQL & "TO_BIRTH              DATE_TYPE, "
   SQL = SQL & "BREED_RATE            MONEY_TYPE, "
   SQL = SQL & "PREGNANT_RATE         MONEY_TYPE, "
   SQL = SQL & "CHILD_RATE            MONEY_TYPE, "
   SQL = SQL & "BIRTH_RATE            MONEY_TYPE, "
   SQL = SQL & "BREED_AMOUNT          MONEY_TYPE, "
   SQL = SQL & "BIRTH_AMOUNT          MONEY_TYPE, "
   SQL = SQL & "DAY_COUNT             MONEY_TYPE, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
         
   SQL = "ALTER TABLE BRTPRM_ITEM  ADD CONSTRAINT BRTPRM_ITEM_ID_PK PRIMARY KEY (BRTPRM_ITEM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BRTPRM_ITEM  ADD CONSTRAINT BRTPRM_ITEM_PARAM_ID_FK FOREIGN KEY (PARAM_ID) REFERENCES PARAMETERS "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR BRTPRM_ITEM_SEQ "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE TABLE RVNPRM_ITEM "
   SQL = SQL & "( "
   SQL = SQL & "RVNPRM_ITEM_ID        ID_TYPE NOT NULL, "
   SQL = SQL & "PARAM_ID              ID_TYPE NOT NULL, "
   SQL = SQL & "FROM_SALE             DATE_TYPE, "
   SQL = SQL & "TO_SALE               DATE_TYPE, "
   SQL = SQL & "UNIT_PRICE            MONEY_TYPE, "
   SQL = SQL & "TOTAL_PRICE           MONEY_TYPE, "
   SQL = SQL & "SALE_AMOUNT           MONEY_TYPE, "
   SQL = SQL & "REVENUE_ID            ID_TYPE, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE RVNPRM_ITEM  ADD CONSTRAINT RVNPRM_ITEM_ID_PK PRIMARY KEY (RVNPRM_ITEM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE RVNPRM_ITEM  ADD CONSTRAINT RVNPRM_ITEM_PARAM_ID_FK FOREIGN KEY (PARAM_ID) REFERENCES PARAMETERS "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE RVNPRM_ITEM  ADD CONSTRAINT RVNPRM_ITEM_REVENUE_ID_FK FOREIGN KEY (REVENUE_ID) REFERENCES REVENUE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR RVNPRM_ITEM_SEQ "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_94 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_95() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_95"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "ALTER TABLE BILLING_DOC ADD SIMULATE_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD BATCH_ID ID_TYPE "
   m_Conn.Execute (SQL)
      
   SQL = "ALTER TABLE INVENTORY_DOC ADD SIMULATE_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
      
   SQL = "ALTER TABLE INVENTORY_DOC ADD BATCH_ID ID_TYPE "
   m_Conn.Execute (SQL)
      
   SQL = "ALTER TABLE BALANCE_ACCUM ADD SIMULATE_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
      
   SQL = "ALTER TABLE CAPITAL_MOVEMENT ADD SIMULATE_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   
   SQL = "ALTER TABLE USEDPRM_ITEM ADD ADG_RATE MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BRTPRM_ITEM ADD AVG_WEIGHT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE ADJPRM_ITEM ADD AVG_WEIGHT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_DOC ADD YYYYMM CODE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD YYYYMM CODE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CAPITAL_MOVEMENT ADD YYYYMM CODE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PAYMENT ADD YYYYMM CODE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALEPRM_ITEM ADD CREDIT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CAPITAL_LOSS ADD SIMULATE_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CAPITAL_LOSS ADD YYYYMM CODE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD YYYYMM2 CODE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_DOC ADD DUE_DATE DATE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_DOC ADD YYYYMM2 CODE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_DOC ADD YYYYMM2 CODE_TYPE "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
      
'   SQL = "UPDATE BILLING_DOC SET SIMULATE_FLAG = 'N' "
'   m_Conn.Execute (SQL)
'
'   SQL = "UPDATE INVENTORY_DOC SET SIMULATE_FLAG = 'N' "
'   m_Conn.Execute (SQL)
'
'   SQL = "UPDATE BALANCE_ACCUM SET SIMULATE_FLAG = 'N' "
'   m_Conn.Execute (SQL)
'
'   SQL = "UPDATE CAPITAL_MOVEMENT SET SIMULATE_FLAG = 'N' "
'   m_Conn.Execute (SQL)
   
'   SQL = "UPDATE INVENTORY_DOC SET YYYYMM = SUBSTR(DOCUMENT_DATE, 1, 7) "
'   m_Conn.Execute (SQL)
'
'   SQL = "UPDATE BILLING_DOC SET YYYYMM = SUBSTR(DOCUMENT_DATE, 1, 7) "
'   m_Conn.Execute (SQL)
'
'   SQL = "UPDATE CAPITAL_MOVEMENT SET YYYYMM = SUBSTR(DOCUMENT_DATE, 1, 7) "
'   m_Conn.Execute (SQL)
'
'   SQL = "UPDATE PAYMENT SET YYYYMM = SUBSTR(PAYMENT_DATE, 1, 7) "
'   m_Conn.Execute (SQL)
   
'   SQL = "UPDATE CAPITAL_LOSS SET SIMULATE_FLAG = 'N' "
'   m_Conn.Execute (SQL)
   
'   SQL = "UPDATE CAPITAL_LOSS SET YYYYMM = SUBSTR(DOCUMENT_DATE, 1, 7) "
'   m_Conn.Execute (SQL)
'
'   SQL = "UPDATE BILLING_DOC SET YYYYMM2 = SUBSTR(DUE_DATE, 1, 7) "
'   m_Conn.Execute (SQL)
'
'   SQL = "UPDATE INVENTORY_DOC SET DUE_DATE = DOCUMENT_DATE "
'   m_Conn.Execute (SQL)
'
'   SQL = "UPDATE INVENTORY_DOC SET YYYYMM2 = SUBSTR(DUE_DATE, 1, 7) "
'   m_Conn.Execute (SQL)
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_95 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_96() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_96"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "CREATE GENERATOR SALEPRM_ITEM_SEQ "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_96 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_97() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_97"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "ALTER TABLE PARAMETERS ADD PIG_STATUS ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PARAMETERS ADD CONSTRAINT PARAMETER_PIG_STATUS_ID_FK FOREIGN KEY (PIG_STATUS) REFERENCES PRODUCT_STATUS "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PARAMETERS ADD PIG_TYPE ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PARAMETERS ADD CONSTRAINT PARAMETER_PIG_TYPE_ID_FK FOREIGN KEY (PIG_TYPE) REFERENCES PRODUCT_TYPE "
   m_Conn.Execute (SQL)
      
   SQL = "CREATE TABLE PARAM_ITEM "
   SQL = SQL & "( "
   SQL = SQL & "PARAM_ITEM_ID        ID_TYPE NOT NULL, "
   SQL = SQL & "PARAM_ID              ID_TYPE NOT NULL, "
   SQL = SQL & "CUSTOMER_ID             ID_TYPE, "
   SQL = SQL & "SALE_RATIO               MONEY_TYPE, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PARAM_ITEM ADD CONSTRAINT PARAM_ITEM_ID_PI PRIMARY KEY (PARAM_ITEM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PARAM_ITEM ADD CONSTRAINT PARAM_ITEM_PARAM_ID_FK FOREIGN KEY (PARAM_ID) REFERENCES PARAMETERS "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PARAM_ITEM ADD CONSTRAINT PARAM_ITEM_CUSTOMER_ID_FK FOREIGN KEY (CUSTOMER_ID) REFERENCES CUSTOMER "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR PARAM_ITEM_SEQ "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE ADJPRM_ITEM ADD EXPENSE_COST MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE ADJPRM_ITEM ADD FEED_COST MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PARAM_ITEM ADD PIG_TYPE ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PARAM_ITEM ADD TRANSFER_RATE MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PARAM_ITEM ADD CONSTRAINT PARAM_ITEM_PIG_TYPE_ID_FK FOREIGN KEY (PIG_TYPE) REFERENCES PRODUCT_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PARAM_ITEM ADD PIG_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PARAM_ITEM ADD CONSTRAINT PARAM_ITEM_PIG_ID_FK FOREIGN KEY (PIG_ID) REFERENCES PART_ITEM "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PARAM_ITEM ADD BUY_DATE DATE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PARAM_ITEM ADD BUY_AVG_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PARAM_ITEM ADD BUY_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PARAM_ITEM ADD BUY_TOTAL_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PARAM_ITEM ADD BUY_AVG_WEIGHT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CAPITAL_MOVEMENT ADD BATCH_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BALANCE_ACCUM ADD BATCH_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CAPITAL_LOSS ADD BATCH_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PARAM_ITEM ADD EXPENSE_DATE DATE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PARAM_ITEM ADD EXPENSE_TYPE ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PARAM_ITEM ADD EXPENSE_NAME DESC_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PARAM_ITEM ADD EXP_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PARAM_ITEM ADD EXP_TOTAL_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PARAM_ITEM ADD EXP_AVG_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EXPENSE_RATIO ADD BATCH_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE TABLE STATUS_TYPE "
   SQL = SQL & "( "
   SQL = SQL & "STATUS_TYPE_ID        ID_TYPE NOT NULL, "
   SQL = SQL & "STATUS_TYPE_NO              CODE_TYPE NOT NULL, "
   SQL = SQL & "STATUS_TYPE_NAME             NAME_TYPE, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE STATUS_TYPE ADD CONSTRAINT STATUS_TYPE_ID_PK PRIMARY KEY (STATUS_TYPE_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR STATUS_TYPE_SEQ "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PRODUCT_STATUS ADD STATUS_TYPE ID_TYPE "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_97 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_98() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_98"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
         
   SQL = "CREATE TABLE MONTHLY_ACCUM "
   SQL = SQL & "( "
   SQL = SQL & "MONTHLY_ACCUM_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "PART_ITEM_ID          ID_TYPE NOT NULL, "
   SQL = SQL & "DOCUMENT_DATE         DATE_TYPE NOT NULL, "
   SQL = SQL & "IMPORT_AMOUNT         MONEY_TYPE, "
   SQL = SQL & "EXPORT_AMOUNT         MONEY_TYPE, "
   SQL = SQL & "BALANCE_AMOUNT1       MONEY_TYPE, "
   SQL = SQL & "BALANCE_AMOUNT2       MONEY_TYPE, "
   SQL = SQL & "LOCATION_ID           ID_TYPE, "
   SQL = SQL & "SIMULATE_FLAG         FLAG_TYPE,"
   SQL = SQL & "BATCH_ID              ID_TYPE,"
   SQL = SQL & "YYYYMM                CODE_TYPE,"
   SQL = SQL & "BIRTH_AMOUNT          MONEY_TYPE,"
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MONTHLY_ACCUM ADD CONSTRAINT MONTHLY_ACCUM_ID_PK PRIMARY KEY (MONTHLY_ACCUM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MONTHLY_ACCUM ADD CONSTRAINT MONTHLY_ACCUM_LOCATION_ID_FK FOREIGN KEY (LOCATION_ID) REFERENCES LOCATION (LOCATION_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MONTHLY_ACCUM ADD CONSTRAINT MONTHLY_ACCUM_PART_ID_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM (PART_ITEM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR MONTHLY_ACCUM_SEQ "
   m_Conn.Execute (SQL)
      
   SQL = "ALTER TABLE BALANCE_ACCUM ADD YYYYMM CODE_TYPE "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_98 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_98_jill() As Boolean
On Error Resume Next 'GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_98_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "CREATE TABLE PACKAGE_TYPE "
   SQL = SQL & "( "
   SQL = SQL & "PACKAGE_TYPE_ID        ID_TYPE NOT NULL, "
   SQL = SQL & "PACKAGE_TYPE_CODE               CODE_TYPE, "
   SQL = SQL & "PACKAGE_TYPE_NAME               NAME_TYPE, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PACKAGE_TYPE ADD CONSTRAINT PACKAGE_TYPE_ID_PK PRIMARY KEY (PACKAGE_TYPE_ID) "
   m_Conn.Execute (SQL)
   
    SQL = "CREATE GENERATOR PACKAGE_TYPE_SEQ "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE TABLE PACKAGE "
   SQL = SQL & "( "
   SQL = SQL & "PKG_ID        ID_TYPE NOT NULL, "
   SQL = SQL & "PKG_CODE               CODE_TYPE, "
   SQL = SQL & "PKG_NAME               NAME_TYPE, "
   SQL = SQL & "PKG_BASIC_FLAG               FLAG_TYPE, "
   SQL = SQL & "PKG_TYPE             ID_TYPE, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PACKAGE ADD CONSTRAINT PKG_ID_PK PRIMARY KEY (PKG_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PACKAGE ADD CONSTRAINT PKG_TYPE_FK FOREIGN KEY (PKG_TYPE) REFERENCES PACKAGE_TYPE "
   m_Conn.Execute (SQL)
   
    SQL = "CREATE GENERATOR PACKAGE_SEQ "
   m_Conn.Execute (SQL)
   
   
   SQL = "CREATE TABLE PACKAGE_DETAIL "
   SQL = SQL & "( "
   SQL = SQL & "PKG_DETAIL_ID        ID_TYPE NOT NULL, "
   SQL = SQL & "PKG_ID               ID_TYPE, "
   SQL = SQL & "STATUS_BUY_ID               ID_TYPE, "
   SQL = SQL & "FROM_WEIGHT             MONEY_TYPE, "
   SQL = SQL & "TO_WEIGHT             MONEY_TYPE, "
   SQL = SQL & "CUT_WEIGHT             MONEY_TYPE, "
   SQL = SQL & "COST_PER_EXCEED             MONEY_TYPE, "
   SQL = SQL & "PEDIGREE_COST             MONEY_TYPE, "
   SQL = SQL & "COST_PER_WEIGHT             MONEY_TYPE, "
   SQL = SQL & "COST_PER_UNIT             MONEY_TYPE, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PACKAGE_DETAIL ADD CONSTRAINT PKG_DETAIL_ID_PK PRIMARY KEY (PKG_DETAIL_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PACKAGE_DETAIL ADD CONSTRAINT PKG_ID_FK FOREIGN KEY (PKG_ID) REFERENCES PACKAGE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PACKAGE_DETAIL ADD CONSTRAINT STATUS_BUY_ID_FK FOREIGN KEY (STATUS_BUY_ID) REFERENCES PRODUCT_STATUS "
   m_Conn.Execute (SQL)
   
    SQL = "CREATE GENERATOR PACKAGE_DETAIL_SEQ "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_98_jill = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_99_jill() As Boolean
On Error Resume Next  'GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_99_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "CREATE TABLE CUSTOMER_PACKAGE "
   SQL = SQL & "( "
   SQL = SQL & "CUSTOMER_PACKAGE_ID        ID_TYPE NOT NULL, "
   SQL = SQL & "CUSTOMER_ID             ID_TYPE, "
   SQL = SQL & "PKG_TYPE             ID_TYPE, "
   SQL = SQL & "PKG_ID             ID_TYPE, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER_PACKAGE ADD CONSTRAINT CUSTOMER_PACKAGE_ID_PK PRIMARY KEY (CUSTOMER_PACKAGE_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER_PACKAGE ADD CONSTRAINT CUSTOMER_PACKAGE_CUSTOMER_ID_FK FOREIGN KEY (CUSTOMER_ID) REFERENCES CUSTOMER "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER_PACKAGE ADD CONSTRAINT CUSTOMER_PACKAGE_PKG_TYPE_FK FOREIGN KEY (PKG_TYPE) REFERENCES PACKAGE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CUSTOMER_PACKAGE ADD CONSTRAINT CUSTOMER_PACKAGE_PKG_ID_FK FOREIGN KEY (PKG_ID) REFERENCES PACKAGE "
   m_Conn.Execute (SQL)
   
    SQL = "CREATE GENERATOR CUSTOMER_PACKAGE_SEQ "
   m_Conn.Execute (SQL)
   
   
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_99_jill = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_100_jill() As Boolean
On Error Resume Next 'GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_100_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "ALTER TABLE DO_ITEM ADD PKG_TYPE ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD CONSTRAINT DO_ITEM_PKG_TYPE_FK FOREIGN KEY (PKG_TYPE) REFERENCES PACKAGE_TYPE "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_100_jill = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_101() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_101"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
            
   SQL = "ALTER TABLE TRNPRM_ITEM ADD LOSS_TYPE ID_TYPE "
   m_Conn.Execute (SQL)
      
   SQL = "ALTER TABLE PARAM_ITEM ADD SHARE_SELL_TYPE ID_TYPE "
   m_Conn.Execute (SQL)
            
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "DELETE FROM PART_LOCATION "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM PART_ITEM WHERE PART_ITEM_ID NOT IN (SELECT PART_ITEM_ID FROM PRTITEM_MAP) AND PIG_FLAG = 'Y' "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE TRNPRM_ITEM SET LOSS_TYPE = 1 "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE PARAM_ITEM SET SHARE_SELL_TYPE = 1 "
   m_Conn.Execute (SQL)
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_101 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_102_jill() As Boolean
On Error Resume Next 'GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_102_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "ALTER TABLE PACKAGE DROP CONSTRAINT  PKG_TYPE_FK  "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_102_jill = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
Public Function Patch_1_0_12_103_jill() As Boolean
On Error Resume Next 'GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_103_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "ALTER TABLE CUSTOMER_PACKAGE DROP CONSTRAINT  CUSTOMER_PACKAGE_PKG_TYPE_FK  "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_103_jill = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
Public Function Patch_1_0_12_104_jill() As Boolean
On Error Resume Next 'GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_104_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "ALTER TABLE DO_ITEM DROP CONSTRAINT  DO_ITEM_PKG_TYPE_FK  "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_104_jill = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
Public Function Patch_1_0_12_105_jill() As Boolean
On Error Resume Next 'GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_105_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "ALTER TABLE DO_ITEM ADD PEDIGREE_COST  MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_105_jill = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
Public Function Patch_1_0_12_106_jill() As Boolean
On Error Resume Next 'GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_106_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   SQL = "ALTER TABLE DO_ITEM ADD DISCOUNT_PERCENT  MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD DISCOUNT_AMOUNT  MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD DISCOUNT_REASON  DESC_TYPE "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_106_jill = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
Public Function Patch_1_0_12_107_jill() As Boolean
On Error Resume Next 'GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_107_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   SQL = "ALTER TABLE DO_ITEM ADD SHOW_AVG  FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_107_jill = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
Public Function Patch_1_0_12_108_jill() As Boolean
On Error Resume Next 'GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_108_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   SQL = "UPDATE DO_ITEM SET SHOW_AVG = 'Y' "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_108_jill = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_109() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean
Dim PaidAmount As Double
Dim UpdateFlag As Boolean

   VersionStr = "1_0_12_109"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
            
   SQL = "ALTER TABLE BRTPRM_ITEM ADD PIGG_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BRTPRM_ITEM ADD PIGL_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BRTPRM_ITEM ADD PIGG_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BRTPRM_ITEM ADD PIGL_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
      
   SQL = "ALTER TABLE BRTPRM_ITEM ADD PIGGL_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BRTPRM_ITEM ADD BREED_PERCENT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MONTHLY_ACCUM ADD SELL_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MONTHLY_ACCUM ADD STATUS_IN_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MONTHLY_ACCUM ADD STATUS_OUT_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MONTHLY_ACCUM ADD BUY_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   Call glbDaily.PatchBirthItemParam(False, glbErrorLog)
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_109 = True
   Exit Function
   
ErrorHandler:
   ''debug.print Err.DESCRIPTION
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_1_0_12_110_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_110_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "CREATE TABLE COST_ACCUM "
   SQL = SQL & "( "
   SQL = SQL & "COST_ACCUM_ID        ID_TYPE NOT NULL, "
   SQL = SQL & "DOCUMENT_DATE             DATE_TYPE, "
   SQL = SQL & "DOCUMENT_TYPE             ID_TYPE, "
   SQL = SQL & "PART_ITEM_ID             ID_TYPE, "
   SQL = SQL & "LOCATION_ID             ID_TYPE, "
   SQL = SQL & "BATCH_ID             ID_TYPE, "
   SQL = SQL & "COST_RAW             MONEY_TYPE, "
   SQL = SQL & "COST_EXP             MONEY_TYPE, "
   SQL = SQL & "YYYYMM             CODE_TYPE, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE COST_ACCUM ADD CONSTRAINT COST_ACCUM_ID_PK PRIMARY KEY (COST_ACCUM_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE COST_ACCUM ADD CONSTRAINT COST_ACCUM_PART_ID_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE COST_ACCUM ADD CONSTRAINT COST_ACCUM_LOCATION_ID_FK FOREIGN KEY (LOCATION_ID) REFERENCES LOCATION "
   m_Conn.Execute (SQL)
   
    SQL = "CREATE GENERATOR COST_ACCUM_SEQ "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_110_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_1_0_12_111_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_111_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "ALTER TABLE COST_ACCUM ADD CUS_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE COST_ACCUM ADD CONSTRAINT COST_ACCUM_CUS_ID_FK FOREIGN KEY (CUS_ID) REFERENCES CUSTOMER "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_111_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_1_0_12_112_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_112_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "ALTER TABLE INVENTORY_DOC ADD CUS_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_DOC ADD CONSTRAINT INVENTORY_DOC_CUS_ID_FK FOREIGN KEY (CUS_ID) REFERENCES CUSTOMER "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_112_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_1_0_12_113_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_113_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "ALTER TABLE COST_ACCUM ADD ITEM_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
      
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_113_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_1_0_12_114_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_114_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "ALTER TABLE COST_ACCUM ADD DOCUMENT_CATEGORY ID_TYPE "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_114_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_1_0_12_115_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_115_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===

   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "DELETE FROM PAYMENT_ITEM WHERE PAYMENT_ID IN (SELECT PAYMENT_ID FROM PAYMENT WHERE PAYMENT_DATE < '2006-08-01 00:00:00') "
   m_Conn.Execute (SQL)

   SQL = "DELETE FROM PAYMENT WHERE PAYMENT_DATE < '2006-08-01 00:00:00' "
   m_Conn.Execute (SQL)
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_115_seub = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
   glbErrorLog.ShowUserError
   
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_1_0_12_116_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_116_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===

   SQL = "ALTER TABLE PARAM_ITEM ADD PIG_STATUS_ID ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE PARAM_ITEM ADD CTRL_FROM_DATE DATE_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE PARAM_ITEM ADD CTRL_TO_DATE DATE_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE PARAM_ITEM ADD CTRL_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)

   Call m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_116_seub = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
   glbErrorLog.ShowUserError
   
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_1_0_12_115_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_115_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "ALTER TABLE COST_ACCUM ADD COST_PB MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_115_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_1_0_12_116_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_116_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "ALTER TABLE COST_ACCUM ADD COST_MOVEMENT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_116_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_1_0_12_117_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_117_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===

   SQL = "ALTER TABLE PARAMETERS ADD ADG MONEY_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE SYSTEM_PARAM ADD FROM_LOCK_DATE DATE_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE SYSTEM_PARAM ADD TO_LOCK_DATE DATE_TYPE "
   m_Conn.Execute (SQL)

   Call m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_117_seub = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
   glbErrorLog.ShowUserError
   
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_1_0_12_118_jill() As Boolean
'On Error GoTo ErrorHandler
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_118_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "CREATE TABLE WEIGHT_RECORD "
   SQL = SQL & "( "
   SQL = SQL & "WEIGHT_RECORD_ID        ID_TYPE NOT NULL, "
   SQL = SQL & "RECORD_DATE             DATE_TYPE, "
   SQL = SQL & "PART_ITEM_ID             ID_TYPE, "
   SQL = SQL & "ITEM_AMOUNT             MONEY_TYPE, "
   SQL = SQL & "AVG_WEIGHT             MONEY_TYPE, "
   SQL = SQL & "WEIGHT_AMOUNT             MONEY_TYPE, "
   SQL = SQL & "PIG_AGE             MONEY_TYPE, "
   SQL = SQL & "PIG_AGE_INT             ID_TYPE, "
   SQL = SQL & "BATCH_ID             ID_TYPE, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE WEIGHT_RECORD ADD CONSTRAINT WEIGHT_RECORD_ID_PK PRIMARY KEY (WEIGHT_RECORD_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE WEIGHT_RECORD ADD CONSTRAINT WEIGHT_RECORD_PART_ID_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE WEIGHT_RECORD ADD CONSTRAINT WEIGHT_RECORD_BATCH_ID_FK FOREIGN KEY (BATCH_ID) REFERENCES BATCH "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR WEIGHT_RECORD_SEQ "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_118_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_1_0_12_119_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_119_seub"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "ALTER TABLE IMPORT_ITEM ADD ACTUAL_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE IMPORT_ITEM ADD ACTUAL_PRICE MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_119_seub = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_1_0_12_120_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_120_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "INSERT INTO REVENUE_TYPE (REVENUE_TYPE_ID, REVENUE_NO, REVENUE_NAME, CREATE_DATE, CREATE_BY, MODIFY_DATE, MODIFY_BY) VALUES (99999, '99999', 'ยอกยกมาลูกหนี้', '2005-10-13 13:49:40.0000', 4893, '2005-10-13 13:49:40.0000', 4893) "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_120_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_1_0_12_121_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_12_121_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   '===
   
   SQL = "ALTER TABLE BALANCE_ACCUM ADD BALANCE_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MOVEMENT_ITEM ADD BALANCE_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CAPITAL_MOVEMENT ADD BALANCE_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LOSS_ITEM ADD BALANCE_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CAPITAL_LOSS ADD BALANCE_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE RECEIPT_ITEM ADD BALANCE_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD BALANCE_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE RO_ITEM ADD BALANCE_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE RECEIPT_CNDN ADD BALANCE_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_ADJUST ADD BALANCE_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BILLING_DOC ADD BALANCE_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE IMPORT_ITEM ADD BALANCE_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EXPORT_ITEM ADD BALANCE_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INVENTORY_DOC ADD BALANCE_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PAYMENT_ITEM ADD BALANCE_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PAYMENT ADD BALANCE_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR BALANCE_COUNT_SEQ "
   m_Conn.Execute (SQL)
   
   Call m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_12_121_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2006_11_07_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_11_07_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   
  SQL = "CREATE TABLE  CONFIG_DOC  "
   SQL = SQL & "( "
   SQL = SQL & "CONFIG_DOC_ID             ID_TYPE NOT NULL, "
   SQL = SQL & "CONFIG_DOC_TYPE                    ID_TYPE , "
   SQL = SQL & "ENTERPRISE_ID             ID_TYPE , "
   SQL = SQL & "PREFIX                    CODE_TYPE , "
   SQL = SQL & "CODE1                    CODE_TYPE , "
   SQL = SQL & "YEAR_TYPE                    ID_TYPE , "
   SQL = SQL & "CODE2                    CODE_TYPE , "
   SQL = SQL & "MONTH_TYPE                    ID_TYPE , "
   SQL = SQL & "CODE3                    CODE_TYPE , "
   SQL = SQL & "DIGIT_AMOUNT                    ID_TYPE , "
   SQL = SQL & "RUNNING_NO                    ID_TYPE , "
   SQL = SQL & "LAST_NO                    CODE_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CONFIG_DOC  ADD CONSTRAINT CONFIG_DOC_PK  PRIMARY KEY (CONFIG_DOC_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR CONFIG_DOC_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_11_07_1_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2006_11_07_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_11_07_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE CONFIG_DOC ADD MM CODE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CONFIG_DOC ADD UPDATE_MONTH_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE CONFIG_DOC SET UPDATE_MONTH_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CONFIG_DOC ADD UPDATE_YEAR_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE CONFIG_DOC SET UPDATE_YEAR_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_11_07_2_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2006_11_10_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_11_10_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE EXPORT_ITEM ADD DISCOUNT_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_11_10_1_jill = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2006_11_29_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_11_29_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
'   SQL = "CREATE TABLE  MASTER_REF "
'   SQL = SQL & "( "
'   SQL = SQL & "KEY_ID      ID_TYPE NOT NULL, "
'   SQL = SQL & "KEY_CODE           CODE_TYPE , "
'   SQL = SQL & "KEY_NAME           NAME_TYPE , "
'   SQL = SQL & "MASTER_AREA           ID_TYPE , "
'   SQL = SQL & "CBO_ID           ID_TYPE , "
'   SQL = SQL & "CASE_NO           CODE_TYPE , "
'   SQL = SQL & "TEMP_ID1           ID_TYPE , "
'   SQL = SQL & "TEMP_ID2           ID_TYPE , "
'   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
'   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
'   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
'   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
'   SQL = SQL & ") "
'   m_Conn.Execute (SQL)
'
'   SQL = "ALTER TABLE MASTER_REF ADD CONSTRAINT KEY_ID_PK PRIMARY KEY (KEY_ID) "
'   m_Conn.Execute (SQL)
'
'   SQL = "CREATE TABLE  CASH_DOC "
'   SQL = SQL & "( "
'   SQL = SQL & "CASH_DOC_ID      ID_TYPE NOT NULL, "
'   SQL = SQL & "DOCUMENT_DATE           DATE_TYPE , "
'   SQL = SQL & "BANK_BRANCH           ID_TYPE , "
'   SQL = SQL & "BANK_ID           ID_TYPE , "
'   SQL = SQL & "BANK_ACCOUNT           ID_TYPE , "
'   SQL = SQL & "ENTERPRISE_ID           ID_TYPE , "
'   SQL = SQL & "DOCUMENT_NO           CODE_TYPE , "
'   SQL = SQL & "DOCUMENT_TYPE           ID_TYPE , "
'   SQL = SQL & "EMP_ID           ID_TYPE , "
'   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
'   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
'   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
'   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
'   SQL = SQL & ") "
'   m_Conn.Execute (SQL)
'
'   SQL = "ALTER TABLE CASH_DOC ADD CONSTRAINT CASH_DOC_ID_PK PRIMARY KEY (CASH_DOC_ID) "
'   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CASH_DOC ADD CONSTRAINT CASH_DOC_BANK_ID_FK FOREIGN KEY (BANK_ID) REFERENCES BANK  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CASH_DOC ADD CONSTRAINT CASH_DOC_BBRANCH_FK FOREIGN KEY (BANK_BRANCH) REFERENCES BANK_BRANCH  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CASH_DOC ADD CONSTRAINT CASH_DOC_BNK_ACCOUNT_FK FOREIGN KEY (BANK_ACCOUNT) REFERENCES BANK_ACCOUNT  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CASH_DOC ADD CONSTRAINT CASH_DOC_EMP_ID_FK FOREIGN KEY (EMP_ID) REFERENCES EMPLOYEE  "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE TABLE  CHEQUE "
   SQL = SQL & "( "
   SQL = SQL & "CHEQUE_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "CHEQUE_NO           CODE_TYPE , "
   SQL = SQL & "CHEQUE_DATE           DATE_TYPE , "
   SQL = SQL & "EFFECTIVE_DATE           DATE_TYPE , "
   SQL = SQL & "CHEQUE_TYPE           ID_TYPE , "
   SQL = SQL & "CHEQUE_STATUS           ID_TYPE , "
   SQL = SQL & "BANK_ID           ID_TYPE , "
   SQL = SQL & "BANK_BRANCH           ID_TYPE , "
   SQL = SQL & "ENTERPRISE_ID           ID_TYPE , "
   SQL = SQL & "DIRECTION           ID_TYPE , "
   SQL = SQL & "CHEQUE_AMOUNT           MONEY_TYPE , "
   SQL = SQL & "APAR_ID           ID_TYPE , "
   SQL = SQL & "BANK_FLAG           FLAG_TYPE , "
   SQL = SQL & "TX_DATE           DATE_TYPE , "
   SQL = SQL & "TX_NO CODE_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CHEQUE ADD CONSTRAINT CHEQUE_ID_PK PRIMARY KEY (CHEQUE_ID) "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CHEQUE ADD CONSTRAINT CHEQUE_BANK_BRANCH_FK FOREIGN KEY (BANK_BRANCH) REFERENCES MASTER_REF  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CHEQUE ADD CONSTRAINT CHEQUE_BANK_ID_FK FOREIGN KEY (BANK_ID) REFERENCES MASTER_REF  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CHEQUE ADD CONSTRAINT CHEQUE_CHEQUE_TYPE_FK FOREIGN KEY (CHEQUE_TYPE) REFERENCES MASTER_REF  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CHEQUE ADD CONSTRAINT CHEQUE_ENTERPRISE_ID_FK FOREIGN KEY (ENTERPRISE_ID) REFERENCES ENTERPRISE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CHEQUE ADD CONSTRAINT CHEQUE_APAR_ID_FK FOREIGN KEY (APAR_ID) REFERENCES CUSTOMER  "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE TABLE CASH_TRAN "
   SQL = SQL & "( "
   SQL = SQL & "CASH_TRAN_ID      ID_TYPE NOT NULL, "
    SQL = SQL & "BILLING_DOC_ID           ID_TYPE , "
   SQL = SQL & "CHECK_ID           ID_TYPE , "
   SQL = SQL & "AMOUNT           MONEY_TYPE , "
   SQL = SQL & "BANK_ID           ID_TYPE , "
    SQL = SQL & "BANK_BRANCH           ID_TYPE , "
   SQL = SQL & "ENTERPRISE_ID           ID_TYPE , "
   SQL = SQL & "TX_TYPE           CODE_TYPE , "
   SQL = SQL & "PAYMENT_TYPE           ID_TYPE , "
   SQL = SQL & "BANK_ACCOUNT           ID_TYPE , "
   SQL = SQL & "CASH_DOC_ID           ID_TYPE , "
   SQL = SQL & "FEE_AMOUNT           MONEY_TYPE , "
   SQL = SQL & "NET_AMOUNT           MONEY_TYPE , "
   SQL = SQL & "TX_DATE           DATE_TYPE , "
   SQL = SQL & "TX_NO           CODE_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE CASH_TRAN ADD CONSTRAINT CASH_TRAN_ID_PK PRIMARY KEY (CASH_TRAN_ID) "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CASH_TRAN ADD CONSTRAINT CASH_TRAN_BANK_AMOUNT_FK FOREIGN KEY (BANK_ACCOUNT) REFERENCES BANK_ACCOUNT  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CASH_TRAN ADD CONSTRAINT CASH_TRAN_BILLING_DOC_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CASH_TRAN ADD CONSTRAINT CASH_TRAN_CHEQUE_FK FOREIGN KEY (CHECK_ID) REFERENCES CHEQUE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CASH_TRAN ADD CONSTRAINT CASH_TRAN_CASH_DOC_FK FOREIGN KEY (CASH_DOC_ID) REFERENCES CASH_DOC  "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR CHEQUE_SEQ "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR CASH_DOC_SEQ "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR CASH_TRAN_SEQ "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_11_29_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_11_29_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_11_29_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "ALTER TABLE CASH_TRAN ADD CUSTOMER_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_11_29_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_11_29_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_11_29_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = " ALTER TABLE CASH_TRAN ADD CONSTRAINT CASH_TRAN_CUSTOMER_ID_FK FOREIGN KEY (CUSTOMER_ID) REFERENCES CUSTOMER  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_11_29_3_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_11_29_4_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_11_29_4_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
'  SQL = " ALTER TABLE CASH_DOC ADD  CUSTOMER_ID ID_TYPE  "
'   m_Conn.Execute (SQL)
   
  SQL = " ALTER TABLE CASH_DOC ADD CONSTRAINT CASH_DOC_CUSTOMER_ID_FK FOREIGN KEY (CUSTOMER_ID) REFERENCES CUSTOMER  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_11_29_4_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2006_11_29_5_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_11_29_5_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = " ALTER TABLE CHEQUE ADD  POST_FLAG FLAG_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_11_29_5_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2006_11_29_6_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_11_29_6_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = " UPDATE CHEQUE SET POST_FLAG = 'N' "
   m_Conn.Execute (SQL)
    
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_11_29_6_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2006_11_29_7_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_11_29_7_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
'   SQL = "CREATE TABLE  CASH_DOC_POST "
'   SQL = SQL & "( "
'   SQL = SQL & "CASH_DOC_POST_ID      ID_TYPE NOT NULL, "
'   SQL = SQL & "CASH_DOC_ID           ID_TYPE , "
'   SQL = SQL & "BANK_BRANCH           ID_TYPE , "
'   SQL = SQL & "BANK_ID           ID_TYPE , "
'   SQL = SQL & "CHEQUE_ID           ID_TYPE , "
'   SQL = SQL & "CHEQUE_AMOUNT           MONEY_TYPE , "
'   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
'   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
'   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
'   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
'   SQL = SQL & ") "
'   m_Conn.Execute (SQL)
'
'   SQL = "ALTER TABLE CASH_DOC_POST ADD CONSTRAINT CASH_DOC_POST_ID_PK PRIMARY KEY (CASH_DOC_POST_ID) "
'   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CASH_DOC_POST ADD CONSTRAINT CASH_DOC_POST_CASH_DOC_ID_FK FOREIGN KEY (CASH_DOC_ID) REFERENCES CASH_DOC  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CASH_DOC_POST ADD CONSTRAINT CASH_DOC_POST_BANK_ID_FK FOREIGN KEY (BANK_ID) REFERENCES BANK  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CASH_DOC_POST ADD CONSTRAINT CASH_DOC_POST_BBRANCH_FK FOREIGN KEY (BANK_BRANCH) REFERENCES BANK_BRANCH  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CASH_DOC_POST ADD CONSTRAINT CASH_DOC_POST_CHEQUE_FK FOREIGN KEY (CHEQUE_ID) REFERENCES CHEQUE  "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR CASH_DOC_POST_SEQ "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_11_29_7_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2006_11_29_8_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_11_29_8_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = " ALTER TABLE CHEQUE DROP CONSTRAINT CHEQUE_BANK_BRANCH_FK  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CHEQUE DROP CONSTRAINT CHEQUE_BANK_ID_FK "
   m_Conn.Execute (SQL)
      
    SQL = " ALTER TABLE CHEQUE ADD CONSTRAINT CHEQUE_BANK_BRANCH_FK FOREIGN KEY (BANK_BRANCH) REFERENCES BANK_BRANCH  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE CHEQUE ADD CONSTRAINT CHEQUE_BANK_ID_FK FOREIGN KEY (BANK_ID) REFERENCES BANK  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_11_29_8_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2006_11_30_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2006_11_30_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
  SQL = "CREATE GENERATOR MASTER_REF_SEQ "
   m_Conn.Execute (SQL)
    
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_11_30_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2006_12_04_1_jill() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2006_12_04_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE BANK_ACCOUNT ADD  BANK_ID ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE BANK_ACCOUNT ADD  BBRANCH_ID ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE BANK_ACCOUNT ADD CONSTRAINT BANK_ACCOUNT_BANK_ID_FK FOREIGN KEY (BANK_ID) REFERENCES BANK  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE BANK_ACCOUNT ADD CONSTRAINT BANK_ACCOUNT_BBRANCH_ID_FK FOREIGN KEY (BBRANCH_ID) REFERENCES BANK_BRANCH  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_12_04_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2006_12_28_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim EI As CExportItem
Dim IsOK As Boolean
   VersionStr = "2006_12_28_1_jill"

   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                              
   Set EI = New CExportItem
   Call glbDaily.PatchPigAgeByExportFood(EI, IsOK, False, glbErrorLog)
   Set EI = Nothing
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2006_12_28_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_05_17_1_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim EI As CExportItem
Dim IsOK As Boolean
   VersionStr = "2007_05_17_1_seub"

   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE BALANCE_ACCUM ADD  PIG_AGE ID_TYPE  "
   m_Conn.Execute (SQL)
   '===
   
   m_Conn.BeginTrans
   HasBegin = True
                                 
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_05_17_1_seub = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_05_17_2_seub() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim Ba As CBalanceAccum
Dim IsOK As Boolean
   
   VersionStr = "2007_05_17_2_seub"

   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   m_Conn.BeginTrans
   HasBegin = True
   
'   Set Ba = New CBalanceAccum
'   Call glbDaily.PatchBAPigAge(Ba, IsOK, False, glbErrorLog)
'   Set Ba = Nothing
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_05_17_2_seub = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_05_22_1_seub() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim Ba As CBalanceAccum
Dim IsOK As Boolean
   
   VersionStr = "2007_05_22_1_seub"

   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = " ALTER TABLE PART_ITEM ADD  FEED_GROUP ID_TYPE  "
   m_Conn.Execute (SQL)
      
   SQL = "ALTER TABLE PART_ITEM ADD CONSTRAINT PART_ITEM_FEED_GROUP_FK FOREIGN KEY (FEED_GROUP) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
    
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_05_22_1_seub = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_07_12_1_jill() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_07_12_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE EXPENSE_TYPE ADD  SELL_PARENT_FLAG FLAG_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_07_12_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_07_19_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2007_07_19_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = "CREATE TABLE  REVENUE_COST "
   SQL = SQL & "( "
   SQL = SQL & "REVENUE_COST_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "REVENUE_NO           CODE_TYPE , "
   SQL = SQL & "REVENUE_DATE           DATE_TYPE , "
   SQL = SQL & "REVENUE_DESC           DESC_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE TABLE  REVENUE_COST_ITEM "
   SQL = SQL & "( "
   SQL = SQL & "REVENUE_COST_ITEM_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "REVENUE_COST_ID           ID_TYPE , "
   SQL = SQL & "REVENUE_TYPE_ID           ID_TYPE , "
   SQL = SQL & "REVENUE_COST_ITEM_AMOUNT           MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR REVENUE_COST_SEQ "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR REVENUE_COST_ITEM_SEQ "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_07_19_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_07_19_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2007_07_19_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
'   SQL = "ALTER TABLE REVENUE_COST ADD CONSTRAINT REVENUE_COST_ID_PK PRIMARY KEY (REVENUE_COST_ID) "
'   m_Conn.Execute (SQL)
'
'   SQL = "ALTER TABLE REVENUE_COST_ITEM ADD CONSTRAINT REVENUE_COST_ITEM_ID_PK PRIMARY KEY (REVENUE_COST_ITEM_ID) "
'   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE REVENUE_COST_ITEM ADD CONSTRAINT REVENUE_CTI_REVENUE_ID_FK FOREIGN KEY (REVENUE_COST_ID) REFERENCES REVENUE_COST  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE REVENUE_COST_ITEM ADD CONSTRAINT REVENUE_CTI_TYPE_ID_FK FOREIGN KEY (REVENUE_TYPE_ID) REFERENCES REVENUE_TYPE  "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_07_19_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_08_20_1_jill() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_08_20_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE REVENUE_COST_ITEM ADD  PIG_STATUS ID_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_08_20_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_08_20_2_jill() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_08_20_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE REVENUE_COST_ITEM ADD CONSTRAINT REVENUE_COST_ITEM_STATUS_FK1 FOREIGN KEY (PIG_STATUS) REFERENCES PRODUCT_STATUS (PRODUCT_STATUS_ID) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_08_20_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_08_20_3_jill() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_08_20_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE REVENUE_COST_ITEM ADD  REVENUE_COST_ITEM_SELL MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_08_20_3_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_08_24_1_jill() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_08_24_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE DO_ITEM ADD  BFOOD_AMOUNT MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE DO_ITEM ADD  BEXPENSE_AMOUNT MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE DO_ITEM ADD  BIRTH_AMOUNT MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE DO_ITEM ADD  FOOD_AMOUNT MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE DO_ITEM ADD  EXPENSE_AMOUNT MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_08_24_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_08_24_2_jill() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_08_24_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE BRTPRM_ITEM ADD  BIRTH_COST MONEY_TYPE  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_08_24_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_09_11_1_jill() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_09_11_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD DO_ID_EX ID_TYPE  "
   m_Conn.Execute (SQL)
   
   SQL = " UPDATE RECEIPT_ITEM SET DO_ID_EX = DO_ID  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM DROP DO_ID  "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD DO_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = " UPDATE RECEIPT_ITEM SET DO_ID = DO_ID_EX  "
   m_Conn.Execute (SQL)
   
   SQL = " UPDATE RECEIPT_ITEM RI SET RI.DO_ID = NULL WHERE (RI.DO_ID NOT IN (SELECT BD.BILLING_DOC_ID FROM BILLING_DOC BD )) "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM ADD CONSTRAINT RECEIPT_ITEM_DO_ID_FK FOREIGN KEY (DO_ID) REFERENCES BILLING_DOC "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE RECEIPT_ITEM DROP DO_ID_EX "
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_09_11_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_09_14_1_jill() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_09_14_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = " DELETE FROM BATCH_ITEM UG WHERE UG.PARAM_ID NOT IN (SELECT PT.PARAM_ID FROM PARAMETERS PT )  "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BATCH_ITEM ADD CONSTRAINT BATCH_ITEM_PARAM_FK1 FOREIGN KEY (PARAM_ID) REFERENCES PARAMETERS "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_09_14_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_09_14_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2007_09_14_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = "CREATE TABLE INTAKE  "
   SQL = SQL & "( "
   SQL = SQL & "INTAKE_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "BATCH_ID           ID_TYPE , "
   SQL = SQL & "DOCUMENT_DATE           DATE_TYPE , "
   SQL = SQL & "PIG_ID           ID_TYPE , "
   SQL = SQL & "PART_ITEM_ID           ID_TYPE , "
   SQL = SQL & "INTAKE_COST           MONEY_TYPE , "                       'มูลค่าเบอร์อาหารนั้นๆที่หมูกิน
   SQL = SQL & "INTAKE_AMOUNT           MONEY_TYPE , "               'จำนวนหมู
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INTAKE ADD CONSTRAINT INTAKE_ID_PK PRIMARY KEY (INTAKE_ID) "
   m_Conn.Execute (SQL)

   SQL = "CREATE GENERATOR INTAKE_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_09_14_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_09_14_3_jill() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_09_14_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE INTAKE ADD CONSTRAINT INTAKE_PIG_ID_FK FOREIGN KEY (PIG_ID) REFERENCES PART_ITEM_ID "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INTAKE ADD CONSTRAINT INTAKE_PART_ITEM_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM_ID "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_09_14_3_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_09_14_4_jill() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_09_14_4_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE INTAKE ADD SIMULATE_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE INTAKE ADD YYYYMM CODE_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_09_14_4_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_09_18_1_jill() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_09_18_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE INTAKE ADD CURRENT_PIG_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_09_18_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_09_25_1_jill() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_09_25_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE WEIGHT_RECORD ADD ADG MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_09_25_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_10_16_1_jill() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_10_16_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE USEDPRM_ITEM ADD INTAKE_TYPE ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_10_16_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_10_16_2_jill() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_10_16_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE USEDPRM_ITEM UI SET UI.INTAKE_TYPE = 1 WHERE ((SELECT  PT.PART_TYPE_NO FROM PART_ITEM PI LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) WHERE PI.PART_ITEM_ID =  UI.FOOD_ID) = 10) "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE USEDPRM_ITEM UI SET UI.INTAKE_TYPE = 2 WHERE ((SELECT  PT.PART_TYPE_NO FROM PART_ITEM PI LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) WHERE PI.PART_ITEM_ID =  UI.FOOD_ID) = 41) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_10_16_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_10_16_3_jill() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_10_16_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE USEDPRM_ITEM UI SET UI.INTAKE_TYPE = 3 WHERE ((SELECT  PT.PART_TYPE_NO FROM PART_ITEM PI LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) WHERE PI.PART_ITEM_ID =  UI.FOOD_ID) = 51) "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE USEDPRM_ITEM UI SET UI.INTAKE_TYPE = 2 WHERE ((SELECT  PT.PART_TYPE_NO FROM PART_ITEM PI LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) WHERE PI.PART_ITEM_ID =  UI.FOOD_ID) = 33) "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_10_16_3_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_10_16_4_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_10_16_4_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE DO_ITEM ADD BMEDICINE_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD MEDICINE_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE DO_ITEM ADD OTHER_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_10_16_4_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_10_16_5_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_10_16_5_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE ADJPRM_ITEM ADD MEDICINE_COST MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_10_16_5_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_10_18_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_10_18_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE PARAM_ITEM ADD YYYYMM CODE_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_10_18_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_10_24_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_10_24_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE PRODUCT_STATUS ADD NON_SALE_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_10_24_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_10_25_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_10_25_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE PRODUCT_STATUS SET NON_SALE_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_10_25_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_10_29_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_10_29_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BATCH ADD BALANCE_CASH MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BATCH ADD CASH_FIRST_MONTH MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_10_29_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_10_30_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_10_30_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE PARAM_ITEM ADD DEPRECIATION_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE PARAM_ITEM SET DEPRECIATION_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_10_30_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_10_30_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_10_30_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE EXPENSE_TYPE SET SELL_PARENT_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EXPENSE_TYPE ADD DEPRECIATION_GOOD_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_10_30_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_10_30_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_10_30_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE EXPENSE_TYPE SET DEPRECIATION_GOOD_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_10_30_3_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_11_12_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2007_11_12_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = "CREATE TABLE GL_AGE_AMOUNT  "
   SQL = SQL & "( "
   SQL = SQL & "GL_AGE_AMOUNT_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "PARAM_ID           ID_TYPE , "
   SQL = SQL & "GL_TYPE           ID_TYPE , "
   SQL = SQL & "GL_AGE           ID_TYPE , "
   SQL = SQL & "GL_AMOUNT           MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE GL_AGE_AMOUNT ADD CONSTRAINT GL_AGE_AMOUNT_ID_PK PRIMARY KEY (GL_AGE_AMOUNT_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR GL_AGE_AMOUNT_SEQ "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE TABLE GL_BACK_AMOUNT  "
   SQL = SQL & "( "
   SQL = SQL & "GL_BACK_AMOUNT_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "PARAM_ID           ID_TYPE , "
   SQL = SQL & "GL_TYPE           ID_TYPE , "
   SQL = SQL & "GL_AGE           ID_TYPE , "
   SQL = SQL & "GL_AMOUNT           MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE GL_BACK_AMOUNT ADD CONSTRAINT GL_BACK_AMOUNT_ID_PK PRIMARY KEY (GL_BACK_AMOUNT_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR GL_BACK_AMOUNT_SEQ "
   m_Conn.Execute (SQL)
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_11_12_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_11_12_2_jill() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_11_12_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE GL_AGE_AMOUNT ADD CONSTRAINT GL_AGE_AMOUNT_PIG_TYPE_FK FOREIGN KEY (GL_TYPE) REFERENCES PIG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE GL_BACK_AMOUNT ADD CONSTRAINT GL_BACK_AMOUNT_PIG_TYPE_FK FOREIGN KEY (GL_TYPE) REFERENCES PIG_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_11_12_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_11_12_3_jill() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_11_12_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE GL_AGE_AMOUNT ADD CONSTRAINT GL_AGE_AMOUNT_PARAM_FK FOREIGN KEY (PARAM_ID) REFERENCES PARAMETERS "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE GL_BACK_AMOUNT ADD CONSTRAINT GL_BACK_AMOUNT_PARAM_FK FOREIGN KEY (PARAM_ID) REFERENCES PARAMETERS "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_11_12_3_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_11_19_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2007_11_19_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = "CREATE TABLE PRICE_ADJUST "
   SQL = SQL & "( "
   SQL = SQL & "PRICE_ADJUST_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "PART_ITEM_ID           ID_TYPE , "
   SQL = SQL & "AVG_PRICE           MONEY_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PRICE_ADJUST ADD CONSTRAINT PRICE_ADJUST_ID_PK PRIMARY KEY (PRICE_ADJUST_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR PRICE_ADJUST_SEQ "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_11_19_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_11_19_2_jill() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_11_19_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE PRICE_ADJUST ADD CONSTRAINT PRICE_ADJUST_PART_ITEM_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_11_19_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_11_26_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_11_26_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE IMPORT_ITEM ADD INTAKE_TYPE ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_11_26_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2007_11_26_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_11_26_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BATCH ADD MEDICINE_FIRST_MONTH MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE BATCH ADD OTHER_FIRST_MONTH MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_11_26_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_11_29_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_11_29_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE ADJPRM_ITEM ADD BIRTH_COST MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_11_29_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_11_29_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_11_29_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE DO_ITEM ADD BBIRTH_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_11_29_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_12_03_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_12_03_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE REVENUE_COST_ITEM ADD PIG_TYPE ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_12_03_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_12_03_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_12_03_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE REVENUE_COST_ITEM ADD CONSTRAINT REVENUE_COST_ITEM_PIG_TYPE_FK FOREIGN KEY (PIG_TYPE) REFERENCES PRODUCT_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_12_03_2_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_12_06_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_12_06_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE SALEPRM_ITEM ADD FROM_SALE DATE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SALEPRM_ITEM ADD TO_SALE DATE_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_12_06_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_12_14_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_12_14_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE MONTHLY_ACCUM ADD ADJUST_IN_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MONTHLY_ACCUM ADD ADJUST_OUT_AMOUNT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_12_14_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_12_17_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_12_17_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE MONTHLY_ACCUM ADD GL_IN MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE MONTHLY_ACCUM ADD GL_OUT MONEY_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_12_17_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2007_12_25_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2007_12_25_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   m_Conn.BeginTrans
   HasBegin = True
   Dim Ei1 As CExportItem
   Dim Rs As ADODB.Recordset
   Dim iCount As Long
   
   Set Ei1 = New CExportItem
   Set Rs = New ADODB.Recordset
   Ei1.EXPORT_ITEM_ID = -1
   Ei1.DOCUMENT_TYPE = 7 'เอกสารโอนเข้าเรือนขาย
   Ei1.PIG_FLAG = "Y"
   Call Ei1.QueryData(1, Rs, iCount)
   
   While Not Rs.EOF
      Call Ei1.PopulateFromRS(1, Rs)
      Ei1.PIG_AGE = GetAge(Ei1.PART_NO, Ei1.DOCUMENT_DATE)
      Ei1.AGE_CODE = GetAgeCode(Ei1.PIG_AGE)
      
      Ei1.AddEditMode = SHOW_EDIT
      Call Ei1.AddEditData
      Rs.MoveNext
   Wend
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2007_12_25_1_jill = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_01_08_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2008_01_08_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "DELETE FROM MONTHLY_ACCUM WHERE PART_ITEM_ID IN (12227,12228,12229,12230,12231,12232,12233,12234,12235,12236,12237,12238)  "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM BALANCE_ACCUM WHERE PART_ITEM_ID IN (12227,12228,12229,12230,12231,12232,12233,12234,12235,12236,12237,12238)  "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM DO_ITEM WHERE PART_ITEM_ID IN (12227,12228,12229,12230,12231,12232,12233,12234,12235,12236,12237,12238)  "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM EXPORT_ITEM WHERE PART_ITEM_ID IN (12227,12228,12229,12230,12231,12232,12233,12234,12235,12236,12237,12238)  "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM IMPORT_ITEM WHERE PART_ITEM_ID IN (12227,12228,12229,12230,12231,12232,12233,12234,12235,12236,12237,12238)  "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM PART_LOCATION WHERE PART_ITEM_ID IN (12227,12228,12229,12230,12231,12232,12233,12234,12235,12236,12237,12238)  "
   m_Conn.Execute (SQL)
   
   SQL = "DELETE FROM PART_ITEM WHERE PART_ITEM_ID IN (12227,12228,12229,12230,12231,12232,12233,12234,12235,12236,12237,12238)  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_01_08_1_jill = True
   Exit Function
   
ErrorHandler:
      ''debug.print m_Conn.Errors(0)
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2008_05_29_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2008_05_29_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "DELETE FROM RECEIPT_ITEM WHERE DO_ID = 27045 "
   m_Conn.Execute (SQL)
      
   SQL = "UPDATE BILLING_DOC SET PAID_AMOUNT = 0 WHERE BILLING_DOC_ID = 27045 "
   m_Conn.Execute (SQL)
      
   SQL = "DELETE FROM RECEIPT_ITEM WHERE DO_ID = 27084 "
   m_Conn.Execute (SQL)
      
   SQL = "ALTER TABLE RECEIPT_ITEM ADD CONSTRAINT RECEIPT_ITEM_BD_ID_FK FOREIGN KEY (BILLING_DOC_ID) REFERENCES BILLING_DOC "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2008_05_29_1_jill = True
   Exit Function
   
ErrorHandler:
      ''debug.print m_Conn.Errors(0)
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2009_06_18_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2009_06_18_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILLING_DOC ADD CLOSE_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2009_06_18_1_jill = True
   Exit Function
   
ErrorHandler:
      ''debug.print m_Conn.Errors(0)
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2009_06_18_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2009_06_18_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE BILLING_DOC  SET CLOSE_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2009_06_18_2_jill = True
   Exit Function
   
ErrorHandler:
      ''debug.print m_Conn.Errors(0)
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2009_07_24_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2009_07_24_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILLING_DOC ADD VALID_DATE DATE_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2009_07_24_1_jill = True
   Exit Function
   
ErrorHandler:
      ''debug.print m_Conn.Errors(0)
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2009_07_24_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2009_07_24_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE BILLING_DOC DROP CLOSE_FLAG "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2009_07_24_2_jill = True
   Exit Function
   
ErrorHandler:
      ''debug.print m_Conn.Errors(0)
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2009_07_24_3_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2009_07_24_3_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE RECEIPT_ITEM ADD VALID_DATE DATE_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2009_07_24_3_jill = True
   Exit Function
   
ErrorHandler:
      ''debug.print m_Conn.Errors(0)
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2013_08_01_1_ging() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2013_08_01_1_ging"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
  SQL = "CREATE TABLE EXPOSE_TYPE   "
   SQL = SQL & "( "
   SQL = SQL & " EXPOSE_TYPE_ID ID_TYPE NOT NULL,  "
   SQL = SQL & "EXPOSE_TYPE_NO CODE_TYPE, "
   SQL = SQL & "EXPOSE_TYPE_NAME CODE_TYPE, "
   SQL = SQL & "CREATE_DATE DATE_TYPE, "
   SQL = SQL & "CREATE_BY ID_TYPE, "
   SQL = SQL & "MODIFY_DATE DATE_TYPE, "
   SQL = SQL & "MODIFY_BY ID_TYPE "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EXPOSE_TYPE ADD PRIMARY KEY (EXPOSE_TYPE_ID)"
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR EXPOSE_TYPE_SEQ"
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EXPORT_ITEM ADD MFG_DATE DATE_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE EXPORT_ITEM ADD EXP_DATE DATE_TYPE "
   m_Conn.Execute (SQL)
   
  SQL = "ALTER TABLE EXPORT_ITEM ADD EXPOSE_TYPE_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE EXPORT_ITEM ADD CONSTRAINT EXPOSE_TYPE_ID_FK FOREIGN KEY (EXPOSE_TYPE_ID) REFERENCES EXPOSE_TYPE"
   m_Conn.Execute (SQL)
  
   SQL = "ALTER TABLE INVENTORY_DOC ADD SURETY_ID ID_TYPE "
   m_Conn.Execute (SQL)

   SQL = "ALTER TABLE INVENTORY_DOC ADD CONSTRAINT SURETY_ID_FK FOREIGN KEY (SURETY_ID) REFERENCES EMPLOYEE"
   m_Conn.Execute (SQL)

   m_Conn.BeginTrans
   HasBegin = True
      
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_08_01_1_ging = True
   Exit Function
   
ErrorHandler:
      
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2013_09_13_1_ging() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2013_09_13_1_ging"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE PART_ITEM ADD REGIST_NUM CODE_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PART_ITEM ADD DRUG_SIZE DESC_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PART_ITEM ADD DRUG_NAME DESC_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PART_ITEM ADD SUPPLIER_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_09_13_1_ging = True
   Exit Function
   
ErrorHandler:
      ''debug.print m_Conn.Errors(0)
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2013_11_04_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2013_11_04_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE IMPORT_ITEM ADD IMPORT_DESC DESC_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2013_11_04_1_jill = True
   Exit Function
   
ErrorHandler:
      ''debug.print m_Conn.Errors(0)
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2014_02_04_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2014_02_04_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE INVENTORY_DOC ADD PO_NO DESC_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2014_02_04_1_jill = True
   Exit Function
   
ErrorHandler:
      ''debug.print m_Conn.Errors(0)
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2017_07_20_dear() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_07_20_dear"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE RECEIPT_ITEM ADD DESCRIPTION_DETAIL DESC_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_07_20_dear = True
   Exit Function
   
ErrorHandler:
      ''debug.print m_Conn.Errors(0)
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_08_30_1_dear() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_08_30_1_dear"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = "CREATE TABLE  AUTHEN_PO "
   SQL = SQL & "( "
   SQL = SQL & "AUTHEN_PO_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "AUTHEN_PO_GROUP           ID_TYPE , "
   SQL = SQL & "AUTHEN_PO_FROM           MONEY_TYPE , "
   SQL = SQL & "AUTHEN_PO_TO           MONEY_TYPE , "
   SQL = SQL & "AUTHEN_PO_DESC           DESC_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   
   SQL = "ALTER TABLE AUTHEN_PO ADD CONSTRAINT AUTHEN_PO_ID_PK PRIMARY KEY (AUTHEN_PO_ID) "
   m_Conn.Execute (SQL)
   
'   SQL = " ALTER TABLE CASH_DOC_POST ADD CONSTRAINT CASH_DOC_POST_CASH_DOC_ID_FK FOREIGN KEY (CASH_DOC_ID) REFERENCES CASH_DOC  "
'   m_Conn.Execute (SQL)
'
'   SQL = " ALTER TABLE CASH_DOC_POST ADD CONSTRAINT CASH_DOC_POST_BANK_ID_FK FOREIGN KEY (BANK_ID) REFERENCES BANK  "
'   m_Conn.Execute (SQL)
'
'   SQL = " ALTER TABLE CASH_DOC_POST ADD CONSTRAINT CASH_DOC_POST_BBRANCH_FK FOREIGN KEY (BANK_BRANCH) REFERENCES BANK_BRANCH  "
'   m_Conn.Execute (SQL)
'
'   SQL = " ALTER TABLE CASH_DOC_POST ADD CONSTRAINT CASH_DOC_POST_CHEQUE_FK FOREIGN KEY (CHEQUE_ID) REFERENCES CHEQUE  "
'   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR AUTHEN_PO_SEQ "
   m_Conn.Execute (SQL)
      
    m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_08_30_1_dear = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function

Public Function Patch_2017_08_30_2_dear() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_08_30_2_dear"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
  
   SQL = "CREATE TABLE  AUTHEN_DETAIL "
   SQL = SQL & "( "
   SQL = SQL & "AUTHEN_DETAIL_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "AUTHEN_PO_ID           ID_TYPE , "
   SQL = SQL & "AUTHEN_AREA          ID_TYPE , "
   SQL = SQL & "AUTHEN_USER         CODE_TYPE , "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   
   SQL = "ALTER TABLE AUTHEN_DETAIL ADD CONSTRAINT AUTHEN_DETAIL_ID_PK PRIMARY KEY (AUTHEN_DETAIL_ID) "
   m_Conn.Execute (SQL)
   
   SQL = " ALTER TABLE AUTHEN_DETAIL ADD CONSTRAINT AUTHEN_PO_ID_FK FOREIGN KEY (AUTHEN_PO_ID) REFERENCES AUTHEN_PO  "
   m_Conn.Execute (SQL)
'
'   SQL = " ALTER TABLE CASH_DOC_POST ADD CONSTRAINT CASH_DOC_POST_BANK_ID_FK FOREIGN KEY (BANK_ID) REFERENCES BANK  "
'   m_Conn.Execute (SQL)
'
'   SQL = " ALTER TABLE CASH_DOC_POST ADD CONSTRAINT CASH_DOC_POST_BBRANCH_FK FOREIGN KEY (BANK_BRANCH) REFERENCES BANK_BRANCH  "
'   m_Conn.Execute (SQL)
'
'   SQL = " ALTER TABLE CASH_DOC_POST ADD CONSTRAINT CASH_DOC_POST_CHEQUE_FK FOREIGN KEY (CHEQUE_ID) REFERENCES CHEQUE  "
'   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR AUTHEN_DETAIL_SEQ "
   m_Conn.Execute (SQL)
      
    m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_08_30_2_dear = True
   Exit Function
   
ErrorHandler:
   If HasBegin Then
      m_Conn.RollbackTrans
   End If
End Function
Public Function Patch_2017_08_31_1_dear() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_08_31_1_dear"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER AUTHEN_PO_BA FOR AUTHEN_PO "
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.AUTHEN_PO_ID  = GEN_ID (AUTHEN_PO_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_08_31_1_dear = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2017_08_31_2_dear() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_08_31_2_dear"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "                 CREATE TRIGGER AUTHEN_DETAIL_BA FOR AUTHEN_DETAIL "
   SQL = SQL & "  ACTIVE BEFORE INSERT "
   SQL = SQL & "  POSITION 0 "
   SQL = SQL & "  AS "
   SQL = SQL & "  BEGIN "
   SQL = SQL & "        NEW.AUTHEN_DETAIL_ID  = GEN_ID (AUTHEN_DETAIL_SEQ, 1); "
   SQL = SQL & "  END; "
   m_Conn.Execute (SQL)
      
   m_Conn.BeginTrans
   HasBegin = True
   
   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_08_31_2_dear = True
   Exit Function
   
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
      
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2017_08_31_3_dear() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_08_31_3_dear"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
SQL = "ALTER TABLE USER_ACCOUNT ADD REAL_NAME DESC_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_08_31_3_dear = True
   Exit Function
ErrorHandler:
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1" & m_Conn.Errors(0)
      glbErrorLog.ShowUserError
   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
'Public Function Patch_2017_11_03_1_jill() As Boolean
'On Error GoTo ErrorHandler
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'
'   VersionStr = "2017_11_03_1_jill"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = "ALTER TABLE DO_ITEM ADD PIG_AGE_DAY ID_TYPE "
'   m_Conn.Execute (SQL)
'
'   m_Conn.BeginTrans
'   HasBegin = True
'
'   SQL = "INSERT INTO PATCH_DELIVERY "
'   SQL = SQL & "( "
'   SQL = SQL & "PATCH_DELIVERY_ID, "
'   SQL = SQL & "PATCH_VERSION, "
'   SQL = SQL & "PATCH_DATE, "
'   SQL = SQL & "PATCH_BY "
'   SQL = SQL & ") VALUES "
'   SQL = SQL & "( "
'   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
'   SQL = SQL & "'" & VersionStr & "', "
'   SQL = SQL & "'" & ServerDateTime & "', "
'   SQL = SQL & "" & glbUser.USER_ID & " "
'   SQL = SQL & ")"
'   m_Conn.Execute (SQL)
'
'   m_Conn.CommitTrans
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2017_11_03_1_jill = True
'   Exit Function
'
'ErrorHandler:
'      ''debug.print m_Conn.Errors(0)
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
'      glbErrorLog.ShowUserError
'
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function
'Public Function Patch_2017_11_03_2_jill() As Boolean
'On Error GoTo ErrorHandler
'Dim SQL As String
'Dim ServerDateTime As String
'Dim HasBegin As Boolean
'Dim VersionStr As String
'
'   VersionStr = "2017_11_03_2_jill"
'   HasBegin = False
'   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
'
'   SQL = "ALTER TABLE EXPORT_ITEM ADD PIG_AGE_DAY ID_TYPE "
'   m_Conn.Execute (SQL)
'
'   m_Conn.BeginTrans
'   HasBegin = True
'
'   SQL = "INSERT INTO PATCH_DELIVERY "
'   SQL = SQL & "( "
'   SQL = SQL & "PATCH_DELIVERY_ID, "
'   SQL = SQL & "PATCH_VERSION, "
'   SQL = SQL & "PATCH_DATE, "
'   SQL = SQL & "PATCH_BY "
'   SQL = SQL & ") VALUES "
'   SQL = SQL & "( "
'   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
'   SQL = SQL & "'" & VersionStr & "', "
'   SQL = SQL & "'" & ServerDateTime & "', "
'   SQL = SQL & "" & glbUser.USER_ID & " "
'   SQL = SQL & ")"
'   m_Conn.Execute (SQL)
'
'   m_Conn.CommitTrans
'   HasBegin = False
'
'   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
'   glbErrorLog.ShowUserError
'
'   Patch_2017_11_03_2_jill = True
'   Exit Function
'
'ErrorHandler:
'      ''debug.print m_Conn.Errors(0)
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
'      glbErrorLog.ShowUserError
'
'   If HasBegin Then
'      m_Conn.RollbackTrans
'      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
'      glbErrorLog.ShowUserError
'   End If
'End Function

Public Function Patch_2017_12_28_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2017_12_28_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE LOCATION ADD LOCK_WEIGHT_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PRODUCT_TYPE ADD LOCK_WEIGHT_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE PRODUCT_STATUS ADD LOCK_WEIGHT_FLAG FLAG_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_12_28_1_jill = True
   Exit Function

ErrorHandler:
      ''debug.print m_Conn.Errors(0)
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2017_12_28_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String

   VersionStr = "2017_12_28_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   SQL = "UPDATE LOCATION SET LOCK_WEIGHT_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE PRODUCT_TYPE SET LOCK_WEIGHT_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   SQL = "UPDATE PRODUCT_STATUS SET LOCK_WEIGHT_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)
   
   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_12_28_2_jill = True
   Exit Function

ErrorHandler:
      ''debug.print m_Conn.Errors(0)
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2018_05_30_1_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2018_05_30_1_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE CUSTOMER ADD TAKE_AWAY_FLAG FLAG_TYPE DEFAULT 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2018_05_30_1_jill = True
   Exit Function

ErrorHandler:
      ''debug.print m_Conn.Errors(0)
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
Public Function Patch_2018_05_30_2_jill() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2018_05_30_2_jill"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE CUSTOMER SET TAKE_AWAY_FLAG = 'N' "
   m_Conn.Execute (SQL)
   
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2018_05_30_2_jill = True
   Exit Function

ErrorHandler:
      ''debug.print m_Conn.Errors(0)
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2018_08_16_1_dear() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2018_08_16_1_dear"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "ALTER TABLE CUSTOMER ADD TAKE_KEY_ACCOUNT FLAG_TYPE DEFAULT 'N' "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2018_08_16_1_dear = True
   Exit Function

ErrorHandler:
      ''debug.print m_Conn.Errors(0)
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function

Public Function Patch_2018_08_16_2_dear() As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
   
   VersionStr = "2018_08_16_2_dear"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   SQL = "UPDATE CUSTOMER SET TAKE_KEY_ACCOUNT = 'N' "
   m_Conn.Execute (SQL)
   
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False

   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2018_08_16_2_dear = True
   Exit Function

ErrorHandler:
      ''debug.print m_Conn.Errors(0)
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.1"
      glbErrorLog.ShowUserError

   If HasBegin Then
      m_Conn.RollbackTrans
      glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " Error.2"
      glbErrorLog.ShowUserError
   End If
End Function
