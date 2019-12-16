Attribute VB_Name = "modMain"
Option Explicit
'"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\somepath\mydb.mdb;User Id=admin;Password=;"

Public Const ROOT_TREE = "Root"

Public Const DUMMY_KEY = 27
Public Const DUMMY_BALANCE_DO_ID = 99999
Public GLB_GRID_COLOR As Long
Public GLB_NORMAL_COLOR As Long
Public GLB_ALERT_COLOR As Long
Public GLB_SHOW_COLOR As Long
Public GLB_FORM_COLOR As Long
Public GLB_HEAD_COLOR As Long
Public GLB_GRIDHD_COLOR As Long
Public GLB_MANDATORY_COLOR As Long

Public Enum FIELD_TYPE
   INT_TYPE = 1
   MONEY_TYPE = 2
   DATE_TYPE = 3
   STRING_TYPE = 4
   BOOLEAN_TYPE = 5
End Enum

Public Enum FIELD_CAT
   ID_CAT = 1
   MODIFY_DATE_CAT = 2
   CREATE_DATE_CAT = 3
   MODIFY_BY_CAT = 4
   CREATE_BY_CAT = 5
   DATA_CAT = 6
   TEMP_CAT = 7
End Enum

Public Enum SHOW_MODE_TYPE
   SHOW_ADD = 1
   SHOW_EDIT = 2
   SHOW_VIEW = 3
   SHOW_VIEW_ONLY = 4
End Enum

Public Enum RESOURCE_TYPE
   HOTEL_RESOURCE = 1
End Enum

Public Enum TEXT_BOX_TYPE
   TEXT_STRING = 1
   TEXT_INTEGER = 2
   TEXT_FLOAT = 3
   TEXT_FLOAT_MONEY = 4
   TEXT_INTEGER_MONEY = 5
End Enum

Public Enum MASTER_TYPE
   CHEQUE_TYPE = 1
   FEED_GROUP = 2
   MEMO_TYPE = 6
   MEMO_STATUS = 7
End Enum

Public Enum CONFIG_DOC_TYPE
   SELL_SO = 1
   SELL_RETURN = 2
   
   BUY_PO_RAW = 11
   BUY_PO_MATERIAL = 12
   BUY_PO_EXPENSE = 13
   BUY_PO_GENERAL = 14
   
   BUY_PO_RAW_AUTO = 21
   BUY_PO_MATERIAL_AUTO = 22
   BUY_PO_EXPENSE_AUTO = 23
   BUY_PO_GENERAL_AUTO = 24
   
   BUY_RO_RAW = 31
   BUY_RO_MATERIAL = 32
   BUY_RO_EXPENSE = 33
   BUY_RO_GENERAL = 34
   
End Enum

Public Enum LANGUAGE_TYPE
   LANG_ENG = 1
   LANG_THAI = 2
End Enum

Public Enum PAYMENT_TYPE
   CASH_PMT = 1
   CREDITCRD_PMT = 2
   CHECK_PMT = 3
   BANKTRF_PMT = 4
   CASHRET_PMT = 5
End Enum

Public Enum CASH_DOC_TYPE
   CASH_TRANSFER = 1
   CASH_DEPOSIT = 2
   CASH_WITHDRAW = 3
   CASH_PITTYCASH = 4
   CASH_WHTHDRAW2 = 5
   CASH_DEPOSIT2 = 6
   POST_CHEQUE = 7
End Enum

Public Enum UNIQUE_TYPE
   EMPCODE_UNIQUE = 1
   EMPNAME_LASTNAME_UNIQUE = 2
   TRUCK_UNIQUE = 3
   DO_PLAN_UNIQUE = 4
   DBN_UNIQUE = 5
   CUSTCODE_UNIQUE = 6
   USERGROUP_UNIQUE = 7
   USERNAME_UNIQUE = 8
   IMPORT_UNIQUE = 9
   EXPORT_UNIQUE = 10
   REPAIR_UNIQUE = 11
   REPAIR_FORMULA_UNIQUE = 12
   SUPPLIER_UNIQUE = 13
   PARTNO_UNIQUE = 14
   QUOATATION_UNIQUE = 15
   TEACHER_UNIQUE = 16
   SUBJECT_UNIQUE = 17
   FACULTY_UNIQUE = 18
   EXPENSE_UNIQUE = 19
   PO_UNIQUE = 20
   CUSTOMER_UNIQUE = 21
   REVENUE_UNIQUE = 22
   BORROW_UNIQUE = 23
   PRDFEATURE_UNIQUE = 24
   JOBPLAN_UNIQUE = 25
   
   PARTTYPE_NO = 26
   PARTTYPE_NAME = 27
   LOCATION_NO = 28
   LOCATION_NAME = 29
   PRODUCTTYPE_NO = 30
   PRODUCTTYPE_NAME = 31
   PRODUCTSTATUS_NO = 32
   PRODUCTSTATUS_NAME = 33
   HOUSE_NO = 34
   HOUSE_NAME = 35
   COUNTRY_NO = 36
   COUNTRY_NAME = 37
   CSTGRADE_NO = 38
   CSTGRADE_NAME = 39
   CSTTYPE_NO = 40
   CSTTYPE_NAME = 41
   SUPPLIERTYPE_NO = 42
   SUPPLIERYPE_NAME = 43
   SUPPLIERGRADE_NO = 44
   SUPPLIERGRADE_NAME = 45
   SUPPLIERSTATUS_NO = 46
   SUPPLIERSTATUS_NAME = 47
   POSITION_NO = 48
   UNIT_NO = 49
   UNIT_NAME = 50
   YEAR_NO = 51
   PARTGROUP_NO = 52
   PARTGROUP_NAME = 53
   LOCATION_NO_EX = 54
   
   PACKAGE_CODE = 55
   PACKAGE_NAME = 56
   PACKAGE_BASIC = 57
   PRICE_ADJUST = 58
    
   EXPOSE_TYPE_NO = 59
   EXPOSE_TYPE_NAME = 60
End Enum

Public Enum NUMBER_TYPE
   PO_NUMBER = 1
   OPERATE_NUMBER = 2
   BORROW_NUMBER = 3
   DEBIT_NOTE_NUMBER = 4
   'bum+
   EXPENSE_NUMBER = 5
   REPAIR_NUMBER = 6
   IMPORT_NUMBER = 7
   EXPORT_NUMBER = 8
   PLAN_NUMBER = 9
   FUEL_NUMBER1 = 10
   FUEL_NUMBER2 = 11
   BILL_NUMBER = 13
   QUOATATION_NUMBER = 14
   REVENUE_NUMBER = 15
   DO_NUMBER = 16
   RECEIPT_NUMBER = 17
   JOBPLAN_NUMBER = 18
   INVOICE_RECEIPT_NUMBER = 19
   DBN_NUMBER = 20
   CDN_NUMBER = 21
   
End Enum

'===================== For clear treeview =========================
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd _
    As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const TV_FIRST As Long = &H1100
Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Const TVM_DELETEITEM As Long = (TV_FIRST + 1)
Const TVGN_ROOT As Long = &H0
Const WM_SETREDRAW As Long = &HB
'===================== For clear treeview =========================

Public Const PROJECT_NAME = "Mittraphap Farm Management"
Public Const GLB_FONT = "JasmineUPC"
Private Const MODULE_NAME = "modMain"

Public glbErrorLog As clsErrorLog
Public glbDatabaseMngr As clsDatabaseMngr
Public glbSetting As clsGlobalSetting
Public glbParameterObj As clsParameter
Public glbUser As CUser
Public glbGroup As CGroup
Public glbAdmin As clsAdmin
Public glbMaster As clsMaster
Public glbDaily As clsDaily
Public glbLegacy As clsLegacy
Public glbEnterPrise As CEnterprise
Public glbAuthenPO As clsAuthenPO

Public CustomerPackage As Collection
Public PackageDetail As Collection
Public T706Collection1 As Collection
Public T706Collection2 As Collection

'Public glbPricePlan As clsPricePlan
'Public glbHR As clsHR
'Public glbInventory As clsInventory
'Public glbLedger As clsLedger
Public glbLoginTracking As CLoginTracking
'Public glbSystemParam As Collection
Public glbAccessRight As Collection

Public Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function VerifyDate(L As Label, D As uctlDate, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   If L Is Nothing Then
      S = ""
   Else
      S = L.Caption
   End If

   If Not D.VerifyDate(NullAllow) Then
      VerifyDate = False
      D.SetFocus
      Call MsgBox("��سҡ�͡������ " & " '" & S & "' " & "���١��ͧ��Фú��ǹ ", vbOKOnly, PROJECT_NAME)
   Else
      VerifyDate = True
   End If
End Function

Public Function VerifyDateInterval(D As Date) As Boolean
Static TempCol As Collection
Dim Sp As CSystemParam
Dim TempStr As String

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadSystemParam(Nothing, TempCol)
   End If
   
   Set Sp = GetSystemParam(TempCol, "DOC_LOCKDATE")
   If (Sp.GetFieldValue("FROM_LOCK_DATE") <= D) And _
      (Sp.GetFieldValue("TO_LOCK_DATE")) >= D Then
      VerifyDateInterval = True
   Else
      VerifyDateInterval = False
      TempStr = "�ѹ����ͧ���������ҧ " & DateToStringExtEx2(Sp.GetFieldValue("FROM_LOCK_DATE")) & " - " & DateToStringExtEx2(Sp.GetFieldValue("TO_LOCK_DATE"))
      glbErrorLog.LocalErrorMsg = TempStr
      Call glbErrorLog.ShowUserError
   End If
End Function

Public Function VerifyTime(L As Label, T As uctlTime, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   If L Is Nothing Then
      S = ""
   Else
      S = L.Caption
   End If

   If Not T.VerifyTime(NullAllow) Then
      VerifyTime = False
      T.SetFocus
      Call MsgBox("��سҡ�͡������ " & " '" & S & "' " & "���١��ͧ��Фú��ǹ ", vbOKOnly, PROJECT_NAME)
   Else
      VerifyTime = True
   End If
End Function

Public Function VerifyTextData(L As Label, T As TextBox, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   If L Is Nothing Then
      S = ""
   Else
      S = L.Caption
   End If
   
   If Not NullAllow Then
      If Len(Trim(T.Text)) = 0 Then
         VerifyTextData = False
         Call MsgBox("��سҡ�͡������ " & " '" & S & "' " & "���١��ͧ��Фú��ǹ ", vbOKOnly, PROJECT_NAME)
         If T.Enabled Then
            T.SetFocus
         End If
         Exit Function
      End If
   End If
   
   If (T.Tag = TEXT_INTEGER) Or (T.Tag = TEXT_FLOAT) Or (T.Tag = TEXT_FLOAT_MONEY) Or (T.Tag = TEXT_INTEGER_MONEY) Then
      If Trim(T.Text) = "" Then
         If NullAllow Then
            VerifyTextData = True
            Exit Function
         End If
      End If
      If IsNumeric(Trim(T.Text)) Then
         If InStr(1, T.Text, ".") <= 0 Then
            If Val(Trim(T.Text)) < 0 Then
               VerifyTextData = False
            Else
               VerifyTextData = True
               Exit Function
            End If
         Else
            If T.Tag = TEXT_INTEGER Then
               VerifyTextData = False
            Else
               If Val(Trim(T.Text)) < 0 Then
                  VerifyTextData = False
               Else
                  VerifyTextData = True
               End If
            End If
            Exit Function
         End If
      End If
      
      VerifyTextData = False
      Call MsgBox("��سҡ�͡������ " & " '" & S & "' " & "���١��ͧ��Фú��ǹ ", vbOKOnly, PROJECT_NAME)
      If T.Enabled Then
         T.SetFocus
      End If
      Exit Function
   ElseIf T.Tag = TEXT_STRING Then
      If (InStr(1, T.Text, ";") > 0) Or (InStr(1, T.Text, "|") > 0) Then
         Call MsgBox("��سҡ�͡������ " & " '" & S & "' " & "���١��ͧ��Фú��ǹ ", vbOKOnly, PROJECT_NAME)
         T.SetFocus
         
         VerifyTextData = False
         Exit Function
      End If
      
      VerifyTextData = True
   End If
End Function

Public Function VerifyTextControl(L As Label, T As uctlTextBox, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   If L Is Nothing Then
      S = ""
   Else
      S = L.Caption
   End If
   
   If Not NullAllow Then
      If Len(Trim(T.Text)) = 0 Then
         VerifyTextControl = False
         Call MsgBox("��سҡ�͡������ " & " '" & S & "' " & "���١��ͧ��Фú��ǹ ", vbOKOnly, PROJECT_NAME)
         If T.Enabled Then
            T.SetFocus
         End If
         Exit Function
      End If
   End If
   
   If (T.Tag = TEXT_INTEGER) Or (T.Tag = TEXT_FLOAT) Or (T.Tag = TEXT_FLOAT_MONEY) Or (T.Tag = TEXT_INTEGER_MONEY) Then
      If Trim(T.Text) = "" Then
         If NullAllow Then
            VerifyTextControl = True
            Exit Function
         End If
      End If
      If IsNumeric(Trim(T.Text)) Then
         If InStr(1, T.Text, ".") <= 0 Then
            If Val(Trim(T.Text)) < 0 Then
               VerifyTextControl = True 'false
               Exit Function 'remove this if false
            Else
               VerifyTextControl = True
               Exit Function
            End If
         Else
            If T.Tag = TEXT_INTEGER Then
               VerifyTextControl = False
            Else
               If Val(Trim(T.Text)) < 0 Then
                  VerifyTextControl = True 'false
                  Exit Function
               Else
                  VerifyTextControl = True
                  Exit Function
               End If
            End If
'            Exit Function
         End If
      End If
      
      VerifyTextControl = False
      Call MsgBox("��سҡ�͡������ " & " '" & S & "' " & "���١��ͧ��Фú��ǹ ", vbOKOnly, PROJECT_NAME)
      If T.Enabled Then
         T.SetFocus
      End If
      Exit Function
   ElseIf T.Tag = TEXT_STRING Then
      If (InStr(1, T.Text, ";") > 0) Or (InStr(1, T.Text, "|") > 0) Then
         Call MsgBox("��سҡ�͡������ " & " '" & S & "' " & "���١��ͧ��Фú��ǹ ", vbOKOnly, PROJECT_NAME)
         T.SetFocus
         
         VerifyTextControl = False
         Exit Function
      End If
      
      VerifyTextControl = True
   End If
End Function

Private Sub GetParentItemDesc(Acc As String, Ri As CRightItem, ReportName As String)
   Ri.DEFAULT_VALUE = "N"
   If Acc = "ADMIN" Then
      Ri.RIGHT_ITEM_DESC = "�к������ż����ҹ"
   ElseIf Acc = "ADMIN_GROUP" Then
      Ri.RIGHT_ITEM_DESC = "�к������š���������ҹ"
   ElseIf Acc = "ADMIN_GROUP_ADD" Then
      Ri.RIGHT_ITEM_DESC = "����ö���������š���������ҹ"
   ElseIf Acc = "ADMIN_GROUP_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "����ö��䢢����š���������ҹ"
   ElseIf Acc = "ADMIN_GROUP_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "����öź�����š���������ҹ"
   
   ElseIf Acc = "ADMIN_USER" Then
      Ri.RIGHT_ITEM_DESC = "�к������ż����ҹ"
   ElseIf Acc = "ADMIN_USER_ADD" Then
      Ri.RIGHT_ITEM_DESC = "����ö���������ż����ҹ"
   ElseIf Acc = "ADMIN_USER_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "����ö��䢢����ż����ҹ"
   ElseIf Acc = "ADMIN_USER_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "����öź�����ż����ҹ"
   
   ElseIf Acc = "ADMIN_REPORT" Then
      Ri.RIGHT_ITEM_DESC = "��§ҹ�к������ż����ҹ"
   ElseIf Mid(Acc, 1, 12) = "ADMIN_REPORT" Then
      Ri.RIGHT_ITEM_DESC = ReportName
   
   ElseIf Acc = "MASTER" Then
      Ri.RIGHT_ITEM_DESC = "�к���������ѡ"
      
   ElseIf Acc = "MASTER_MAIN" Then
      Ri.RIGHT_ITEM_DESC = "�к���������ѡ��ǹ��ҧ"
   ElseIf Acc = "MASTER_MAIN_ADD" Then
      Ri.RIGHT_ITEM_DESC = "����ö������������ѡ��ǹ��ҧ"
   ElseIf Acc = "MASTER_MAIN_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "����ö��䢢�������ѡ��ǹ��ҧ"
   ElseIf Acc = "MASTER_MAIN_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "����öź��������ѡ��ǹ��ҧ"
      
   ElseIf Acc = "MASTER_INVENTORY" Then
      Ri.RIGHT_ITEM_DESC = "�к���������ѡ��ѧ"
   ElseIf Acc = "MASTER_INVENTORY_ADD" Then
      Ri.RIGHT_ITEM_DESC = "����ö������������ѡ��ѧ"
   ElseIf Acc = "MASTER_INVENTORY_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "����ö��䢢�������ѡ��ѧ"
   ElseIf Acc = "MASTER_INVENTORY_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "����öź��������ѡ��ѧ"
      
   ElseIf Acc = "MASTER_PIG" Then
      Ri.RIGHT_ITEM_DESC = "�к���������ѡ�к��������ء�"
   ElseIf Acc = "MASTER_PIG_ADD" Then
      Ri.RIGHT_ITEM_DESC = "����ö������������ѡ�к��������ء�"
   ElseIf Acc = "MASTER_PIG_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "����ö��䢢�������ѡ�к��������ء�"
   ElseIf Acc = "MASTER_PIG_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "����öź��������ѡ�к��������ء�"
      
   ElseIf Acc = "MASTER_LEDGER" Then
      Ri.RIGHT_ITEM_DESC = "�к���������ѡ�к��ѭ��"
   ElseIf Acc = "MASTER_LEDGER_ADD" Then
      Ri.RIGHT_ITEM_DESC = "����ö������������ѡ�к��ѭ��"
   ElseIf Acc = "MASTER_LEDGER_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "����ö��䢢�������ѡ�к��ѭ��"
   ElseIf Acc = "MASTER_LEDGER_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "����öź��������ѡ�к��ѭ��"
      
   ElseIf Acc = "MASTER_PACKAGE" Then
      Ri.RIGHT_ITEM_DESC = "�к���������ѡ��õ���Ҥ��Թ���"
   ElseIf Acc = "MASTER_PACKAGE_ADD" Then
      Ri.RIGHT_ITEM_DESC = "����ö������������ѡ��õ���Ҥ��Թ���"
   ElseIf Acc = "MASTER_PACKAGE_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "����ö��䢢�������ѡ��õ���Ҥ��Թ���"
   ElseIf Acc = "MASTER_PACKAGE_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "����öź��������ѡ��õ���Ҥ��Թ���"
      
   ElseIf Acc = "MAIN" Then
      Ri.RIGHT_ITEM_DESC = "�к���������ǹ��ҧ"
   
   ElseIf Acc = "MAIN_CUSTOMER" Then
      Ri.RIGHT_ITEM_DESC = "�к��������١���"
   ElseIf Acc = "MAIN_CUSTOMER_ADD" Then
      Ri.RIGHT_ITEM_DESC = "����ö�����������١���"
   ElseIf Acc = "MAIN_CUSTOMER_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "����ö��䢢������١���"
   ElseIf Acc = "MAIN_CUSTOMER_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "����öź�������١���"
      
   ElseIf Acc = "MAIN_SUPPLIER" Then
      Ri.RIGHT_ITEM_DESC = "�к������ūѾ���������"
   ElseIf Acc = "MAIN_SUPPLIER_ADD" Then
      Ri.RIGHT_ITEM_DESC = "����ö���������ūѾ���������"
   ElseIf Acc = "MAIN_SUPPLIER_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "����ö��䢢����ūѾ���������"
   ElseIf Acc = "MAIN_SUPPLIER_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "����öź�����ūѾ���������"
      
   ElseIf Acc = "MAIN_ENTERPRISE" Then
      Ri.RIGHT_ITEM_DESC = "�к������ź���ѷ"
   ElseIf Acc = "MAIN_ENTERPRISE_ADD" Then
      Ri.RIGHT_ITEM_DESC = "����ö���������ź���ѷ"
   ElseIf Acc = "MAIN_ENTERPRISE_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "����ö��䢢����ź���ѷ"
   ElseIf Acc = "MAIN_ENTERPRISE_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "����öź�����ź���ѷ"

   ElseIf Acc = "MAIN_EMPLOYEE" Then
      Ri.RIGHT_ITEM_DESC = "�к������ž�ѡ�ҹ"
   ElseIf Acc = "MAIN_EMPLOYEE_ADD" Then
      Ri.RIGHT_ITEM_DESC = "����ö���������ž�ѡ�ҹ"
   ElseIf Acc = "MAIN_EMPLOYEE_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "����ö��䢢����ž�ѡ�ҹ"
   ElseIf Acc = "MAIN_EMPLOYEE_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "����öź�����ž�ѡ�ҹ"
      
   ElseIf Acc = "MAIN_REPORT" Then
      Ri.RIGHT_ITEM_DESC = "��§ҹ�к���������ǹ��ҧ"
   ElseIf Mid(Acc, 1, 11) = "MAIN_REPORT" Then
      Ri.RIGHT_ITEM_DESC = ReportName
      
   ElseIf Acc = "INVENTORY" Then
      Ri.RIGHT_ITEM_DESC = "�к������Ť�ѧ"
   
   ElseIf Acc = "INVENTORY_PART" Then
      Ri.RIGHT_ITEM_DESC = "�к��������ѵ�شԺ"
   ElseIf Acc = "INVENTORY_PART_ADD" Then
      Ri.RIGHT_ITEM_DESC = "����ö�����������ѵ�شԺ"
   ElseIf Acc = "INVENTORY_PART_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "����ö��䢢������ѵ�شԺ"
   ElseIf Acc = "INVENTORY_PART_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "����öź�������ѵ�شԺ"
   
   ElseIf Acc = "INVENTORY_IMPORT" Then
      Ri.RIGHT_ITEM_DESC = "�к������š�ù�����ѵ�شԺ"
   ElseIf Acc = "INVENTORY_IMPORT_ADD" Then
      Ri.RIGHT_ITEM_DESC = "����ö���������š�ù�����ѵ�شԺ"
   ElseIf Acc = "INVENTORY_IMPORT_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "����ö��䢢����š�ù�����ѵ�شԺ"
   ElseIf Acc = "INVENTORY_IMPORT_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "����öź�����š�ù�����ѵ�شԺ"
   
   ElseIf Acc = "INVENTORY_EXPORT" Then
      Ri.RIGHT_ITEM_DESC = "�к������š���ԡ�����ѵ�شԺ"
   ElseIf Acc = "INVENTORY_EXPORT_ADD" Then
      Ri.RIGHT_ITEM_DESC = "����ö���������š���ԡ�����ѵ�شԺ"
   ElseIf Acc = "INVENTORY_EXPORT_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "����ö��䢢����š���ԡ�����ѵ�شԺ"
   ElseIf Acc = "INVENTORY_EXPORT_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "����öź�����š���ԡ�����ѵ�شԺ"
   
   ElseIf Acc = "INVENTORY_TRANSFER" Then
      Ri.RIGHT_ITEM_DESC = "�к������š���͹�ѵ�شԺ"
   ElseIf Acc = "INVENTORY_TRANSFER_ADD" Then
      Ri.RIGHT_ITEM_DESC = "����ö���������š���͹�ѵ�شԺ"
   ElseIf Acc = "INVENTORY_TRANSFER_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "����ö��䢢����š���͹�ѵ�شԺ"
   ElseIf Acc = "INVENTORY_TRANSFER_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "����öź�����š���͹�ѵ�شԺ"
   
   ElseIf Acc = "INVENTORY_ADJUST" Then
      Ri.RIGHT_ITEM_DESC = "�к������š�û�Ѻ�ʹ�ѵ�شԺ"
   ElseIf Acc = "INVENTORY_ADJUST_ADD" Then
      Ri.RIGHT_ITEM_DESC = "����ö���������š�û�Ѻ�ʹ�ѵ�شԺ"
   ElseIf Acc = "INVENTORY_ADJUST_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "����ö��䢢����š�û�Ѻ�ʹ�ѵ�شԺ"
   ElseIf Acc = "INVENTORY_ADJUST_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "����öź�����š�û�Ѻ�ʹ�ѵ�شԺ"
   
   ElseIf Acc = "INVENTORY_REPORT" Then
      Ri.RIGHT_ITEM_DESC = "��§ҹ�к������Ť�ѧ"
   ElseIf Mid(Acc, 1, 16) = "INVENTORY_REPORT" Then
      Ri.RIGHT_ITEM_DESC = ReportName
   
   
   ElseIf Acc = "PIG" Then
      Ri.RIGHT_ITEM_DESC = "�к������ź������ء�"
   
   ElseIf Acc = "PIG_WEEK" Then
      Ri.RIGHT_ITEM_DESC = "�к��������ѻ�����Դ�ء�"
   ElseIf Acc = "PIG_WEEK_ADD" Then
      Ri.RIGHT_ITEM_DESC = "����ö�����������ѻ�����Դ�ء�"
   ElseIf Acc = "PIG_WEEK_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "����ö��䢢������ѻ�����Դ�ء�"
   ElseIf Acc = "PIG_WEEK_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "����öź�������ѻ�����Դ�ء�"
   
   ElseIf Acc = "PIG_IMPORT" Then
      Ri.RIGHT_ITEM_DESC = "�к������š�ù�����ء�"
   ElseIf Acc = "PIG_IMPORT_ADD" Then
      Ri.RIGHT_ITEM_DESC = "����ö���������š�ù�����ء�"
   ElseIf Acc = "PIG_IMPORT_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "����ö��䢢����š�ù�����ء�"
   ElseIf Acc = "PIG_IMPORT_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "����öź�����š�ù�����ء�"
   
   ElseIf Acc = "PIG_BIRTH" Then
      Ri.RIGHT_ITEM_DESC = "�к��������ءä�ʹ"
   ElseIf Acc = "PIG_BIRTH_ADD" Then
      Ri.RIGHT_ITEM_DESC = "����ö�����������ءä�ʹ"
   ElseIf Acc = "PIG_BIRTH_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "����ö��䢢������ءä�ʹ"
   ElseIf Acc = "PIG_BIRTH_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "����öź�������ءä�ʹ"
   
   ElseIf Acc = "PIG_TRANSFER" Then
      Ri.RIGHT_ITEM_DESC = "�к������š���͹�ء�"
   ElseIf Acc = "PIG_TRANSFER_ADD" Then
      Ri.RIGHT_ITEM_DESC = "����ö���������š���͹�ء�"
   ElseIf Acc = "PIG_TRANSFER_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "����ö��䢢����š���͹�ء�"
   ElseIf Acc = "PIG_TRANSFER_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "����öź�����š���͹�ء�"
   
   ElseIf Acc = "PIG_ADJUST" Then
      Ri.RIGHT_ITEM_DESC = "�к������š�û�Ѻ�ʹ�ء�"
   ElseIf Acc = "PIG_ADJUST_ADD" Then
      Ri.RIGHT_ITEM_DESC = "����ö���������š�û�Ѻ�ʹ�ء�"
   ElseIf Acc = "PIG_ADJUST_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "����ö��䢢����š�û�Ѻ�ʹ�ء�"
   ElseIf Acc = "PIG_ADJUST_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "����öź�����š�û�Ѻ�ʹ�ء�"
   
   ElseIf Acc = "PIG_REPORT" Then
      Ri.RIGHT_ITEM_DESC = "��§ҹ�к������ź������ء�"
   ElseIf Mid(Acc, 1, 10) = "PIG_REPORT" Then
      Ri.RIGHT_ITEM_DESC = ReportName
      
      
   ElseIf Acc = "PACKAGE" Then
      Ri.RIGHT_ITEM_DESC = "�к������š�õ���Ҥ��Թ���"
   ElseIf Acc = "PACKAGE_ADD" Then
      Ri.RIGHT_ITEM_DESC = "����ö���������š�õ���Ҥ��Թ���"
   ElseIf Acc = "PACKAGE_EDIT" Then
      Ri.RIGHT_ITEM_DESC = "����ö��䢢����š�õ���Ҥ��Թ���"
   ElseIf Acc = "PACKAGE_DELETE" Then
      Ri.RIGHT_ITEM_DESC = "����öź�����š�õ���Ҥ��Թ���"
   
      
   ElseIf Acc = "LEDGER" Then
      Ri.RIGHT_ITEM_DESC = "�к��ѭ��"
   ElseIf Acc = "LEDGER_SELL" Then
      Ri.RIGHT_ITEM_DESC = "�к��ѭ�բ��"
   ElseIf Acc = "LEDGER_COST" Then
      Ri.RIGHT_ITEM_DESC = "�к��鹷ع"
   ElseIf Acc = "LEDGER_CASH" Then
      Ri.RIGHT_ITEM_DESC = "�Թʴ˹�ҿ����"
   ElseIf Acc = "LEDGER_REPORT" Then
      Ri.RIGHT_ITEM_DESC = "��§ҹ�к��ѭ��"
   
   
   
   ElseIf Len(ReportName) > 0 Then
      Ri.RIGHT_ITEM_DESC = ReportName
   Else
      Ri.RIGHT_ITEM_DESC = ""
   End If
End Sub

Private Function GetParentKey(Acc As String, TopFlag As Boolean) As String
Dim I As Long
Dim j As Long

   For I = 1 To Len(Acc)
      If Mid(Acc, I, 1) = "_" Then
         j = I
      End If
   Next I
   
   If j > 1 Then
      GetParentKey = Mid(Acc, 1, j - 1)
      TopFlag = False
   Else
      GetParentKey = ""
      TopFlag = True
   End If
End Function

Public Function CreatePermissionNode(Acc As String, ParentID As Long, ReportName As String) As Boolean
Dim ParentKey As String
Dim TopFlag As Boolean
Dim TempParentID As Long
Dim CreateFlag As Boolean
Dim Ri As CRightItem
Dim TempRs As ADODB.Recordset
Dim iCount As Long
   
   'Create node here
   Set Ri = New CRightItem
   Set TempRs = New ADODB.Recordset
   TempParentID = 0
   
   Ri.RIGHT_ID = -1
   Ri.RIGHT_ITEM_NAME = Acc
   Call Ri.QueryData(1, TempRs, iCount)
   If TempRs.EOF Then
      ParentKey = GetParentKey(Acc, TopFlag)
      If Not TopFlag Then
         Call CreatePermissionNode(ParentKey, TempParentID, ReportName)
         Ri.PARENT_ID = TempParentID
      End If
      
      Ri.AddEditMode = SHOW_ADD
      Call GetParentItemDesc(Acc, Ri, ReportName)
      Call Ri.AddEditData
      ParentID = Ri.RIGHT_ID
   Else
      Call Ri.PopulateFromRS(1, TempRs)
      ParentID = Ri.RIGHT_ID
   End If
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   Set Ri = Nothing
End Function

Public Function VerifyAccessRight(Acc As String, Optional ReportName As String = "", Optional msgType As Long) As Boolean
Dim R As CGroupRight
Dim iCount As Long
Dim TempParentID As Long
Dim FoundFlag As Boolean
   
   If glbUser.REAL_USER_ID = 0 Then
      VerifyAccessRight = True
      Exit Function
   End If
   
   Call glbDaily.StartTransaction
   Call CreatePermissionNode(Acc, TempParentID, ReportName)
   Call glbDaily.CommitTransaction
   
   FoundFlag = False
   For Each R In glbAccessRight
      If R.RIGHT_ITEM_NAME = Acc Then
         FoundFlag = True
         If R.RIGHT_STATUS = "Y" Then
            VerifyAccessRight = True
            Exit For
         Else
            VerifyAccessRight = False
            Exit For
         End If
      End If
   Next R
   
   VerifyAccessRight = True
   
   If Not VerifyAccessRight Then
      VerifyAccessRight = False
      If Not msgType = 2 Then
         glbErrorLog.LocalErrorMsg = "�������ö��ҹ�������ǹ��������ͧ�ҡ���Է��������§ -> " & Acc
         glbErrorLog.ShowUserError
      End If
   Else
      VerifyAccessRight = True
   End If
End Function

Public Function VerifyCombo(L As Label, C As ComboBox, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   If L Is Nothing Then
      S = ""
   Else
      S = L.Caption
   End If
   
   If Not NullAllow Then
      If Len(C.Text) = 0 Then
         VerifyCombo = False
         Call MsgBox("��سҡ�͡������ " & " '" & S & "' " & "���١��ͧ��Фú��ǹ ", vbOKOnly, PROJECT_NAME)
         If C.Enabled And C.Visible Then
            C.SetFocus
         End If
         Exit Function
      End If
   End If
   
   VerifyCombo = True
End Function

Public Function VerifyComboEx(C As ComboBox, Optional NullAllow As Boolean = False) As Boolean
Dim S As String
   
   If Not NullAllow Then
      If Len(C.Text) = 0 Then
         VerifyComboEx = False
         Call MsgBox("��سҡ�͡������ " & " '" & S & "' " & "���١��ͧ��Фú��ǹ ", vbOKOnly, PROJECT_NAME)
         If C.Enabled Then
            C.SetFocus
         End If
         Exit Function
      End If
   End If
   
   VerifyComboEx = True
End Function

Public Function VerifyItem(C As Collection, T As Object, Idx As Long) As Boolean
Dim I As Long
Dim Count As Long

   If C.Count <= 0 Then
      VerifyItem = True
      Exit Function
   End If
   
   For I = 1 To C.Count
      If C.Item(I).CURRENT_FLAG = "Y" Then
         Count = Count + 1
      End If
   Next I
   
   If Count <> 1 Then
      Call MsgBox("��س����͡����������դ�һѨ�غѹ 1 ��¡��", vbOKOnly, PROJECT_NAME)
   
      T.Tabs.Item(Idx).Selected = True
      VerifyItem = False
      Exit Function
   End If
   
   VerifyItem = True
End Function

Public Sub SetTextLenType(T As TextBox, TT As TEXT_BOX_TYPE, L As Long)
   If TT = TEXT_FLOAT_MONEY Or TT = TEXT_INTEGER_MONEY Then
      T.Alignment = 1
   End If
   
   T.Tag = TT
   T.MaxLength = L
End Sub

Public Function ChangeQuote(StrQ As String) As String
   ChangeQuote = Replace(StrQ, "'", "''")
End Function

Public Function NVLI(Value As Variant, I As Long) As Long
On Error Resume Next

   If IsNull(Value) Then
      NVLI = I
   Else
      NVLI = Value
   End If
End Function

Public Function NVLD(Value As Variant, I As Double) As Double
On Error Resume Next

   If IsNull(Value) Then
      NVLD = I
   Else
      NVLD = Value
   End If
End Function

Public Function NVLS(Value As Variant, S As String) As String
On Error Resume Next

   If IsNull(Value) Then
      NVLS = S
'   ElseIf IsEmpty(Value) Then
'      NVLS = S
   Else
      NVLS = Value
   End If
End Function

Public Function EmptyToString(Value As String, S As String) As String
On Error Resume Next

   If Value = "" Then
      EmptyToString = S
   Else
      EmptyToString = Value
   End If
End Function

Public Function CryptString(strInput As String, strKey As String, action As Boolean)
Dim I As Integer, C As Integer
Dim strOutput As String

If Len(strKey) Then
    For I = 1 To Len(strInput)
        C = Asc(Mid$(strInput, I, 1))
        If action Then
            C = C + Asc(Mid$(strKey, (I Mod Len(strKey)) + 1, 1))
        Else: C = C - Asc(Mid$(strKey, (I Mod Len(strKey)) + 1, 1))
        End If
        strOutput = strOutput & Chr$(C And &HFF)
    Next I
Else
    strOutput = strInput
End If
CryptString = strOutput
End Function

Public Function EncryptText(PText As String) As String
   EncryptText = CryptString(PText, "GENETICOTHELLO", True)
End Function

Public Function DecryptText(CText As String) As String
   DecryptText = CryptString(CText, "GENETICOTHELLO", False)
End Function

Public Function EnableForm(Frm As Form, En As Boolean)
   If Frm Is Nothing Then
      Exit Function
   End If
   
   Frm.Enabled = En
   If En Then
      Screen.MousePointer = vbArrow
   Else
      Frm.Refresh
      DoEvents
      Screen.MousePointer = 11
   End If
End Function

Public Function IntToThaiMonth(M As Long) As String
   If glbParameterObj Is Nothing Then
      Exit Function
   End If
   
   If M = 1 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "���Ҥ�"
      Else
         IntToThaiMonth = "January"
      End If
   ElseIf M = 2 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "����Ҿѹ��"
      Else
         IntToThaiMonth = "February"
      End If
      
   ElseIf M = 3 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "�չҤ�"
      Else
         IntToThaiMonth = "March"
      End If
      
   ElseIf M = 4 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "����¹"
      Else
         IntToThaiMonth = "April"
      End If
      
   ElseIf M = 5 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "����Ҥ�"
      Else
         IntToThaiMonth = "May"
      End If
      
   ElseIf M = 6 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "�Զع�¹"
      Else
         IntToThaiMonth = "June"
      End If
      
   ElseIf M = 7 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "�á�Ҥ�"
      Else
         IntToThaiMonth = "July"
      End If
      
   ElseIf M = 8 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "�ԧ�Ҥ�"
      Else
         IntToThaiMonth = "August"
      End If
      
   ElseIf M = 9 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "�ѹ��¹"
      Else
         IntToThaiMonth = "September"
      End If
      
   ElseIf M = 10 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "���Ҥ�"
      Else
         IntToThaiMonth = "October"
      End If
      
   ElseIf M = 11 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "��Ȩԡ�¹"
      Else
         IntToThaiMonth = "November"
      End If
      
   ElseIf M = 12 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonth = "�ѹ�Ҥ�"
      Else
         IntToThaiMonth = "December"
      End If
   Else
      IntToThaiMonth = ""
   End If
End Function

Public Function IntToThaiMonthEx(M As Long) As String
   If glbParameterObj Is Nothing Then
      Exit Function
   End If
   
   If M = 1 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonthEx = "�.�."
      Else
         IntToThaiMonthEx = "January"
      End If
   ElseIf M = 2 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonthEx = "�.�."
      Else
         IntToThaiMonthEx = "February"
      End If
      
   ElseIf M = 3 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonthEx = "��.�."
      Else
         IntToThaiMonthEx = "March"
      End If
      
   ElseIf M = 4 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonthEx = "��.�."
      Else
         IntToThaiMonthEx = "April"
      End If
      
   ElseIf M = 5 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonthEx = "�.�."
      Else
         IntToThaiMonthEx = "May"
      End If
      
   ElseIf M = 6 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonthEx = "��.�."
      Else
         IntToThaiMonthEx = "June"
      End If
      
   ElseIf M = 7 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonthEx = "�.�."
      Else
         IntToThaiMonthEx = "July"
      End If
      
   ElseIf M = 8 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonthEx = "�.�."
      Else
         IntToThaiMonthEx = "August"
      End If
      
   ElseIf M = 9 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonthEx = "�.�."
      Else
         IntToThaiMonthEx = "September"
      End If
      
   ElseIf M = 10 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonthEx = "�.�."
      Else
         IntToThaiMonthEx = "October"
      End If
      
   ElseIf M = 11 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonthEx = "�.�."
      Else
         IntToThaiMonthEx = "November"
      End If
      
   ElseIf M = 12 Then
      If glbParameterObj.Language = 1 Then
         IntToThaiMonthEx = "�.�."
      Else
         IntToThaiMonthEx = "December"
      End If
   Else
      IntToThaiMonthEx = ""
   End If
End Function

Public Function DateToStringMonthYearExt(D As Date) As String
   If D < 0 Then
      DateToStringMonthYearExt = ""
      Exit Function
   End If
   
   DateToStringMonthYearExt = " " & IntToThaiMonth(Month(D)) & " " & Format(Year(D) + 543, "0000")
End Function

Public Function DateToStringExt(D As Date) As String
   If D < 0 Then
      DateToStringExt = "-"
      Exit Function
   Else
      DateToStringExt = Day(D) & " " & IntToThaiMonth(Month(D)) & " " & Format(Year(D) + 543, "0000")
   End If
End Function

Public Function DateToStringExtEx(D As Date) As String
   If D < 0 Then
      DateToStringExtEx = ""
      Exit Function
   End If
   
   DateToStringExtEx = Day(D) & " " & IntToThaiMonth(Month(D)) & " " & Format(Year(D) + 543, "0000") & _
                     " " & Format(Hour(D), "00") & ":" & Format(Minute(D), "00") & ":" & Format(Second(D), "00")
End Function

Public Function DateToStringIntEx2(D As Date, Minute As Long, Second As Long) As String
   DateToStringIntEx2 = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-" & Format(Day(D), "00") & " " & _
   Format(Minute, "00") & ":" & Format(Second, "00") & ":00"
End Function

Public Function DateToStringExtEx2(D As Date) As String
   If D > 0 Then
      DateToStringExtEx2 = Format(Day(D), "00") & "/" & Format(Month(D), "00") & "/" & Format(Year(D) + 543, "0000")
   Else
      DateToStringExtEx2 = ""
   End If
End Function
Public Function DateToStringExtEx4(D As Date) As String
   If D > 0 Then
      DateToStringExtEx4 = Format(Day(D), "00") & "/" & Format(Month(D), "00") & "/" & Format(Year(D), "0000")
   Else
      DateToStringExtEx4 = ""
   End If
End Function

Public Function DateToStringExtEx3(D As Date) As String
   If D > 0 Then
      DateToStringExtEx3 = Format(Day(D), "00") & "/" & Format(Month(D), "00") & "/" & Format(Year(D) + 543, "0000")
      DateToStringExtEx3 = DateToStringExtEx3 & " " & Format(Hour(D), "00") & ":" & Format(Minute(D), "00") & ":" & Format(Second(D), "00")
   Else
      DateToStringExtEx3 = ""
   End If
End Function

Public Function DateToStringIntEx3(D As Date) As String
   DateToStringIntEx3 = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-" & Format(Day(D), "00")
End Function

Public Function InternalDateToStringEx4(D As String) As String
Dim T As Date
   T = InternalDateToDate(D)
   If T > 0 Then
      InternalDateToStringEx4 = Format(Day(T), "00") & "/" & Format(Month(T), "00") & "/" & Format(Year(T) + 543, "0000")
   Else
      InternalDateToStringEx4 = ""
   End If
End Function



Public Function DateToStringInt(D As Date) As String
   If D = -1 Then
      DateToStringInt = "9999-99-99 99:99:99"
   ElseIf D = -2 Then
      DateToStringInt = "0000-00-00 00:00:00"
   Else
      DateToStringInt = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-" & Format(Day(D), "00") & _
                     " " & Format(Hour(D), "00") & ":" & Format(Minute(D), "00") & ":" & Format(Second(D), "00")
   End If
End Function
Public Function DateToStringIntEndMonth(D As Date) As String
   DateToStringIntEndMonth = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-31" & _
                     " 00:00:00"
End Function
Public Function DateToStringIntToSumFarm(D As Date) As String
   If D = -1 Then
      DateToStringIntToSumFarm = "99999999"
   ElseIf D = -2 Then
      DateToStringIntToSumFarm = "00000000"
   Else
      DateToStringIntToSumFarm = Format(Year(D), "0000") & Format(Month(D), "00") & Format(Day(D), "00")
   End If
End Function
Public Function DateToStringIntEx(D As Date) As String
   DateToStringIntEx = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-" & Format(Day(D), "00") & _
                     " 23:59:59"
End Function

Public Function DateToStringIntHi(D As Date) As String
   If D > 0 Then
      DateToStringIntHi = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-" & Format(Day(D), "00") & _
                     " 23:59:59"
   Else
      DateToStringIntHi = "9999" & "-" & "99" & "-" & "99" & _
                     " 99:99:99"
   End If
End Function

Public Function DateToStringIntLow(D As Date) As String
   If D = -1 Then
      DateToStringIntLow = "9999" & "-" & "99" & "-" & "99" & _
                     " 99:99:99"
   ElseIf D = -2 Then
      DateToStringIntLow = "0000" & "-" & "00" & "-" & "00" & _
                     " 00:00:00"
   Else
      DateToStringIntLow = Format(Year(D), "0000") & "-" & Format(Month(D), "00") & "-" & Format(Day(D), "00") & _
                        " 00:00:00"
   End If
End Function
Public Function InternalDateToDate(IntDate As String) As Date
Dim DStr As Long
Dim D As Long
Dim MStr As String
Dim M As Long
Dim YStr As String
Dim Y As Long

Dim HHStr As Long
Dim HH As Long
Dim MMStr As String
Dim MM As Long
Dim SSStr As String
Dim SS As Long

   If (IntDate = "") Or (IntDate = "9999-99-99 99:99:99") Then
      InternalDateToDate = -1
      Exit Function
   End If
   
   If (IntDate = "") Or (IntDate = "0000-00-00 00:00:00") Then
      InternalDateToDate = -2
      Exit Function
   End If
   
   If Len(IntDate) < 19 Then
      InternalDateToDate = Now
      Exit Function
   End If
   
   YStr = Mid(IntDate, 1, 4)
   MStr = Mid(IntDate, 6, 2)
   DStr = Mid(IntDate, 9, 2)
   
'   If Not IsNumeric(YStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(MStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(DStr) Then
'      Exit Function
'   End If
   
   HHStr = Mid(IntDate, 12, 2)
   MMStr = Mid(IntDate, 15, 2)
   SSStr = Mid(IntDate, 18, 2)
   
'   If Not IsNumeric(HHStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(MMStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(SSStr) Then
'      Exit Function
'   End If
   
   HH = Val(HHStr)
   MM = Val(MMStr)
   SS = Val(SSStr)
   
   Y = Val(YStr)
   M = Val(MStr)
   D = Val(DStr)
   
   InternalDateToDate = DateSerial(Y, M, D) + TimeSerial(HH, MM, SS)
End Function

Public Function InternalDateToDateEx(IntDate As String) As Date
Dim DStr As Long
Dim D As Long
Dim MStr As String
Dim M As Long
Dim YStr As String
Dim Y As Long

Dim HHStr As Long
Dim HH As Long
Dim MMStr As String
Dim MM As Long
Dim SSStr As String
Dim SS As Long

   If (IntDate = "") Or (IntDate = "9999-99-99 99:99:99") Then
      InternalDateToDateEx = -1
      Exit Function
   End If
   
   If (IntDate = "") Or (IntDate = "0000-00-00 00:00:00") Then
      InternalDateToDateEx = -1
      Exit Function
   End If
   
   If Len(IntDate) < 8 Then
      InternalDateToDateEx = Now
      Exit Function
   End If
   
   YStr = Mid(IntDate, 1, 4)
   MStr = Mid(IntDate, 5, 2)
   DStr = Mid(IntDate, 7, 2)
   
'   If Not IsNumeric(YStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(MStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(DStr) Then
'      Exit Function
'   End If
   
   HHStr = "00"
   MMStr = "00"
   SSStr = "00"
   
'   If Not IsNumeric(HHStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(MMStr) Then
'      Exit Function
'   End If
'
'   If Not IsNumeric(SSStr) Then
'      Exit Function
'   End If
   
   HH = Val(HHStr)
   MM = Val(MMStr)
   SS = Val(SSStr)
   
   Y = Val(YStr) - 543
   M = Val(MStr)
   D = Val(DStr)
   
   InternalDateToDateEx = DateSerial(Y, M, D) + TimeSerial(HH, MM, SS)
End Function

Public Function InternalDateToDateEx2(IntDate As String) As Date
Dim DStr As Long
Dim D As Long
Dim MStr As String
Dim M As Long
Dim YStr As String
Dim Y As Long

Dim HHStr As Long
Dim HH As Long
Dim MMStr As String
Dim MM As Long
Dim SSStr As String
Dim SS As Long

   If (IntDate = "") Or (IntDate = "9999-99-99 99:99:99") Then
      InternalDateToDateEx2 = -1
      Exit Function
   End If
   
   If (IntDate = "") Or (IntDate = "0000-00-00 00:00:00") Then
      InternalDateToDateEx2 = -1
      Exit Function
   End If
   
   If Len(IntDate) < 10 Then
      InternalDateToDateEx2 = Now
      Exit Function
   End If
   
   YStr = Mid(IntDate, 1, 4)
   MStr = Mid(IntDate, 6, 2)
   DStr = Mid(IntDate, 9, 2)
      
   HHStr = "00"
   MMStr = "00"
   SSStr = "00"
      
   HH = Val(HHStr)
   MM = Val(MMStr)
   SS = Val(SSStr)
   
   Y = Val(YStr)
   M = Val(MStr)
   D = Val(DStr)
   
   InternalDateToDateEx2 = DateSerial(Y, M, D) + TimeSerial(HH, MM, SS)
End Function

Public Function ReFormatDate(DStr As String) As String
Dim YYYY As String
Dim MM As String
Dim DD As String

   YYYY = Mid(DStr, 5, 4)
   MM = Mid(DStr, 3, 2)
   DD = Mid(DStr, 1, 2)
   
   ReFormatDate = YYYY & MM & DD
End Function

Public Sub InitTextBox(T As TextBox, Msg As String, Optional Password As String = "")
   T.PasswordChar = Password
   T.FontSize = 12
   T.FontName = "MS Sans Serif"
   T.Text = Msg
   T.BackColor = GLB_GRID_COLOR
   'T.FontBold = True
End Sub

Public Sub ShowTotalLabel(L As Label, Value As Long)
   L.Caption = "��� = " & Value
End Sub

Public Sub ClearTreeView(ByVal tvHwnd As Long)
Dim lNodeHandle As Long

    'Turn off redrawing on the Treeview for more speed improvements
    SendMessageLong tvHwnd, WM_SETREDRAW, False, 0

    Do
        lNodeHandle = SendMessageLong(tvHwnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
         If lNodeHandle > 0 Then
            SendMessageLong tvHwnd, TVM_DELETEITEM, 0, lNodeHandle
         Else
            Exit Do
         End If
    Loop

    'Turn on redrawing on the Treeview
    SendMessageLong tvHwnd, WM_SETREDRAW, True, 0
End Sub

Public Sub InitCombo(C As ComboBox)
   C.FontSize = 12
   C.FontName = "MS Sans Serif"
   C.BackColor = GLB_GRID_COLOR
End Sub

Public Function VerifyGrid(S As String) As Boolean
   If S = "" Then
      VerifyGrid = False
      glbErrorLog.LocalErrorMsg = "��س����͡�����ŷ���ͧ��á�͹"
      glbErrorLog.ShowUserError
   Else
      VerifyGrid = True
   End If
End Function

Public Function ConfirmDelete(S As String) As Boolean
   glbErrorLog.LocalErrorMsg = "��ҹ��ͧ��è�ź������ " & S & "' ���������"
   If glbErrorLog.AskMessage = vbNo Then
      ConfirmDelete = False
      Exit Function
   Else
      ConfirmDelete = True
   End If
End Function

Public Sub InitFormHeader(L As Label, Caption As String)
   L.Caption = Caption
   L.FontBold = True
   L.FontSize = 20
   L.FontName = GLB_FONT
   L.Alignment = 2
   L.ForeColor = RGB(0, 10, 0)
End Sub

Public Sub InitDialogHeader(L As Label, Caption As String)
   L.Caption = Caption
   L.FontBold = True
   L.FontSize = 16
   L.FontName = GLB_FONT
   L.Alignment = 2
End Sub

Public Sub InitNormalLabel(L As Label, Caption As String, Optional Color As Long = 0)
   L.Caption = Caption
   L.FontBold = False
   L.FontSize = 14
   L.FontBold = True
   L.FontName = GLB_FONT
   L.BackStyle = 0
   L.ForeColor = Color
End Sub

Public Sub InitOption(O As OptionButton, Caption As String)
   O.Caption = Caption
   O.FontSize = 14
   O.FontBold = True
   O.FontName = GLB_FONT
   O.BackColor = GLB_FORM_COLOR
End Sub

Public Sub InitOptionEx(O As SSOption, Caption As String)
   O.Caption = Caption
   O.Font.Size = 14
   O.Font.Bold = True
   O.Font.Name = GLB_FONT
   O.BackColor = GLB_FORM_COLOR
   O.BackStyle = ssTransparent
End Sub

Public Sub InitCheckBox(C As SSCheck, Caption As String)
   C.Caption = Caption
   C.FontSize = 14
   C.FontBold = True
   C.FontName = GLB_FONT
   C.BackColor = GLB_FORM_COLOR
   C.BackStyle = ssTransparent
   C.TripleState = True
End Sub

Public Sub InitMainButton(B As SSCommand, Caption As String, Optional Color As Double = &HFFFFFF)
   B.Caption = Caption
   B.Font.Bold = True
   B.Font.Size = 14
   B.Font.Name = GLB_FONT
   B.Font3D = ssInsetLight
   B.BackColor = RGB(255, 255, 255)
   B.ButtonStyle = ssWin95 '= ssActiveBorders
   B.MousePointer = ssCustom
   B.MouseIcon = LoadPicture(glbParameterObj.ButtonCursor)
End Sub

Public Sub InitHeaderFooter(H As SSPanel, F As SSPanel)
'   H.PICTURE = LoadPicture("D:\Picture\WINPricing100\header.gif")
   If Not (F Is Nothing) Then
'      F.PICTURE = LoadPicture("D:\Picture\WINPricing100\footer.gif")
   End If
End Sub

Public Sub InitMainButtonOld(B As CommandButton, Caption As String, Optional Color As Double = &HFFFFFF)
   B.Caption = Caption
   B.Font.Bold = True
   B.Font.Size = 14
   B.Font.Name = GLB_FONT
   B.BackColor = GLB_FORM_COLOR
End Sub

Public Sub SetSelect(T As TextBox)
   T.SelStart = 0
   T.SelLength = Len(T.Text)
End Sub

Public Sub InitDialogButton(B As CommandButton, Caption As String)
   B.Caption = Caption
   B.FontBold = True
   B.FontSize = 14
   B.FontName = GLB_FONT
   
   B.BackColor = &HFFFFFF
End Sub

Public Sub ReleaseAll()
   Set glbErrorLog = Nothing
   Set glbDatabaseMngr = Nothing
   Set glbParameterObj = Nothing
   Set glbUser = Nothing
   Set glbGroup = Nothing

   Set glbSetting = Nothing
   Set glbDaily = Nothing
   Set glbAdmin = Nothing
   Set glbMaster = Nothing
   Set glbLegacy = Nothing
   Set glbLoginTracking = Nothing
   Set glbEnterPrise = Nothing
   Set glbAccessRight = Nothing
   
   Set CustomerPackage = Nothing
   Set PackageDetail = Nothing
   Set T706Collection1 = Nothing
   Set T706Collection2 = Nothing
End Sub

Public Sub SetEnableDisableTextBox(T As TextBox, En As Boolean)
   If En Then
      T.Enabled = True
      T.BackColor = GLB_GRID_COLOR
   Else
      T.Enabled = False
      T.BackColor = &H8000000F
   End If
End Sub

Public Sub SetEnableDisableComboBox(T As ComboBox, En As Boolean)
   If En Then
      T.Enabled = True
      T.BackColor = GLB_GRID_COLOR
   Else
      T.Enabled = False
      T.BackColor = &H8000000F
   End If
End Sub

Public Sub SetEnableDisableButton(B As SSCommand, En As Boolean)
   If En Then
      B.Enabled = True
      B.BackColor = GLB_GRID_COLOR
   Else
      B.Enabled = False
      B.BackColor = &H8000000F
   End If
End Sub

Public Function ConfirmExit(HasEdit As Boolean) As Boolean
   If Not HasEdit Then
      ConfirmExit = True
   Else
      glbErrorLog.LocalErrorMsg = "��ҹ��ͧ��è��͡�ҡ�����������ա�úѹ�֡���������������"
      If glbErrorLog.AskMessage = vbYes Then
         ConfirmExit = True
      Else
         ConfirmExit = False
      End If
   End If
End Function

Public Function ThaiBaht(ByVal pamt As Double) As String
Dim valstr As String, vLen As Integer, vno As Integer, syslge As String
Dim I As Integer, j As Integer, v As Integer
Dim wnumber(10) As String, wdigit(10) As String, spcdg(5) As String
Dim vword(20) As String

 If pamt <= 0# Then
   ThaiBaht = ""
   Exit Function
 End If
 valstr = Trim(Format$(pamt, "##########0.00"))
 vLen = Len(valstr) - 3
 For I = 1 To 20
     vword(I) = ""
 Next I
wnumber(1) = "˹��": wnumber(2) = "�ͧ": wnumber(3) = "���": wnumber(4) = "���"
wnumber(5) = "���": wnumber(6) = "ˡ": wnumber(7) = "��": wnumber(8) = "Ỵ"
wnumber(9) = "���": wdigit(1) = "�ҷ": wdigit(2) = "�Ժ": wdigit(3) = "����": wdigit(4) = "�ѹ"
wdigit(5) = "����": wdigit(6) = "�ʹ": wdigit(7) = "��ҹ": spcdg(1) = "ʵҧ��": spcdg(2) = "���"
spcdg(3) = "���": spcdg(4) = "��ǹ"
For I = 1 To vLen
    vno = Int(Val(Mid$(valstr, I, 1)))
    If vno = 0 Then
        vword(I) = ""
        If (vLen - I + 1) = 7 Then
            vword(I) = wdigit(7)             '--��ҹ
        End If
    Else
        If (vLen - I + 1) > 7 Then
            j = vLen - I - 5               '--�Թ��ѡ��ҹ
        Else
            j = vLen - I + 1               '--��ѡ�ʹ
        End If
        vword(I) = wnumber(vno) + wdigit(j) '-30�֧90
        If vno = 1 And j = 2 Then
            vword(I) = wdigit(2)             '--�Ժ
        End If
        If vno = 2 And j = 2 Then
            vword(I) = spcdg(3) + wdigit(j)  '--����Ժ
        End If
        If j = 1 Then                       ' ������ -->����Ժ���
            vword(I) = wnumber(vno)
            If vno = 1 And vLen > 1 Then
                If Mid$(valstr, I - 1, 1) <> "0" Then
                    vword(I) = spcdg(2)
                End If
            End If
        End If
        If j = 7 Then         '-��ѡ�ó� 11,111,111.00 �Ժ���
            vword(I) = wnumber(vno) + wdigit(j)   '-��ҹ
            If vno = 1 And vLen > 7 Then
                If Mid$(valstr, I - 1, 1) <> "0" Then
                    vword(I) = spcdg(2) + wdigit(j)
                End If
            End If
        End If
    End If
Next I
    
If Int(pamt) > 0 Then
       vword(vLen) = vword(vLen) + wdigit(1)
End If
 '--------------�ȹ��� --------------
valstr = Mid$(valstr, vLen + 2, 2)
vLen = Len(valstr)
For I = 1 To vLen
    vno = Int(Val(Mid$(valstr, I, 1)))
    If vno = 0 Then
           vword(I + 10) = ""
    Else
           j = vLen - I + 1
           vword(I + 10) = wnumber(vno) + wdigit(j)
        If vno = 1 And j = 2 Then
              vword(I + 10) = wdigit(2)
        End If
        If vno = 2 And j = 2 Then
              vword(I + 10) = spcdg(3) + wdigit(j)
        End If
        If j = 1 Then
            If vno = 1 And Int(Val(Mid$(valstr, I - 1, 1))) <> 0 Then
                 vword(I + 10) = spcdg(2)
            Else
                 vword(I + 10) = wnumber(vno)
            End If
        End If
    End If
Next I
If pamt <> 0 Then
    If Val(valstr) = 0 Then
        vword(13) = spcdg(4)
    Else
        vword(13) = spcdg(1)
    End If
End If

 '*** ������ó�����ҡ ��е�ͧ��õѴ����¤
 valstr = ""
 For I = 1 To 20
    'IF LEN(valstr) < 70 AND LEN(valstr + vword(i)) > 70 Then
    '   valstr = valstr + REPLICATE(" ",70 - LEN(valstr))
    'END IF
    valstr = valstr + vword(I)
 Next I
 'valstr='('+valstr+')'
 ThaiBaht = (valstr)
End Function
Function ThaiBahtEng(dblValue As Double) As String
Static ones(0 To 9) As String
Static teens(0 To 9) As String
Static tens(0 To 9) As String
Static thousands(0 To 4) As String
Dim I As Integer, nPosition As Integer
Dim nDigit As Integer, bAllZeros As Integer
Dim strResult As String, strTemp As String
Dim tmpBuff As String

ones(0) = "zero"
ones(1) = "one"
ones(2) = "two"
ones(3) = "three"
ones(4) = "four"
ones(5) = "five"
ones(6) = "six"
ones(7) = "seven"
ones(8) = "eight"
ones(9) = "nine"

teens(0) = "ten"
teens(1) = "eleven"
teens(2) = "twelve"
teens(3) = "thirteen"
teens(4) = "fourteen"
teens(5) = "fifteen"
teens(6) = "sixteen"
teens(7) = "seventeen"
teens(8) = "eighteen"
teens(9) = "nineteen"

tens(0) = ""
tens(1) = "ten"
tens(2) = "twenty"
tens(3) = "thirty"
tens(4) = "forty"
tens(5) = "fifty"
tens(6) = "sixty"
tens(7) = "seventy"
tens(8) = "eighty"
tens(9) = "ninty"

thousands(0) = ""
thousands(1) = "thousand"
thousands(2) = "million"
thousands(3) = "billion"
thousands(4) = "trillion"

'Trap errors
On Error GoTo ThaiBahtEngError
'Get fractional part
'strResult = "and " & Format((dblValue - Int(dblValue)) * 100, "00") &"/100"
'strResult = "baht only)"
strResult = "only"
'Convert rest to string and process each digit
strTemp = CStr(Int(dblValue))
'Iterate through string
For I = Len(strTemp) To 1 Step -1
'Get value of this digit
nDigit = Val(Mid$(strTemp, I, 1))
'Get column position
nPosition = (Len(strTemp) - I) + 1
'Action depends on 1's, 10's or 100's column
Select Case (nPosition Mod 3)
Case 1 '1's position
bAllZeros = False
If I = 1 Then
tmpBuff = ones(nDigit) & " "
ElseIf Mid$(strTemp, I - 1, 1) = "1" Then
tmpBuff = teens(nDigit) & " "
I = I - 1 'Skip tens position
ElseIf nDigit > 0 Then
tmpBuff = ones(nDigit) & " "
Else
'If next 10s & 100s columns are also
'zero, then don't show 'thousands'
bAllZeros = True
If I > 1 Then
If Mid$(strTemp, I - 1, 1) <> "0" Then
bAllZeros = False
End If
End If
If I > 2 Then
If Mid$(strTemp, I - 2, 1) <> "0" Then
bAllZeros = False
End If
End If
tmpBuff = ""
End If
If bAllZeros = False And nPosition > 1 Then
tmpBuff = tmpBuff & thousands(nPosition / 3) & " "
End If
strResult = tmpBuff & strResult
Case 2 'Tens position
If nDigit > 0 Then
strResult = tens(nDigit) & " " & strResult
End If
Case 0 'Hundreds position
If nDigit > 0 Then
strResult = ones(nDigit) & " hundred " & strResult
End If
End Select
Next I
'Convert first letter to upper case
If Len(strResult) > 0 Then
strResult = UCase$(Left$(strResult, 1)) & Mid$(strResult, 2)
End If

EndThaiBahtEng:
'Return result
ThaiBahtEng = strResult
Exit Function

ThaiBahtEngError:
strResult = "#Error#"
Resume EndThaiBahtEng
End Function
Public Function WildCard(WStr As String, SubLen As Long, NewStr As String) As Boolean
Dim Tmp As String
   Tmp = Trim(WStr)
   If Tmp = "" Then
      WildCard = False
      Exit Function
   End If
   
   If Mid(Tmp, Len(Tmp)) = "%" Then
      SubLen = Len(Tmp) - 1
      NewStr = Mid(Tmp, 1, SubLen)
      
      WildCard = True
   Else
      WildCard = False
   End If
End Function

Public Function FormatString(S As String, Patch As String, L As Long) As String
Dim Temp As String
Dim Start As Long
Dim I As Long
Dim j As Long

   Temp = Space(L)
   Call Replace(Temp, " ", Patch)
   j = 0
   Start = (L - Len(S)) \ 2
   
   For I = 1 To L
      If I < Start Then
         Mid(Temp, I) = Patch
      Else
         If I > Start + Len(S) Then
            Mid(Temp, I) = Patch
         Else
            j = j + 1
            Mid(Temp, I) = Mid(S, j)
         End If
      End If
   Next I
   
   FormatString = Temp
End Function

Public Function FormatNumber(N As Variant, Optional ZeroString As String = "0.00", Optional DecimalPoint As Long = 2, Optional NullFlag As Boolean = False) As String
Dim T As Double
Dim Temp As String
Dim I As Long

   If IsNull(N) Then
      T = 0
   Else
      T = N
   End If
   
   If DecimalPoint < 2 Then
      DecimalPoint = 2
   End If
   
   Temp = ""
   For I = 1 To DecimalPoint
      Temp = Temp & "0"
   Next I
      
   If T = 0 Then
      If ZeroString = "-" Then
         FormatNumber = ZeroString
      ElseIf NullFlag Then
         FormatNumber = ""
      Else
         FormatNumber = "0." & Temp
      End If
   ElseIf T > 0 Then
      FormatNumber = Format(T, "#,##0." & Temp)
   ElseIf T < 0 Then
      FormatNumber = "(" & Format(-1 * T, "#,##0." & Temp) & ")"
   End If
End Function

Public Function FormatNumberInt(N As Variant, Optional ZeroString As String = "0") As String
Dim T As Double

   If IsNull(N) Then
      T = 0
   Else
      T = N
   End If
   
   If T = 0 Then
      FormatNumberInt = ZeroString
   ElseIf T > 0 Then
      FormatNumberInt = Format(T, "#,##0")
   ElseIf T < 0 Then
      FormatNumberInt = "(" & Format(-1 * T, "#,##0") & ")"
   End If
End Function
Public Function FormatNumberToNull(N As Variant, Optional DecimalPoint As Long = 2, Optional Quat As Boolean = True, Optional ZeroString As String = "") As String
Dim T As Double
Dim TempStr As String
Dim I As Long

   TempStr = "."
   For I = 1 To DecimalPoint
      TempStr = TempStr & "0"
   Next I
   If DecimalPoint = 0 Then
       TempStr = ""
   End If
   
   If IsNull(N) Then
      T = 0
   Else
      T = N
   End If
   
   If T = 0 Then
      If ZeroString = "0" Then
         FormatNumberToNull = ZeroString & TempStr
      Else
         FormatNumberToNull = ZeroString
      End If
   ElseIf Quat Then
      FormatNumberToNull = Format(T, "#,##0" & TempStr)
   Else
      FormatNumberToNull = Format(T, "0" & TempStr)
   End If
End Function

Public Function ReverseFormatNumber(N As String) As Double
   ReverseFormatNumber = Val(Replace(N, ",", ""))
End Function

Public Function IDToListIndex(Cbo As ComboBox, ID As Long) As Long
Dim I As Long
Dim Temp As String

   IDToListIndex = -1
   For I = 0 To Cbo.ListCount - 1
      If InStr(Cbo.ItemData(I), ":") <= 0 Then
         Temp = Cbo.ItemData(I)
      Else
         Temp = Mid(Cbo.ItemData(I), 1, InStr(Cbo.ItemData(I), ":") - 1)
      End If
      If Temp = ID Then
         IDToListIndex = I
      End If
   Next I
End Function
Public Sub Main()
On Error GoTo ErrorHandler
Dim I As Long
Dim TempDB As String
   
   GLB_GRID_COLOR = RGB(255, 255, 250)
   GLB_NORMAL_COLOR = RGB(0, 0, 0)
   GLB_ALERT_COLOR = RGB(255, 0, 0)
   GLB_FORM_COLOR = RGB(180, 200, 200)
   GLB_HEAD_COLOR = GLB_FORM_COLOR
   GLB_GRIDHD_COLOR = RGB(149, 194, 240)
   GLB_SHOW_COLOR = RGB(0, 0, 240)
   GLB_MANDATORY_COLOR = RGB(0, 0, 255)

   Set glbSetting = New clsGlobalSetting
   Set glbParameterObj = New clsParameter
   Set glbUser = New CUser
   Set glbGroup = New CGroup
   
   
   Set glbErrorLog = New clsErrorLog
   glbErrorLog.DayKeepLog = 10
   glbErrorLog.LogFileMode = LOG_CURRENT_DATE
   
   glbErrorLog.ModuleName = MODULE_NAME
   glbErrorLog.RoutineName = "Main"
   glbErrorLog.MsgBoxTitle = PROJECT_NAME
   
   If App.PrevInstance = True Then
      glbErrorLog.LocalErrorMsg = "����������١�ѹ��͹˹�ҹ������"
      glbErrorLog.ShowUserError

      Set glbErrorLog = Nothing
      Exit Sub
   End If
   
   Load frmSplash
   frmSplash.Show 0
   frmSplash.Refresh
   
   If Command = "1" Or Command = "" Then
      TempDB = glbParameterObj.DBFile
   Else
      TempDB = glbParameterObj.DBFileAPX
   End If
   
   Set glbDatabaseMngr = New clsDatabaseMngr
   If Not glbDatabaseMngr.ConnectDatabase(TempDB, glbParameterObj.UserName, glbParameterObj.Password, glbErrorLog) Then
      frmDBSetting.UserName = glbParameterObj.UserName
      frmDBSetting.Password = glbParameterObj.Password
      frmDBSetting.FileDb = glbParameterObj.DBFile
      frmDBSetting.Header = " �������ö���͵�Ͱҹ�������� "

      Load frmDBSetting
      frmDBSetting.Show 1
      If frmDBSetting.OKClick Then
         glbParameterObj.UserName = frmDBSetting.UserName
         glbParameterObj.Password = frmDBSetting.Password
         
         If Command = "1" Or Command = "" Then
            glbParameterObj.DBFile = frmDBSetting.FileDb
         Else
            glbParameterObj.DBFileAPX = frmDBSetting.FileDb
         End If
      Else
         Unload frmDBSetting
         Set frmDBSetting = Nothing

         Unload frmSplash
         Set frmSplash = Nothing

         Call ReleaseAll
         End
      End If
      Unload frmDBSetting
      Set frmDBSetting = Nothing
   End If
   
'   If Not glbDatabaseMngr.ConnectLegacyDatabase(glbParameterObj.DBFile, glbParameterObj.UserName, glbParameterObj.Password, glbErrorLog) Then
'      ''debug.print "Error"
'   End If
   
'   If Not glbDatabaseMngr.ConnectAgentServer(glbParameterObj.LicenseIP, glbParameterObj.LicensePort, glbErrorLog) Then
'      frmAgentSetting.Port = glbParameterObj.LicensePort
'      frmAgentSetting.IP = glbParameterObj.LicenseIP
'      frmAgentSetting.Header = " �������ö�������͡Ѻ��ૹ������������� "
'
'      Load frmAgentSetting
'      frmAgentSetting.Show 1
'
'      If frmAgentSetting.OKClick Then
'         glbParameterObj.LicenseIP = frmAgentSetting.IP
'         glbParameterObj.LicensePort = frmAgentSetting.Port
'      Else
'         Unload frmAgentSetting
'         Set frmAgentSetting = Nothing
'
'         Unload frmSplash
'         Set frmSplash = Nothing
'
'         Call ReleaseAll
'         End
'      End If
'      Unload frmAgentSetting
'      Set frmAgentSetting = Nothing
'   End If
   Unload frmSplash
   Set frmSplash = Nothing
   
   Set glbDaily = New clsDaily
   Set glbAdmin = New clsAdmin
   Set glbMaster = New clsMaster
   Set glbLegacy = New clsLegacy
   Set glbLoginTracking = New CLoginTracking
   Set glbEnterPrise = New CEnterprise
   Set glbAccessRight = New Collection
   Set glbAuthenPO = New clsAuthenPO
   
   Set CustomerPackage = New Collection
   Set PackageDetail = New Collection
   Set T706Collection1 = New Collection
   Set T706Collection2 = New Collection
'   Call PatchDB
'   Call ReleaseAll
'   Exit Sub
   
   Load frmWinPricingMain
   frmWinPricingMain.Show

   Exit Sub
   
ErrorHandler:
   If glbErrorLog Is Nothing Then
      MsgBox Err.DESCRIPTION
   Else
      glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   End If
   
End Sub

Public Sub InitOrderType(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("������ҡ"))
   C.ItemData(1) = 1
   
   C.AddItem (MapText("�ҡ仹���"))
   C.ItemData(2) = 2
End Sub

Public Function GetItem(Col As Collection, Idx As Long, RealIndex As Long) As Object
Dim I As Long
Dim Count As Long

   Count = 0
   For I = 1 To Col.Count
      If Col.Item(I).Flag <> "D" Then
         Count = Count + 1
      End If
      If Count = Idx Then
         RealIndex = I
         Set GetItem = Col.Item(I)
         Exit Function
      End If
   Next I
   
   Set GetItem = Nothing
End Function

Public Function CountItem(Col As Collection) As Long
Dim I As Long
Dim Count As Long
   Count = 0
   For I = 1 To Col.Count
      If Col.Item(I).Flag <> "D" Then
         Count = Count + 1
      End If
   Next I
   
   CountItem = Count
End Function

Public Function VSP_CalTable(ByVal pRaw As String, ByVal pWidth As Long, ByRef pPer() As Long) As String
On Error GoTo ErrorHandler
Dim strTemp As String
Dim I As Long
Dim Count As Long
Dim iPer As Long
Dim tPer As Long
Dim Total As Long
Dim Prefix() As String
Dim Value() As Long
Dim iTemp As Long
   
   pRaw = Trim$(pRaw)
   If Len(pRaw) <= 0 Then
      VSP_CalTable = ""
      Exit Function
   End If
   Count = 0
   iPer = 1
   Total = 0
   strTemp = ""
   While iPer <= Len(pRaw)
      If Val(Mid$(pRaw, iPer, 1)) <= 0 Then
         strTemp = strTemp & Mid$(pRaw, iPer, 1)
      Else
         Count = Count + 1
         ReDim Preserve Prefix(Count)
         ReDim Preserve Value(Count)
         Prefix(Count) = strTemp
         tPer = InStr(iPer, pRaw, "|")
         If tPer <= 0 Then tPer = InStr(iPer, pRaw, ";")

         Value(Count) = Val(Mid$(pRaw, iPer, tPer - iPer))
         Total = Total + Value(Count)
         iPer = tPer
         strTemp = ""
      End If
      iPer = iPer + 1
   Wend
   strTemp = ""
   ReDim pPer(Count)
   For I = 1 To Count - 1
      iTemp = CLng((Value(I) * pWidth) / Total)
      strTemp = strTemp & Trim$(Prefix(I)) & Trim$(Str$(iTemp)) & "|"
      If I = 1 Then
         pPer(I - 1) = iTemp
      Else
         pPer(I - 1) = pPer(I - 2) + iTemp
      End If
   Next I
   strTemp = strTemp & Trim$(Prefix(I)) & CLng(((Value(I) * pWidth) / Total)) & ";"
   If I > 1 Then
      iTemp = CLng((Value(I) * pWidth) / Total)
      pPer(I - 1) = pPer(I - 2) + iTemp
   End If
   VSP_CalTable = strTemp

   Exit Function
ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Function

Public Function Check2Flag(A As Long) As String
   If A = ssCBChecked Then
      Check2Flag = "Y"
   Else
      Check2Flag = "N"
   End If
End Function

Public Function CheckUniqueNs(UnqType As UNIQUE_TYPE, Key As String, ID As Long, Optional TempID As Long = -1, Optional FieldNameExTendValue As String) As Boolean
On Error GoTo ErrorHandler
Dim TableName As String
Dim FieldName1 As String
Dim FieldName2 As String
Dim FieldNameExTend As String
Dim Flag As Boolean
Dim Count As Long

   CheckUniqueNs = False
'
'   TEACHER_UNIQUE = 16
'   SUBJECT_UNIQUE = 17
'   FACULTY_UNIQUE = 18
   
   Flag = False
   If UnqType = TEACHER_UNIQUE Then
      TableName = "TEACHER"
      FieldName1 = "TEACHER_CODE"
      FieldName2 = "TEACHER_ID"
      Flag = True
   ElseIf UnqType = USERGROUP_UNIQUE Then
      TableName = "USER_GROUP"
      FieldName1 = "GROUP_NAME"
      FieldName2 = "GROUP_ID"
      Flag = True
   ElseIf UnqType = SUBJECT_UNIQUE Then
      TableName = "SUBJECT"
      FieldName1 = "SUBJECT_CODE"
      FieldName2 = "SUBJECT_ID"
      Flag = True
   ElseIf UnqType = PRDFEATURE_UNIQUE Then
      TableName = "PRDFEATURE_NAME"
      FieldName1 = "PRODUCT_CODE"
      FieldName2 = "PRDFEATURE_NAME_ID"
      Flag = True
   ElseIf UnqType = FACULTY_UNIQUE Then
      TableName = "FACULTY"
      FieldName1 = "FACULTY_CODE"
      FieldName2 = "FACULTY_ID"
      Flag = True
   ElseIf UnqType = DBN_UNIQUE Then
      TableName = "BILL"
      FieldName1 = "BILL_NO"
      FieldName2 = "BILL_ID"
      Flag = True
   ElseIf UnqType = EMPCODE_UNIQUE Then
      TableName = "EMPLOYEE"
      FieldName1 = "EMP_CODE"
      FieldName2 = "EMP_ID"
      Flag = True
   ElseIf UnqType = USERNAME_UNIQUE Then
      TableName = "USER_ACCOUNT"
      FieldName1 = "USER_NAME"
      FieldName2 = "USER_ID"
      Flag = True
   ElseIf UnqType = REPAIR_UNIQUE Then
      TableName = "REPAIR_DATA"
      FieldName1 = "REPAIR_NUM"
      FieldName2 = "REPAIR_ID"
      Flag = True
   ElseIf UnqType = IMPORT_UNIQUE Then
      TableName = "INVENTORY_DOC"
      FieldName1 = "DOCUMENT_NO"
      FieldName2 = "INVENTORY_DOC_ID"
      Flag = True
   ElseIf UnqType = EXPORT_UNIQUE Then
      TableName = "INVENTORY_DOC"
      FieldName1 = "DOCUMENT_NO"
      FieldName2 = "INVENTORY_DOC_ID"
      Flag = True
   ElseIf UnqType = REPAIR_FORMULA_UNIQUE Then
      TableName = "REPAIR_FORMULA"
      FieldName1 = "FORMULA_CODE"
      FieldName2 = "FORMULA_ID"
      Flag = True
   ElseIf UnqType = SUPPLIER_UNIQUE Then
      TableName = "SUPPLIER"
      FieldName1 = "SUPPLIER_CODE"
      FieldName2 = "SUPPLIER_ID"
      Flag = True
   ElseIf UnqType = PARTNO_UNIQUE Then
      TableName = "PART_ITEM"
      FieldName1 = "PART_NO"
      FieldName2 = "PART_ITEM_ID"
      Flag = True
   ElseIf UnqType = QUOATATION_UNIQUE Then
      TableName = "QUOATATION"
      FieldName1 = "QUOATATION_NO"
      FieldName2 = "QUOATATION_ID"
      Flag = True
   ElseIf UnqType = EXPENSE_UNIQUE Then
      TableName = "EXPENSE_GROUP"
      FieldName1 = "GROUP_NO"
      FieldName2 = "EXPENSE_GROUP_ID"
      Flag = True
   ElseIf UnqType = REVENUE_UNIQUE Then
      TableName = "REVENUE_GROUP"
      FieldName1 = "GROUP_NO"
      FieldName2 = "REVENUE_GROUP_ID"
      Flag = True
   ElseIf UnqType = PO_UNIQUE Then
      TableName = "PURCHASE_ORDER"
      FieldName1 = "PO_NO"
      FieldName2 = "PO_ID"
      Flag = True
   ElseIf UnqType = CUSTOMER_UNIQUE Then
      TableName = "PATIENT"
      FieldName1 = "PATIENT_CODE"
      FieldName2 = "PATIENT_ID"
      Flag = True
   ElseIf UnqType = BORROW_UNIQUE Then
      TableName = "EMP_RECEIVABLE"
      FieldName1 = "BORROW_NO"
      FieldName2 = "EMP_RECEIVABLE_ID"
      Flag = True
   ElseIf UnqType = TRUCK_UNIQUE Then
      TableName = "RESOURCE"
      FieldName1 = "RESOURCE_NO"
      FieldName2 = "RESOURCE_ID"
      Flag = True
   ElseIf UnqType = JOBPLAN_UNIQUE Then
      TableName = "JOB_PLAN"
      FieldName1 = "PLAN_NO"
      FieldName2 = "JOB_PLAN_ID"
      Flag = True
   ElseIf UnqType = PARTTYPE_NO Then
      TableName = "PART_TYPE"
      FieldName1 = "PART_TYPE_NO"
      FieldName2 = "PART_TYPE_ID"
      Flag = True
   ElseIf UnqType = PARTTYPE_NAME Then
      TableName = "PART_TYPE"
      FieldName1 = "PART_TYPE_NAME"
      FieldName2 = "PART_TYPE_ID"
      Flag = True
   ElseIf UnqType = LOCATION_NO Then
      TableName = "LOCATION"
      FieldName1 = "LOCATION_NO"
      FieldName2 = "LOCATION_ID"
      Flag = True
   ElseIf UnqType = LOCATION_NO_EX Then
      TableName = "LOCATION"
      FieldName1 = "LOCATION_NO"
      FieldName2 = "LOCATION_TYPE"
      Flag = True
   ElseIf UnqType = LOCATION_NAME Then
      TableName = "LOCATION"
      FieldName1 = "LOCATION_NAME"
      FieldName2 = "LOCATION_ID"
      Flag = True
   ElseIf UnqType = PRODUCTTYPE_NO Then
      TableName = "PRODUCT_TYPE"
      FieldName1 = "PRODUCT_TYPE_NO"
      FieldName2 = "PRODUCT_TYPE_ID"
      Flag = True
   ElseIf UnqType = PRODUCTTYPE_NAME Then
      TableName = "PRODUCT_TYPE"
      FieldName1 = "PRODUCT_TYPE_NAME"
      FieldName2 = "PRODUCT_TYPE_ID"
      Flag = True
   ElseIf UnqType = PRODUCTSTATUS_NO Then
      TableName = "PRODUCT_STATUS"
      FieldName1 = "PRODUCT_STATUS_NO"
      FieldName2 = "PRODUCT_STATUS_ID"
      Flag = True
   ElseIf UnqType = PRODUCTSTATUS_NAME Then
      TableName = "PRODUCT_STATUS"
      FieldName1 = "PRODUCT_STATUS_NAME"
      FieldName2 = "PRODUCT_STATUS_ID"
      Flag = True
   ElseIf UnqType = HOUSE_NO Then
      TableName = "HOUSE"
      FieldName1 = "HOUSE_NO"
      FieldName2 = "HOUSE_ID"
      Flag = True
   ElseIf UnqType = HOUSE_NAME Then
      TableName = "HOUSE"
      FieldName1 = "HOUSE_NAME"
      FieldName2 = "HOUSE_ID"
      Flag = True
   ElseIf UnqType = COUNTRY_NO Then
      TableName = "COUNTRY"
      FieldName1 = "COUNTRY_NO"
      FieldName2 = "COUNTRY_ID"
      Flag = True
   ElseIf UnqType = COUNTRY_NAME Then
      TableName = "COUNTRY"
      FieldName1 = "COUNTRY_NAME"
      FieldName2 = "COUNTRY_ID"
      Flag = True
   ElseIf UnqType = CSTGRADE_NO Then
      TableName = "CUSTOMER_GRADE"
      FieldName1 = "CSTGRADE_NO"
      FieldName2 = "CSTGRADE_ID"
      Flag = True
   ElseIf UnqType = CSTGRADE_NAME Then
      TableName = "CUSTOMER_GRADE"
      FieldName1 = "CSTGRADE_NAME"
      FieldName2 = "CSTGRADE_ID"
      Flag = True
   ElseIf UnqType = CSTTYPE_NO Then
      TableName = "CUSTOMER_TYPE"
      FieldName1 = "CSTTYPE_NO"
      FieldName2 = "CSTTYPE_ID"
      Flag = True
   ElseIf UnqType = CSTTYPE_NAME Then
      TableName = "CUSTOMER_TYPE"
      FieldName1 = "CSTTYPE_NAME"
      FieldName2 = "CSTTYPE_ID"
      Flag = True
   ElseIf UnqType = CUSTCODE_UNIQUE Then
      TableName = "CUSTOMER"
      FieldName1 = "CUSTOMER_CODE"
      FieldName2 = "CUSTOMER_ID"
      Flag = True
   ElseIf UnqType = SUPPLIERGRADE_NO Then
      TableName = "SUPPLIER_GRADE"
      FieldName1 = "SUPPLIER_GRADE_NO"
      FieldName2 = "SUPPLIER_GRADE_ID"
      Flag = True
   ElseIf UnqType = SUPPLIERGRADE_NAME Then
      TableName = "SUPPLIER_GRADE"
      FieldName1 = "SUPPLIER_GRADE_NAME"
      FieldName2 = "SUPPLIER_GRADE_ID"
      Flag = True
   ElseIf UnqType = SUPPLIERTYPE_NO Then
      TableName = "SUPPLIER_TYPE"
      FieldName1 = "SUPPLIER_TYPE_NO"
      FieldName2 = "SUPPLIER_TYPE_ID"
      Flag = True
   ElseIf UnqType = SUPPLIERYPE_NAME Then
      TableName = "SUPPLIER_TYPE"
      FieldName1 = "SUPPLIER_TYPE_NAME"
      FieldName2 = "SUPPLIER_TYPE_ID"
      Flag = True
   ElseIf UnqType = SUPPLIERSTATUS_NO Then
      TableName = "SUPPLIER_STATUS"
      FieldName1 = "SUPPLIER_STATUS_NO"
      FieldName2 = "SUPPLIER_STATUS_ID"
      Flag = True
   ElseIf UnqType = SUPPLIERSTATUS_NAME Then
      TableName = "SUPPLIER_STATUS"
      FieldName1 = "SUPPLIER_STATUS_NAME"
      FieldName2 = "SUPPLIER_STATUS_ID"
      Flag = True
   ElseIf UnqType = POSITION_NO Then
      TableName = "EMP_POSITION"
      FieldName1 = "POSITION_NAME"
      FieldName2 = "POSITION_ID"
      Flag = True
   ElseIf UnqType = UNIT_NO Then
      TableName = "UNIT"
      FieldName1 = "UNIT_NO"
      FieldName2 = "UNIT_ID"
      Flag = True
   ElseIf UnqType = UNIT_NAME Then
      TableName = "UNIT"
      FieldName1 = "UNIT_NAME"
      FieldName2 = "UNIT_ID"
      Flag = True
   ElseIf UnqType = YEAR_NO Then
      TableName = "YEAR_SEQ"
      FieldName1 = "YEAR_NO"
      FieldName2 = "YEAR_SEQ_ID"
      Flag = True
   ElseIf UnqType = PARTGROUP_NO Then
      TableName = "PART_GROUP"
      FieldName1 = "PART_GROUP_NO"
      FieldName2 = "PART_GROUP_ID"
      Flag = True
   ElseIf UnqType = PARTGROUP_NAME Then
      TableName = "PART_GROUP"
      FieldName1 = "PART_GROUP_NAME"
      FieldName2 = "PART_GROUP_ID"
      Flag = True
   ElseIf UnqType = DO_PLAN_UNIQUE Then
      TableName = "BILLING_DOC"
      FieldName1 = "DOCUMENT_NO"
      FieldName2 = "BILLING_DOC_ID"
      Flag = True
    ElseIf UnqType = PACKAGE_CODE Then
      TableName = "PACKAGE"
      FieldName1 = "PKG_CODE"
      FieldName2 = "PKG_ID"
      Flag = True
    ElseIf UnqType = PACKAGE_NAME Then
      TableName = "PACKAGE"
      FieldName1 = "PKG_NAME"
      FieldName2 = "PKG_ID"
      Flag = True
   ElseIf UnqType = PACKAGE_BASIC Then
      TableName = "PACKAGE"
      FieldName1 = "PKG_BASIC_FLAG"
      FieldName2 = "PKG_ID"
      FieldNameExTend = "PKG_TYPE"
      Flag = True
   ElseIf UnqType = PRICE_ADJUST Then
      TableName = "PRICE_ADJUST"
      FieldName1 = "PART_ITEM_ID"
      FieldName2 = "PRICE_ADJUST_ID"
      Flag = True
 ElseIf UnqType = EXPOSE_TYPE_NO Then
      TableName = "EXPOSE_TYPE"
      FieldName1 = "EXPOSE_TYPE_NO"
      FieldName2 = "EXPOSE_TYPE_ID"
      Flag = True
  ElseIf UnqType = EXPOSE_TYPE_NAME Then
      TableName = "EXPOSE_TYPE"
      FieldName1 = "EXPOSE_TYPE_NAME"
      FieldName2 = "EXPOSE_TYPE_ID"
      Flag = True
  End If
   
   If Flag Then
      Count = glbDatabaseMngr.CountRecord(TableName, FieldName1, FieldName2, Key, ID, glbErrorLog, FieldNameExTend, FieldNameExTendValue)
      If Count <> 0 Then
         CheckUniqueNs = False
      Else
         CheckUniqueNs = True
      End If
   End If
      
   Exit Function
ErrorHandler:
   glbErrorLog.LocalErrorMsg = "Runtime error."
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
   CheckUniqueNs = False
End Function

Public Function Check2FlagConvert(A As Long) As String
   If A = 1 Then
      Check2FlagConvert = "N"
   Else
      Check2FlagConvert = "Y"
   End If
End Function

Public Function FlagToCheck(F As String) As Long
   If F = "Y" Then
      FlagToCheck = 1
   Else
      FlagToCheck = 0
   End If
End Function

Public Function Minus2Zero(A As Double) As Double
   If A < 0 Then
      Minus2Zero = 0
   Else
      Minus2Zero = A
   End If
End Function

Public Function Zero2One(A As Double) As Long
   If A = 0 Then
      Zero2One = 1
   Else
      Zero2One = A
   End If
End Function

Public Function Minus2Flag(A As Double) As String
   If A < 0 Then
      Minus2Flag = "Y"
   Else
      Minus2Flag = "N"
   End If
End Function

Public Function AdjustPage(Vsp As VSPrinter, Header As String, Body As String, offset As Long, Optional TestFlag As Boolean = False, Optional SpaceCount As Long) As Boolean
Dim TempStr As String

   TempStr = Header & Body
   Vsp.CalcTable = TempStr
   
   If (Vsp.Y1 + offset - SpaceCount) > (Vsp.PageHeight - Vsp.MarginBottom) Then
      If Not TestFlag Then
         Vsp.NewPage
      End If
      AdjustPage = True
   Else
      AdjustPage = False
   End If
End Function
Public Sub AddMemoNote()
Dim itemcount As Long
Dim OKClick As Boolean

   frmAddEditMemoNote.HeaderText = MapText("���� MEMO")
   frmAddEditMemoNote.ShowMode = SHOW_ADD
   Load frmAddEditMemoNote
   frmAddEditMemoNote.Show 1
   
   OKClick = frmAddEditMemoNote.OKClick
   
   Unload frmAddEditMemoNote
   Set frmAddEditMemoNote = Nothing
   
End Sub

Public Function PatchTable(Vsp As VSPrinter, Header As String, Body As String, offset As Long, Optional EnableFlag As Boolean = True, Optional SpaceCount As Long = 0) As Boolean
Dim TempStr As String
   
   If Not EnableFlag Then
      PatchTable = True
      Exit Function
   End If
   
   TempStr = Header & Body
   Vsp.CalcTable = TempStr
   
   While Not AdjustPage(Vsp, Header, Body, offset, True, SpaceCount)
      Call Vsp.AddTable(Header, "", Body)
   Wend
End Function
Public Function PatchWildCard(T As String) As String
   If Len(Trim(T)) <> 0 Then
      PatchWildCard = T & "%"
   Else
      PatchWildCard = T
   End If
End Function

Public Sub PatchDB()
Dim p As CPatch

   Set p = New CPatch
   
   If Not p.IsPatch("1_0_12_2") Then
      Call p.Patch_1_0_12_2
   End If
   
   If Not p.IsPatch("1_0_12_3") Then
      Call p.Patch_1_0_12_3
   End If
   
   If Not p.IsPatch("1_0_12_9") Then
      Call p.Patch_1_0_12_9
   End If
   
   If Not p.IsPatch("1_0_12_10") Then
      Call p.Patch_1_0_12_10
   End If
   
   If Not p.IsPatch("1_0_12_11") Then
      Call p.Patch_1_0_12_11
   End If
   
   If Not p.IsPatch("1_0_12_12") Then
      Call p.Patch_1_0_12_12
   End If
   
   If Not p.IsPatch("1_0_12_13") Then
      Call p.Patch_1_0_12_13
   End If
   
   If Not p.IsPatch("1_0_12_14") Then
      Call p.Patch_1_0_12_14
   End If
   
   If Not p.IsPatch("1_0_12_15") Then
      Call p.Patch_1_0_12_15
   End If
   
   If Not p.IsPatch("1_0_12_16") Then
      Call p.Patch_1_0_12_16
   End If
   
   If Not p.IsPatch("1_0_12_17") Then
      Call p.Patch_1_0_12_17
   End If
      
   If Not p.IsPatch("1_0_12_18") Then
      Call p.Patch_1_0_12_18
   End If
   
   If Not p.IsPatch("1_0_12_19") Then
      Call p.Patch_1_0_12_19
   End If
   
   If Not p.IsPatch("1_0_12_20") Then
      Call p.Patch_1_0_12_20
   End If
   
   If Not p.IsPatch("1_0_12_21") Then
      Call p.Patch_1_0_12_21
   End If
   
   If Not p.IsPatch("1_0_12_22") Then
      Call p.Patch_1_0_12_22
   End If
   
   If Not p.IsPatch("1_0_12_23") Then
      Call p.Patch_1_0_12_23
   End If
   
   If Not p.IsPatch("1_0_12_24") Then
      Call p.Patch_1_0_12_24
   End If
   
   If Not p.IsPatch("1_0_12_25") Then
      Call p.Patch_1_0_12_25
   End If
   
   If Not p.IsPatch("1_0_12_26") Then
      Call p.Patch_1_0_12_26
   End If
   
   If Not p.IsPatch("1_0_12_27") Then
      Call p.Patch_1_0_12_27
   End If
   
   If Not p.IsPatch("1_0_12_28") Then
      Call p.Patch_1_0_12_28
   End If
   
   If Not p.IsPatch("1_0_12_29") Then
      Call p.Patch_1_0_12_29
   End If
   
   If Not p.IsPatch("1_0_12_30") Then
      Call p.Patch_1_0_12_30
   End If
   
   If Not p.IsPatch("1_0_12_31") Then
      Call p.Patch_1_0_12_31
   End If
   
   If Not p.IsPatch("1_0_12_32") Then
      Call p.Patch_1_0_12_32
   End If
   
   If Not p.IsPatch("1_0_12_33") Then
      Call p.Patch_1_0_12_33
   End If
   
   If Not p.IsPatch("1_0_12_34") Then
      Call p.Patch_1_0_12_34
   End If
   
   If Not p.IsPatch("1_0_12_35") Then
      Call p.Patch_1_0_12_35
   End If
   
   If Not p.IsPatch("1_0_12_36") Then
      Call p.Patch_1_0_12_36
   End If
   
   If Not p.IsPatch("1_0_12_37") Then
      Call p.Patch_1_0_12_37
   End If
   
   If Not p.IsPatch("1_0_12_38") Then
      Call p.Patch_1_0_12_38
   End If
   
   If Not p.IsPatch("1_0_12_39") Then
      Call p.Patch_1_0_12_39
   End If
   
   If Not p.IsPatch("1_0_12_40") Then
      Call p.Patch_1_0_12_40
   End If
   
   If Not p.IsPatch("1_0_12_41") Then
      Call p.Patch_1_0_12_41
   End If
   
   If Not p.IsPatch("1_0_12_42") Then
      Call p.Patch_1_0_12_42
   End If
   
   If Not p.IsPatch("1_0_12_43") Then
      Call p.Patch_1_0_12_43
   End If
      
   If Not p.IsPatch("1_0_12_44") Then
      Call p.Patch_1_0_12_44
   End If
   
   If Not p.IsPatch("1_0_12_45") Then
      Call p.Patch_1_0_12_45
   End If
   
   If Not p.IsPatch("1_0_12_46") Then
      Call p.Patch_1_0_12_46
   End If
   
   If Not p.IsPatch("1_0_12_47") Then
      Call p.Patch_1_0_12_47
   End If
   
   If Not p.IsPatch("1_0_12_48") Then
      Call p.Patch_1_0_12_48
   End If
   
   If Not p.IsPatch("1_0_12_49") Then
      Call p.Patch_1_0_12_49
   End If
   
   If Not p.IsPatch("1_0_12_50") Then
      Call p.Patch_1_0_12_50
   End If
   
   If Not p.IsPatch("1_0_12_51") Then
      Call p.Patch_1_0_12_51
   End If
   
   If Not p.IsPatch("1_0_12_52") Then
      Call p.Patch_1_0_12_52
   End If
   
   If Not p.IsPatch("1_0_12_53") Then
      Call p.Patch_1_0_12_53
   End If
   
   If Not p.IsPatch("1_0_12_54") Then
      Call p.Patch_1_0_12_54
   End If
   
   If Not p.IsPatch("1_0_12_55") Then
      Call p.Patch_1_0_12_55
   End If
   
   If Not p.IsPatch("1_0_12_56") Then
      Call p.Patch_1_0_12_56
   End If
   
   If Not p.IsPatch("1_0_12_57") Then
      Call p.Patch_1_0_12_57
   End If
   
   If Not p.IsPatch("1_0_12_59") Then
      Call p.Patch_1_0_12_59
   End If

   If Not p.IsPatch("1_0_12_60") Then
      Call p.Patch_1_0_12_60
   End If

   If Not p.IsPatch("1_0_12_61") Then
      Call p.Patch_1_0_12_61
   End If

   If Not p.IsPatch("1_0_12_62") Then
      Call p.Patch_1_0_12_62
   End If

   If Not p.IsPatch("1_0_12_63") Then
      Call p.Patch_1_0_12_63
   End If
   
   If Not p.IsPatch("1_0_12_64") Then
      Call p.Patch_1_0_12_64
   End If
   
   If Not p.IsPatch("1_0_12_65") Then
      Call p.Patch_1_0_12_65
   End If
   
   If Not p.IsPatch("1_0_12_66") Then
      Call p.Patch_1_0_12_66
   End If
   
   If Not p.IsPatch("1_0_12_67") Then
      Call p.Patch_1_0_12_67
   End If
   
   If Not p.IsPatch("1_0_12_68") Then
      Call p.Patch_1_0_12_68
   End If
   
   If Not p.IsPatch("1_0_12_69") Then
      Call p.Patch_1_0_12_69
   End If
      
   If Not p.IsPatch("1_0_12_70") Then
      Call p.Patch_1_0_12_70
   End If
      
   If Not p.IsPatch("1_0_12_71") Then
      Call p.Patch_1_0_12_71
   End If
      
   If Not p.IsPatch("1_0_12_72") Then
      Call p.Patch_1_0_12_72
   End If
      
   If Not p.IsPatch("1_0_12_73") Then
      Call p.Patch_1_0_12_73
   End If
      
   If Not p.IsPatch("1_0_12_74") Then
      Call p.Patch_1_0_12_74
   End If
      
   If Not p.IsPatch("1_0_12_75") Then
      Call p.Patch_1_0_12_75
   End If
      
   If Not p.IsPatch("1_0_12_76") Then
      Call p.Patch_1_0_12_76
   End If
      
   If Not p.IsPatch("1_0_12_77") Then
      Call p.Patch_1_0_12_77
   End If
      
   If Not p.IsPatch("1_0_12_78") Then
      Call p.Patch_1_0_12_78
   End If
      
   If Not p.IsPatch("1_0_12_79") Then
      Call p.Patch_1_0_12_79
   End If
      
   If Not p.IsPatch("1_0_12_80") Then
      Call p.Patch_1_0_12_80
   End If
      
   If Not p.IsPatch("1_0_12_81") Then
      Call p.Patch_1_0_12_81
   End If
      
   If Not p.IsPatch("1_0_12_82") Then
      Call p.Patch_1_0_12_82
   End If
      
   If Not p.IsPatch("1_0_12_83") Then
      Call p.Patch_1_0_12_83
   End If
      
   If Not p.IsPatch("1_0_12_84") Then
      Call p.Patch_1_0_12_84
   End If
      
   If Not p.IsPatch("1_0_12_85") Then
      Call p.Patch_1_0_12_85
   End If
      
   If Not p.IsPatch("1_0_12_86") Then
      Call p.Patch_1_0_12_86
   End If
      
   If Not p.IsPatch("1_0_12_87") Then
      Call p.Patch_1_0_12_87
   End If

   If Not p.IsPatch("1_0_12_88") Then
      Call p.Patch_1_0_12_88
   End If
      
   If Not p.IsPatch("1_0_12_89") Then
      Call p.Patch_1_0_12_89
   End If

   If Not p.IsPatch("1_0_12_91") Then
      Call p.Patch_1_0_12_91
   End If
      
   If Not p.IsPatch("1_0_12_92") Then
      Call p.Patch_1_0_12_92
   End If
      
   If Not p.IsPatch("1_0_12_93") Then
      Call p.Patch_1_0_12_93
   End If

   If Not p.IsPatch("1_0_12_94") Then
      Call p.Patch_1_0_12_94
   End If

   If Not p.IsPatch("1_0_12_95") Then
      Call p.Patch_1_0_12_95
   End If

   If Not p.IsPatch("1_0_12_96") Then
      Call p.Patch_1_0_12_96
   End If
      
   If Not p.IsPatch("1_0_12_97") Then
      Call p.Patch_1_0_12_97
   End If
      
   If Not p.IsPatch("1_0_12_98") Then
      Call p.Patch_1_0_12_98
   End If
      
   If Not p.IsPatch("1_0_12_98_jill") Then
      Call p.Patch_1_0_12_98_jill
   End If
      
   If Not p.IsPatch("1_0_12_99_jill") Then
      Call p.Patch_1_0_12_99_jill
   End If
      
   If Not p.IsPatch("1_0_12_100_jill") Then
      Call p.Patch_1_0_12_100_jill
   End If
      
   If Not p.IsPatch("1_0_12_101") Then
      Call p.Patch_1_0_12_101
   End If
      
   If Not p.IsPatch("1_0_12_102_jill") Then
      Call p.Patch_1_0_12_102_jill
   End If
      
   If Not p.IsPatch("1_0_12_103_jill") Then
      Call p.Patch_1_0_12_103_jill
   End If
      
   If Not p.IsPatch("1_0_12_104_jill") Then
      Call p.Patch_1_0_12_104_jill
   End If
      
   If Not p.IsPatch("1_0_12_105_jill") Then
      Call p.Patch_1_0_12_105_jill
   End If
      
   If Not p.IsPatch("1_0_12_106_jill") Then
      Call p.Patch_1_0_12_106_jill
   End If
      
   If Not p.IsPatch("1_0_12_107_jill") Then
      Call p.Patch_1_0_12_107_jill
   End If
            
   If Not p.IsPatch("1_0_12_108_jill") Then
      Call p.Patch_1_0_12_108_jill
   End If
                        
   If Not p.IsPatch("1_0_12_109") Then
      Call p.Patch_1_0_12_109
   End If
                        
   If Not p.IsPatch("1_0_12_110_jill") Then
      Call p.Patch_1_0_12_110_jill
   End If
                        
   If Not p.IsPatch("1_0_12_111_jill") Then
      Call p.Patch_1_0_12_111_jill
   End If
   
   If Not p.IsPatch("1_0_12_112_jill") Then
      Call p.Patch_1_0_12_112_jill
   End If
                        
   If Not p.IsPatch("1_0_12_113_jill") Then
      Call p.Patch_1_0_12_113_jill
   End If
                        
   If Not p.IsPatch("1_0_12_114_jill") Then
      Call p.Patch_1_0_12_114_jill
   End If
                        
   If Not p.IsPatch("1_0_12_115_jill") Then
      Call p.Patch_1_0_12_115_jill
   End If
                        
   If Not p.IsPatch("1_0_12_116_jill") Then
      Call p.Patch_1_0_12_116_jill
   End If
                        
   If Not p.IsPatch("1_0_12_116_seub") Then
      Call p.Patch_1_0_12_116_seub
   End If
                        
   If Not p.IsPatch("1_0_12_117_seub") Then
      Call p.Patch_1_0_12_117_seub
   End If
   
   If Not p.IsPatch("1_0_12_118_jill") Then
      Call p.Patch_1_0_12_118_jill
   End If
                        
   If Not p.IsPatch("1_0_12_119_seub") Then
      Call p.Patch_1_0_12_119_seub
   End If
                        
   If Not p.IsPatch("1_0_12_120_jill") Then
      Call p.Patch_1_0_12_120_jill
   End If
   
   If Not p.IsPatch("2006_11_07_1_jill") Then '1
      Call p.Patch_2006_11_07_1_jill
   End If
   
   If Not p.IsPatch("2006_11_07_2_jill") Then '2
      Call p.Patch_2006_11_07_2_jill
   End If
          
   If Not p.IsPatch("2006_11_10_1_jill") Then '3
      Call p.Patch_2006_11_10_1_jill
   End If
          
   If Not p.IsPatch("2006_11_29_1_jill") Then '4
      Call p.Patch_2006_11_29_1_jill
   End If
   
   If Not p.IsPatch("2006_11_29_2_jill") Then '5
      Call p.Patch_2006_11_29_2_jill
   End If
   
   If Not p.IsPatch("2006_11_29_3_jill") Then '6
      Call p.Patch_2006_11_29_3_jill
   End If
   
   If Not p.IsPatch("2006_11_29_4_jill") Then '7
      Call p.Patch_2006_11_29_4_jill
   End If
   
   If Not p.IsPatch("2006_11_29_5_jill") Then '8
      Call p.Patch_2006_11_29_5_jill
   End If
   
   If Not p.IsPatch("2006_11_29_6_jill") Then '9
      Call p.Patch_2006_11_29_6_jill
   End If
   
   If Not p.IsPatch("2006_11_29_7_jill") Then '10
      Call p.Patch_2006_11_29_7_jill
   End If
   
   If Not p.IsPatch("2006_11_29_8_jill") Then '11
      Call p.Patch_2006_11_29_8_jill
   End If
   
   If Not p.IsPatch("2006_11_30_1_jill") Then '12
      Call p.Patch_2006_11_30_1_jill
   End If
   
   If Not p.IsPatch("2006_12_04_1_jill") Then '13
      Call p.Patch_2006_12_04_1_jill
   End If
   
   If Not p.IsPatch("2006_12_28_1_jill") Then '14
      Call p.Patch_2006_12_28_1_jill
   End If
   
   If Not p.IsPatch("2007_05_17_1_seub") Then '15
      Call p.Patch_2007_05_17_1_seub
   End If
   
   If Not p.IsPatch("2007_05_17_2_seub") Then '16
      Call p.Patch_2007_05_17_2_seub
   End If
   
   If Not p.IsPatch("2007_05_22_1_seub") Then '17
      Call p.Patch_2007_05_22_1_seub
   End If
   
   If Not p.IsPatch("2007_07_12_1_jill") Then '18
      Call p.Patch_2007_07_12_1_jill
   End If
   
   If Not p.IsPatch("2007_07_19_1_jill") Then '19
      Call p.Patch_2007_07_19_1_jill
   End If
   
   If Not p.IsPatch("2007_07_19_2_jill") Then '20
      Call p.Patch_2007_07_19_2_jill
   End If

   If Not p.IsPatch("2007_08_20_1_jill") Then '21
      Call p.Patch_2007_08_20_1_jill
   End If
   
   If Not p.IsPatch("2007_08_20_2_jill") Then '22
      Call p.Patch_2007_08_20_2_jill
   End If
   
   If Not p.IsPatch("2007_08_20_3_jill") Then '23
      Call p.Patch_2007_08_20_3_jill
   End If
   
   If Not p.IsPatch("2007_08_24_1_jill") Then '24
      Call p.Patch_2007_08_24_1_jill
   End If
   
   If Not p.IsPatch("2007_08_24_2_jill") Then '25
      Call p.Patch_2007_08_24_2_jill
   End If
   
   If Not p.IsPatch("2007_09_11_1_jill") Then '26
      Call p.Patch_2007_09_11_1_jill
   End If
   
   If Not p.IsPatch("2007_09_14_1_jill") Then '27
      Call p.Patch_2007_09_14_1_jill
   End If
   
   If Not p.IsPatch("2007_09_14_2_jill") Then '28
      Call p.Patch_2007_09_14_2_jill
   End If
   
   If Not p.IsPatch("2007_09_14_3_jill") Then '29
      Call p.Patch_2007_09_14_3_jill
   End If
   
   If Not p.IsPatch("2007_09_14_4_jill") Then '30
      Call p.Patch_2007_09_14_4_jill
   End If
   
   If Not p.IsPatch("2007_09_18_1_jill") Then '31
      Call p.Patch_2007_09_18_1_jill
   End If
   
   If Not p.IsPatch("2007_09_25_1_jill") Then '32
      Call p.Patch_2007_09_25_1_jill
   End If
   
   If Not p.IsPatch("2007_10_16_1_jill") Then '33
      Call p.Patch_2007_10_16_1_jill
   End If
   
   If Not p.IsPatch("2007_10_16_2_jill") Then '34
      Call p.Patch_2007_10_16_2_jill
   End If
   
   If Not p.IsPatch("2007_10_16_3_jill") Then '35
      Call p.Patch_2007_10_16_3_jill
   End If
      
   If Not p.IsPatch("2007_10_16_4_jill") Then '36
      Call p.Patch_2007_10_16_4_jill
   End If
   
   If Not p.IsPatch("2007_10_16_5_jill") Then '37
      Call p.Patch_2007_10_16_5_jill
   End If
   
   If Not p.IsPatch("2007_10_18_1_jill") Then '38
      Call p.Patch_2007_10_18_1_jill
   End If
   
   If Not p.IsPatch("2007_10_24_1_jill") Then '39
      Call p.Patch_2007_10_24_1_jill
   End If
   
   If Not p.IsPatch("2007_10_25_1_jill") Then '40
      Call p.Patch_2007_10_25_1_jill
   End If
   
   If Not p.IsPatch("2007_10_29_1_jill") Then '41
      Call p.Patch_2007_10_29_1_jill
   End If
   
   If Not p.IsPatch("2007_10_30_1_jill") Then '42
      Call p.Patch_2007_10_30_1_jill
   End If
   
   If Not p.IsPatch("2007_10_30_2_jill") Then '43
      Call p.Patch_2007_10_30_2_jill
   End If
   
   If Not p.IsPatch("2007_10_30_3_jill") Then '44
      Call p.Patch_2007_10_30_3_jill
   End If
   
   If Not p.IsPatch("2007_11_12_1_jill") Then '45
      Call p.Patch_2007_11_12_1_jill
   End If
   
   If Not p.IsPatch("2007_11_12_2_jill") Then '46
      Call p.Patch_2007_11_12_2_jill
   End If
   
   If Not p.IsPatch("2007_11_12_3_jill") Then '47
      Call p.Patch_2007_11_12_3_jill
   End If
   
   If Not p.IsPatch("2007_11_19_1_jill") Then '48
      Call p.Patch_2007_11_19_1_jill
   End If
   
   If Not p.IsPatch("2007_11_19_2_jill") Then '49
      Call p.Patch_2007_11_19_2_jill
   End If
   
   If Not p.IsPatch("2007_11_26_1_jill") Then '50
      Call p.Patch_2007_11_26_1_jill
   End If
   
   If Not p.IsPatch("2007_11_26_2_jill") Then '51
      Call p.Patch_2007_11_26_2_jill
   End If
   
   If Not p.IsPatch("2007_11_29_1_jill") Then '52
      Call p.Patch_2007_11_29_1_jill
   End If
   
   If Not p.IsPatch("2007_11_29_2_jill") Then '53
      Call p.Patch_2007_11_29_2_jill
   End If
   
   If Not p.IsPatch("2007_12_03_1_jill") Then '54
      Call p.Patch_2007_12_03_1_jill
   End If
   
   If Not p.IsPatch("2007_12_03_2_jill") Then '55
      Call p.Patch_2007_12_03_2_jill
   End If
   
   If Not p.IsPatch("2007_12_06_1_jill") Then '56
      Call p.Patch_2007_12_06_1_jill
   End If
   
   If Not p.IsPatch("2007_12_14_1_jill") Then '57
      Call p.Patch_2007_12_14_1_jill
   End If
   
   If Not p.IsPatch("2007_12_17_1_jill") Then '58
      Call p.Patch_2007_12_17_1_jill
   End If
   
   If Not p.IsPatch("2007_12_25_1_jill") Then '59
      Call p.Patch_2007_12_25_1_jill
   End If
   
'   If Not p.IsPatch("2008_01_08_1_jill") Then '60          '����Ѻ MA2 ���ҧ����� ��ѭ�� ���ҧ�ѻ���� �Դ ��ӫ�͹
'      Call p.Patch_2008_01_08_1_jill
'   End If
   
'   If Not p.IsPatch("2008_05_29_1_jill") Then '61          '����Ѻ MP ���ҧ ����� ��ѭ�� RECEIPT_ITEM ����� BILLING_DOC_ID
'      Call p.Patch_2008_05_29_1_jill
'   End If
   
   If Not p.IsPatch("2009_06_18_1_jill") Then '62
      Call p.Patch_2009_06_18_1_jill
   End If
   
   If Not p.IsPatch("2009_06_18_2_jill") Then '63
      Call p.Patch_2009_06_18_2_jill
   End If
   
   If Not p.IsPatch("2009_07_24_1_jill") Then '64
      Call p.Patch_2009_07_24_1_jill
   End If
   
   If Not p.IsPatch("2009_07_24_2_jill") Then '65
      Call p.Patch_2009_07_24_2_jill
   End If
   
   If Not p.IsPatch("2009_07_24_3_jill") Then '66
      Call p.Patch_2009_07_24_3_jill
   End If
   
   If Not p.IsPatch("2013_08_01_1_ging") Then '67
      Call p.Patch_2013_08_01_1_ging
   End If
   
   If Not p.IsPatch("2013_09_13_1_ging") Then '68
      Call p.Patch_2013_09_13_1_ging
   End If
   
   If Not p.IsPatch("2013_11_04_1_jill") Then '69
      Call p.Patch_2013_11_04_1_jill
   End If
   
   If Not p.IsPatch("2014_02_04_1_jill") Then '70
      Call p.Patch_2014_02_04_1_jill
   End If
   
   If Not p.IsPatch("2017_07_20_dear") Then '71
      Call p.Patch_2017_07_20_dear
   End If
   
   If Not p.IsPatch("2017_08_30_1_dear") Then '72
      Call p.Patch_2017_08_30_1_dear
   End If
   
   If Not p.IsPatch("2017_08_30_2_dear") Then '73
      Call p.Patch_2017_08_30_2_dear
   End If
   
   If Not p.IsPatch("2017_08_31_1_dear") Then '74
      Call p.Patch_2017_08_31_1_dear
   End If
   
   If Not p.IsPatch("2017_08_31_2_dear") Then '75
      Call p.Patch_2017_08_31_2_dear
   End If
   
   If Not p.IsPatch("2017_08_31_3_dear") Then '76
      Call p.Patch_2017_08_31_3_dear
   End If
   
'   If Not p.IsPatch("2017_11_03_1_jill") Then '77
'      Call p.Patch_2017_11_03_1_jill
'   End If
'
'   If Not p.IsPatch("2017_11_03_2_jill") Then '78
'      Call p.Patch_2017_11_03_2_jill
'   End If
   
   If Not p.IsPatch("2017_12_28_1_jill") Then '79
      Call p.Patch_2017_12_28_1_jill
   End If
   
   If Not p.IsPatch("2017_12_28_2_jill") Then '80
      Call p.Patch_2017_12_28_2_jill
   End If
   
   If Not p.IsPatch("2018_05_30_1_jill") Then '81
      Call p.Patch_2018_05_30_1_jill
   End If
   
   If Not p.IsPatch("2018_05_30_2_jill") Then '82
      Call p.Patch_2018_05_30_2_jill
   End If
   
   If Not p.IsPatch("2018_08_16_1_dear") Then '83
      Call p.Patch_2018_08_16_1_dear
   End If
   
   If Not p.IsPatch("2018_08_16_2_dear") Then '84
      Call p.Patch_2018_08_16_2_dear
   End If
   
   Set p = Nothing
End Sub
Public Function MyDiffEx(ByVal D1 As Double, ByVal D2 As Double) As Double
   If D2 = 0 Then
      MyDiffEx = 0
   Else
      MyDiffEx = D1 / D2
   End If
End Function
Public Function MyDiffEx2(ByVal D1 As Double, ByVal D2 As Double) As Double
   If D2 = 0 Then
      MyDiffEx2 = 0
   Else
      MyDiffEx2 = CDbl(Format(D1 / D2, "0.0000"))
   End If
End Function

Public Function MyDiff(ByVal D1 As Double, ByVal D2 As Double) As Double
   If D2 = 0 Then
      MyDiff = 0
   Else
      MyDiff = CDbl(Format(D1 / D2, "0.00"))
   End If
End Function

'Public Sub CheckMemo(TriggerCode As Long)
'Dim M As CMemo
'Dim TempRs As ADODB.Recordset
'Dim ItemCount As Long
'
'   Set M = New CMemo
'   Set TempRs = New ADODB.Recordset
'
'   M.MEMO_ID = -1
'   M.MEMO_STATUS = "N"
'   M.ASSIGN_TO = glbUser.REAL_USER_ID
'   M.FROM_DATE = Now
'   M.TO_DATE = DateAdd("H", 1, M.FROM_DATE)
'   M.TRIGGER_CODE = TriggerCode
'   Call M.QueryData2(TempRs, ItemCount)
'
'   If ItemCount > 0 Then
'      glbErrorLog.LocalErrorMsg = "����¡������͹���֧��˹����� ��ҹ��ͧ��èд���¡��������� ?"
'      If glbErrorLog.AskMessage = vbYes Then
'         frmMemo.MemoStatus = "N"
'         frmMemo.HeaderText = "��Ǩ�ͺ��¡������͹"
'         Load frmMemo
'         frmMemo.Show 1
'
'         Unload frmMemo
'         Set frmMemo = Nothing
'      End If
'   End If
'
'   If TempRs.State = adStateOpen Then
'      TempRs.Close
'   End If
'   Set TempRs = Nothing
'   Set M = Nothing
'End Sub
'
'Public Sub PatchDB()
'Dim p As CPatch
'
'   Set p = New CPatch
'
'   If Not p.IsPatch("3_0_12_19") Then
'      Call p.Patch_3_0_12_19
'   End If
'
'   If Not p.IsPatch("3_0_12_20") Then
'      Call p.Patch_3_0_12_20
'   End If
'
'   If Not p.IsPatch("3_0_12_21") Then
'      Call p.Patch_3_0_12_21
'   End If
'
'   If Not p.IsPatch("3_0_12_22") Then
'      Call p.Patch_3_0_12_22
'   End If
'
'   If Not p.IsPatch("3_0_12_23") Then
'      Call p.Patch_3_0_12_23
'   End If
'
'   Set p = Nothing
'End Sub
'
'Public Function DOType2Flag(DoType As Long) As String
'   If DoType = 1 Then
'      DOType2Flag = "N"
'   ElseIf DoType = 2 Then
'      DOType2Flag = "Y"
'   Else
'      DOType2Flag = ""
'
'   End If
'End Function

Public Function PackAddress(Rs As ADODB.Recordset) As String
Dim AddressStr As String

   AddressStr = ""
   
   If NVLS(Rs("HOME_NO1"), "") <> "" Then
      AddressStr = AddressStr & NVLS(Rs("HOME_NO1"), "") & " "
   End If

   If NVLS(Rs("MOO1"), "") <> "" Then
      AddressStr = AddressStr & "����." & NVLS(Rs("MOO1"), "") & " "
   End If

   If NVLS(Rs("SOI1"), "") <> "" Then
      AddressStr = AddressStr & "���." & NVLS(Rs("SOI1"), "") & " "
   End If

   If NVLS(Rs("ROAD1"), "") <> "" Then
      AddressStr = AddressStr & "�." & NVLS(Rs("ROAD1"), "") & " "
   End If

   If NVLS(Rs("KWANG1"), "") <> "" Then
      AddressStr = AddressStr & "�ǧ" & NVLS(Rs("KWANG1"), "") & " "
   End If

   If NVLS(Rs("KHATE1"), "") <> "" Then
      AddressStr = AddressStr & "ࢵ" & NVLS(Rs("KHATE1"), "") & " "
   End If

   If NVLS(Rs("PROVINCE"), "") <> "" Then
      AddressStr = AddressStr & "�." & NVLS(Rs("PROVINCE"), "") & " "
   End If

   If NVLS(Rs("ZIPCODE1"), "") <> "" Then
      AddressStr = AddressStr & " " & NVLS(Rs("ZIPCODE1"), "") & " "
   End If

   PackAddress = AddressStr
End Function

Public Function MapText(Msg As String) As String
   MapText = Msg
End Function

Public Function SetReportConfig(Vsp As VSPrinter, ReportClassName As String) As Boolean
Dim I As Long
Dim Count As Long
Dim Rp As CReportConfig
Dim TempRs As ADODB.Recordset
Dim Rps As Collection
Dim iCount As Long

   If Rps Is Nothing Then
      Set TempRs = New ADODB.Recordset
      
      Set Rps = New Collection
      Set Rp = New CReportConfig
      
      Rp.REPORT_CONFIG_ID = -1
      Call Rp.QueryData(TempRs, iCount)
      Set Rp = Nothing
      
      While Not TempRs.EOF
         Set Rp = New CReportConfig
         
         Call Rp.PopulateFromRS(1, TempRs)
         Call Rps.Add(Rp)
         
         Set Rp = Nothing
         TempRs.MoveNext
      Wend
      
      Set Rp = Nothing
      If TempRs.State = adStateOpen Then
         TempRs.Close
      End If
      Set TempRs = Nothing
   End If
   
   SetReportConfig = False
   For Each Rp In Rps
      If Rp.REPORT_KEY = ReportClassName Then
         Vsp.PaperSize = Rp.PAPER_SIZE
         Vsp.ORIENTATION = Rp.ORIENTATION
         Vsp.MarginBottom = Rp.MARGIN_BOTTOM * 567
         Vsp.MarginFooter = Rp.MARGIN_FOOTER * 567
         Vsp.MarginHeader = Rp.MARGIN_HEADER * 567
         Vsp.MarginLeft = Rp.MARGIN_LEFT * 567
         Vsp.MarginRight = Rp.MARGIN_RIGHT * 567
         Vsp.MarginTop = Rp.MARGIN_TOP * 567
'         Vsp.FontName = Rp.FONT_NAME
         If Rp.FONT_SIZE > 0 Then
            Vsp.FontSize = Rp.FONT_SIZE
         End If
         Vsp.MarginLeft = Rp.MARGIN_LEFT * 567
         Vsp.MarginRight = Rp.MARGIN_RIGHT * 567
         If Rp.PAPER_HEIGHT > 0 Then
            Vsp.PaperWidth = Rp.PAPER_HEIGHT * 567
         End If
         If Rp.PAPER_WIDTH > 0 Then
            Vsp.PaperHeight = Rp.PAPER_HEIGHT * 567
         End If
         
         SetReportConfig = True
         Exit Function
      End If
   Next Rp
   Set Rps = Nothing
End Function

Public Function GetBalanceItem(Col As Collection, PartItemID As Long, LocationID As Long, DocDate As Date) As Object
Dim D As Object
Dim Key As String
Dim MaxSeq As Long
Dim I As Long
Dim MaxIndex As Long
Static II As CImportItem
Dim MaxDate As Date

   MaxDate = -2
   For Each D In Col
      If (DateToStringInt(D.DOCUMENT_DATE) < DateToStringInt(DocDate)) And (D.PART_ITEM_ID = PartItemID) And (D.LOCATION_ID = LocationID) Then
         If DateToStringInt(D.DOCUMENT_DATE) > DateToStringInt(MaxDate) Then
            MaxDate = InternalDateToDate(DateToStringInt(D.DOCUMENT_DATE))
         End If
      End If
   Next D

   I = 0
   MaxSeq = -1
   MaxIndex = -1
   For Each D In Col
      I = I + 1

      If (D.PART_ITEM_ID = PartItemID) And (D.LOCATION_ID = LocationID) And _
         (DateToStringInt(D.DOCUMENT_DATE) = DateToStringInt(MaxDate)) Then
            If D.TRANSACTION_SEQ > MaxSeq Then
               MaxSeq = D.TRANSACTION_SEQ
               MaxIndex = I
            End If
      End If
   Next D

   If MaxIndex > 0 Then
      Set GetBalanceItem = Col(MaxIndex)
   Else
      If II Is Nothing Then
         Set II = New CImportItem
      End If
      Set GetBalanceItem = II
   End If
End Function

Public Function GetDoItem(m_TempCol As Collection, TempKey As String) As CDoItem
On Error Resume Next
Dim EI As CDoItem
Static TempEi As CDoItem

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CDoItem
      End If
      Set GetDoItem = TempEi
   Else
      Set GetDoItem = EI
   End If
End Function
Public Function GetParamItem(m_TempCol As Collection, TempKey As String) As CParamItem
On Error Resume Next
Dim EI As CParamItem
Static TempEi As CParamItem

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CParamItem
      End If
      Set GetParamItem = TempEi
   Else
      Set GetParamItem = EI
   End If
End Function

Public Function GetRoItem(m_TempCol As Collection, TempKey As String) As CROItem
On Error Resume Next
Dim EI As CROItem
Static TempEi As CROItem

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CROItem
      End If
      Set GetRoItem = TempEi
   Else
      Set GetRoItem = EI
   End If
End Function


Public Function GetProductStatus(m_TempCol As Collection, TempKey As String) As CProductStatus
On Error Resume Next
Dim EI As CProductStatus
Static TempEi As CProductStatus

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CProductStatus
      End If
      Set GetProductStatus = TempEi
   Else
      Set GetProductStatus = EI
   End If
End Function

Public Function GetReceiptItem(m_TempCol As Collection, TempKey As String) As CReceiptItem
On Error Resume Next
Dim EI As CReceiptItem
Static TempEi As CReceiptItem

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CReceiptItem
      End If
      Set GetReceiptItem = TempEi
   Else
      Set GetReceiptItem = EI
   End If
End Function

Public Function GetImportItem(m_TempCol As Collection, TempKey As String) As CImportItem
On Error Resume Next
Dim EI As CImportItem
Static TempEi As CImportItem

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CImportItem
      End If
      Set GetImportItem = TempEi
   Else
      Set GetImportItem = EI
   End If
End Function

Public Function GetMonthlyAccum(m_TempCol As Collection, TempKey As String) As CMonthlyAccum
On Error Resume Next
Dim EI As CMonthlyAccum
Static TempEi As CMonthlyAccum

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CMonthlyAccum
      End If
      Set GetMonthlyAccum = TempEi
   Else
      Set GetMonthlyAccum = EI
   End If
End Function
Public Function GetMonthlyAccumEx(m_TempCol As Collection, TempKey As String) As CMonthlyAccum
On Error Resume Next
Dim EI As CMonthlyAccum
Static TempEi As CMonthlyAccum
   
   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CMonthlyAccum
      End If
      Set GetMonthlyAccumEx = TempEi
   Else
      Set GetMonthlyAccumEx = EI
   End If
End Function

Public Function GetLocation(m_TempCol As Collection, TempKey As String) As CLocation
On Error Resume Next
Dim EI As CLocation
Static TempEi As CLocation

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CLocation
      End If
      Set GetLocation = TempEi
   Else
      Set GetLocation = EI
   End If
End Function

Public Function GetBalanceAccum(m_TempCol As Collection, TempKey As String) As CBalanceAccum
On Error Resume Next
Dim EI As CBalanceAccum
Static TempEi As CBalanceAccum

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CBalanceAccum
      End If
      Set GetBalanceAccum = TempEi
   Else
      Set GetBalanceAccum = EI
   End If
End Function
Public Function GetBankBranch(m_TempCol As Collection, TempKey As String) As CBankBranch
On Error Resume Next
Dim EI As CBankBranch
Static TempEi As CBankBranch

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CBankBranch
      End If
      Set GetBankBranch = TempEi
   Else
      Set GetBankBranch = EI
   End If
End Function
Public Function GetBankAccount(m_TempCol As Collection, TempKey As String) As CBankAccount
On Error Resume Next
Dim EI As CBankAccount
Static TempEi As CBankAccount

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
'         Set TempEi = New CBankAccount
      End If
      Set GetBankAccount = TempEi
   Else
      Set GetBankAccount = EI
   End If
End Function

Public Function GetBillingDoc(m_TempCol As Collection, TempKey As String) As CBillingDoc
On Error Resume Next
Dim EI As CBillingDoc
Static TempEi As CBillingDoc

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CBillingDoc
      End If
      Set GetBillingDoc = TempEi
   Else
      Set GetBillingDoc = EI
   End If
End Function
Public Function GetRegion(m_TempCol As Collection, TempKey As String) As CRegion
On Error Resume Next
Dim EI As CRegion
Static TempEi As CRegion

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CRegion
      End If
      Set GetRegion = TempEi
   Else
      Set GetRegion = EI
   End If
End Function

Public Function GetMovementSearch1(m_TempCol As Collection, TempKey As String) As CMovementItemSearch1
On Error Resume Next
Dim EI As CMovementItemSearch1
Static TempEi As CMovementItemSearch1

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
'      If TempEi Is Nothing Then
'         Set TempEi = New CMovementItemSearch1
'      End If
      Set GetMovementSearch1 = TempEi
   Else
      Set GetMovementSearch1 = EI
   End If
End Function
Public Function GetCostAccumSearch(m_TempCol As Collection, TempKey As String) As CCost_Accum
On Error Resume Next
Dim EI As CCost_Accum
Static TempEi As CCost_Accum

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
'      If TempEi Is Nothing Then
'         Set TempEi = New CCostAccumSearch1
'      End If
      Set GetCostAccumSearch = TempEi
   Else
      Set GetCostAccumSearch = EI
   End If
End Function
Public Function GetCostAccum(m_TempCol As Collection, TempKey As String) As CCost_Accum
On Error Resume Next
Dim EI As CCost_Accum
Static TempEi As CCost_Accum

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CCost_Accum
      End If
      Set GetCostAccum = TempEi
   Else
      Set GetCostAccum = EI
   End If
End Function

Public Function GetCExportId(m_TempCol As Collection, TempKey As String) As CExportId
On Error Resume Next
Dim EI As CExportId
Static TempEi As CExportId
   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
'      If TempEi Is Nothing Then
'         Set TempEi = New CExportId
'      End If
      Set GetCExportId = TempEi
   Else
      Set GetCExportId = EI
   End If
End Function

Public Function GetCapitalMovement(m_TempCol As Collection, TempKey As String) As CCapitalMovement
On Error Resume Next
Dim EI As CCapitalMovement
Static TempEi As CCapitalMovement

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CCapitalMovement
      End If
      Set GetCapitalMovement = TempEi
   Else
      Set GetCapitalMovement = EI
   End If
End Function

Public Function GetMovementItem(m_TempCol As Collection, TempKey As String) As CMovementItem
On Error Resume Next
Dim EI As CMovementItem
Static TempEi As CMovementItem

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CMovementItem
      End If
      Set GetMovementItem = TempEi
   Else
      Set GetMovementItem = EI
   End If
End Function

Public Function GetLossItem(m_TempCol As Collection, TempKey As String) As CLossItem
On Error Resume Next
Dim EI As CLossItem
Static TempEi As CLossItem

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CLossItem
      End If
      Set GetLossItem = TempEi
   Else
      Set GetLossItem = EI
   End If
End Function

Public Function GetRevenueType(m_TempCol As Collection, TempKey As String) As CRevenueType
On Error Resume Next
Dim EI As CRevenueType
Static TempEi As CRevenueType

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CRevenueType
      End If
      Set GetRevenueType = TempEi
   Else
      Set GetRevenueType = EI
   End If
End Function

Public Function GetPartItem(m_TempCol As Collection, TempKey As String) As CPartItem
On Error Resume Next
Dim EI As CPartItem
Static TempEi As CPartItem

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CPartItem
      End If
      Set GetPartItem = TempEi
   Else
      Set GetPartItem = EI
   End If
End Function

Public Function GetParameter(m_TempCol As Collection, TempKey As String) As CParameter
On Error Resume Next
Dim EI As CParameter
Static TempEi As CParameter

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CParameter
      End If
      Set GetParameter = TempEi
   Else
      Set GetParameter = EI
   End If
End Function

Public Function GetPopulation(m_TempCol As Collection, TempKey As String) As CPopulation
On Error Resume Next
Dim EI As CPopulation
Static TempEi As CPopulation

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CPopulation
      Else
         TempEi.PIG_ID = -1
      End If
      Set GetPopulation = TempEi
   Else
      Set GetPopulation = EI
      EI.PIG_ID = TempKey
   End If
End Function

Public Function GetPopulationEx(m_TempCol As Collection, TempKey As String) As CPopulation
On Error Resume Next
Dim EI As CPopulation

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      Set GetPopulationEx = Nothing
   Else
      Set GetPopulationEx = EI
   End If
End Function
Public Function GetGlAgeAmount(m_TempCol As Collection, TempKey As String) As CGLAgeAmount
On Error Resume Next
Dim EI As CGLAgeAmount

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      Set GetGlAgeAmount = Nothing
   Else
      Set GetGlAgeAmount = EI
   End If
End Function

Public Function GetCustomer(m_TempCol As Collection, TempKey As String) As CCustomer
On Error Resume Next
Dim EI As CCustomer
Static TempEi As CCustomer

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CCustomer
      End If
      Set GetCustomer = TempEi
   Else
      Set GetCustomer = EI
   End If
End Function
Public Function GetEmployee(m_TempCol As Collection, TempKey As String) As CEmployee
On Error Resume Next
Dim EI As CEmployee
Static TempEi As CEmployee

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CEmployee
      End If
      Set GetEmployee = TempEi
   Else
      Set GetEmployee = EI
   End If
End Function
Public Function GetAccount(m_TempCol As Collection, TempKey As String) As CAccount
On Error Resume Next
Dim EI As CAccount
Static TempEi As CAccount

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CAccount
      End If
      Set GetAccount = TempEi
   Else
      Set GetAccount = EI
   End If
End Function

Public Function GetSystemParam(m_TempCol As Collection, TempKey As String) As CSystemParam
On Error Resume Next
Dim EI As CSystemParam
Static TempEi As CSystemParam

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CSystemParam
      End If
      Set GetSystemParam = TempEi
   Else
      Set GetSystemParam = EI
   End If
End Function
Public Function GetExportItem(TempCol As Collection, TempKey As String) As CExportItem
On Error Resume Next
Dim EI As CExportItem
Static TempEi As CExportItem

   Set EI = TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CExportItem
      End If
      Set GetExportItem = TempEi
   Else
      Set GetExportItem = EI
   End If
End Function

Public Function GetExpenseRatio(TempCol As Collection, TempKey As String) As CExpenseRatio
On Error Resume Next
Dim EI As CExpenseRatio
Static TempEi As CExpenseRatio

   Set EI = TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CExpenseRatio
      End If
      Set GetExpenseRatio = TempEi
   Else
      Set GetExpenseRatio = EI
   End If
End Function

Public Function GetAge(PigCode As String, TempDate As Date) As Double
Dim Yw As CYearWeek
Static TempCol As Collection
Dim WeekCount As Double
   
   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadYearWeek(Nothing, TempCol)
   End If
   
   For Each Yw In TempCol
      If Yw.YEAR_NO & Format(Yw.WEEK_NO, "00") = PigCode Then
         If Yw.FROM_DATE < 0 Then
            WeekCount = 999
         Else
            WeekCount = DateDiff("W", Yw.FROM_DATE, TempDate)
         End If
         Exit For
      End If
   Next Yw
   
   GetAge = WeekCount
End Function
Public Function GetAgeDay(PigCode As String, TempDate As Date) As Double
Dim Yw As CYearWeek
Static TempCol As Collection
Dim DayCount As Double
   
   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadYearWeek(Nothing, TempCol)
   End If
   
   For Each Yw In TempCol
      If Yw.YEAR_NO & Format(Yw.WEEK_NO, "00") = PigCode Then
         If Yw.FROM_DATE < 0 Then
            DayCount = 999
         Else
            DayCount = DateDiff("d", Yw.TO_DATE, TempDate)
         End If
         Exit For
      End If
   Next Yw
   
   GetAgeDay = DayCount
  
End Function

Private Function CompareAge(A1 As Double, A2 As Double) As Boolean
   If A2 <= 0 Then
      CompareAge = (A1 >= A2)
   Else
      CompareAge = (A1 > A2)
   End If
End Function

Public Function GetAgeCode(AGE As Double) As String
Dim Ar As CAgeRange
Static TempCol As Collection
Dim AgeCode As String

   If TempCol Is Nothing Then
      Set TempCol = New Collection
      Call LoadAgeRange(Nothing, TempCol)
   End If
   
   For Each Ar In TempCol
      If CompareAge(AGE, Ar.FROM_WEEK) And (AGE <= Ar.TO_WEEK) Then
         AgeCode = Ar.AGE_RANGE_NO
         Exit For
      End If
   Next Ar
   
   GetAgeCode = AgeCode
End Function

Private Function CompareKey(D1 As CPartItem, D2 As CPartItem, TempID As Long) As Boolean
Dim OrderID As Long
Dim OrderType As Long
Dim TempResult As Boolean
   

   If D1.PIG_AGE = D2.PIG_AGE Then
      TempResult = D1.PIG_TYPE > D2.PIG_TYPE
   Else
      TempResult = D1.PIG_AGE < D2.PIG_AGE
   End If
         
   CompareKey = TempResult
End Function

Public Sub Selectionsort(List As Collection, MIN As Long, MAX As Long, TempID As Long)
Dim I As Long
Dim j As Long
Dim best_value As CPartItem
Dim Temp As CPartItem
Dim best_j As Integer

   Set best_value = New CPartItem
    For I = MIN To MAX - 1
        Set best_value = List(I)
        best_j = I
        For j = I + 1 To MAX
            If CompareKey(List(j), best_value, TempID) Then
                Set best_value = List(j)
                best_j = j
            End If
        Next j
        
        Set Temp = List(I)
        List.Remove (best_j)
        If best_j > List.Count Then
         Call List.Add(Temp, , , best_j - 1)
      Else
         Call List.Add(Temp, , best_j)
      End If
    
        List.Remove (I)
        If I > List.Count Then
         Call List.Add(best_value, , , I - 1)
      Else
         Call List.Add(best_value, , I)
      End If
    Next I
    Set best_value = Nothing
End Sub

Public Function LastDayOfMonth(ByVal ValidDate As Date) As Byte
Dim LastDay As Byte
   LastDay = DatePart("d", DateAdd("d", -1, DateAdd("m", 1, DateAdd("d", -DatePart("d", ValidDate) + 1, ValidDate))))
   LastDayOfMonth = LastDay
End Function

Public Sub LoadPictureFromFile(FileName As String, Pc As PictureBox)
On Error Resume Next
    If Dir(FileName) <> "" Then
      Pc.Picture = LoadPicture(FileName)
   End If
End Sub

Public Sub GetFirstLastDate(D As Date, FD As Date, Ld As Date)
Dim MM As Long
Dim DD1 As Long
Dim DD2 As Long
Dim YYYY As Long

   MM = Month(D)
   DD1 = 1
   DD2 = LastDayOfMonth(D)
   YYYY = Year(D)
   
   FD = DateSerial(YYYY, MM, DD1)
   Ld = DateSerial(YYYY, MM, DD2)
End Sub

Public Sub CalculateIncludePrice(Ivd As CInventoryDoc, TotalUnit As Double, DeliveryFee As Double)
Dim II As CImportItem
Dim AvgFee As Double
Dim SumUnit As Double

   SumUnit = 0
   For Each II In Ivd.ImportExports
      If II.Flag <> "D" Then
         SumUnit = SumUnit + II.IMPORT_AMOUNT
      End If
   Next II

   If SumUnit > 0 Then
      AvgFee = MyDiffEx(DeliveryFee, SumUnit)
   Else
      AvgFee = 0
   End If
      
   For Each II In Ivd.ImportExports
      If II.Flag <> "D" Then
         II.TOTAL_INCLUDE_PRICE = II.TOTAL_ACTUAL_PRICE + (AvgFee * II.IMPORT_AMOUNT)
         II.INCLUDE_UNIT_PRICE = MyDiffEx(II.TOTAL_INCLUDE_PRICE, II.IMPORT_AMOUNT)
         
         If II.Flag <> "A" Then
            II.Flag = "E"
         End If
      End If
   Next II
End Sub

Public Sub StartExportFile(Vsp As VSPrinter)
   Vsp.ExportFile = ""
   Vsp.ExportFile = glbParameterObj.ReportFile
   Vsp.ExportFormat = vpxPlainHTML
End Sub

Public Sub CloseExportFile(Vsp As VSPrinter)
   Vsp.ExportFile = ""
End Sub

Public Sub BalanceAccum2ImportItem(Ba As CBalanceAccum, II As CImportItem)
   II.LOCATION_ID = Ba.LOCATION_ID
   II.PART_ITEM_ID = Ba.PART_ITEM_ID
   II.CURRENT_AMOUNT = Ba.BALANCE_AMOUNT
   II.NEW_PRICE = MyDiffEx(Ba.TOTAL_INCLUDE_PRICE, Ba.BALANCE_AMOUNT)
   II.TOTAL_INCLUDE_PRICE = Ba.TOTAL_INCLUDE_PRICE
   II.TX_TYPE = "I"
End Sub

Public Function DocType2Set(DocType As Long) As String
   If DocType = 10 Then
      DocType2Set = "(10)"
   ElseIf DocType = 13 Then
      DocType2Set = "(13)"
   Else
      DocType2Set = "(10, 13)"
   End If
End Function

Public Function BillingDocType2Set(DocType As Long) As String
   If DocType = 1 Then
      BillingDocType2Set = "(1)"
   ElseIf DocType = 2 Then
      BillingDocType2Set = "(2)"
   Else
      BillingDocType2Set = "(1, 2)"
   End If
End Function

Public Sub PopulateInternalField(ShowMode As SHOW_MODE_TYPE, O As Object)
Dim Tf As CTableField
Dim TempID As Long
Dim InternalDate As String

   For Each Tf In O.m_FieldList
      If Tf.FieldCat = ID_CAT Then
         If ShowMode = SHOW_ADD Then
            Call glbDatabaseMngr.GetSeqID(O.SequenceName, TempID, glbErrorLog)
            Call Tf.SetValue(TempID)
         End If
      ElseIf Tf.FieldCat = CREATE_DATE_CAT Then
         If ShowMode = SHOW_ADD Then
            Call glbDatabaseMngr.GetServerDateTime(InternalDate, glbErrorLog)
            Call Tf.SetValue(InternalDateToDate(InternalDate))
         End If
      ElseIf Tf.FieldCat = MODIFY_DATE_CAT Then
         If ShowMode = SHOW_EDIT Then
            Call glbDatabaseMngr.GetServerDateTime(InternalDate, glbErrorLog)
            Call Tf.SetValue(InternalDateToDate(InternalDate))
         End If
      ElseIf Tf.FieldCat = CREATE_BY_CAT Then
         If ShowMode = SHOW_ADD Then
            Call Tf.SetValue(glbUser.USER_ID)
         End If
      ElseIf Tf.FieldCat = MODIFY_BY_CAT Then
         If ShowMode = SHOW_EDIT Then
            Call Tf.SetValue(glbUser.USER_ID)
         End If
      End If
   Next Tf
End Sub

Public Function GenerateInsertSQL(O As Object) As String
Dim Tf As CTableField
Dim SQL As String
Dim Sep As String

   SQL = "INSERT INTO " & O.TableName & vbCrLf & " (" & vbCrLf
   For Each Tf In O.m_FieldList
      If Tf.FieldCat <> TEMP_CAT Then
         If Tf.FieldCat = MODIFY_BY_CAT Then
            Sep = "" & vbCrLf & ") " & vbCrLf & "VALUES " & vbCrLf & "(" & vbCrLf
         Else
            Sep = ", " & vbCrLf
         End If
         
         SQL = SQL & Tf.FieldName & Sep
      End If
   Next Tf
   
   For Each Tf In O.m_FieldList
      If Tf.FieldCat <> TEMP_CAT Then
         If Tf.FieldCat = MODIFY_BY_CAT Then
            Tf.SetValue (99999)
            Sep = "" & vbCrLf & ")"
         ElseIf Tf.FieldCat = CREATE_BY_CAT Then
            Tf.SetValue (99999)
            Sep = ", " & vbCrLf
         Else
            Sep = ", " & vbCrLf
         End If
'''debug.print "---" & Tf.FieldName
         SQL = SQL & Tf.TransformToSQLString & Sep
'''debug.print "---" & Tf.FieldName
      End If
   Next Tf
   
   GenerateInsertSQL = SQL
End Function

Public Function GenerateUpdateSQL(O As Object) As String
Dim Tf As CTableField
Dim SQL As String
Dim Sep As String
Dim TempKeyName As String
Dim TempKeyVal As Long

   SQL = "UPDATE " & O.TableName & " SET" & vbCrLf
   For Each Tf In O.m_FieldList
      If Tf.FieldCat <> TEMP_CAT Then
         If Tf.FieldCat = ID_CAT Then
            TempKeyName = Tf.FieldName
            TempKeyVal = Tf.GetValue
         Else
            If Tf.FieldCat = MODIFY_BY_CAT Then
               Tf.SetValue (99999)
               Sep = "" & vbCrLf
            ElseIf Tf.FieldCat = CREATE_BY_CAT Then
               Tf.SetValue (99999)
               Sep = ", " & vbCrLf
            Else
               Sep = ", " & vbCrLf
            End If
            
            SQL = SQL & Tf.FieldName & " = " & Tf.TransformToSQLString & Sep
         End If
      End If
   Next Tf
      
   SQL = SQL & "WHERE " & TempKeyName & " = " & TempKeyVal
   
   GenerateUpdateSQL = SQL
End Function
Public Function GenerateSearchLike(StartWith As String, SearchIn As String, SubLen As Long, NewStr As String) As String
    Dim WhereStr As String
    Dim StartStringNo As Long
    Dim I As Long
    StartStringNo = 1
    WhereStr = " " & StartWith & "((SUBSTR(" & SearchIn & "," & StartStringNo & "," & StartStringNo + SubLen - 1 & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
    For I = 2 To 50
        StartStringNo = StartStringNo + 1
        WhereStr = WhereStr & " OR " & "(SUBSTR(" & SearchIn & "," & StartStringNo & "," & StartStringNo + SubLen - 1 & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
    Next I
    WhereStr = WhereStr & ")"
    GenerateSearchLike = WhereStr
End Function

Public Function GetNextID(OldID As Long, Col As Collection) As Long
Dim O As Object
Dim I As Long

   I = 0
   For Each O In Col
      I = I + 1
      If (I > OldID) And (O.Flag <> "D") Then
         GetNextID = I
         Exit Function
      End If
   Next O
   GetNextID = OldID
End Function

Public Function GetCollectionFromType(B As CBatch, TempID As Long) As Collection
   If TempID = 1 Then
      Set GetCollectionFromType = B.BirthItems
   ElseIf TempID = 2 Then
      Set GetCollectionFromType = B.FoodItems
   ElseIf TempID = 3 Then
      Set GetCollectionFromType = B.TransferItems
   ElseIf TempID = 4 Then
      Set GetCollectionFromType = B.SaleItems
   ElseIf TempID = 5 Then
      Set GetCollectionFromType = B.WeightItems
   ElseIf TempID = 6 Then
      Set GetCollectionFromType = B.Feeds
   ElseIf TempID = 7 Then
      Set GetCollectionFromType = B.Balances
   ElseIf TempID = 9 Then
      Set GetCollectionFromType = B.Revenues
   ElseIf TempID = 10 Then
      Set GetCollectionFromType = B.CustRatios
   ElseIf TempID = 11 Then
      Set GetCollectionFromType = B.ChangePigTypes
   ElseIf TempID = 12 Then
      Set GetCollectionFromType = B.BuyItems
   ElseIf TempID = 13 Then
      Set GetCollectionFromType = B.ExpenseSharingItems
   ElseIf TempID = 14 Then
      Set GetCollectionFromType = B.PigAdjItems
   ElseIf TempID = 15 Then
      Set GetCollectionFromType = B.ManagementExpenses
   ElseIf TempID = 16 Then
      Set GetCollectionFromType = B.Glages
   ElseIf TempID = 17 Then
      Set GetCollectionFromType = B.GLbacks
   Else
      ''debug.print
   End If
End Function

Public Function GenerateBalanceAmt(PartItemID As Long, FromDate As Date, ToDate As Date, Optional HouseGroupID As Long = -1, Optional LocationID As Long = -1, Optional BatchID As Long = -1, Optional PigAge As Long = -1, Optional PigCode As String = "", Optional DateCount As Long) As Double
Dim TempSum As Double
Dim Ba As CBalanceAccum
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim NewDate As Date
Dim TmpDateCount As Long

   Set Ba = New CBalanceAccum
   Set TempRs = New ADODB.Recordset

   NewDate = FromDate 'mcolParam("FROM_DATE")
   TempSum = 0
   TmpDateCount = 0
   While NewDate <= ToDate 'mcolParam("TO_DATE")
      If (PigAge < 0) Or ((PigAge >= 0) And (PigAge = GetAge(PigCode, NewDate))) Then
         TmpDateCount = TmpDateCount + 1
         
         Ba.PART_ITEM_ID = PartItemID
         Ba.FROM_DATE = -1
         Ba.TO_DATE = NewDate
         Ba.HOUSE_GROUP_ID = HouseGroupID
         Ba.LOCATION_ID = LocationID
         Ba.BATCH_ID = BatchID
         Call Ba.QueryData(2, TempRs, iCount)
         If Not TempRs.EOF Then
            Call Ba.PopulateFromRS(2, TempRs)
         End If
         '''debug.print NewDate & "-" & Ba.BALANCE_AMOUNT
         TempSum = TempSum + (Ba.BALANCE_AMOUNT)
      End If
      
      NewDate = DateAdd("D", 1, NewDate)
   Wend
   
   Set TempRs = Nothing
   Set Ba = Nothing
   
   GenerateBalanceAmt = TempSum
   DateCount = TmpDateCount
End Function
Public Function GenerateBalanceAmtFollowDate(PartItemID As Long, FromDate As Date, ToDate As Date, Optional HouseGroupID As Long = -1, Optional LocationID As Long = -1, Optional BatchID As Long = -1, Optional FromDate1 As Date = -1, Optional ToDate1 As Date = -1) As Double
Dim TempSum As Double
Dim Ba As CBalanceAccum
Dim TempRs As ADODB.Recordset
Dim iCount As Long
Dim NewDate As Date
   
   Set Ba = New CBalanceAccum
   Set TempRs = New ADODB.Recordset
   
   NewDate = FromDate 'mcolParam("FROM_DATE")
   TempSum = 0
   While NewDate <= ToDate 'mcolParam("TO_DATE")
      Ba.PART_ITEM_ID = PartItemID
      Ba.FROM_DATE = -1
      Ba.TO_DATE = NewDate
      Ba.HOUSE_GROUP_ID = HouseGroupID
      Ba.LOCATION_ID = LocationID
      Ba.BATCH_ID = BatchID
      Call Ba.QueryData(2, TempRs, iCount)
      If Not TempRs.EOF Then
         Call Ba.PopulateFromRS(2, TempRs)
      End If
      TempSum = TempSum + (Ba.BALANCE_AMOUNT)
      
      NewDate = DateAdd("D", 1, NewDate)
   Wend
   
   Set TempRs = Nothing
   Set Ba = Nothing
   
   GenerateBalanceAmtFollowDate = TempSum
End Function

Public Function GetObject(ClassName As String, m_TempCol As Collection, TempKey As String, Optional SetNew As Boolean = True) As Object
On Error Resume Next
Dim EI As Object
Dim TempEi As Object

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing And SetNew Then
         Set TempEi = CreateObject(ClassName)
         If TempEi Is Nothing Then
            Set TempEi = GetNewClass(ClassName)
         End If
      End If
      Set GetObject = TempEi
   Else
      Set GetObject = EI
   End If
End Function
Public Function GetNewClass(ClassName As String) As Object
   If ClassName = "CPartType" Then
      Static m_CPartType As CPartType
      If m_CPartType Is Nothing Then
         Set m_CPartType = New CPartType
      End If
      Set GetNewClass = m_CPartType
   ElseIf ClassName = "CCostSearch1" Then
      Static m_CCostSearch1 As CCostSearch1
      If m_CCostSearch1 Is Nothing Then
         Set m_CCostSearch1 = New CCostSearch1
      End If
      Set GetNewClass = m_CCostSearch1
   ElseIf ClassName = "CIntake" Then
      Static m_CIntake As CIntake
      If m_CIntake Is Nothing Then
         Set m_CIntake = New CIntake
      End If
      Set GetNewClass = m_CIntake
   ElseIf ClassName = "CPriceAdjust" Then
      Static m_CPriceAdjust As CPriceAdjust
      If m_CPriceAdjust Is Nothing Then
         Set m_CPriceAdjust = New CPriceAdjust
      End If
      Set GetNewClass = m_CPriceAdjust
   ElseIf ClassName = "CDoItem" Then
      Static m_CDoItem As CDoItem
      If m_CDoItem Is Nothing Then
         Set m_CDoItem = New CDoItem
      End If
      Set GetNewClass = m_CDoItem
   End If
End Function
Public Function ConvertDocToConfigNo(DocKind As Long, DocType As Long, DocSubType As Long, Optional ReceiptType As Long = -1) As Long
   If DocKind = 1 Then
      If DocType = 1 Then
         ConvertDocToConfigNo = 3
      ElseIf ReceiptType = 1 Then
         ConvertDocToConfigNo = 3
      ElseIf ReceiptType = 3 Then
         ConvertDocToConfigNo = 3
      End If
   ElseIf DocKind = 2 Then
      If DocType = 2 Then
         ConvertDocToConfigNo = 81
      End If
   ElseIf DocKind = 3 Then
      If DocType = 1 Then
         ConvertDocToConfigNo = 50
      End If
   End If
End Function

Public Function GetCashTran(m_TempCol As Collection, TempKey As String) As CCashTran
On Error Resume Next
Dim EI As CCashTran
Static TempEi As CCashTran

   Set EI = m_TempCol(TempKey)
   If EI Is Nothing Then
      If TempEi Is Nothing Then
         Set TempEi = New CCashTran
      End If
      Set GetCashTran = TempEi
   Else
      Set GetCashTran = EI
   End If
End Function
Public Function CashDocType2Text(ID As CASH_DOC_TYPE) As String
   If ID = CASH_DEPOSIT Then
      CashDocType2Text = "㺹ӽҡ�Թ"
   ElseIf ID = CASH_PITTYCASH Then
      CashDocType2Text = "��ԡ/�������Թʴ����"
   ElseIf ID = CASH_TRANSFER Then
      CashDocType2Text = "��͹�Թ�����ҧ�ѭ��"
   ElseIf ID = CASH_WITHDRAW Then
      CashDocType2Text = "㺶͹�Թ (���������Թʴ����)"
   ElseIf ID = CASH_WHTHDRAW2 Then
      CashDocType2Text = "㺶͹�Թ/�͹�Թ (�����)"
   ElseIf ID = CASH_DEPOSIT2 Then
      CashDocType2Text = "㺹ӽҡ�Թ/�͹�Թ (�����)"
   ElseIf ID = POST_CHEQUE Then
      CashDocType2Text = "��׹�ѹ������Ѻ�Թ"
   End If
End Function
Public Function CheckHaveValue(OldCheckHaveValue As Boolean, Amt As Double) As Boolean
   If (Amt <> 0) Or OldCheckHaveValue Then
      CheckHaveValue = True
   End If
End Function

