VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_ModuleName As String
Private m_ADOConn As ADODB.Connection
Private m_Rs As ADODB.Recordset

Public Function QueryLoginTracking(Sp As CLoginTracking, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long

   RName = "QueryLoginTracking"
   QueryLoginTracking = False

   IsOK = True
    Call Sp.QueryData(Rs, ItemCount)

   QueryLoginTracking = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryLoginTracking = False
End Function

Public Function QueryUserAccount(Ua As CUserAccount, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long

   RName = "QueryUserAccount"
   QueryUserAccount = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)

   Dim bIsOk As Boolean
   Dim iCount As Long

'   If Ua.QueryFlag = 1 Then
'      Dim cLogin As CLoginTracking
'      Set cLogin = New CLoginTracking
'      cLogin.USER_ID = Ua.USER_ID
'      cLogin.LOGIN_FROM_DATE = -1
'      cLogin.LOGIN_TO_DATE = -1
'      cLogin.OrderBy = 1
'      Call cLogin.QueryData(m_Rs, iCount)
'      Set cLogin = Nothing
'
'      Set Ua.Logins = Nothing
'      Set Ua.Logins = New Collection
'      While Not m_Rs.EOF
'         Set cLogin = New CLoginTracking
'         Call cLogin.PopulateFromRS(m_Rs)
'         Call Ua.Logins.Add(cLogin)
'         Set cLogin = Nothing
'
'         m_Rs.MoveNext
'      Wend
'   End If

   QueryUserAccount = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryUserAccount = False
End Function

Public Function AddEditUserAccount(Ua As CUserAccount, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditUserAccount"
   AddEditUserAccount = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditUserAccount = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditUserAccount = False
End Function

Public Function DeleteUserAccount(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CUserAccount

   RName = "DeleteUserAccount"
   DeleteUserAccount = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Set cData = New CUserAccount
   cData.USER_ID = UID
   Call cData.DeleteData(1)
   Set cData = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteUserAccount = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteUserAccount = True
End Function

Public Function DBLogin(UserName As String, Password As String, CanLogin As Boolean, User As CUser, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim ItemCount As Long
Dim LoginID As Long

   RName = "Login"
   DBLogin = False

   SQL1 = "SELECT COUNT(*) FROM USER_ACCOUNT "
   If Not glbDatabaseMngr.GetRs(SQL1, "", False, ItemCount, m_Rs, ErrorObj) Then
      Exit Function
   End If
   If NVLI(m_Rs(0), 0) <= 0 Then 'ถ้าไม่มีผู้ใช้อยุในดาตาเบสให้ล็อคอินอะไรก็ได้
      If UserName <> "" Then
         ErrorObj.LocalErrorMsg = "ไม่พบบัญชีรายชื่อผู้ใช้นี้ในฐานข้อมูล"

         DBLogin = True
         CanLogin = False
         Exit Function
      End If

      If Not glbDatabaseMngr.Login(UserName, 0, CanLogin, ErrorObj) Then
         Exit Function
      End If
      If Not CanLogin Then
         DBLogin = True
         Exit Function
      End If

      User.USER_ID = 0
      User.GROUP_ID = 0
      User.GROUP_NAME = ""
      User.USER_NAME = UserName

      DBLogin = True
      CanLogin = True
      Exit Function
   End If

   SQL1 = "SELECT UD.*, GD.* FROM USER_ACCOUNT UD, USER_GROUP GD "
   SQL1 = SQL1 & "WHERE (GD.GROUP_ID = UD.GROUP_ID) "
   SQL1 = SQL1 & "AND (UD.USER_NAME = '" & ChangeQuote(UserName) & "' ) "
   SQL1 = SQL1 & "AND (UD.USER_PASSWORD = '" & ChangeQuote(EncryptText(Password)) & "' ) "

   If Not glbDatabaseMngr.GetRs(SQL1, "", False, ItemCount, m_Rs, ErrorObj) Then
      Exit Function
   End If
   If m_Rs.EOF Then
      ErrorObj.LocalErrorMsg = "ไม่พบบัญชีรายชื่อผู้ใช้นี้ในฐานข้อมูล"

      DBLogin = True
      CanLogin = False
      Exit Function
   End If

   If NVLS(m_Rs("USER_STATUS"), "Y") <> "Y" Then
      ErrorObj.LocalErrorMsg = "รายชื่อของท่านขณะนี้ไม่สามารถเข้าใช้งานระบบได้ โปรดติดต่อผู้บริหารระบบ "

      DBLogin = True
      CanLogin = False
      Exit Function
   End If

   User.USER_ID = NVLI(m_Rs("USER_ID"), 0)
   User.USER_NAME = NVLS(m_Rs("USER_NAME"), "")
   User.GROUP_ID = NVLS(m_Rs("GROUP_ID"), "")
   User.GROUP_NAME = NVLS(m_Rs("GROUP_NAME"), "")
   User.REAL_USER_ID = NVLI(m_Rs("USER_ID"), 0)
   User.EXCEPTION_FLAG = NVLS(m_Rs("EXCEPTION_FLAG"), "N")
   If Not glbDatabaseMngr.Login(UserName, NVLI(m_Rs("USER_ID"), 0), CanLogin, ErrorObj) Then
      Exit Function
   End If
   If Not CanLogin Then
      DBLogin = True
      Exit Function
   End If
   
   m_ADOConn.BeginTrans

   glbLoginTracking.LOGIN_DATE = Now
   glbLoginTracking.LOGOUT_DATE = -1
   glbLoginTracking.IP_ADDRESS = glbDatabaseMngr.m_Winsock.LocalIP
   glbLoginTracking.USER_NAME = User.USER_NAME
   glbLoginTracking.USER_ID = User.USER_ID
   glbLoginTracking.AddEditMode = SHOW_ADD
   Call glbLoginTracking.AddEditData

   m_ADOConn.CommitTrans
   
   CanLogin = True
   DBLogin = True
   Exit Function

ErrorHandler:
   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   DBLogin = False
End Function

Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   m_ModuleName = "clsAdmin"
   Set m_Rs = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
End Sub

Public Function ChangePassword(UserID As Long, OldPassword As String, NewPassword As String, IsCanChange As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim UserName As String
Dim ItemCount As Long

   RName = "ChangePassword"
   ChangePassword = False
   HasBegin = False

   SQL = "SELECT * FROM USER_ACCOUNT WHERE (USER_ID = " & UserID & ") "
   If Not glbDatabaseMngr.GetRs(SQL, "", False, ItemCount, m_Rs, ErrorObj) Then
      ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
      Exit Function
   End If
   If m_Rs.EOF Then
      ErrorObj.LocalErrorMsg = "ไม่มีผู้ใช้นี้อยู่ในฐานข้อมูล "
      IsCanChange = False
      ChangePassword = True
      Exit Function
   End If
   UserName = NVLS(m_Rs("USER_NAME"), "")

   SQL = "SELECT * FROM USER_ACCOUNT "
   SQL = SQL & "WHERE (USER_PASSWORD = '" & ChangeQuote(EncryptText(OldPassword)) & "') AND "
   SQL = SQL & "(USER_ID = " & UserID & ") "
   If Not glbDatabaseMngr.GetRs(SQL, "", False, ItemCount, m_Rs, ErrorObj) Then
      ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
      Exit Function
   End If
   If m_Rs.EOF Then
      ErrorObj.LocalErrorMsg = "ท่านใส่รหัสผ่านเก่าไม่ถูกต้อง "
      IsCanChange = False
      ChangePassword = True
      Exit Function
   End If

   m_ADOConn.BeginTrans
   HasBegin = True

   SQL = "UPDATE USER_ACCOUNT SET "
   SQL = SQL & "USER_PASSWORD = '" & ChangeQuote(EncryptText(NewPassword)) & "' "
   SQL = SQL & "WHERE (USER_ID = " & UserID & ") "
   Call m_ADOConn.Execute(SQL)

   Call m_ADOConn.CommitTrans
   HasBegin = False

   IsCanChange = True
   ChangePassword = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION & SQL
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = ""
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   ChangePassword = False
End Function

'==
Public Function QueryUserGroup(Ua As CUserGroup, Rs As ADODB.Recordset, ItemCount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long
Dim bIsOk As Boolean

   RName = "QueryUserGroup"
   QueryUserGroup = False

   IsOK = True

   Call Ua.QueryData(1, Rs, ItemCount)
   
   If Ua.QueryFlag = 1 Then
      Dim Gr As CGroupRight
      Set Gr = New CGroupRight
      Gr.GROUP_ID = Ua.GROUP_ID
      Call Gr.QueryData(2, m_Rs, iCount)
      Set Gr = Nothing

      Set Ua.RightItems = Nothing
      Set Ua.RightItems = New Collection

      While Not m_Rs.EOF
         Set Gr = New CGroupRight
         Call Gr.PopulateFromRS(2, m_Rs)
         Call Ua.RightItems.Add(Gr)
         Set Gr = Nothing

         m_Rs.MoveNext
      Wend
   End If
   
   QueryUserGroup = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryUserGroup = False
End Function

Public Function AddEditUserGroup(Ua As CUserGroup, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditUserGroup"
   AddEditUserGroup = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
      HasBegin = True
   End If
   
   Call Ua.AddEditData

   Dim Gr As CGroupRight
   For Each Gr In Ua.RightItems
      Gr.GROUP_ID = Ua.GROUP_ID
      If Gr.Flag = "A" Then
         Gr.AddEditMode = SHOW_ADD
         Call Gr.AddEditData
      ElseIf Gr.Flag = "E" Then
         Gr.AddEditMode = SHOW_EDIT
         Call Gr.AddEditData
      ElseIf Gr.Flag = "D" Then
         Call Gr.DeleteData
      End If
   Next Gr
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditUserGroup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditUserGroup = False
End Function

Public Function DeleteUserGroup(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CUserGroup

   RName = "DeleteUserGroup"
   DeleteUserGroup = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Set cData = New CUserGroup
   cData.GROUP_ID = UID
   Call cData.DeleteData(1)
   Set cData = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteUserGroup = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteUserGroup = True
End Function

