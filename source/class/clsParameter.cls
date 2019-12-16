VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const WIN_MIS_REVISION = "3"
Private Const APP_REVISION = "12"

Private Const REGKEY = "Software\WIN\MittraphapFarm"
Private Reg As clsRegistry

Public DBFile As String
Public DBFileAP As String
Public DBFileAPX As String

Public UserName As String
Public Password As String

Public LicenseIP As String
Public LicensePort As Long
Public Version As String
Public Language As Long

Public AppName As String
Public AppVersion As String
Public ResourceType As RESOURCE_TYPE

Public Drive As String
Public Directory As String
Public ImportYear As Long
Public ImportSemester As Long
Public FilterExt As String
Public NewExt As String
Public UsedCount As Long
Public Temp As Long
Public Area As Long

Public Drive2 As String
Public Directory2 As String
Public ImportYear2 As Long
Public ImportSemester2 As Long
Public FilterExt2 As String
Public NewExt2 As String
Public UsedCount2 As Long
Public Temp2 As Long
Public Area2 As Long

Public Rs232PortNo As String
Public TimeDelay As String

Public MainPicture As String
Public MainButton As String
Public NormalButton1 As String
Public NormalForm1 As String
Public ButtonCursor As String
Public ApplicationPath As String
Public Programowner As String
Public ReportFile As String
Public ReceiptVoccherPic1 As String
Public DOFormPic1 As String
Public BLExportFileName As String
Public SumFarmExportFileName As String

Private Sub Class_Initialize()
Dim ErrCode As Integer

   Set Reg = New clsRegistry
   Reg.RgtRootKey = &H80000002 'HKEY_LOCAL_MACHINE
   Reg.RgtSubKey = REGKEY

   ApplicationPath = App.Path
   ReportFile = ApplicationPath & "\" & "temp_report.html"
   
   AppName = "GeneticSoft WIN Pricing"
   AppVersion = App.Major & "." & App.Minor & "." & WIN_MIS_REVISION & "." & App.Revision & "   (อินเตอร์เบส 6.0)"
   'DBFile = App.Path & "\IBData\MITTRAPHAP_FARM.GDB"
   DBFile = "C:\GDB\mittraphap_farm.gdb"
   
   UserName = "SYSDBA"
   Password = "masterkey"
   
   LicenseIP = "127.0.0.1"
   LicensePort = 8888
   BLExportFileName = ""
   SumFarmExportFileName = ""
   
   Call Reg.RgtGetStringValue("BLExportFileName", BLExportFileName, ErrCode)
   Call Reg.RgtGetStringValue("SumFarmExportFileName", SumFarmExportFileName, ErrCode)
   
   Call Reg.RgtGetStringValue("DBFile", DBFile, ErrCode)
   If Not (Command = "" Or Command = "1") Then
      Call Reg.RgtGetStringValue("DBFileAP" & Command, DBFileAPX, ErrCode)
   End If
   Call Reg.RgtGetStringValue("UserName", UserName, ErrCode)
   Call Reg.RgtGetStringValue("Password", Password, ErrCode)
   Password = DecryptText(Password)
   Call Reg.RgtGetStringValue("LicenseIP", LicenseIP, ErrCode)
   Call Reg.RgtGetLongValue("LicensePort", LicensePort, ErrCode)
   Call Reg.RgtGetLongValue("Language", Language, ErrCode)
   Language = 1
   
   Call Reg.RgtGetStringValue("Drive", Drive, ErrCode)
   Call Reg.RgtGetStringValue("Directory", Directory, ErrCode)
   Call Reg.RgtGetLongValue("ImportYear", ImportYear, ErrCode)
   Call Reg.RgtGetLongValue("ImportSemester", ImportSemester, ErrCode)
   Call Reg.RgtGetStringValue("FilterExt", FilterExt, ErrCode)
   Call Reg.RgtGetStringValue("NewExt", NewExt, ErrCode)
   Call Reg.RgtGetLongValue("Day", UsedCount, ErrCode)
   Call Reg.RgtGetLongValue("Area", Area, ErrCode)
   
   Call Reg.RgtGetStringValue("Drive2", Drive2, ErrCode)
   Call Reg.RgtGetStringValue("Directory2", Directory2, ErrCode)
   Call Reg.RgtGetLongValue("ImportYear2", ImportYear2, ErrCode)
   Call Reg.RgtGetLongValue("ImportSemester2", ImportSemester2, ErrCode)
   Call Reg.RgtGetStringValue("FilterExt2", FilterExt2, ErrCode)
   Call Reg.RgtGetStringValue("NewExt2", NewExt2, ErrCode)
   Call Reg.RgtGetLongValue("Day2", UsedCount2, ErrCode)
   Call Reg.RgtGetLongValue("Area2", Area2, ErrCode)
   
   Call Reg.RgtGetStringValue("Rs232PortNo", Rs232PortNo, ErrCode)
   Call Reg.RgtGetStringValue("TimeDelay", TimeDelay, ErrCode)
   If Trim(Rs232PortNo) = "" Then
      Rs232PortNo = 1
      TimeDelay = 300
   End If
   
   Version = App.Major & "." & App.Minor & "." & APP_REVISION & "." & App.Revision
   
   ResourceType = HOTEL_RESOURCE
   Temp = 0
   
   MainPicture = App.Path & "\Pictures\main1.jpg"
   MainButton = App.Path & "\Pictures\retangle_button1.jpg"
   NormalButton1 = App.Path & "\Pictures\normal_button2.jpg"
   NormalForm1 = App.Path & "\Pictures\dialog_form1.jpg"
   ButtonCursor = App.Path & "\Pictures\H_POINT.CUR"
   
   ReceiptVoccherPic1 = App.Path & "\Pictures\receipt_vocher.jpg"
   DOFormPic1 = App.Path & "\Pictures\do_form.jpg"
End Sub

Private Sub Class_Terminate()
Dim ErrCode As Integer
   
   Call Reg.RgtSetStringValue("DBFile", DBFile, ErrCode)
   If Not (Command = "" Or Command = "1") Then
      Call Reg.RgtSetStringValue("DBFileAP" & Command, DBFileAPX, ErrCode)
   End If
   Call Reg.RgtSetStringValue("UserName", UserName, ErrCode)
   Password = EncryptText(Password)
   Call Reg.RgtSetStringValue("Password", Password, ErrCode)
   Call Reg.RgtSetStringValue("LicenseIP", LicenseIP, ErrCode)
   Call Reg.RgtSetLongValue("LicensePort", LicensePort, ErrCode)
   
   Call Reg.RgtSetStringValue("Drive", Drive, ErrCode)
   Call Reg.RgtSetStringValue("Directory", Directory, ErrCode)
   Call Reg.RgtSetLongValue("ImportYear", ImportYear, ErrCode)
   Call Reg.RgtSetLongValue("ImportSemester", ImportSemester, ErrCode)
   Call Reg.RgtSetStringValue("FilterExt", FilterExt, ErrCode)
   Call Reg.RgtSetStringValue("NewExt", NewExt, ErrCode)
   Call Reg.RgtSetLongValue("Day", UsedCount, ErrCode)
   Call Reg.RgtSetLongValue("Area", Area, ErrCode)
   Call Reg.RgtSetLongValue("Language", Language, ErrCode)
   
   Call Reg.RgtSetStringValue("Drive2", Drive2, ErrCode)
   Call Reg.RgtSetStringValue("Directory2", Directory2, ErrCode)
   Call Reg.RgtSetLongValue("ImportYear2", ImportYear2, ErrCode)
   Call Reg.RgtSetLongValue("ImportSemester2", ImportSemester2, ErrCode)
   Call Reg.RgtSetStringValue("FilterExt2", FilterExt2, ErrCode)
   Call Reg.RgtSetStringValue("NewExt2", NewExt2, ErrCode)
   Call Reg.RgtSetLongValue("Day2", UsedCount2, ErrCode)
   Call Reg.RgtSetLongValue("Area2", Area2, ErrCode)
   Call Reg.RgtSetStringValue("BLExportFileName", BLExportFileName, ErrCode)
   Call Reg.RgtSetStringValue("SumFarmExportFileName", SumFarmExportFileName, ErrCode)
   
   If Val(Rs232PortNo) > 0 Then
      Call Reg.RgtSetStringValue("Rs232PortNo", Rs232PortNo, ErrCode)
      Call Reg.RgtSetStringValue("TimeDelay", TimeDelay, ErrCode)
   End If
   
   Set Reg = Nothing
End Sub
