VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public GROUP_ID As Long
Public GROUP_NAME As String
Public GROUP_RIGHT As String
Public GROUP_DESC As String

Public OrderBy As Long
Public OrderType As Long
Public QueryFlag As Long
Public ShowMode As SHOW_MODE_TYPE

Public RightDatas As Collection

Private Sub Class_Initialize()
   GROUP_ID = -1
   GROUP_NAME = ""
   GROUP_RIGHT = ""
   OrderBy = -1
   OrderType = -1
   QueryFlag = -1
   
   Set RightDatas = New Collection
End Sub

Private Sub Class_Terminate()
   Set RightDatas = Nothing
End Sub
