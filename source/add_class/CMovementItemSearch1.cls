VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMovementItemSearch1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public HOUSE_ID As Long
Public PIG_ID As Long
Public PART_ITEM_ID As Long
Public EXPENSE_TYPE As Long
Public CAPITAL_AMOUNT As Double

Public Sub PrintDebug(DocNo As String, TempStr As String, TxAmount As Double, TxType As String)
   If (HOUSE_ID = 255) And (PIG_ID = 9323) And (PART_ITEM_ID = 7263) Then
      ''debug.print TxType & "=" & DocNo & "=" & GetKey1 & "=" & CAPITAL_AMOUNT & "=" & TempStr & "=" & TxAmount
   End If
End Sub

Public Function GetKey1() As String
   GetKey1 = HOUSE_ID & "-" & PIG_ID & "-" & PART_ITEM_ID
End Function

Public Function GetKey2() As String
   GetKey2 = HOUSE_ID & "-" & PIG_ID
End Function

Public Function GetKey3() As String
   GetKey3 = HOUSE_ID & "-" & PIG_ID & "-" & EXPENSE_TYPE
End Function

