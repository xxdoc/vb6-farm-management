VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPopulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Flag As String

Public PIG_ID As Long
Public PIG_NO As String
Public PIG_TYPE As String
Public PIG_NAME As String

Public CURRENT_AGE As Double
Public CURRENT_AMOUNT As Double
Public TOTAL_WEIGHT As Double
Public AVG_WEIGHT As Double

Public USED_AMOUNT As Double
Public TOTAL_PRICE As Double
Public FEED_COST As Double
Public MEDICINE_COST As Double
Public EXPENSE_COST As Double
Public BIRTH_COST As Double

Public PIG_STATUS As Long
Public LEFT_AMOUNT  As Double

Private Sub Class_Initialize()
   Flag = "I"
End Sub
