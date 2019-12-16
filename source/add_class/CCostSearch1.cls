VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCostSearch1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public PIG_ID As Long

Public FROM_BIRTH As Date
Public TO_BIRTH As Date

Public BFOOD_AMOUNT As Double                                     'ยกมาอาหาร
Public FOOD_AMOUNT As Double

Public BMEDICINE_AMOUNT As Double                                     'ยกมา ยา+วัคซีน
Public MEDICINE_AMOUNT As Double

Public BEXPENSE_AMOUNT As Double                              'ยกมาค่าใช้จ่ายผลิต
Public EXPENSE_AMOUNT As Double

Public BBIRTH_AMOUNT As Double                                    'ยกมาเกิด
Public BIRTH_AMOUNT As Double

Public OTHER_AMOUNT As Double

Public CURRENT_AMOUNT As Double
Public COST_AMOUNT As Double
Public COST_PER_AMOUNT As Double
Public Function GetKey1() As String
   GetKey1 = Trim(Str(PIG_ID))
End Function
