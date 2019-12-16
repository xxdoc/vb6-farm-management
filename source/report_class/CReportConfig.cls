VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public REPORT_CONFIG_ID As Long
Public REPORT_CLASS_NAME As String
Public PAPER_SIZE As Double
Public ORIENTATION As Double
Public MARGIN_BOTTOM As Double
Public MARGIN_FOOTER As Double
Public MARGIN_HEADER As Double
Public MARGIN_LEFT As Double
Public MARGIN_RIGHT As Double
Public MARGIN_TOP As Double
Public FONT_NAME As String
Public FONT_SIZE As Double
Public MASTER_FLAG As String
Public PAPER_WIDTH As Double
Public PAPER_HEIGHT As Double

Public OrderBy As Long
Public OrderType As Long


