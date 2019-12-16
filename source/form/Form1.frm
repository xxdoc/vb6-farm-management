VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5490
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   705
      Left            =   870
      TabIndex        =   0
      Top             =   420
      Width           =   2505
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Rs As ADODB.Recordset

Private Sub Command1_Click()
Dim II As CImportItem
Dim iCount As Long

   Call glbDaily.StartTransaction
   
   Set II = New CImportItem
   II.IMPORT_ITEM_ID = -1
   II.PIG_FLAG = ""
   Call II.QueryDataPatch(1, m_Rs, iCount)

   While Not m_Rs.EOF
      Call II.PopulateFromRS(1, m_Rs)
      
      II.AddEditMode = SHOW_ADD
      Call II.AddEditDataExt
      m_Rs.MoveNext
   Wend
   
   Call glbDaily.CommitTransaction
   Set II = Nothing
End Sub

Private Sub Form_Load()
   Set m_Rs = New ADODB.Recordset
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set m_Rs = Nothing
End Sub
