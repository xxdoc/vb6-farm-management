VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSComm32.Ocx"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10575
   LinkTopic       =   "Form2"
   ScaleHeight     =   6600
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResult 
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   5880
      Width           =   5895
   End
   Begin VB.TextBox txtSent 
      Height          =   1575
      Left            =   3840
      TabIndex        =   3
      Top             =   4080
      Width           =   6015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Write"
      Height          =   1575
      Left            =   480
      TabIndex        =   2
      Top             =   4200
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Read"
      Height          =   1575
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Timer Timer1 
      Left            =   6840
      Top             =   840
   End
   Begin VB.TextBox txtRXTX 
      Height          =   1575
      Left            =   3720
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1440
      Width           =   6015
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   8400
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private intPortID As Integer ' Ex. 1, 2, 3, 4 for COM1 - COM4
Private lngStatus As Long
Private strError  As String
Private strData   As String
Private Sub Command1_Click()
   ' Set modem control lines.
    lngStatus = CommSetLine(intPortID, LINE_RTS, True)
    lngStatus = CommSetLine(intPortID, LINE_DTR, True)
    
    ' Read maximum of 64 bytes from serial port.
    lngStatus = CommRead(intPortID, strData, 64)
    If lngStatus > 0 Then
        ' Process data.
    ElseIf lngStatus < 0 Then
        ' Handle error.
    End If

    ' Reset modem control lines.
    lngStatus = CommSetLine(intPortID, LINE_RTS, False)
    lngStatus = CommSetLine(intPortID, LINE_DTR, False)
   
   DoEvents
   
   txtRXTX.Text = strData
   
   Me.Refresh
   
End Sub

Private Sub Command2_Click()
   ' Set modem control lines.
    lngStatus = CommSetLine(intPortID, LINE_RTS, True)
    lngStatus = CommSetLine(intPortID, LINE_DTR, True)
   
    ' Write data to serial port.
    strData = txtSent.Text
    If Len(strData) <= 0 Then
      'strData = "R" + "L" + "0" + Chr$(13)
      'strData = "O0W0" + Chr$(13)
      strData = "R" + Chr$(13)
   Else
      'strData = txtSent.Text + Chr$(13)
    End If
    lngSize = Len(strData)
    lngStatus = CommWrite(intPortID, strData)
    If lngStatus <> lngSize Then
      ' Handle error.
      Debug.Print
    End If
   
   lngStatus = CommRead(intPortID, strData, 64)
    If lngStatus > 0 Then
        ' Process data.
    ElseIf lngStatus < 0 Then
        ' Handle error.
    End If
    
    If Len(strData) > 0 Then
      txtResult.Text = strData
   Else
      txtResult.Text = "NO Data"
   End If
   
   'Call Command1_Click
   
   ' Reset modem control lines.
    lngStatus = CommSetLine(intPortID, LINE_RTS, False)
    lngStatus = CommSetLine(intPortID, LINE_DTR, False)

End Sub

Private Sub Form_Load()
   intPortID = 1
   
    ' Initialize Communications
    Call CommClose(intPortID)
    
    lngStatus = CommOpen(intPortID, "COM" & CStr(intPortID), "baud=9600 parity=N data=8 stop=1")
    
    If lngStatus <> 0 Then
   ' Handle error.
        lngStatus = CommGetError(strError)
         MsgBox "COM Error: " & strError
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' Close communications.
   Call CommClose(intPortID)
End Sub
