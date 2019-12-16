VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   6225
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   675
      Left            =   0
      TabIndex        =   9
      Top             =   -120
      Width           =   6225
      Begin VB.Label lblHeader 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   150
         TabIndex        =   10
         Top             =   210
         Width           =   5955
      End
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1410
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2190
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2190
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000000&
      Height          =   1755
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   6225
      Begin VB.OptionButton radEng 
         Height          =   435
         Left            =   3630
         TabIndex        =   4
         Top             =   1200
         Width           =   2025
      End
      Begin VB.OptionButton radThai 
         Height          =   435
         Left            =   1440
         TabIndex        =   3
         Top             =   1200
         Width           =   2025
      End
      Begin prjWINPricing300.uctlTextBox txtPassword 
         Height          =   375
         Left            =   1380
         TabIndex        =   2
         Top             =   690
         Width           =   4455
         _ExtentX        =   0
         _ExtentY        =   0
      End
      Begin prjWINPricing300.uctlTextBox txtUserName 
         Height          =   375
         Left            =   1380
         TabIndex        =   1
         Top             =   300
         Width           =   4455
         _ExtentX        =   0
         _ExtentY        =   0
      End
      Begin VB.Label lblPassword 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label lblUserName 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   150
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public OKClick As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      MsgBox Me.Name
   End If
End Sub

Private Sub cmdCancel_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub cmdOK_Click()
Dim IsCanLogin As Boolean

   Call EnableForm(Me, False)
   If Not glbAdmin.DBLogin(txtUserName.Text, txtPassword.Text, IsCanLogin, glbUser, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)

      Call EnableForm(Me, True)
      txtUserName.SetFocus
      Exit Sub
   End If
   
   If Not IsCanLogin Then
      glbErrorLog.ShowUserError
      
      Call EnableForm(Me, True)
      txtUserName.SetFocus
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
   
   OKClick = True
   Unload Me
End Sub

Private Sub Form_Load()
   Me.BackColor = GLB_FORM_COLOR
   Frame1.BackColor = GLB_FORM_COLOR
   Frame2.BackColor = GLB_HEAD_COLOR
   lblHeader.BackColor = GLB_HEAD_COLOR
   
   OKClick = False
   
   Call InitOption(radThai, "ไทย")
   Call InitOption(radEng, "English")
   radEng.Enabled = False
   radThai.Value = True
'   If glbParameterObj.Language = 1 Then
'      radThai.Value = True
'   Else
'      radEng.Value = True
'   End If
   
   Call InitNormalLabel(lblUserName, GetTextMessage("TEXT-KEY77"))
   Call InitNormalLabel(lblPassword, GetTextMessage("TEXT-KEY173"))
   Call txtUserName.SetTextLenType(TEXT_STRING, glbSetting.USERNAME_TYPE)
   txtUserName.SetTextType (1)
   Call txtPassword.SetTextLenType(TEXT_STRING, glbSetting.PASSWORD_TYPE)
   txtPassword.PasswordChar = "*"
   
   Call InitDialogButton(cmdOK, GetTextMessage("TEXT-KEY92"))
   Call InitDialogButton(cmdCancel, GetTextMessage("TEXT-KEY165"))
   
   Call InitDialogHeader(lblHeader, GetTextMessage("TEXT-KEY443"))
End Sub

Private Sub radEng_Click()
   glbParameterObj.Language = 2
End Sub

Private Sub radThai_Click()
   glbParameterObj.Language = 1
End Sub

Private Sub txtPassword_GotFocus()
'   Call SetSelect(txtPassword)
End Sub

Private Sub txtUserName_GotFocus()
'   Call SetSelect(txtUserName)
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
