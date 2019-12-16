VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditCustomerPackage 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditCustomerPackage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2595
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   4577
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox CboPackageType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   5355
      End
      Begin prjFarmManagement.uctlTextLookup uctlPackage 
         Height          =   435
         Left            =   2520
         TabIndex        =   1
         Top             =   780
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin VB.Label lblPackageType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   420
         Width           =   2205
      End
      Begin Threed.SSCommand cmdNext 
         Height          =   525
         Left            =   2505
         TabIndex        =   2
         Top             =   1530
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCustomerPackage.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblPackage 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   2205
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   4155
         TabIndex        =   3
         Top             =   1530
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCustomerPackage.frx":0BE4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5805
         TabIndex        =   4
         Top             =   1530
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditCustomerPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE
Public ParentShowMode As SHOW_MODE_TYPE
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection

Private m_Package As Collection


Public ParentForm As Form
Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub InitFormLayout()
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Call InitNormalLabel(lblPackage, MapText("แบบการตั้งราคา"))
   Call InitNormalLabel(lblPackageType, MapText("ประเภทการตั้งราคา"))
   
   Call InitCombo(CboPackageType)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdNext.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdNext, MapText("ถัดไป"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim ItemCount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim Di As CCustomerPackage
         
         Set Di = TempCollection.Item(ID)
         
         CboPackageType.ListIndex = IDToListIndex(CboPackageType, Di.PKG_TYPE)
         uctlPackage.MyCombo.ListIndex = IDToListIndex(uctlPackage.MyCombo, Di.PKG_ID)
         
         
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Public Function GetNextID(OldID As Long, Col As Collection) As Long
Dim O As Object
Dim I As Long

   I = 0
   For Each O In Col
      I = I + 1
      If (I > OldID) And (O.Flag <> "D") Then
         GetNextID = I
         Exit Function
      End If
   Next O
   GetNextID = OldID
End Function


Private Sub cmdNext_Click()
Dim NewID As Long
   If Not SaveData Then
      Exit Sub
   End If
   
   Call ParentForm.ShowCustomerPackageGrid
   
   If ShowMode = SHOW_EDIT Then
      NewID = GetNextID(ID, TempCollection)
      If ID = NewID Then
         glbErrorLog.LocalErrorMsg = "ถึงเรคคอร์ดสุดท้ายแล้ว"
         glbErrorLog.ShowUserError
         
         Call ParentForm.ShowCustomerPackageGrid
         Exit Sub
      End If

      ID = NewID
   ElseIf ShowMode = SHOW_ADD Then
        uctlPackage.MyCombo.ListIndex = -1
        CboPackageType.ListIndex = -1
   End If
   Call QueryData(True)
   Call ParentForm.ShowCustomerPackageGrid
   
End Sub

Private Sub cmdOK_Click()
   If Not cmdOK.Enabled Then
      Exit Sub
   End If
   
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long
Dim I As Long
   If Not VerifyCombo(lblPackage, uctlPackage.MyCombo, False) Then
      Exit Function
   End If
    
    If Not VerifyCombo(lblPackageType, CboPackageType, False) Then
      Exit Function
   End If
    
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   
   Dim Di As CCustomerPackage
   Dim CheckDetail As CCustomerPackage
   I = 0
   For Each CheckDetail In TempCollection
    I = I + 1
    If CheckDetail.PKG_TYPE = CboPackageType.ItemData(Minus2Zero(CboPackageType.ListIndex)) And ID <> I Then
        glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & CboPackageType.Text & " " & MapText("อยู่ในระบบแล้ว")
        glbErrorLog.ShowUserError
        Exit Function
    End If
   Next
   If ShowMode = SHOW_ADD Then
      Set Di = New CCustomerPackage
      
      Di.Flag = "A"
      Call TempCollection.Add(Di)
   Else
      Set Di = TempCollection.Item(ID)
      If Di.Flag <> "A" Then
         Di.Flag = "E"
      End If
   End If


   Di.PKG_TYPE = CboPackageType.ItemData(Minus2Zero(CboPackageType.ListIndex))
   Di.PKG_ID = uctlPackage.MyCombo.ItemData(Minus2Zero(uctlPackage.MyCombo.ListIndex))
    Di.PKG_NAME = uctlPackage.MyCombo.Text
    Di.PACKAGE_TYPE_NAME = CboPackageType.Text
       
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call InitPackageType(CboPackageType)
      
      
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         Call QueryData(True)
      End If
      
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
      KeyCode = 0
'   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   End If
End Sub

Private Sub Form_Load()

   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   
   Set m_Package = New Collection
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_Package = Nothing
   
End Sub
Private Sub uctlPackage_Change()
    m_HasModify = True
End Sub

Private Sub CboPackageType_Click()
Dim TempID As Long
    TempID = CboPackageType.ItemData(Minus2Zero(CboPackageType.ListIndex))
    If TempID > 0 Then
      Call LoadPackage(uctlPackage.MyCombo, m_Package, TempID, "N")
      Set uctlPackage.MyCollection = m_Package
    End If
    m_HasModify = True
End Sub
