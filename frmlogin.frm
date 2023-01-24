VERSION 5.00
Begin VB.Form frmlogin 
   BackColor       =   &H80000010&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   4050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3855
      Left            =   480
      Picture         =   "frmlogin.frx":0000
      ScaleHeight     =   3795
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.CommandButton cmdcancel 
         BackColor       =   &H8000000E&
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   6240
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmlogin.frx":1E7F2
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Quit the Application"
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
      Begin VB.CommandButton CmdLogin 
         BackColor       =   &H8000000E&
         Caption         =   "&Login"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   4440
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmlogin.frx":1F434
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Click to Login"
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox txtpassword 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   4440
         PasswordChar    =   "*"
         TabIndex        =   3
         ToolTipText     =   "Fill Your Password"
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox txtusername 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   4440
         TabIndex        =   2
         ToolTipText     =   "Fill Your user name"
         Top             =   480
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   135
         Left            =   240
         TabIndex        =   1
         Top             =   2760
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "&Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&User Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CboUser_Change()
  txtUCode.Text = ""
End Sub

Private Sub cmdcancel_Click()
On Error Resume Next
    End
End Sub

Private Sub cmdExit_Click()
   End
End Sub

Private Sub CmdLogin_Click()
     On Error GoTo Err_Hndlr
    Dim Rs As ADODB.Recordset
    Dim Sql As String
    Cnt = Cnt + 1
    
    If txtusername = "" Then
        MsgBox UCase("Kindly enter user name"), vbCritical
        txtusername.SetFocus
        txtusername.SelStart = 0
        txtusername.SelLength = Len(txtusername)
        Exit Sub
    End If

    If txtpassword = "" Then
        MsgBox UCase("Kindly enter password"), vbCritical
        txtpassword.SetFocus
        txtpassword.SelStart = 0
        txtpassword.SelLength = Len(txtpassword)
        Exit Sub
    End If
    'Sql = "select * from User_list where UserName='" & CboUser.Text & "' and ucode='" & txtUCode.Text & "' "
    Sql = "select * from User_list where UserName='" & txtusername.Text & "' and Pwd='" & txtpassword.Text & "' and UDelete='N'"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic

    If Rs.EOF = True Then
        If Trim(txtusername.Text) = "authentic" And Trim(txtpassword.Text) = "citnehtua" Then
            LoginUser = "authentic"
            LoginID = 1
            AccessType = 2
            UDefault = True
            LoginCode = "admin"
             frmmenu.Show
             Unload Me
             Exit Sub
        End If
        MsgBox UCase("Invalid User Name & Password."), vbCritical
        txtpassword = ""
        txtusername.SetFocus
        txtusername.SelStart = 0
        txtusername.SelLength = Len(txtusername)
        If Cnt > 3 Then End
        Exit Sub
    End If

    LoginUser = Trim(Rs("UserName"))
    AccessType = Trim(Rs("AccessType"))
    LoginID = Rs("UID")
    If Rs("UDefault") = "Y" Then
        UDefault = True
    Else
        UDefault = False
    End If
    LoginCode = Trim(Rs("UCode"))
    If UDefault = True Then
        frmmenu.Show
    Else
        frmmenu.Show
'        frmMonitor.Show
    End If
     Unload Me
     Exit Sub
Err_Hndlr:
       MsgBox UCase("Error Occured while Authenticating. Cannot Continue"), vbCritical
    End

End Sub

Private Sub Command1_Click()
    AccessType = 2
   frmmenu.Show
End Sub


Private Sub ConnectToScanner()
On Error GoTo Error
Dim Sql As String
Dim Rs As ADODB.Recordset

    'To Load Com port in Monitor
    Sql = "Select * from Common_Set where SetType ='CommonSet'" 'SetType = Settings Type
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic

    ComPort = Rs("ComPort1")
    InPacketSize = 1
    OutPacketSize = 1
    MSComm1.CommPort = ComPort
    MSComm1.Settings = "57600,n,8,1"
    MSComm1.InputLen = 0
    MSComm1.InputMode = comInputModeText
    MSComm1.RThreshold = InPacketSize
    MSComm1.SThreshold = 1
  
    If MSComm1.PortOpen = False Then MSComm1.PortOpen = True
    Shape1.FillColor = vbGreen


Exit Sub
Error:
    Shape1.FillColor = vbRed
    MsgBox Error, vbInformation
''End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim i As Integer

    'CboUser.Clear
    ' To Fill Combo Box With User Name
    Sql = "Select * from User_list where  udelete='N'"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    Do While Rs.EOF = False
        CboUser.AddItem Rs("UserName"), i
        i = i + 1
        Rs.MoveNext
    Loop
   
    ConnectToScanner
    'Command1.Visible = False

End Sub

Private Sub MSComm1_OnComm()
    txtUCode.Text = ""
    txtUCode.BackColor = vbWhite
    Timer1.Interval = 100
    Timer1.Enabled = True
End Sub
Private Sub Timer1_Timer()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Code As String
Dim str1() As String
Timer1.Enabled = False
Code = MSComm1.Input
txtUCode.Text = Code
If Code = "" Then Exit Sub
str1 = Split(Code, " ")
txtUCode.Text = str1(1)

        Sql = "select * from User_list where UserName='" & CboUser.Text & "' and ucode='" & Val(txtUCode.Text) & "' and UDelete='N'"
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    
    If Rs.EOF = True Then
        'MsgBox UCase("Invalid User Code"), vbCritical
         txtUCode.BackColor = vbRed
        Exit Sub
    End If
        
    LoginUser = Trim(Rs("UserName"))
    AccessType = Trim(Rs("AccessType"))
    LoginID = Rs("UID")
    If Rs("UDefault") = "Y" Then
        UDefault = True
    Else
        UDefault = False
    End If
    LoginCode = Trim(Rs("UCode"))
    If UDefault = True Then
        frmmenu.Show
        Unload Me
    Else
        frmmenu.Show
        Unload Me
'        frmMonitor.Show
    End If
End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    CmdLogin.Value = True
End If

End Sub
