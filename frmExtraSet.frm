VERSION 5.00
Begin VB.Form frmExtraSet 
   Caption         =   "Extra Settings"
   ClientHeight    =   8205
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14085
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   14085
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000010&
      Height          =   7215
      Left            =   1320
      ScaleHeight     =   7155
      ScaleWidth      =   11880
      TabIndex        =   0
      Top             =   120
      Width           =   11940
      Begin VB.Frame Frame6 
         Caption         =   "Send Mail"
         Height          =   5415
         Left            =   5520
         TabIndex        =   33
         Top             =   120
         Width           =   6255
         Begin VB.CheckBox chkTomailbypass 
            BackColor       =   &H00808080&
            Caption         =   "Bypass"
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   0
            Left            =   2040
            TabIndex        =   60
            Top             =   2280
            Width           =   975
         End
         Begin VB.CheckBox chkTomailbypass 
            BackColor       =   &H00808080&
            Caption         =   "Bypass"
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   5
            Left            =   2040
            TabIndex        =   59
            Top             =   4680
            Width           =   975
         End
         Begin VB.CheckBox chkTomailbypass 
            BackColor       =   &H00808080&
            Caption         =   "Bypass"
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   4
            Left            =   2040
            TabIndex        =   58
            Top             =   4200
            Width           =   975
         End
         Begin VB.CheckBox chkTomailbypass 
            BackColor       =   &H00808080&
            Caption         =   "Bypass"
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   3
            Left            =   2040
            TabIndex        =   57
            Top             =   3720
            Width           =   975
         End
         Begin VB.CheckBox chkTomailbypass 
            BackColor       =   &H00808080&
            Caption         =   "Bypass"
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   2
            Left            =   2040
            TabIndex        =   56
            Top             =   3240
            Width           =   975
         End
         Begin VB.CheckBox chkTomailbypass 
            BackColor       =   &H00808080&
            Caption         =   "Bypass"
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   1
            Left            =   2040
            TabIndex        =   55
            Top             =   2760
            Width           =   975
         End
         Begin VB.TextBox txtTomail 
            Height          =   360
            Index           =   6
            Left            =   3120
            TabIndex        =   53
            Top             =   4680
            Width           =   3015
         End
         Begin VB.TextBox txtTomail 
            Height          =   360
            Index           =   5
            Left            =   3120
            TabIndex        =   51
            Top             =   4200
            Width           =   3015
         End
         Begin VB.TextBox txtTomail 
            Height          =   360
            Index           =   4
            Left            =   3120
            TabIndex        =   49
            Top             =   3720
            Width           =   3015
         End
         Begin VB.TextBox txtTomail 
            Height          =   360
            Index           =   3
            Left            =   3120
            TabIndex        =   45
            Top             =   3240
            Width           =   3015
         End
         Begin VB.TextBox txtTomail 
            Height          =   360
            Index           =   2
            Left            =   3120
            TabIndex        =   44
            Top             =   2760
            Width           =   3015
         End
         Begin VB.TextBox txtTomail 
            Height          =   360
            Index           =   1
            Left            =   3120
            TabIndex        =   43
            Top             =   2280
            Width           =   3015
         End
         Begin VB.TextBox txtTomail 
            Height          =   360
            Index           =   0
            Left            =   3120
            TabIndex        =   41
            Text            =   "ABC@GMAIL"
            Top             =   1800
            Width           =   3015
         End
         Begin VB.TextBox txtfromPassword 
            Height          =   360
            Left            =   3120
            TabIndex        =   37
            Text            =   "PASSWORD"
            Top             =   1320
            Width           =   3015
         End
         Begin VB.TextBox txtfromMail 
            Height          =   360
            Left            =   3120
            TabIndex        =   36
            Text            =   "mAIL"
            Top             =   840
            Width           =   3015
         End
         Begin VB.TextBox txtapilink 
            Height          =   360
            Left            =   3120
            TabIndex        =   35
            Text            =   "www.MindaRikaSendEmail.com"
            Top             =   360
            Width           =   3015
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Bypass"
            Height          =   255
            Left            =   1320
            TabIndex        =   34
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "TO MAIL 7"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   17
            Left            =   120
            TabIndex        =   54
            Top             =   4680
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "TO MAIL 6"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   16
            Left            =   120
            TabIndex        =   52
            Top             =   4200
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "TO MAIL 5"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   15
            Left            =   120
            TabIndex        =   50
            Top             =   3720
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "TO MAIL 2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   14
            Left            =   120
            TabIndex        =   48
            Top             =   2280
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "TO MAIL 3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   13
            Left            =   120
            TabIndex        =   47
            Top             =   2760
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "TO MAIL 4"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   12
            Left            =   120
            TabIndex        =   46
            Top             =   3240
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "TO MAIL 1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   11
            Left            =   120
            TabIndex        =   42
            Top             =   1800
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "Web Api Link"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   10
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "FROM MAIL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   9
            Left            =   120
            TabIndex        =   39
            Top             =   840
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "FROM PASSWORD"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   38
            Top             =   1320
            Width           =   2895
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1815
         Left            =   120
         TabIndex        =   24
         Top             =   2040
         Width           =   5295
         Begin VB.CheckBox ChkComPortBP2 
            BackColor       =   &H00808080&
            Caption         =   "ByPass"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   360
            Left            =   2040
            TabIndex        =   30
            Top             =   1320
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.ComboBox CboCom2 
            Height          =   360
            Left            =   3480
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   1320
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txtPLC_Port 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2520
            TabIndex        =   27
            Text            =   "1234"
            Top             =   840
            Width           =   2655
         End
         Begin VB.TextBox txtPLC_IP 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2520
            TabIndex        =   26
            Text            =   "192.168.1.12"
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "Melexis ComPort"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   360
            Index           =   6
            Left            =   120
            TabIndex        =   31
            Top             =   1320
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "PLC Port"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   360
            Index           =   4
            Left            =   120
            TabIndex        =   28
            Top             =   840
            Width           =   2295
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "PLC IP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   360
            Index           =   3
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Other"
         Height          =   855
         Left            =   120
         TabIndex        =   15
         Top             =   4680
         Width           =   5295
         Begin VB.CheckBox ChkNetworkDBBP 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00808080&
            Caption         =   "SQL DB ByPass"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   360
            Left            =   120
            TabIndex        =   23
            Top             =   1440
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox txtNetworkDB 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2160
            TabIndex        =   20
            Text            =   "M1"
            Top             =   840
            Visible         =   0   'False
            Width           =   7935
         End
         Begin VB.TextBox txtPrinterName1 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2160
            TabIndex        =   18
            Text            =   "ZDesigner ZD230-203dpi ZPL"
            Top             =   360
            Width           =   3015
         End
         Begin VB.TextBox txtMachineNo 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2160
            TabIndex        =   17
            Text            =   "M1"
            Top             =   360
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   "Provider=MSDASQL;DRIVER=Sql Server;SERVER=ServerName; DATABASE=DataBase Name; UID=UserName; PWD=Password;"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   1200
            Width           =   10215
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "SQL DB Path"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   360
            Index           =   8
            Left            =   120
            TabIndex        =   21
            Top             =   840
            Visible         =   0   'False
            Width           =   1950
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            Caption         =   "Printer Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   360
            Index           =   5
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   1485
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "Machine No"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   360
            Index           =   2
            Left            =   210
            TabIndex        =   16
            Top             =   360
            Visible         =   0   'False
            Width           =   1995
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "ComPort Selection"
         Height          =   975
         Left            =   120
         TabIndex        =   11
         Top             =   3960
         Visible         =   0   'False
         Width           =   5295
         Begin VB.CheckBox ChkComPortBP1 
            BackColor       =   &H00808080&
            Caption         =   "ByPass"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   360
            Left            =   1920
            TabIndex        =   22
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox CboCom1 
            Height          =   360
            Left            =   3360
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "PLC ComPort"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   360
            Index           =   50
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000008&
         BorderStyle     =   0  'None
         Height          =   1440
         Left            =   120
         TabIndex        =   8
         Top             =   5640
         Width           =   11655
         Begin VB.CommandButton cmdSave 
            BackColor       =   &H80000014&
            Height          =   1005
            Left            =   3600
            Picture         =   "frmExtraSet.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Save to Modulaor.Mdb"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdClose 
            BackColor       =   &H80000012&
            Height          =   1005
            Left            =   5880
            Picture         =   "frmExtraSet.frx":3D04
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Close Screen"
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Time Shift"
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   5295
         Begin VB.CheckBox ChkShift4 
            Caption         =   "Check For General"
            Height          =   255
            Left            =   1200
            TabIndex        =   14
            Top             =   0
            Width           =   1935
         End
         Begin VB.TextBox txtShift1 
            Height          =   360
            Left            =   2160
            TabIndex        =   4
            Text            =   "06:00 AM"
            Top             =   360
            Width           =   2415
         End
         Begin VB.TextBox txtShift2 
            Height          =   360
            Left            =   2160
            TabIndex        =   3
            Text            =   "02:00 PM"
            Top             =   840
            Width           =   2415
         End
         Begin VB.TextBox txtShift3 
            Height          =   360
            Left            =   2160
            TabIndex        =   2
            Text            =   "09:00 PM"
            Top             =   1320
            Width           =   2415
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "Shift 3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   7
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "Shift 2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "Shift 1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   57
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   1935
         End
      End
   End
End
Attribute VB_Name = "frmExtraSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Enum BasicAction
    Load = 1
    Save = 2
End Enum

Private Sub CmdClose_Click()
    frmmenu.Show
    Unload Me
End Sub

Private Sub CmdSave_Click()
On Error GoTo Error
'Dim Sql As String
'Dim Rs As ADODB.Recordset

    If ChkShift4.Value = 0 Then
        If getShiftValid(txtShift1, txtShift2, txtShift3) = False Then Exit Sub
    End If
    
    If ValidLen(1, 3, txtMachineNo) = False Then Exit Sub
    

    If ValidLen(1, 255, txtNetworkDB) = False Then Exit Sub

    ExtraSetting Save
    
    ResetForm

    MsgBox "Saved Successfully", vbInformation

    ExtraSetting Load
    
Exit Sub
Error:
MsgBox Error, vbInformation
End Sub

''Private Sub Command1_Click()
''    txtMachineNo = getShift
''
''    'PrintLabel
''End Sub

Private Sub Form_Load()
On Error GoTo Error

'Advance
Me.WindowState = 2
Me.BackColor = &H80000010
Picture1.BorderStyle = 1
Picture1.Appearance = 0
Picture1.BackColor = vbButtonFace
Picture1.Left = (Screen.Width - Picture1.Width) / 2
Picture1.Top = (Screen.Height - Picture1.Height) / 2 - 400

    'My cboCom Load
    For i = 0 To 19
        CboCom1.AddItem "Com " & i + 1, i
        CboCom1.ListIndex = 0
        CboCom2.AddItem "Com " & i + 1, i
        CboCom2.ListIndex = 0
    Next
    
    ExtraSetting Load


Exit Sub
Error:
MsgBox Error, vbInformation
End Sub

Private Sub ExtraSetting(Action As BasicAction)
On Error GoTo Error
Dim Sql As String
Dim Rs As ADODB.Recordset

    Sql = "Select * from Common_Set where SetType ='CommonSet'" 'SetType = Settings Type
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic

    If Action = 1 Then
        CboCom1.ListIndex = (Rs("ComPort1") - 1)
       
        ChkComPortBP1.Value = Rs("ComPortBP1")
        txtShift1 = Rs("Shift1")
        txtShift2 = Rs("Shift2")
        txtShift3 = Rs("Shift3")
        ChkShift4.Value = Rs("Shift4")
        txtMachineNo = Rs("MachineNo")
        txtPrinterName1 = Rs("PrinterName1")
        txtNetworkDB = Rs("NetworkDB")
        ChkNetworkDBBP.Value = Rs("NetworkDBBP")
        txtPLC_IP = Rs("PLC_IP")
        txtPLC_Port = Rs("PLC_Port")
        CboCom2.ListIndex = (Rs("ComPort2") - 1)
        ChkComPortBP2.Value = Rs("ComPortBP2")
        txtapilink.Text = Rs("WebApiLink")
        txtfromMail.Text = Rs("SenderEmail")
        txtfromPassword.Text = Rs("SenderPassword")
        txtTomail(0).Text = Rs("ToEmail1")
        txtTomail(1).Text = Rs("ToEmail2")
        txtTomail(2).Text = Rs("ToEmail3")
        txtTomail(3).Text = Rs("ToEmail4")
        txtTomail(4).Text = Rs("ToEmail5")
        txtTomail(5).Text = Rs("ToEmail6")
        txtTomail(6).Text = Rs("ToEmail7")
        Check1.Value = Val(Rs("EmailBypass"))
        For i = 0 To 5
         chkTomailbypass(i).Value = Val(Rs("EmailBypass" & i + 1))
        Next
    ElseIf Action = 2 Then
        Rs("ComPort1") = CboCom1.ListIndex + 1
        Rs("ComPort2") = CboCom2.ListIndex + 1
        Rs("ComPortBP2") = ChkComPortBP2.Value
        Rs("ComPortBP1") = ChkComPortBP1.Value
        Rs("Shift1") = Trim(txtShift1.Text)
        Rs("Shift2") = Trim(txtShift2.Text)
        Rs("Shift3") = Trim(txtShift3.Text)
        Rs("Shift4") = Val(ChkShift4.Value)
        Rs("MachineNo") = Trim$(txtMachineNo)
        Rs("PrinterName1") = Trim$(txtPrinterName1)
        Rs("NetworkDB") = Trim$(txtNetworkDB)
        Rs("NetworkDBBP") = ChkNetworkDBBP.Value
        Rs("PLC_IP") = Trim$(txtPLC_IP.Text)
        Rs("PLC_Port") = Trim$(txtPLC_Port.Text)
        Rs("EmailBypass") = Check1.Value
        Rs("WebApiLink") = txtapilink.Text
        Rs("SenderEmail") = txtfromMail.Text
        Rs("SenderPassword") = txtfromPassword.Text
        Rs("ToEmail1") = txtTomail(0).Text
        Rs("ToEmail2") = txtTomail(1).Text
        Rs("ToEmail3") = txtTomail(2).Text
        Rs("ToEmail4") = txtTomail(3).Text
        Rs("ToEmail5") = txtTomail(4).Text
        Rs("ToEmail6") = txtTomail(5).Text
        Rs("ToEmail7") = txtTomail(6).Text
        For i = 0 To 5
           Rs("EmailBypass" & i + 1) = chkTomailbypass(i).Value
        Next
        
        Rs.Update
    End If

Exit Sub
Error:
MsgBox "Error in ComPort Setting" & vbNewLine & Error, vbInformation
End Sub

Private Function ValidLen(ByVal Min, Max As Long, Text As TextBox) As Boolean

    If Trim(Text) = "" Or (Len(Text) < Min Or Len(Text) > Max) Then
        MsgBox ("Kindly Enter Between " & Min & " To " & Max & " Characters"), vbCritical
        Text.SetFocus
        Text.SelStart = 0
        Text.SelLength = Len(Text)
        Text.BackColor = vbRed
        ValidLen = False
    Else
        Text.BackColor = vbWhite
        ValidLen = True
    End If

End Function

Private Function ValidEntry(ByVal Min, Max As Double, Text As TextBox) As Boolean

    If IsNumeric(Text) = False Or (Val(Text) < Min Or Val(Text) > Max) Then
        MsgBox ("Kindly Enter Between " & Min & " To " & Max), vbInformation
        Text.SetFocus
        Text.SelStart = 0
        Text.SelLength = Len(Text)
        Text.BackColor = vbRed
        ValidEntry = False
    Else
        Text.BackColor = vbWhite
        ValidEntry = True
    End If

End Function

Private Function getShiftValid(sTime1, sTime2, sTime3 As String) As Boolean
On Error GoTo Error
Dim AmPm(2) As String
Dim TmpLan(2) As String
'Const MaxTime = "11:59 PM"

    TmpLan(0) = Len(sTime1)
    TmpLan(1) = Len(sTime2)
    TmpLan(2) = Len(sTime3)
    AmPm(0) = Right(sTime1, 2)
    AmPm(1) = Right(sTime2, 2)
    AmPm(2) = Right(sTime3, 2)

    For i = 0 To 2

        If (TmpLan(i) > "8" Or TmpLan(i) < "8") Then
            MsgBox "Please Enter a Valid Time", vbInformation
            Exit Function

        End If

        If (AmPm(i) = "AM" Or AmPm(i) = "PM") Then
        Else
            MsgBox "Enter a Valid Time With AM/PM", vbInformation
            Exit Function

        End If

    Next

    If TimeValue(sTime2) < TimeValue(sTime1) Then
        MsgBox "Shift 2 Time Should Be Greater Then Shift 1", vbInformation
        Exit Function

    End If

    If TimeValue(sTime3) < TimeValue(sTime2) Then
        MsgBox "Shift 3 Time Should Be Greater Then Shift 2", vbInformation
        Exit Function

    End If

    getShiftValid = True

Exit Function
Error:
MsgBox "Error Found in Time Shift Validation" & vbNewLine & Error, vbInformation
End Function

Private Sub ResetForm()
Dim txt As Control

For Each txt In Me
    If TypeOf txt Is TextBox Then
        txt.Text = ""
    End If
Next

For Each txt In Me
    If TypeOf txt Is CheckBox Then
        txt.Value = 0
    End If
Next

End Sub

