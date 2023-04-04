VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmmenu 
   Caption         =   "Select Your Menu"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   ScaleHeight     =   7770
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Manual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Chane Parameters of Testing"
      Top             =   4440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdBackup 
      Caption         =   "&Data Backup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12000
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmmenu.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Gives backup for main Modulator_Testing.mdb"
      Top             =   3240
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton cmdReadLabel 
      Caption         =   "&Read Bar CodeLabel"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   12000
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmmenu.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   " Read BarCode on product"
      Top             =   2040
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000010&
      Height          =   5775
      Left            =   240
      ScaleHeight     =   5715
      ScaleWidth      =   8280
      TabIndex        =   0
      Top             =   240
      Width           =   8340
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3000
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmmenu.frx":1298
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Gives backup for main Modulator_Testing.mdb"
         Top             =   4200
         UseMaskColor    =   -1  'True
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Counter Setting"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   720
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmmenu.frx":3A0F
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Gives backup for main Modulator_Testing.mdb"
         Top             =   4200
         UseMaskColor    =   -1  'True
         Width           =   2175
      End
      Begin VB.PictureBox PictureWI 
         Height          =   2895
         Left            =   960
         ScaleHeight     =   2835
         ScaleWidth      =   9195
         TabIndex        =   15
         Top             =   5280
         Visible         =   0   'False
         Width           =   9255
         Begin VB.Image Image1 
            Height          =   735
            Left            =   0
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.CommandButton CmdPrintLabel 
         Caption         =   "&Print BarCodeLabel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3000
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmmenu.frx":5E8F
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Print Duplicate Label"
         Top             =   3000
         UseMaskColor    =   -1  'True
         Width           =   2175
      End
      Begin VB.CommandButton cmdReports1 
         Caption         =   "&Production Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   720
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmmenu.frx":695A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Gives backup for main Modulator_Testing.mdb"
         Top             =   3000
         UseMaskColor    =   -1  'True
         Width           =   2175
      End
      Begin VB.CommandButton CmdGenSet 
         Caption         =   "&General Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   720
         Picture         =   "frmmenu.frx":7355
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Chane Parameters of Testing"
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CommandButton cmdReports 
         Caption         =   "&Excel Report Generation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3240
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmmenu.frx":8499
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Gives backup for main Modulator_Testing.mdb"
         Top             =   4320
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.PictureBox Picture2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   720
         ScaleHeight     =   1035
         ScaleWidth      =   4395
         TabIndex        =   7
         Top             =   480
         Width           =   4455
         Begin VB.ComboBox CboModelName 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   480
            Width           =   4215
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Model Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   1245
         End
      End
      Begin VB.CommandButton cmdUserConfig 
         Caption         =   "&User Configuration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   5280
         MaskColor       =   &H00FFC0C0&
         Picture         =   "frmmenu.frx":8E94
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Change / Add / Delete User"
         Top             =   1680
         UseMaskColor    =   -1  'True
         Width           =   2175
      End
      Begin VB.CommandButton cmdsettings 
         Caption         =   "&Change Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3000
         Picture         =   "frmmenu.frx":A6C0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Chane Parameters of Testing"
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CommandButton cmdMonitor 
         Caption         =   "&Testing Screen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   5280
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmmenu.frx":B804
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Testing Screen"
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   2175
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   5280
         Picture         =   "frmmenu.frx":C37D
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exit the Application"
         Top             =   3000
         Width           =   2175
      End
   End
   Begin MSComDlg.CommonDialog CDIALBOX 
      Left            =   13320
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FSO As New FileSystemObject

Private Sub cmdBackup_Click()
    On Error Resume Next
    DataBackup
End Sub


Public Sub cmdExit_Click()
    'On Error Resume Next
    If MsgBox(UCase("Are you sure to quit the application"), vbYesNo + vbInformation) = vbYes Then
        Con.Close
        End
    Else
        Cancel = 1
    End If
End Sub

Private Sub CmdGenSet_Click()
    frmExtraSet.Show
    Unload Me
End Sub

Private Sub cmdMonitor_Click()
    Unload Me
    frmMonitor.Show
End Sub

Private Sub CmdPrintLabel_Click()
    frmPrintLabel.Show
    Unload Me
End Sub

Private Sub cmdReports_Click()
Dim FolderName As String

FolderName = App.Path & "\ExCelReport\"
'Shell "explorer " & FolderName,  vbNormalFocus
Shell "Explorer.exe /e,/root,""" & FolderName & """", vbNormalFocus   'Working

''    frmReport.Show
''    Unload Me
End Sub

Private Sub cmdReports1_Click()
    frmReportType.Show
    Unload Me
End Sub

Private Sub cmdsettings_Click()
    frmsettings.Show
    Unload Me
End Sub

Private Sub cmdUserConfig_Click()
    frmSetUser.Show
    Unload Me
End Sub

Private Sub cmdReadLabel_Click()
'    frmReadLabel.Show
frmFullReport.Show
    Me.Hide
End Sub

Private Sub Command1_Click()
    frmmanual.Show
    Unload Me
End Sub

Private Sub Command2_Click()
    frmSettings2.Show
    Unload Me
End Sub

Private Sub Command3_Click()
    frmmanual.Show
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo Error

'Advance
Me.WindowState = 2
Me.BackColor = vbButtonFace '&H80000010
Picture1.BorderStyle = 1
Picture1.Appearance = 0
Picture1.BackColor = &H80000010 'vbButtonFace
Picture1.Left = (Screen.Width - Picture1.Width) / 2
Picture1.Top = (Screen.Height - Picture1.Height) / 2 - 400
'
'Image1.Width = PictureWI.Width
'Image1.Height = PictureWI.Height
'Image1.Stretch = True


LoadModelCombo cbomodelname
ModelName = GetSetting(App.Title, "LastModel", "LastModel")
LastModel ModelName, cbomodelname
    
UserAccess 'Loas Controls Using User
    
Exit Sub
Error:
MsgBox Error, vbInformation
End Sub
Private Sub CboModelName_Click()

ModelName = cbomodelname.Text
SaveSetting App.Title, "LastModel", "LastModel", ModelName
'ModelPicture Image1, ModelName
End Sub
Private Sub ModelPicture(mImage As Image, mPictureName As String)
On Error Resume Next

mImage.Stretch = True
mImage.Picture = LoadPicture(App.Path & "\WI\" & mPictureName & ".jpg")

End Sub
Private Sub LoadModelCombo(Combo As ComboBox)
Dim Sql As String
Dim rs As ADODB.Recordset
Dim i As Integer

    Combo.Clear
    Sql = "Select * from Model_Set order by ModelName"
    Set rs = New ADODB.Recordset
    rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    Do While rs.EOF = False
        Combo.AddItem rs("ModelName")
        rs.MoveNext
    Loop
    
End Sub

Private Sub LastModel(ByVal Model As String, Combo As ComboBox)
Dim Sql As String
Dim rs As ADODB.Recordset

    Sql = "Select * from Model_Set where ModelName='" & Model & "'"
    Set rs = New ADODB.Recordset
    rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    If rs.EOF = False Then
        Combo.Text = Model
    Else
        Combo.ListIndex = 0
    End If

End Sub

Private Sub DataBackup()
    
     Dim Str As String
     CDIALBOX.Filter = "*.mdb|*.mdb"
     CDIALBOX.CancelError = True
     CDIALBOX.DialogTitle = "Save File"
     CDIALBOX.Action = 2
     'CDIALBOX.ShowSave
     If Trim(CDIALBOX.FileTitle) = "" Then
        Exit Sub
     End If
    
    On Error GoTo Err_Hndlr1
     
     Str = CDIALBOX.FileName
     Set FSO = New FileSystemObject
     FSO.CopyFile App.Path & "\database\DB_MS_Ford_Key.mdb", Str, True
     MsgBox UCase("Database Backup Completed"), vbOKOnly
    Exit Sub

Err_Hndlr1:
    MsgBox UCase("Error while taking Back up"), vbCritical
    
End Sub

Private Sub UserAccess()

If AccessType = "0" Then 'Disable or Hide For Operators
    cmdsettings.Enabled = False
    'cmdUserConfig.Enabled = False
    cmdReports1.Enabled = False
    cmdBackup.Enabled = False
    CmdGenSet.Enabled = False
    CmdPrintLabel.Visible = False
    Command3.Visible = False
    Command2.Visible = False
    
ElseIf AccessType = "1" Then 'Disable or Hide for AccessType 1
'    cmdsettings.Enabled = False
    'CmdGenSet.Enabled = False
    cmdsettings.Enabled = True
    'cmdUserConfig.Enabled = False
    cmdReports1.Enabled = True
    cmdBackup.Enabled = False
    CmdGenSet.Enabled = True
    CmdPrintLabel.Visible = False
    Command3.Visible = False
    Command2.Visible = False
ElseIf AccessType = "2" Then 'Show All Which Will Disable or Hide For One

End If

End Sub

Private Sub MMControl1_Done(NotifyCode As Integer)

End Sub
