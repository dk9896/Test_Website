VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmSetUser 
   Caption         =   "Set User"
   ClientHeight    =   8115
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14910
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
   ScaleHeight     =   8115
   ScaleWidth      =   14910
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   120
      ScaleHeight     =   5235
      ScaleWidth      =   12795
      TabIndex        =   0
      Top             =   120
      Width           =   12855
      Begin VSFlex7Ctl.VSFlexGrid VSFUsers 
         Height          =   3975
         Left            =   5865
         TabIndex        =   13
         Top             =   840
         Width           =   6780
         _cx             =   11959
         _cy             =   7011
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   12494734
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   12494734
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   25
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   400
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmSetUser.frx":0000
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Frame Frame1 
         Height          =   3165
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   5610
         Begin VB.TextBox txtRPwd 
            Height          =   390
            IMEMode         =   3  'DISABLE
            Left            =   2040
            PasswordChar    =   "*"
            TabIndex        =   20
            Top             =   2640
            Width           =   3450
         End
         Begin VB.TextBox txtPwd 
            Height          =   390
            IMEMode         =   3  'DISABLE
            Left            =   2040
            PasswordChar    =   "*"
            TabIndex        =   19
            Top             =   2040
            Width           =   3450
         End
         Begin VB.TextBox txtuid 
            Enabled         =   0   'False
            Height          =   375
            Left            =   4320
            TabIndex        =   18
            Top             =   840
            Width           =   990
         End
         Begin VB.TextBox txtUCode 
            Height          =   390
            IMEMode         =   3  'DISABLE
            Left            =   2040
            TabIndex        =   17
            Top             =   840
            Width           =   1170
         End
         Begin VB.TextBox txtUName 
            Height          =   390
            Left            =   2040
            TabIndex        =   15
            Top             =   240
            Width           =   3390
         End
         Begin VB.ComboBox CboAccessType 
            Height          =   360
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1440
            Width           =   2295
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "User ID "
            Height          =   420
            Left            =   3360
            TabIndex        =   21
            Top             =   840
            Width           =   870
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "User Name "
            Height          =   420
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   1755
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Password "
            Height          =   420
            Left            =   120
            TabIndex        =   11
            Top             =   2040
            Width           =   1755
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "ReType Password "
            Height          =   420
            Left            =   120
            TabIndex        =   10
            Top             =   2640
            Width           =   1755
         End
         Begin VB.Label Code 
            Alignment       =   1  'Right Justify
            Caption         =   "Code "
            Height          =   420
            Left            =   120
            TabIndex        =   9
            Top             =   840
            Width           =   1710
         End
         Begin VB.Label lblAccessType 
            Alignment       =   1  'Right Justify
            Caption         =   "AccessType "
            Height          =   420
            Left            =   120
            TabIndex        =   8
            Top             =   1440
            Width           =   1740
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   1
         Top             =   3840
         Width           =   5610
         Begin VB.CommandButton CmdSave 
            Caption         =   "&Save"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   2760
            Picture         =   "frmSetUser.frx":00AE
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Width           =   1230
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   4080
            Picture         =   "frmSetUser.frx":11F2
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            Width           =   1230
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   120
            Picture         =   "frmSetUser.frx":2336
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   240
            Width           =   1230
         End
         Begin VB.CommandButton cmdReset 
            Caption         =   "&Reset"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   1440
            Picture         =   "frmSetUser.frx":347A
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   1230
         End
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "User Configuration"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   12495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Edit User Double Click on that particular row"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   12
         Top             =   4920
         Width           =   4605
      End
   End
End
Attribute VB_Name = "frmSetUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UID As Integer
Dim CloseScreen As Boolean

Private Sub cmdClose_Click()

CloseScreen = True
CloseMe

End Sub

Private Sub CloseMe()

Unload Me
frmmenu.Show

End Sub

Private Sub Form_Unload(Cancel As Integer)

If CloseScreen = False Then
    CloseMe
Else
    CloseScreen = False
End If

End Sub

Private Sub cmdDelete_Click()
 On Error GoTo Err_Hndlr
    Dim Sql As String
    Dim Rs As ADODB.Recordset
    
    If MsgBox(UCase("Are you sure to delete the user"), vbYesNo + vbDefaultButton2) = vbYes Then
        Sql = "Select * from user_list where UID=" & Val(txtuid) & ""
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
        If Rs.EOF = False Then
            Sql = "update User_list set udelete='Y' where UID=" & Val(txtuid) & ""
        Else
            Sql = "delete from User_list where UID=" & Val(txtuid) & ""
        End If
        Con.Execute Sql
        Rs.Update
        MsgBox UCase("User Deleted Successfully"), vbInformation
        ResetForm
    End If
    Exit Sub
Err_Hndlr:
    MsgBox UCase("Error in Deleting User."), vbCritical
End Sub

Private Sub cmdReset_Click()
    On Error Resume Next
    ResetForm
End Sub
Private Sub CmdSave_Click()
    On Error GoTo Err_Hndlr
    Dim Rs As ADODB.Recordset
    Dim Sql As String
   
'    Dim FlagCheck As Boolean
    
    If CheckValidEntry = False Then Exit Sub
    
    If Trim(txtuid) = "" Then
        FlagCheck = True
    ElseIf Trim(txtUName) = Trim(txtUNameOld) Then
        UID = Trim(txtuid)
        FlagCheck = False
    Else
        FlagCheck = True
    End If
    
    
    If Trim(txtuid) = "" Then
        FlagCheck = True
    ElseIf Trim(txtUCode) = Trim(txtUCodeOld) Then
        UID = Trim(txtuid)
        FlagCheck = False
    Else
        FlagCheck = True
    End If
    
        Sql = "Select * from User_list where UCode='" & Trim(txtUCode) & "' and UDelete='N'"
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
       
    If Trim(txtuid) = "" Then
        Sql = "Select max(UID) from User_list"
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
        If IsNull(Rs(0)) = True Then
            UID = 1
        Else
            UID = Rs(0) + 1
        End If
    End If
    
''    Rs("AccessType") = CboAccessType.ListIndex
    
    If Trim(txtuid) = "" Then
        Sql = "Insert into User_list (username, ucode, pwd, udefault, uid, udelete, AccessType) values ('" & Trim(txtUName) & "', '" & Trim(txtUCode) & "', '" & txtPwd & "', 'Y', " & UID & ", 'N', '" & CboAccessType.ListIndex & "')"
    Else
        Sql = "update User_list set username='" & Trim(txtUName) & "', ucode='" & Trim(txtUCode) & "', pwd='" & txtPwd & "', AccessType='" & CboAccessType.ListIndex & "' where UID=" & txtuid & ""
    End If
    Con.Execute Sql
    MsgBox UCase("Saved Successfully"), vbInformation
    FillUserGrid
    ResetForm
    Exit Sub
Err_Hndlr:
    MsgBox UCase("Error while saving Data. Try to save it again"), vbCritical
End Sub

Private Sub Form_Load()
On Error Resume Next
    
'Advance
Me.WindowState = 2
Me.BackColor = &H80000010
Picture1.BorderStyle = 1
Picture1.Appearance = 0
Picture1.BackColor = vbButtonFace
Picture1.Left = (Screen.Width - Picture1.Width) / 2
Picture1.Top = (Screen.Height - Picture1.Height) / 2 - 400

'    SetSize Me
    
    VSFUsers.Rows = 1
    VSFUsers.ColHidden(3) = True
    VSFUsers.ColHidden(5) = True
    FillUserGrid
    cmdDelete.Enabled = False
    
      CboAccessType.AddItem "Normal", 0
      CboAccessType.AddItem "Medium", 1
      CboAccessType.AddItem "Full Access", 2
      CboAccessType.ListIndex = 0
      
      If AccessType > 1 Then
        CboAccessType.Visible = True
        lblAccessType.Visible = True
      Else
        CboAccessType.Visible = False
        lblAccessType.Visible = False
      End If

End Sub
Private Function CheckValidEntry() As Boolean
    On Error Resume Next
    If Trim(txtUName) = "" Or Len(Trim(txtUName)) > 50 Then
        MsgBox UCase("Enter Valid User Name. Max Len 50"), vbCritical
        txtUName.SetFocus
        txtUName.SelStart = 0
        txtUName.SelLength = Len(txtUName)
        CheckValidEntry = False
        Exit Function
    End If

    If Trim(txtUCode) = "" Or Len(Trim(txtUCode)) > 10 Then
        MsgBox UCase("Enter Valid User Name. Max Len 10"), vbCritical
        txtUCode.SetFocus
        txtUCode.SelStart = 0
        txtUCode.SelLength = Len(txtUCode)
        CheckValidEntry = False
        Exit Function
    End If

    If Trim(txtPwd) = "" Or Len(txtPwd) > 50 Then
        MsgBox UCase("Enter Valid Password. Max Len 50"), vbCritical
        txtPwd.SetFocus
        txtPwd.SelStart = 0
        txtPwd.SelLength = Len(txtPwd)
        CheckValidEntry = False
        Exit Function
    End If
    
    If Trim(txtRPwd) = "" Then
        MsgBox UCase("Password Not Confirmed"), vbCritical
        txtRPwd.SetFocus
        CheckValidEntry = False
        Exit Function
    End If

    If txtPwd <> txtRPwd Then
        MsgBox UCase("Password Not Confirmed"), vbCritical
        txtRPwd = ""
        txtPwd.SetFocus
        txtPwd.SelStart = 0
        txtPwd.SelLength = Len(txtPwd)
        CheckValidEntry = False
        Exit Function
    End If

    CheckValidEntry = True
End Function
Private Sub ResetForm()
    On Error Resume Next
    Dim Obj As Object
    
    For Each Obj In Controls
        If TypeOf Obj Is TextBox Then
             Obj.Text = ""
        End If
    Next
    
    'If cmbTeam.ListCount > 0 Then cmbTeam.ListIndex = 0
    
    txtUName.SetFocus
    FillUserGrid
    cmdDelete.Enabled = False
End Sub

Private Sub FillUserGrid()
    On Error Resume Next
    Dim Row As Integer
    Dim Rs As ADODB.Recordset
    Dim Sql As String
    
    VSFUsers.Rows = 1
    Row = 0
    
    If AccessType = 0 Then
        Sql = "Select * from User_list where UDelete='N' and UID=" & LoginID & ""
        txtUName.Enabled = False
        txtUCode.Enabled = False
    ElseIf AccessType = 1 Then
        Sql = "Select * from User_list where UDelete='N' and AccessType ='0'"
    Else
        Sql = "Select * from User_list where UDelete='N' "
    End If
    
'    Sql = "Select * from User_list where UDelete='N' "
    ' "and (UDefault='N' or UID=" & LoginID & ") order by username"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    Rs.MoveFirst
    Do While Rs.EOF = False
        VSFUsers.Rows = VSFUsers.Rows + 1
        Row = Row + 1
        VSFUsers.TextMatrix(Row, 0) = Trim(Row)
        VSFUsers.TextMatrix(Row, 1) = Trim(Rs("UserName"))
        VSFUsers.TextMatrix(Row, 2) = Trim(Rs("UCode"))
        VSFUsers.TextMatrix(Row, 3) = Trim(Rs("Pwd"))
        VSFUsers.TextMatrix(Row, 4) = Trim(Rs("UID"))
        VSFUsers.TextMatrix(Row, 5) = Trim(Rs("AccessType"))
        Rs.MoveNext
    Loop
    
End Sub

Private Sub txtPwd_GotFocus()
    On Error Resume Next
    txtPwd.SetFocus
    txtPwd.SelStart = 0
    txtPwd.SelLength = Len(txtPwd)
End Sub

Private Sub txtRPwd_GotFocus()
 On Error Resume Next
    txtRPwd.SetFocus
    txtRPwd.SelStart = 0
    txtRPwd.SelLength = Len(txtRPwd)
End Sub


Private Sub txtUName_GotFocus()
    On Error Resume Next
    txtUName.SetFocus
    txtUName.SelStart = 0
    txtUName.SelLength = Len(txtUName)
End Sub

Private Sub VSFUsers_DblClick()
    On Error Resume Next
    Dim Obj As Object
    Dim Row As Integer
    
    If VSFUsers.Rows = 1 Then Exit Sub
    If VSFUsers.Row < 1 Then Exit Sub
    
    Row = VSFUsers.Row
    
    For Each Obj In Controls
        If TypeOf Obj Is TextBox Then
             Obj.Text = ""
        End If
    Next
    
    txtUName = VSFUsers.TextMatrix(Row, 1)
    txtUNameOld = VSFUsers.TextMatrix(Row, 1)
    txtUCode = VSFUsers.TextMatrix(Row, 2)
    txtUCodeOld = VSFUsers.TextMatrix(Row, 2)
    txtPwd = VSFUsers.TextMatrix(Row, 3)
    txtRPwd = VSFUsers.TextMatrix(Row, 3)
    txtuid = VSFUsers.TextMatrix(Row, 4)
    
    CboAccessType.ListIndex = VSFUsers.TextMatrix(Row, 5)
    
    txtUName.SetFocus
    txtUName.SelStart = 0
    txtUName.SelLength = Len(txtUName)
    
    UID = Trim(VSFUsers.TextMatrix(Row, 4))
    
    If (LoginUser = (txtUName.Text) Or AccessType < "2") Then
        cmdDelete.Enabled = False
    Else
        cmdDelete.Enabled = True
    End If
End Sub



