VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3390
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7230
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   0
      ScaleHeight     =   3315
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.Label lblCompanyProduct 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "AKZONOBEL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2595
         TabIndex        =   5
         Top             =   600
         Width           =   1965
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Platform Widows"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4425
         TabIndex        =   4
         Top             =   2220
         Width           =   2550
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version1.02"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5625
         TabIndex        =   3
         Top             =   2580
         Width           =   1350
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         Caption         =   "Copyright Authetic Engineers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   2
         Top             =   2940
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "With Barcode Printer and data traceability feature"
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Top             =   1080
         Width           =   3495
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   480
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   2160
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flagok As Boolean

Private Sub Form_Load()
On Error GoTo Err_Handler


lblVersion.Caption = "Version : " & App.Major & "." & App.Minor & ".0." & App.Revision
lblCompanyProduct.Caption = Replace$(App.Title, "_", " ")

    DirCheck 'Check Dir or Database
    
    DateFormatCheck
    
    flagok = True
    MakeConn
'    LoadConnSetting
'    If SQLbypass = 0 Then
'    SqlConn
'    End If
    'Conn_2007
    
    Make_Column
    
    Timer1.Enabled = True
   
Exit Sub
Err_Handler:
    flagok = False
    MsgBox UCase(Err.Description & vbCrLf & "Contact Vendor!"), vbCritical
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
     If flagok = True Then
        Timer1.Enabled = False
        frmlogin.Show
        
        'Unload Me
     End If
End Sub
Private Sub LoadConnSetting()
Dim Sql As String
Dim Rs As ADODB.Recordset

    Sql = "Select * from Common_Set where SetType ='CommonSet'" 'SetType = Settings Type
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic

        SQLpath = Rs("NetworkDB")
        SQLbypass = Rs("NetworkDBBP")


End Sub
Private Sub DirCheck()
Dim FSO As New FileSystemObject
Dim DirName(5) As String
Dim DirPath As String
    'Auther: Naveen Soni
    'Contact: 8287330444
    DirName(0) = App.Path
    DirName(1) = "Database"
    DirName(2) = "PrnFiles"
    DirName(3) = "Chart"
    DirName(4) = "ExCelMaster"
''    DirName(4) = "Pictures"


    For i = 1 To 4
        DirPath = DirName(0) & "\" & DirName(i)
        If FSO.FolderExists(DirPath) = False Then
           FSO.CreateFolder DirPath
        End If
    Next

    If FSO.FileExists(App.Path & "\Database\" & App.Title & "_DB.mdb") = False Then
       MsgBox "Database Not Found in Directory", vbCritical, "Database Not Found"
       End
    End If

End Sub

Private Sub DateFormatCheck()
Dim strFormat As String

If Day("01-02-03") = 1 And Month("01-02-03") = 2 Then '"DD/MM/yyyy"
    strFormat = "DD/MM/YYYY"
    Exit Sub
ElseIf Day("01-02-03") = 2 And Month("01-02-03") = 1 Then '"MM/DD/yyyy"
    strFormat = "MM/DD/YYYY"
ElseIf Day("01-02-03") = 3 And Month("01-02-03") = 2 Then '"YYYY/MM/DD"
    strFormat = "YYYY/MM/DD"
Else
    strFormat = "Unknown Format"
End If

If MsgBox("Date Format Error :" & vbCrLf & "Current Date Format is : " & strFormat & vbCrLf & _
        "Set Date Format To : DD/MM/YYYY" & vbCrLf & _
        "Do You Want To Set the Date Format?", vbInformation + vbYesNo) = vbYes Then
    Shell "control.exe intl.cpl", vbMaximizedFocus
    End
Else
    End
End If

End Sub
