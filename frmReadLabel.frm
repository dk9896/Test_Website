VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmReadLabel 
   BorderStyle     =   0  'None
   Caption         =   "Read Bar Code "
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14205
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   14205
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CDialFile 
      Left            =   120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open File"
      FontSize        =   10
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   4695
      Left            =   480
      ScaleHeight     =   4635
      ScaleWidth      =   13155
      TabIndex        =   0
      Top             =   720
      Width           =   13215
      Begin VB.PictureBox Picture2 
         Height          =   2535
         Left            =   720
         ScaleHeight     =   2475
         ScaleWidth      =   11790
         TabIndex        =   8
         Top             =   1800
         Width           =   11850
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   2
            Left            =   10920
            TabIndex        =   20
            Text            =   "Text1"
            Top             =   1920
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   1
            Left            =   6600
            TabIndex        =   19
            Text            =   "Text1"
            Top             =   1320
            Width           =   1935
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   0
            Left            =   2400
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   1320
            Width           =   1935
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Index           =   2
            Left            =   5760
            TabIndex        =   17
            Top             =   120
            Width           =   2415
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   10320
            Picture         =   "frmReadLabel.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   120
            Width           =   1125
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   8280
            TabIndex        =   11
            Top             =   120
            Width           =   1935
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Index           =   0
            Left            =   360
            TabIndex        =   9
            Top             =   120
            Width           =   5295
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Pulses"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   5520
            TabIndex        =   22
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Current"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   21
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Customer Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   6000
            TabIndex        =   16
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Operator Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   8280
            TabIndex        =   12
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "ModelName -Part No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   10
            Top             =   600
            Width           =   3015
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   1575
         Left            =   720
         ScaleHeight     =   1515
         ScaleWidth      =   11790
         TabIndex        =   1
         Top             =   120
         Width           =   11850
         Begin VB.TextBox txtBarcodeRead 
            Alignment       =   2  'Center
            BackColor       =   &H80000016&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   30
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1065
            Left            =   240
            TabIndex        =   6
            ToolTipText     =   "DDMMYY-**** (DATE - NUMBER 4 DIGIT) "
            Top             =   240
            Width           =   4215
         End
         Begin VB.TextBox txtShowDate 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   7680
            TabIndex        =   5
            Top             =   840
            Width           =   2055
         End
         Begin VB.TextBox txtShowNo 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   7680
            TabIndex        =   4
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "&OK"
            DisabledPicture =   "frmReadLabel.frx":1144
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   9960
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   840
            Width           =   1335
         End
         Begin VB.CommandButton cmdNewRead 
            Caption         =   "&New Read?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   9960
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   375
            Index           =   3
            Left            =   6960
            TabIndex        =   14
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "S.No."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   375
            Index           =   2
            Left            =   6960
            TabIndex        =   13
            Top             =   240
            Width           =   735
         End
         Begin VB.Shape Shape5 
            Height          =   1335
            Left            =   120
            Top             =   120
            Width           =   6735
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "CLICK in the Window and then sCaN THE BARCODE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   4440
            TabIndex        =   7
            Top             =   240
            Width           =   2415
         End
      End
   End
End
Attribute VB_Name = "frmReadLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim readok As Boolean
Dim barcode As String
Dim SnoDecode As Integer
Dim DateDecode As String
'Dim Filestr As String
' format ddmmyy:lineno10:opcode:counterxxxx
         
Private Sub cmdClose_Click()
    frmmenu.Show
    Me.Hide
End Sub


'Private Sub cmdLoadFile_Click()
'    On Error Resume Next
'     CDialFile.Filter = "*.mdb|*.mdb"
'     CDialFile.Action = 1
'     CDialFile.CancelError = True
'     CDialFile.DialogTitle = "Load File"
'     If Trim(CDialFile.FileTitle) = "" Then
'        Exit Sub
'     End If
'    On Error GoTo Err_Hndlr1
'
'     Filestr = CDialFile.FileName
'     Searchstr = CDialFile.FileTitle
'     txtFile.Text = Searchstr
'     MsgBox UCase("File loaded"), vbOKOnly
'
'    Exit Sub
'
'Err_Hndlr1:
'    MsgBox UCase("Error while loading file"), vbCritical
'
'
'End Sub

Private Sub cmdNewRead_Click()
    resetall
End Sub

Private Sub cmdOK_Click()
'    Dim Str(10) As Integer
    If (Trim(txtBarcodeRead) <> "") And (Len(Trim(txtBarcodeRead)) = 11 Or Len(Trim(txtBarcodeRead)) = 17) Then
     readok = True
    Else
        readok = False
        txtBarcodeRead.Text = ""
        resetall
        Exit Sub
    End If
    barcode = Trim(txtBarcodeRead.Text)
   If Len(Trim(txtBarcodeRead)) = 11 Then
            DateDecode = Left(barcode, 6)
            SnoDecode = Right(barcode, 4)
    ElseIf Len(Trim(txtBarcodeRead)) = 17 Then
            DateDecode = Left(barcode, 6) 'Mid(barcode, 26, 6) ' Left(barcode, 6)
            SnoDecode = Right(barcode, 4)
    End If
    
   
    displaydata
     
End Sub



Private Sub Form_Load()
  '  frmReadLabel.WindowState = 2
    frmReadLabel.Left = (Screen.Width - frmReadLabel.Width) / 2
    frmReadLabel.Top = (Screen.Height - frmReadLabel.Height) / 2
  
    Me.BackColor = RGB(142, 167, 190)
    cmdOk.Enabled = True
    txtBarcodeRead.BackColor = RGB(142, 167, 190)
   '  txtFile.Text = Searchstr
   
End Sub


Private Sub displaydata()
 On Error GoTo Err_Hndlr
 
   'MakeConn
   Sql = "select * from SpeedSensor_Ok where Sdate =" & Trim(DateDecode) & " and SerialNo =" & Trim(SnoDecode)
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    txtShowNo.Text = Rs("SerialNo")
    txtShowDate.Text = Rs("Rdate")
   
        
    Text3(0).Text = Rs("ModelDesc")
    Text3(1).Text = Rs("OperatorName")
    Text3(2).Text = Rs("CustName")
    
    
    Text1(0).Text = Rs("pulses")
    Text1(1).Text = Rs("current")
    'Text1(2).Text = Rs("fault")
   
    Exit Sub
    
Err_Hndlr:
    MsgBox UCase("Record with this No. does not exist"), vbCritical
   
End Sub

Private Sub resetall()
    Dim i As Integer
    
        txtBarcodeRead.Text = ""
        txtShowNo.Text = ""
        txtShowDate.Text = ""
     
        
    For i = 0 To 2
       Text1(i).Text = ""
       Text1(i).BackColor = frmReadLabel.BackColor
       Text1(i).Enabled = True
        
       Text3(i).Text = ""
       Text3(i).BackColor = frmReadLabel.BackColor
       Text3(i).Enabled = True
           
    Next
        
    cmdOk.Enabled = True

End Sub

