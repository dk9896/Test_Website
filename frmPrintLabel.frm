VERSION 5.00
Object = "{97C0E9D8-AD04-4920-9B7A-4B99616579F9}#2.0#0"; "TextPrinter.ocx"
Begin VB.Form frmPrintLabel 
   Caption         =   "Label Printing"
   ClientHeight    =   8955
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   12150
   StartUpPosition =   3  'Windows Default
   Begin TextPrinter.JustPrinter JustPrinter1 
      Height          =   615
      Left            =   3240
      TabIndex        =   9
      Top             =   600
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin VB.PictureBox Picture1 
      Height          =   6255
      Left            =   2400
      ScaleHeight     =   6195
      ScaleWidth      =   10995
      TabIndex        =   0
      Top             =   2040
      Width           =   11055
      Begin VB.CommandButton cmdprint 
         Caption         =   "&Print"
         Height          =   2055
         Left            =   7080
         Picture         =   "frmPrintLabel.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   840
         Width           =   2055
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFC0C0&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "h:mm:ss AMPM"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   5415
         Left            =   1080
         Picture         =   "frmPrintLabel.frx":BE86
         ScaleHeight     =   5355
         ScaleWidth      =   5835
         TabIndex        =   2
         Top             =   360
         Width           =   5895
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tick For 2D Print"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3120
            TabIndex        =   8
            Top             =   960
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox txtCopyNo 
            Alignment       =   2  'Center
            BackColor       =   &H80000001&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   495
            Left            =   3840
            TabIndex        =   4
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox txtDatePr 
            Alignment       =   2  'Center
            BackColor       =   &H80000001&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   495
            Left            =   3840
            TabIndex        =   3
            Top             =   0
            Width           =   2055
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "S.N0."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   6
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3000
            TabIndex        =   5
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   915
         Left            =   7080
         Picture         =   "frmPrintLabel.frx":31050
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3120
         Width           =   2085
      End
   End
End
Attribute VB_Name = "frmPrintLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
Private Sub Check1_Click()

'    If Check1.Value = "1" Then
'        Printtype = "2D"
'    Else
'        Printtype = "1D"
'    End If

End Sub

Private Sub cmdClose_Click()
    CopyLabel = False
    frmmenu.Show
    Unload Me
End Sub

Private Sub CmdPrint_Click()
If ValidEntry(1, 9999999, txtCopyNo) = False Then Exit Sub
    
    CopyLabel = True
    PrintLabel JustPrinter1
End Sub

Private Sub Form_Load()
    frmPrintLabel.WindowState = 2
    Picture1.BackColor = RGB(142, 167, 190)
    txtDatePr.Text = Date
    txtDatePr.Locked = True
    LoadSettingsData

End Sub
Private Sub LoadSettingsData()
On Error GoTo Error
Dim Str() As String
Dim Rs As ADODB.Recordset
Dim Sql As String


    Sql = "Select * from Model_Set where ModelName='" & ModelName & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
        
    txtModelDesc = Rs("ModelDesc")

    Printtype = Rs("Printtype")
    IDNo = Rs("IDNo")
    LastPartNo = Rs("LastPartNo")
    PartNo = Rs("PartNo")
    Darkness = Rs("Darkness")
    Vendorcode = Rs("VendorCode")
    Linecode = Rs("Linecode")
   
    Sql = "Select * from Common_Set where SetType ='CommonSet'" 'SetType = Settings Type
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    PRINTERNAME = Rs("PrinterName1")
    
Exit Sub
Error:
ErrorLog Err.Number, Err.Description, Erl, Me.Name, "LoadSettingsData"
Resume Next
End Sub
Private Function ValidEntry(Min, Max As Double, Text As TextBox) As Boolean

    If Trim(Text) = "" Or (Val(Text) < Min Or Val(Text) > Max) Then
        MsgBox ("Kindly Enter Between " & Min & " To " & Max), vbInformation
        Text.SetFocus
'        Text.BackColor = vbRed
        ValidEntry = False
    Else
'        Text.BackColor = vbWhite
        ValidEntry = True
    End If

End Function
