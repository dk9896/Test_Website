VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{97C0E9D8-AD04-4920-9B7A-4B99616579F9}#2.0#0"; "TextPrinter.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmmanual 
   Caption         =   "Form1"
   ClientHeight    =   9210
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14400
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
   ScaleHeight     =   9210
   ScaleWidth      =   14400
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   18720
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame13 
      Caption         =   "Frame13"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   9240
      TabIndex        =   2
      Top             =   7920
      Visible         =   0   'False
      Width           =   1935
      Begin VB.Timer Timer4 
         Left            =   120
         Top             =   1440
      End
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Left            =   1320
         Top             =   960
      End
      Begin VB.Timer Timer3 
         Left            =   840
         Top             =   960
      End
      Begin VB.Timer Timer1 
         Left            =   120
         Top             =   960
      End
      Begin VB.Timer Timer2 
         Left            =   480
         Top             =   960
      End
      Begin MSCommLib.MSComm MSComm1 
         Left            =   120
         Top             =   240
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin TextPrinter.JustPrinter JustPrinter1 
         Height          =   495
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
      End
   End
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
      Height          =   9735
      Left            =   360
      ScaleHeight     =   9675
      ScaleWidth      =   17835
      TabIndex        =   0
      Top             =   120
      Width           =   17895
      Begin VB.Frame Frame1 
         Height          =   1215
         Left            =   14280
         TabIndex        =   27
         Top             =   3720
         Width           =   3375
         Begin VB.TextBox txtChargerSupply 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   1920
            TabIndex        =   30
            Text            =   "00.00"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdChargerSupply 
            Caption         =   "Send"
            Height          =   375
            Left            =   1920
            TabIndex        =   29
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Charger Supply"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Online Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   14280
         TabIndex        =   18
         Top             =   1440
         Width           =   3375
         Begin VB.TextBox txtOD4 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   22
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   1800
            Width           =   1110
         End
         Begin VB.TextBox txtOD3 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   1320
            Width           =   1110
         End
         Begin VB.TextBox txtOD1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   360
            Width           =   1125
         End
         Begin VB.TextBox txtOd2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   1920
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   840
            Width           =   1110
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Output Current"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Output Voltage"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Input Current"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Input Voltage"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Frame10"
         Height          =   1455
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   5775
         Begin VB.Frame Frame5 
            Height          =   975
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   5415
            Begin VB.TextBox Text1 
               Height          =   375
               Left            =   3120
               Locked          =   -1  'True
               TabIndex        =   13
               Top             =   480
               Width           =   2175
            End
            Begin VB.TextBox txtIP_Host 
               Height          =   375
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   12
               Text            =   "127.0.0.1"
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox txtPort 
               Height          =   375
               Left            =   2400
               Locked          =   -1  'True
               TabIndex        =   11
               Text            =   "1232"
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox txtIP 
               Height          =   375
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   10
               Text            =   "127.0.0.1"
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label Label2 
               Caption         =   "IP Host"
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
               Left            =   1440
               TabIndex        =   16
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label16 
               Caption         =   "PORT:"
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
               Left            =   2520
               TabIndex        =   15
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label1 
               Caption         =   "IP M/C"
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
               Index           =   4
               Left            =   240
               TabIndex        =   14
               Top             =   240
               Width           =   495
            End
         End
      End
      Begin VB.TextBox txtCommandLine 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   120
         TabIndex        =   5
         Text            =   "CommandLine"
         Top             =   8880
         Width           =   17535
      End
      Begin VB.CommandButton CmdClose 
         Caption         =   "Close"
         Height          =   1005
         Left            =   15960
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmmanual.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   7440
         UseMaskColor    =   -1  'True
         Width           =   1635
      End
      Begin VSFlex7Ctl.VSFlexGrid VSFInput 
         Height          =   7695
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   6495
         _cx             =   11456
         _cy             =   13573
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   16744576
         ForeColorFixed  =   16777215
         BackColorSel    =   16744576
         ForeColorSel    =   -2147483637
         BackColorBkg    =   16777215
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
         AllowUserResizing=   0
         SelectionMode   =   2
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   500
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   400
         RowHeightMax    =   0
         ColWidthMin     =   400
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmmanual.frx":0C42
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
         FillStyle       =   1
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   0   'False
         WordWrap        =   -1  'True
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
      Begin VSFlex7Ctl.VSFlexGrid VSFOutput 
         Height          =   7695
         Left            =   6720
         TabIndex        =   7
         Top             =   1080
         Width           =   7455
         _cx             =   13150
         _cy             =   13573
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   16744576
         ForeColorFixed  =   16777215
         BackColorSel    =   16744576
         ForeColorSel    =   -2147483637
         BackColorBkg    =   16777215
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
         AllowUserResizing=   0
         SelectionMode   =   2
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   500
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   400
         RowHeightMax    =   0
         ColWidthMin     =   400
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmmanual.frx":0D61
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
         FillStyle       =   1
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   0   'False
         WordWrap        =   -1  'True
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
      Begin VB.Label Label17 
         Caption         =   "PLC Comm"
         Height          =   255
         Left            =   14280
         TabIndex        =   17
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Shape ShapePLCState 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   16680
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Manual Screen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   17655
      End
   End
End
Attribute VB_Name = "frmmanual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MsgCode As Integer
Dim SetVacuum As Integer

'----------------
Dim PLC_Communication_Error As Boolean
Dim MsgText() As String
Dim MsgColor() As Integer
Dim MsgCount As Integer
Private Sub Load_IO_File()
On Error Resume Next
Dim iFile As Integer
Dim s As String
Dim sTextLines() As String
Dim strArray() As String
Dim WorkFile As String
VSFInput.Cols = 3
VSFInput.Rows = 33
VSFInput.RowHeightMin = 600
VSFOutput.RowHeightMin = 600
VSFOutput.Cols = 4
VSFOutput.Rows = 33
VSFInput.TextMatrix(0, 0) = "No."
VSFInput.TextMatrix(0, 1) = "Input Desc"
VSFInput.TextMatrix(0, 2) = "Status"
VSFOutput.TextMatrix(0, 0) = "No."
VSFOutput.TextMatrix(0, 1) = "Output Desc"
VSFOutput.TextMatrix(0, 2) = "Get Val"
VSFOutput.TextMatrix(0, 3) = "Set Val"
VSFInput.ColWidth(0) = 600
VSFInput.ColWidth(1) = 4500
VSFInput.ColWidth(2) = 1000
VSFOutput.ColWidth(0) = 600
VSFOutput.ColWidth(1) = 4500
VSFOutput.ColWidth(2) = 800
VSFOutput.ColWidth(3) = 1000
'VSFOutput.Col
    WorkFile = App.Path & "\IOList.csv"
    'Read the entire file
   iFile = FreeFile
   Open WorkFile For Input As #iFile
        s = Input(LOF(iFile), iFile)
   Close iFile
   'Split the results into separate lines
   sTextLines = Split(s, vbCrLf)

    mscount = UBound(sTextLines)
    For i = 0 To mscount
     If mscount < VSFInput.Rows Then
        strArray = Split(sTextLines(i), ",")
        VSFInput.TextMatrix(i + 1, 0) = "X" & strArray(0)
        VSFOutput.TextMatrix(i + 1, 0) = "Y" & strArray(2)
        VSFInput.TextMatrix(i + 1, 1) = strArray(1)
        VSFOutput.TextMatrix(i + 1, 1) = strArray(3)
        VSFOutput.Cell(flexcpChecked, i + 1, 3) = flexUnchecked
        VSFOutput.Cell(flexcpAlignment, i + 1, 3) = flexAlignCenterCenter
        
     End If
    Next

ErrorHandler:
Close iFile
End Sub
Private Sub ChkY003_Click()

If ChkY003.Value = 1 Then
    Timer5.Interval = 500
    Timer5.Enabled = True
End If

End Sub

Private Sub ChkY004_Click()
If ChkY004.Value = 1 Then
    Timer5.Interval = 500
    Timer5.Enabled = True
End If
End Sub

Private Sub cmdChargerSupply_Click()
txtChargerSupply.Text = Format(txtChargerSupply.Text, "00.00")
If Val(txtChargerSupply.Text) <= 20 Then
PLcdata(245) = Val(txtChargerSupply.Text) * 100
Else
MsgBox "Please Enter Value Between 0 to 20"
End If

End Sub

Private Sub CmdClose_Click()
If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
frmmenu.Show
Unload Me

End Sub

Private Sub Command1_Click()
Dim TempDegree As Integer

'If ValidEntry(0, 999, txtInputVoltage) = False Then Exit Sub

'SetVacuum = Format(Val(txtInputVoltage), "0")

End Sub

Private Function ValidEntry(Min, Max As Double, Text As TextBox) As Boolean

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

Private Sub ConnectToPLC()
On Error GoTo Error
Dim Sql As String
Dim Rs As ADODB.Recordset

    'To Load Com port in Monitor
   Sql = "Select * from Common_Set where SetType ='CommonSet'" 'SetType = Settings Type
   Set Rs = New ADODB.Recordset
   Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   Dim ComPort(3) As Integer
   Dim ComPortBP(3) As Integer
   ComPort(1) = Rs("ComPort1")
''    ComPort(2) = Rs("ComPort2")
    ComPortBP(1) = Rs("ComPortBP1")
''      ComPortBP(2) = Rs("ComPortBP2")
   PrinterName = Rs("PrinterName1")
   Initialise
   Winsock1.Protocol = sckTCPProtocol
   txtIP.Text = Winsock1.LocalIP
   txtIP_Host = Rs("PLC_IP") '"192.168.1.30"
   txtPort = Rs("PLC_Port")
Exit Sub
Error:
If Err.Number = 8002 Then
    MsgBox "Com Port " & ComPort(Erl) & " Not Working", vbInformation
ElseIf Err.Number = 8005 Then
    MsgBox "Com Port " & ComPort(Erl) & " Already Open", vbInformation
Else
    MsgBox Error, vbInformation
End If
End Sub

Private Sub Form_Load()
On Error GoTo Error

'Advance
Me.WindowState = 2
Picture1.BorderStyle = 1
Picture1.Appearance = 0
Picture1.BackColor = Me.BackColor
Picture1.Left = (Screen.Width - Picture1.Width) / 2
Picture1.Top = (Screen.Height - Picture1.Height) / 2 - 400
Load_IO_File
Call Load_Message_File
ConnectToPLC

Timer1.Enabled = True
Timer1.Interval = 1000
Timer2.Enabled = True
Timer2.Interval = 1000
Timer3.Interval = 500
Timer3.Enabled = True

Exit Sub
Error:
MsgBox Error, vbInformation
End Sub

Private Sub AssignPLCdata()
On Error GoTo Error
   txtOD1.Text = Format(PLcdata(102) / 100, "0.00")
   txtOd2.Text = Format(PLcdata(103) / 1000, "0.000")
   txtOD3.Text = Format(PLcdata(104) / 100, "0.00")
   txtOD4.Text = Format(PLcdata(105) / 1000, "0.000")

'txtOd1 = Format(PLcdata(10), "0")
'txtOd2 = Format(PLcdata(11) / 1000, "0.000")
'txtOd3 = Format(PLcdata(12) / 100, "0.00")
'txtOd4 = Format(PLcdata(13) / 100, "0.00")

MsgCode = PLcdata(108)
j = 0
For i = 0 To 2
    KaKoTora PLcdata(190 + i), &H1, VSFInput, 1 + j, 2
    KaKoTora PLcdata(190 + i), &H2, VSFInput, 2 + j, 2
    KaKoTora PLcdata(190 + i), &H4, VSFInput, 3 + j, 2
    KaKoTora PLcdata(190 + i), &H8, VSFInput, 4 + j, 2
    KaKoTora PLcdata(190 + i), &H10, VSFInput, 5 + j, 2
    KaKoTora PLcdata(190 + i), &H20, VSFInput, 6 + j, 2
    KaKoTora PLcdata(190 + i), &H40, VSFInput, 7 + j, 2
    KaKoTora PLcdata(190 + i), &H80, VSFInput, 8 + j, 2
    KaKoTora PLcdata(190 + i), &H100, VSFInput, 9 + j, 2
    KaKoTora PLcdata(190 + i), &H200, VSFInput, 10 + j, 2
    KaKoTora PLcdata(190 + i), &H400, VSFInput, 11 + j, 2
    KaKoTora PLcdata(190 + i), &H800, VSFInput, 12 + j, 2
    KaKoTora PLcdata(190 + i), &H1000, VSFInput, 13 + j, 2
    KaKoTora PLcdata(190 + i), &H2000, VSFInput, 14 + j, 2
    KaKoTora PLcdata(190 + i), &H4000, VSFInput, 15 + j, 2
    j = j + 15
Next
    KaKoTora PLcdata(190 + 3), &H1, VSFInput, 1 + j, 2
    KaKoTora PLcdata(190 + 3), &H2, VSFInput, 2 + j, 2
    KaKoTora PLcdata(190 + 3), &H4, VSFInput, 3 + j, 2
    
j = 0
For i = 0 To 2
    KaKoTora PLcdata(194 + i), &H1, VSFOutput, 1 + j, 2
    KaKoTora PLcdata(194 + i), &H2, VSFOutput, 2 + j, 2
    KaKoTora PLcdata(194 + i), &H4, VSFOutput, 3 + j, 2
    KaKoTora PLcdata(194 + i), &H8, VSFOutput, 4 + j, 2
    KaKoTora PLcdata(194 + i), &H10, VSFOutput, 5 + j, 2
    KaKoTora PLcdata(194 + i), &H20, VSFOutput, 6 + j, 2
    KaKoTora PLcdata(194 + i), &H40, VSFOutput, 7 + j, 2
    KaKoTora PLcdata(194 + i), &H80, VSFOutput, 8 + j, 2
    KaKoTora PLcdata(194 + i), &H100, VSFOutput, 9 + j, 2
    KaKoTora PLcdata(194 + i), &H200, VSFOutput, 10 + j, 2
    KaKoTora PLcdata(194 + i), &H400, VSFOutput, 11 + j, 2
    KaKoTora PLcdata(194 + i), &H800, VSFOutput, 12 + j, 2
    KaKoTora PLcdata(194 + i), &H1000, VSFOutput, 13 + j, 2
    KaKoTora PLcdata(194 + i), &H2000, VSFOutput, 14 + j, 2
    KaKoTora PLcdata(194 + i), &H4000, VSFOutput, 15 + j, 2
    j = j + 15
Next
    KaKoTora PLcdata(194 + 3), &H1, VSFOutput, 1 + j, 2
    KaKoTora PLcdata(194 + 3), &H2, VSFOutput, 2 + j, 2
    KaKoTora PLcdata(194 + 3), &H4, VSFOutput, 3 + j, 2
    
   'txtILLH.Text = PLcdata(150) / 1000
   'txtILRH.Text = PLcdata(151) / 1000
    ''txtOd1.Text = PLcdata(185) / 1000
    ''txtOd2.Text = PLcdata(186) / 1000
    ''txtOd3.Text = PLcdata(187) / 1000
    ''txtOd4.Text = PLcdata(188) / 1000
    'plcdata(186) = odil
   'plcdata(187) = odmvd
   
Exit Sub
Error:
ErrorLog Err.Number, Err.Description, Erl, Me.Name, "AssignPLCData"
Resume Next
End Sub

Private Sub KaKoTora(Data As Integer, reg As Integer, MyCtrl As VSFlexGrid, Row As Integer, Col As Integer)

If ((Data And reg) = reg) Then
    MyCtrl.Cell(flexcpBackColor, Row, Col) = vbGreen
Else
    MyCtrl.Cell(flexcpBackColor, Row, Col) = vbRed
End If

End Sub

Private Sub LoadData()
On Error GoTo Error
PLcdata(240) = 2

If flexgridvalue(VSFOutput, 1) = 1 And flexgridvalue(VSFOutput, 2) = 1 Then
    VSFOutput.Cell(flexcpChecked, 1, 3) = flexUnchecked
    VSFOutput.Cell(flexcpChecked, 2, 3) = flexUnchecked
End If
PLcdata(241) = &H1 * flexgridvalue(VSFOutput, 1)
PLcdata(241) = PLcdata(241) + &H2 * flexgridvalue(VSFOutput, 2)
PLcdata(241) = PLcdata(241) + &H4 * flexgridvalue(VSFOutput, 3)
PLcdata(241) = PLcdata(241) + &H8 * flexgridvalue(VSFOutput, 4)
PLcdata(241) = PLcdata(241) + &H10 * flexgridvalue(VSFOutput, 5)
PLcdata(241) = PLcdata(241) + &H20 * flexgridvalue(VSFOutput, 6)
PLcdata(241) = PLcdata(241) + &H40 * flexgridvalue(VSFOutput, 7)
PLcdata(241) = PLcdata(241) + &H80 * flexgridvalue(VSFOutput, 8)

PLcdata(241) = PLcdata(241) + &H100 * flexgridvalue(VSFOutput, 9)
PLcdata(241) = PLcdata(241) + &H200 * flexgridvalue(VSFOutput, 10)
PLcdata(241) = PLcdata(241) + &H400 * flexgridvalue(VSFOutput, 11)
PLcdata(241) = PLcdata(241) + &H800 * flexgridvalue(VSFOutput, 12)
PLcdata(241) = PLcdata(241) + &H1000 * flexgridvalue(VSFOutput, 13)
PLcdata(241) = PLcdata(241) + &H2000 * flexgridvalue(VSFOutput, 14)
PLcdata(241) = PLcdata(241) + &H4000 * flexgridvalue(VSFOutput, 15)

PLcdata(242) = &H1 * flexgridvalue(VSFOutput, 16)
PLcdata(242) = PLcdata(242) + &H2 * flexgridvalue(VSFOutput, 17)
PLcdata(242) = PLcdata(242) + &H4 * flexgridvalue(VSFOutput, 18)
PLcdata(242) = PLcdata(242) + &H8 * flexgridvalue(VSFOutput, 19)
PLcdata(242) = PLcdata(242) + &H10 * flexgridvalue(VSFOutput, 20)
PLcdata(242) = PLcdata(242) + &H20 * flexgridvalue(VSFOutput, 21)
PLcdata(242) = PLcdata(242) + &H40 * flexgridvalue(VSFOutput, 22)
PLcdata(242) = PLcdata(242) + &H80 * flexgridvalue(VSFOutput, 23)

PLcdata(242) = PLcdata(242) + &H100 * flexgridvalue(VSFOutput, 24)
PLcdata(242) = PLcdata(242) + &H200 * flexgridvalue(VSFOutput, 25)
PLcdata(242) = PLcdata(242) + &H400 * flexgridvalue(VSFOutput, 26)
PLcdata(242) = PLcdata(242) + &H800 * flexgridvalue(VSFOutput, 27)
PLcdata(242) = PLcdata(242) + &H1000 * flexgridvalue(VSFOutput, 28)
PLcdata(242) = PLcdata(242) + &H2000 * flexgridvalue(VSFOutput, 29)
PLcdata(242) = PLcdata(242) + &H4000 * flexgridvalue(VSFOutput, 30)

PLcdata(243) = &H1 * flexgridvalue(VSFOutput, 31)
PLcdata(243) = PLcdata(243) + &H2 * flexgridvalue(VSFOutput, 32)
PLcdata(243) = PLcdata(243) + &H4 * flexgridvalue(VSFOutput, 33)
PLcdata(243) = PLcdata(243) + &H8 * flexgridvalue(VSFOutput, 34)
PLcdata(243) = PLcdata(243) + &H10 * flexgridvalue(VSFOutput, 35)
PLcdata(243) = PLcdata(243) + &H20 * flexgridvalue(VSFOutput, 36)
PLcdata(243) = PLcdata(243) + &H40 * flexgridvalue(VSFOutput, 37)
PLcdata(243) = PLcdata(243) + &H80 * flexgridvalue(VSFOutput, 38)

PLcdata(243) = PLcdata(243) + &H100 * flexgridvalue(VSFOutput, 39)
PLcdata(243) = PLcdata(243) + &H200 * flexgridvalue(VSFOutput, 40)
PLcdata(243) = PLcdata(243) + &H400 * flexgridvalue(VSFOutput, 41)
PLcdata(243) = PLcdata(243) + &H800 * flexgridvalue(VSFOutput, 42)
PLcdata(243) = PLcdata(243) + &H1000 * flexgridvalue(VSFOutput, 43)
PLcdata(243) = PLcdata(243) + &H2000 * flexgridvalue(VSFOutput, 44)
PLcdata(243) = PLcdata(243) + &H4000 * flexgridvalue(VSFOutput, 45)

PLcdata(244) = &H1 * flexgridvalue(VSFOutput, 46)
PLcdata(244) = PLcdata(244) + &H1 * flexgridvalue(VSFOutput, 47)
PLcdata(244) = PLcdata(244) + &H2 * flexgridvalue(VSFOutput, 48)

Exit Sub
Error:
ErrorLog Err.Number, Err.Description, Erl, Me.Name, "LoadData"
Resume Next
End Sub
Private Function flexgridvalue(vsf As VSFlexGrid, Row As Integer) As Integer
   If VSFOutput.Cell(flexcpChecked, Row, 3) = flexChecked Then
     flexgridvalue = 1
   Else
     flexgridvalue = 0
   End If
End Function
Private Sub Timer2_Timer()
'On Error Resume Next


    Static TOGGLE As Boolean
    TOGGLE = Not (TOGGLE)
    Timer2.Interval = 400
    
    With txtCommandLine
        .BorderStyle = 1
        .Alignment = 2
        .FontBold = True
        .FontSize = 16
    End With
        
    Text1.Text = WinsockStstus(Winsock1.State)


    If Winsock1.State = 7 Then
        ShapePLCState.BackColor = vbGreen
    Else
        ShapePLCState.BackColor = vbRed
    End If
    Dim Description As String
    
    Select Case Winsock1.State
        Case 0
            Description = "Connection Closed"
        Case 1
            Description = "Connection Open"
        Case 2
            Description = "Listening For Incomming Connections"
        Case 3
            Description = "Connection Pending"
        Case 4
            Description = "Resolving Remote Host Name"
        Case 5
            Description = "Remote Host Name Successfully Resolved"
        Case 6
            Description = "Connecting-Remote Host"
        Case 7
            Description = "Connected-Remote Host"
            RetryCount = 5
        Case 8
            Description = "Connection is Closing"
        Case 9
            Description = "Connection Error"
        Case Else
            Description = "Connection Status Error"
    End Select

    If PLC_Communication_Error = True Then
       txtCommandLine.ForeColor = vbRed
       txtCommandLine.Text = "communication error"
        Exit Sub
    End If
    
    If TOGGLE = True Then
        If MsgCode >= 1 And MsgCode <= MsgCount Then
            txtCommandLine.Text = MsgText(MsgCode)

            Select Case MsgColor(MsgCode)
                Case 1
                    txtCommandLine.ForeColor = vbBlue
                Case 2
                    txtCommandLine.ForeColor = vbRed
                Case Else
                    txtCommandLine.ForeColor = vbBlack
            End Select
        Else
            txtCommandLine.Text = ""
        End If
    Else
        txtCommandLine.Text = ""
    End If

End Sub

Private Sub Load_Message_File()
On Error Resume Next
Dim iFile As Integer
Dim s As String
Dim sTextLines() As String
Dim strArray() As String
Dim WorkFile As String

    WorkFile = App.Path & "\Messages.csv"

    'Read the entire file
   iFile = FreeFile
   Open WorkFile For Input As #iFile
        s = Input(LOF(iFile), iFile)
   Close iFile
   'Split the results into separate lines
   sTextLines = Split(s, vbCrLf)

    MsgCount = UBound(sTextLines)
    ReDim MsgText(UBound(sTextLines))
    ReDim MsgColor(UBound(sTextLines))

    For i = 0 To MsgCount
        strArray = Split(sTextLines(i), ",")
        MsgText(i) = strArray(1)
        MsgColor(i) = strArray(2)
    Next

ErrorHandler:
Close iFile
End Sub

''Private Sub Timer3_Timer()
''Dim OutBuffer As Variant
''    LoadData
''    SendData
''    OutBuffer = DataOut()
''    MSComm1.Output = OutBuffer
''    Timer3.Interval = 200
''End Sub


Private Sub VSFOutput_CellChanged(ByVal Row As Long, ByVal Col As Long)
 If Col = 2 And Row > 0 Then
   If VSFOutput.Cell(flexcpBackColor, Row, Col) = vbGreen Then
      VSFOutput.Cell(flexcpChecked, Row, 3) = flexChecked
   Else
      VSFOutput.Cell(flexcpChecked, Row, 3) = flexUnchecked
   End If
 End If
End Sub

Private Function cmdCon()
   Winsock1.Close
   Winsock1.RemoteHost = txtIP_Host.Text
   Winsock1.RemotePort = txtPort.Text
   Winsock1.Connect
End Function

Private Function WinsockStstus(ByVal Value As Integer)
Dim Description As String
   Select Case Value
      Case 0
        Description = "Connection Closed"
      Case 1
        Description = "Connection Open"
      Case 2
        Description = "Listening For Incomming Connections"
      Case 3
        Description = "Connection Pending"
      Case 4
        Description = "Resolving Remote Host Name"
      Case 5
        Description = "Remote Host Name Successfully Resolved"
      Case 6
        Description = "Connecting To Remote Host"
      Case 7
        Description = "Connected To Remote Host"
        RetryCount = 0
      Case 8
        Description = "Connection is Closing"
      Case 9
        Description = "Connection Error"
      Case Else
        Description = "Connection Status Error"
   End Select
   WinsockStstus = Description
End Function

Private Sub Timer1_Timer()
   If (Winsock1.State = 7) And (CommandOn = False) Then
      Timer1.Enabled = False
      Select Case CommandType
         Case 1
            Call GetReadArray(StdReadStartAddress, StdReadCount, ReadArray)
            Winsock1.SendData ReadArray
            CVRead = CVRead + 1
            CommandOn = True
            Timer5.Interval = 800
            Timer5.Enabled = True
         Case 2
            Call GetWriteArray(StdWriteStartAddress, StdWriteCount, WriteArray)
            Winsock1.SendData WriteArray
            CommandOn = True
            Timer5.Interval = 800
            Timer5.Enabled = True
         Case 3
            Call GetReadArray((ExtendedReadStartAddress + (ExtendedReadCount * CVExtPktNo)), ExtendedReadCount, ReadArray)
            Winsock1.SendData ReadArray
            CommandOn = True
            Timer5.Interval = 800
            Timer5.Enabled = True
         Case Else
            CommandType = 1
      End Select
      Exit Sub
   Else
      Timer1.Enabled = True
      Timer1.Interval = 100
   End If

   If (Winsock1.State <> 7) Then 'And (WinSock1.State <> 6) Then
      Timer1.Interval = 1000
      Call cmdCon
   Else
      CommandOn = False
      Timer1.Interval = 1000
   End If
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
   LoadData
   Timer3.Interval = 150
End Sub

Private Sub Timer5_Timer()
PLC_Communication_Error = True
CommandOn = False
CommandType = 1
Timer1.Enabled = True
Timer1.Interval = 80
Timer5.Interval = 500
Timer5.Enabled = True
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim SocketData() As Byte
Dim RegData, A, B, C As String
Dim i, j, K, l, M, n, ExpectedArraySize, ExtndedReadFrom, ExpectedLength As Integer
Dim Idata As Long
Dim Idata1 As Long

   Timer5.Enabled = False
   PLC_Communication_Error = False
   Winsock1.GetData SocketData
   CommandOn = False
   PlcCommCheck = False
   Select Case CommandType
      Case 1
         K = StdReadCount * 2
         ExpectedArraySize = K + 10
         If UBound(SocketData) = ExpectedArraySize Then
            If (SocketData(0) = &HD0) And (SocketData(3) = &HFF) And (SocketData(4) = &HFF) And (SocketData(5) = 3) Then
               j = 11
               For i = StdReadStartAddress To (StdReadStartAddress + StdReadCount - 1)
                  M = CInt(SocketData(j + 1))
                  n = CInt(SocketData(j))
                  Idata = (M * 256) + n
                  If Idata > 32767 Then
                     Idata1 = Idata - 65536
                  Else
                     Idata1 = Idata
                  End If
                  PLcdata(i) = CInt(Idata1)
                  j = j + 2
               Next
               If CVRead = 1 Then CommandType = 2
               If ((CVRead >= WriteDelayCount) And ((PLcdata(StdReadStartAddress + StdReadCount - 1) = 0) Or (ExtendedRequired = False))) Then CVRead = 0
               If ((ExtendedRequired = True) And (PLcdata(StdReadStartAddress + StdReadCount - 1) > 0)) Then
                  CommandType = 3
                  CVExtPktNo = 0
               End If
               AssignPLCdata
            Else
               RejCnt = RejCnt + 1
            End If
         Else
            RejCnt = RejCnt + 1
         End If
      Case 2
         If (UBound(SocketData) = 10 And (SocketData(0) = &HD0) And (SocketData(3) = &HFF) And (SocketData(4) = &HFF) And (SocketData(5) = 3)) Then
            CommandType = 1
         Else
            RejCnt = RejCnt + 1
         End If
      Case 3
         K = ExtendedReadCount * 2
         ExpectedArraySize = K + 10
         If UBound(SocketData) = ExpectedArraySize Then
         If (SocketData(0) = &HD0) And (SocketData(3) = &HFF) And (SocketData(4) = &HFF) And (SocketData(5) = 3) Then
            j = 11
            ExtendReadFrom = ExtendedReadStartAddress + (ExtendedReadCount * CVExtPktNo)
            For i = ExtendReadFrom To (ExtendReadFrom + ExtendedReadCount - 1)
               M = CInt(SocketData(j + 1))
               n = CInt(SocketData(j))
               Idata = (M * 256) + n
               If Idata > 32767 Then
                  Idata1 = Idata - 65536
               Else
                  Idata1 = Idata
               End If
               PLcdata(i) = CInt(Idata1)
               j = j + 2
            Next
            CVExtPktNo = CVExtPktNo + 1
            If (CVExtPktNo >= NoOfExtendedPackets) Then
               CVExtPktNo = 0
               If (CVRead = 1) Then
                  CommandType = 2
               Else
                  CommandType = 1
               End If
               If ((CVRead >= WriteDelayCount)) Then CVRead = 0
            End If
         Else
            RejCnt = RejCnt + 1
         End If
      Else
         RejCnt = RejCnt + 1
      End If
   End Select
 
   ' txtModelName = CommandType
   ' txtOd4 = UBound(SocketData)
   text2 = CommandType & "+" & CVExtPktNo
   Timer1.Interval = 10
   Timer1.Enabled = True
End Sub

