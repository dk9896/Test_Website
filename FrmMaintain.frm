VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form FrmMaintain 
   Caption         =   "Maintenance Screen "
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11370
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
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   6615
      Left            =   120
      ScaleHeight     =   6555
      ScaleWidth      =   11115
      TabIndex        =   0
      Top             =   120
      Width           =   11175
      Begin VB.Frame Frame1 
         Height          =   735
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   3615
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1080
            TabIndex        =   18
            Text            =   "0"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2520
            TabIndex        =   17
            Text            =   "0"
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "OUT"
            Height          =   255
            Left            =   600
            TabIndex        =   20
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "IN"
            Height          =   255
            Index           =   90
            Left            =   2160
            TabIndex        =   19
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   7200
         TabIndex        =   11
         Top             =   3480
         Width           =   3735
         Begin VB.TextBox txtRow 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   28
            Text            =   "0"
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton Command2 
            Caption         =   "DEL"
            Height          =   375
            Left            =   2880
            TabIndex        =   27
            Top             =   1200
            Width           =   735
         End
         Begin VB.CommandButton Command1 
            Caption         =   "CLR"
            Height          =   375
            Left            =   720
            TabIndex        =   26
            Top             =   1200
            Width           =   735
         End
         Begin VB.OptionButton OptPC 
            Caption         =   "PC"
            Height          =   360
            Left            =   2520
            TabIndex        =   25
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton OptPLC 
            Caption         =   "PLC"
            Height          =   375
            Left            =   1560
            TabIndex        =   24
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.CommandButton Command4 
            Caption         =   "ADD"
            Height          =   375
            Left            =   2880
            TabIndex        =   13
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtDataRegister 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1560
            TabIndex        =   12
            Text            =   "0"
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Row"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Data Register"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Data Type"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H0080FF80&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   7200
         ScaleHeight     =   2475
         ScaleWidth      =   3675
         TabIndex        =   4
         Top             =   960
         Width           =   3735
         Begin VB.TextBox txtOperatorName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   7
            Text            =   "Admin"
            Top             =   120
            Width           =   1935
         End
         Begin VB.TextBox txtDate 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   6
            Text            =   "dd/mm/yyyy"
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox txtTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFF00&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   420
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   5
            Text            =   "hh:mm AA"
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Operator Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   9
            Top             =   600
            Width           =   435
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   135
            TabIndex        =   8
            Top             =   1080
            Width           =   480
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   7200
         ScaleHeight     =   1155
         ScaleWidth      =   3675
         TabIndex        =   2
         Top             =   5280
         Width           =   3735
         Begin VB.CommandButton CMDCLOSE 
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Left            =   1200
            MaskColor       =   &H00000000&
            Picture         =   "FrmMaintain.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   1400
         End
      End
      Begin VB.TextBox txtModelName 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Communication Testing"
         Top             =   120
         Width           =   10875
      End
      Begin VSFlex7Ctl.VSFlexGrid VSFData 
         Height          =   4815
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   3615
         _cx             =   6376
         _cy             =   8493
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   51
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   400
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmMaintain.frx":0C42
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
      Begin VSFlex7Ctl.VSFlexGrid VSFWatch 
         Height          =   4815
         Left            =   3840
         TabIndex        =   23
         Top             =   1680
         Width           =   3255
         _cx             =   5741
         _cy             =   8493
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   51
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   400
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmMaintain.frx":0CAE
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "WATCH LIST"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3840
         TabIndex        =   22
         Top             =   960
         Width           =   3255
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   7680
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer5 
      Interval        =   100
      Left            =   7440
      Top             =   6480
   End
End
Attribute VB_Name = "FrmMaintain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Row As Long
Dim Col As Long

Dim MaxRow As Long
Dim PCReg As Long
Dim PLCReg As Long
Private Sub CmdClose_Click()
Unload Me
End Sub
Private Sub Command1_Click()
VSFWatch.Rows = 1
End Sub
Private Sub Command2_Click()
Row = Val(txtRow)
If Row < 1 Then Exit Sub

VSFWatch.RemoveItem Row
txtRow = 0
RowCount
SaveCounterValue
End Sub

Private Sub Command4_Click()
Dim A As Integer
Dim B As Long

If OptPC.Value = True Then A = 0
If OptPLC.Value = True Then A = 1

B = Val(txtDataRegister)

Select Case A
    Case 1
        If B > PLCReg Then Exit Sub
    Case 0
        If B > PCReg Then Exit Sub
End Select

VSFWatch.Rows = VSFWatch.Rows + 1

VSFWatch.TextMatrix(VSFWatch.Rows - 1, 0) = VSFWatch.Rows - 1
VSFWatch.TextMatrix(VSFWatch.Rows - 1, 1) = A
VSFWatch.TextMatrix(VSFWatch.Rows - 1, 2) = B

SaveCounterValue

End Sub

Private Sub Form_Load()
On Error GoTo Error

    LoadGrid
    
    showDataIn
    showDataOut
    Timer5.Enabled = True
    
    txtDate = Format(Date, "dd/mm/yyyy")
    txtTime = Time
    txtOperatorName.Text = UCase(LoginUser)
  
    GetCounterValue
    
Exit Sub
Error:
MsgBox Error, vbInformation
End Sub

Private Sub LoadGrid()
On Error Resume Next
Dim A, B, C As Long

PLCReg = UBound(PLcdata)
PCReg = UBound(Pcdata)
If PCReg > PLCReg Then MaxRow = PCReg Else MaxRow = PLCReg

Text4 = StdReadCount  'PLCReg
Text5 = StdWriteCount '90 'PCReg

MaxRow = MaxRow + 1
VSFData.Rows = MaxRow

With VSFWatch
    .Cols = 5
    .Rows = 1
    .ScrollBars = flexScrollBarVertical
    .Editable = flexEDNone
    .HighLight = flexHighlightNever
    .ExtendLastCol = True
    
    .ColHidden(1) = True
'    .ColHidden(2) = True
    
    .TextMatrix(0, 2) = "Reg"
    .TextMatrix(0, 3) = "Type"
    .TextMatrix(0, 4) = "Data"
    
    
    For Col = 0 To .Cols - 1
        .ColAlignment(Col) = flexAlignCenterCenter
        .FixedAlignment(Col) = flexAlignCenterCenter
        .ColWidth(Col) = 600
    Next
    
'    .ColWidth(0) = 800
    
End With


End Sub

Private Sub Timer5_Timer()
'On Error Resume Next
Timer5.Enabled = False

txtDate = Format(Date, "dd/mm/yyyy")
txtTime = Time
txtOperatorName.Text = UCase(LoginUser)
    
showDataIn
showDataOut
Timer5.Enabled = True
  
End Sub

Private Sub showDataOut()
On Error GoTo Err_Hndlr

For Row = 1 To Val(Text5) 'UBound(PcData)
    VSFData.TextMatrix(Row, 0) = Row - 1
    VSFData.TextMatrix(Row, 1) = PLcdata(200 + (Row - 1))
Next

If VSFWatch.Rows > 1 Then
    For Row = 1 To VSFWatch.Rows - 1
        Call ShowReg(VSFWatch, Row, 1, 2, 3, 4)
    Next
End If

Exit Sub
    
Err_Hndlr:
Resume Next
End Sub

Private Sub ShowReg(ByVal Grid As VSFlexGrid, ByVal Row, Col1, Col2, Col3, Col4 As Long)
Dim A As Long
Dim B As Long
Dim C As Long
Dim D As Long

A = Val(Grid.TextMatrix(Row, Col1))
B = Val(Grid.TextMatrix(Row, Col2))
C = Val(Grid.TextMatrix(Row, Col3))
D = Val(Grid.TextMatrix(Row, Col4))

Select Case A
    Case 0
'        Grid.TextMatrix(Row, Col3) = "PC" '& 200 + B  '& '"=" & PcData(Reg)
'        Grid.TextMatrix(Row, Col4) = PcData(B)
    Case 1
        Grid.TextMatrix(Row, Col3) = "PLC" '& 100 + B  '& '"=" & PcData(Reg)
        Grid.TextMatrix(Row, Col4) = PLcdata(B)
    Case Else
        Grid.TextMatrix(Row, Col3) = "ERR"
        Grid.TextMatrix(Row, Col4) = "ERR"
End Select

End Sub


Private Sub showDataIn()
On Error Resume Next

For Row = 1 To UBound(PLcdata)
    VSFData.TextMatrix(Row, 0) = Row - 1
    VSFData.TextMatrix(Row, 2) = PLcdata(Row - 1)
Next
    
End Sub

Private Sub SaveCounterValue()
Dim SaveRow As Integer

SaveRow = VSFWatch.Rows - 1
SaveSetting App.Title, "Maintain", "WatchRow", SaveRow

If SaveRow > 0 Then
    For Row = 1 To SaveRow
        SaveSetting App.Title, "Maintain", "Type" & Row, VSFWatch.TextMatrix(Row, 1)
        SaveSetting App.Title, "Maintain", "Reg" & Row, VSFWatch.TextMatrix(Row, 2)
    Next
End If

End Sub

Private Sub GetCounterValue()
On Error Resume Next
Dim SaveRow As Integer

SaveRow = Val(GetSetting(App.Title, "Maintain", "WatchRow"))

If SaveRow > 0 Then
    For Row = 1 To SaveRow
        VSFWatch.Rows = VSFWatch.Rows + 1
        VSFWatch.TextMatrix(Row, 1) = Val(GetSetting(App.Title, "Maintain", "Type" & Row))
        VSFWatch.TextMatrix(Row, 2) = Val(GetSetting(App.Title, "Maintain", "Reg" & Row))
    Next
End If

RowCount

End Sub

Private Sub VSFWatch_Click()

Row = VSFWatch.Row

If Row < 1 Then Exit Sub
txtRow = Row

End Sub

Private Sub RowCount()

With VSFWatch
    For Row = 1 To .Rows - 1
        .TextMatrix(Row, 0) = Row
    Next
End With

End Sub
