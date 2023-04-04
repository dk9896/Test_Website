VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form frmSettings2 
   BackColor       =   &H80000010&
   Caption         =   "Setting Test Parameters"
   ClientHeight    =   10335
   ClientLeft      =   120
   ClientTop       =   900
   ClientWidth     =   14610
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10335
   ScaleWidth      =   14610
   Begin VB.PictureBox Picture1 
      Height          =   7935
      Left            =   120
      ScaleHeight     =   7875
      ScaleWidth      =   12795
      TabIndex        =   0
      Top             =   360
      Width           =   12855
      Begin VB.TextBox txtProductionCounter 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   38
         Text            =   "000000"
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "SET"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   37
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "RESET"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   36
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtlinespeed 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   35
         Text            =   "000000"
         Top             =   6120
         Width           =   735
      End
      Begin VB.CommandButton cmdsetlinespeed 
         Caption         =   "SET"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   34
         Top             =   6120
         Width           =   735
      End
      Begin VB.CommandButton cmdresetTarget 
         Caption         =   "RESET"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   31
         Top             =   5160
         Width           =   975
      End
      Begin VB.CommandButton cmdResetNG 
         Caption         =   "RESET"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   29
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdsetNG 
         Caption         =   "SET"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   28
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox txtngcount 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   27
         Text            =   "000000"
         Top             =   4200
         Width           =   855
      End
      Begin VB.CommandButton cmdresetOK 
         Caption         =   "RESET"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   25
         Top             =   3240
         Width           =   975
      End
      Begin VB.CommandButton cmdsetOK 
         Caption         =   "SET"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   24
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox txtokcount 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   23
         Text            =   "000000"
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton cmdresetCoupler 
         Caption         =   "RESET"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   21
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmdsetCoupler 
         Caption         =   "SET"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   20
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtSetCoupler 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   19
         Text            =   "000000"
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton cmdsaveCoupler 
         Caption         =   "SAVE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   18
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtSaveCoupler 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   17
         Text            =   "000000"
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton cmdresetBatch 
         Caption         =   "RESET"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6240
         TabIndex        =   15
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton cmdsetBatch 
         Caption         =   "SET"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         TabIndex        =   14
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtSetbatch 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   13
         Text            =   "000000"
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdsaveBatch 
         Caption         =   "SAVE"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   12
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtsaveBatch 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Text            =   "000000"
         Top             =   1440
         Width           =   855
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   7920
         TabIndex        =   7
         Top             =   0
         Width           =   4815
         Begin VB.TextBox txtModelName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   360
            Left            =   1320
            TabIndex        =   8
            Top             =   240
            Width           =   3345
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model Name"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   7
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   11040
         TabIndex        =   5
         Top             =   6480
         Width           =   1695
         Begin VB.CommandButton CmdClose 
            Caption         =   "&Close"
            Height          =   810
            Left            =   120
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmSettings2.frx":116A
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Close Screen"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Existing Models"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   5415
         Left            =   7920
         TabIndex        =   1
         Top             =   840
         Width           =   4785
         Begin VSFlex7Ctl.VSFlexGrid VSFModel 
            Height          =   4365
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   4515
            _cx             =   7964
            _cy             =   7699
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
            BackColorBkg    =   -2147483638
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
            Rows            =   50
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   400
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmSettings2.frx":1DAC
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
            TabBehavior     =   1
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
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To Edit Model Double Click or Press Enter on Model"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   465
            Left            =   480
            TabIndex        =   4
            Top             =   6720
            Width           =   3705
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Double Click on the Row to get details"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   9
            Left            =   360
            TabIndex        =   3
            Top             =   4920
            Width           =   3915
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Label Label9 
         Caption         =   "Production Counter"
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
         Left            =   240
         TabIndex        =   39
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Qty per hour"
         Height          =   375
         Left            =   3120
         TabIndex        =   33
         Top             =   6240
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Line Speed"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   32
         Top             =   6120
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Target Counter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "NG Counter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "OK Counter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Coupler Counter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Batch Counter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmSettings2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Row As Long
Dim Col As Long

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub DeleteCSV(ByVal FileName As String)
Dim FSO As New FileSystemObject
Dim FilePath As String
    
    FilePath = App.Path & "\ExCelMaster\" & FileName & ".csv"
    
    If FSO.FileExists(FilePath) = True Then
        FSO.DeleteFile FilePath, True
    End If

End Sub

Private Sub WriteCSV(ByVal Grid As VSFlexGrid, ByVal FileName As String)
On Error GoTo Error
Dim Row, Col As Long
Dim strData As String
Dim strLine As String
Dim FilePath As String
    
    FilePath = App.Path & "\ExCelMaster\" & FileName & ".csv"
    
    For Row = 0 To Grid.Rows - 1
        strLine = ""
        For Col = 0 To Grid.Cols - 1
            If Col <> 0 Then strLine = strLine & ","
            strLine = strLine & Trim(Grid.TextMatrix(Row, Col))
        Next
        strData = strData & strLine & vbNewLine
    Next
    
    'Print Report Into File
    Open FilePath$ For Output As #1
        Print #1, strData
    Close #1

Exit Sub
Error:
MsgBox Error, vbInformation
End Sub

Private Sub ReadCSV(ByVal Grid As VSFlexGrid, ByVal FileName As String)
On Error Resume Next
Dim iFile As Integer
Dim Row, Col As Long
Dim strData As String
Dim strLine() As String
Dim strArray() As String
Dim FilePath As String

    FilePath = App.Path & "\ExCelMaster\" & FileName & ".csv"

    'Read the entire file
    iFile = FreeFile
    Open FilePath For Input As #iFile
        strData = Input(LOF(iFile), iFile)
    Close iFile
    'Split the results into separate lines
    strLine = Split(strData, vbCrLf)
    
    For Row = 0 To UBound(strLine)
        strArray = Split(strLine(Row), ",")
        For Col = 0 To UBound(strArray)
            Grid.TextMatrix(Row, Col) = strArray(Col)
        Next
    Next

ErrorHandler:
Close iFile
End Sub


Private Sub LoadGrid()

With VSFData1
    .Rows = 11
    .Cols = 7
    .RowHeight(0) = 500
    .ColWidth(0) = 400
    .Editable = flexEDKbdMouse
    .ExtendLastCol = True

    .TextMatrix(0, 0) = "Sn"
    .TextMatrix(0, 1) = "Pressure"
    .TextMatrix(0, 2) = "Flow Min"
    .TextMatrix(0, 3) = "Flow Max"
    .TextMatrix(0, 4) = "Vacuum"
    .TextMatrix(0, 5) = "Flow Min"
    .TextMatrix(0, 6) = "Flow Max"
    
    For Row = 1 To .Rows - 1
        .TextMatrix(Row, 0) = Row
    Next
    
    For Col = 1 To .Cols - 1
        .ColWidth(Col) = 900
        .ColAlignment(Col) = flexAlignCenterCenter
    Next
    
End With


End Sub

Private Sub Command11_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command10_Click()

End Sub

Private Sub cmdresetBatch_Click()
If MsgBox("Do you want to reset batch counter", vbYesNo) = vbYes Then
 SaveSetting App.Title, ModelName, "BatchCounter", 0
 txtSetbatch.Text = 0
End If
End Sub

Private Sub cmdresetCoupler_Click()
If MsgBox("Do you want to reset coupler counter", vbYesNo) = vbYes Then
 SaveSetting App.Title, ModelName, "CouplerCounter", 0
 txtSetCoupler.Text = 0
End If
End Sub

Private Sub cmdResetNG_Click()
If MsgBox("Do you want to reset NG counter", vbYesNo) = vbYes Then
  SaveSetting App.Title, ModelName, "NGCounter", 0
  txtngcount.Text = 0
End If
End Sub

Private Sub cmdresetOK_Click()
If MsgBox("Do you want to reset OK counter", vbYesNo) = vbYes Then
  SaveSetting App.Title, ModelName, "OkCounter", 0
  txtngcount.Text = 0
End If
End Sub

Private Sub cmdresetTarget_Click()
If MsgBox("Do you want to reset Target counter", vbYesNo) = vbYes Then
    SaveSetting App.Title, ModelName, "TargetProduction", 0
End If
End Sub

Private Sub cmdsaveBatch_Click()
Dim Rs As ADODB.Recordset
Dim Sql As String
If MsgBox("Do you want to save batch counter", vbYesNo) = vbYes Then
    Sql = "Select * from Model_Set where ModelName ='" & Trim(txtModelName.Text) & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    Rs("batchcounter") = Val(txtsaveBatch.Text)
    Rs.Update
    'txtSaveCoupler.Text = Rs("CouplerCounter")
End If
End Sub

Private Sub cmdsaveCoupler_Click()
Dim Rs As ADODB.Recordset
Dim Sql As String
If MsgBox("Do you want to save Coupler Counter", vbYesNo) = vbYes Then
    Sql = "Select * from Model_Set where ModelName ='" & Trim(txtModelName.Text) & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    Rs("CouplerCounter") = Val(txtSaveCoupler.Text)
    Rs.Update
End If
End Sub

Private Sub cmdsetBatch_Click()
If MsgBox("Do you want to set batch counter", vbYesNo) = vbYes Then
  SaveSetting App.Title, ModelName, "BatchCounter", Val(txtSetbatch.Text)
End If
End Sub

Private Sub cmdsetCoupler_Click()
If MsgBox("Do you want to set Coupler counter", vbYesNo) = vbYes Then
 SaveSetting App.Title, ModelName, "CouplerCounter", Val(txtSetCoupler.Text)
End If
End Sub

Private Sub cmdsetlinespeed_Click()
If MsgBox("Do you want to set linespeed", vbYesNo) = vbYes Then
ExtraSetting Save
End If
End Sub

Private Sub cmdsetNG_Click()
If MsgBox("Do you want to set NG counter", vbYesNo) = vbYes Then
 SaveSetting App.Title, ModelName, "NGCounter", Val(txtngcount.Text)
End If
End Sub

Private Sub cmdsetOK_Click()
If MsgBox("Do you want to set OK counter", vbYesNo) = vbYes Then
SaveSetting App.Title, ModelName, "OkCounter", Val(txtokcount.Text)
End If
End Sub

Private Sub Command1_Click()
Dim Rs As ADODB.Recordset
Dim Sql As String
If MsgBox("Do you want to ReSet  Production counter to 0 ?", vbYesNo) = vbYes Then
    Sql = "Select * from Model_Set where ModelName ='" & Trim(txtModelName.Text) & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    Rs("productioncounter") = 0
    Rs.Update
End If
End Sub

'''Private Sub Command4_Click()
''''Dim X, Y As Integer
'''
'''VSFVolt.Rows = ((Val(txtVacFillTime) / Val(txtVacHoldTime))) + 2 '(((Val(txtTestTravel)) * 2) + 1) + 1
'''
'''For i = 1 To VSFVolt.Rows - 1
'''    'VSFVolt.Rows = VSFVolt.Rows + 1
''''    X = ((i * 2) - 1): Y = (i * 2)
'''    VSFVolt.TextMatrix(i, 0) = Format((i - 1) * Val(txtVacHoldTime), "0") 'Format((i - 1) / 2, "0.0") 'i - 1
''''    VSFVolt.TextMatrix(i, 1) = 0 'Format(((X / 100) * 2.45) - 0.2, "0.000")
''''    VSFVolt.TextMatrix(i, 2) = 5 'Format(((Y / 100) * 2.47) + 0.2, "0.000")
''''    VSFVolt.TextMatrix(i, 3) = 0 'Format(((X / 100) * 1.45) - 0.2, "0.000")
''''    VSFVolt.TextMatrix(i, 4) = 5 'Format(((Y / 100) * 1.47) + 0.2, "0.000")
'''Next
'''
'''
'''End Sub

Private Sub VSFModel_DblClick()
Dim Row As Integer

Row = VSFModel.Row
txtModelName = Trim(VSFModel.TextMatrix(Row, 1))

If Row >= 1 Then LoadData
    
End Sub

Private Sub FillModelGrid()
Dim Sql As String
Dim Rs As ADODB.Recordset
Dim Row As Integer
    
    VSFModel.Rows = 1
    
    Sql = "Select * from Model_Set order by ModelName"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    
    Do While Rs.EOF = False
        VSFModel.Rows = VSFModel.Rows + 1
        Row = VSFModel.Rows - 1
        VSFModel.TextMatrix(Row, 0) = Trim(Row)
        VSFModel.TextMatrix(Row, 1) = Trim(Rs("ModelName"))
        Rs.MoveNext
    Loop
    
End Sub

Private Sub cmdAddRow_Click()

    VSFModel.Rows = VSFModel.Rows + 1
    VSFModel.Select VSFModel.Rows - 1, 1
    VSFModel.TopRow = VSFModel.Rows - 1
    VSFModel.Cell(flexcpBackColor, VSFModel.Rows - 1, 1, VSFModel.Rows - 1, VSFModel.Cols - 1) = RGB(220, 220, 220)
    VSFModel.LeftCol = 0
    VSFModel.SetFocus
    VSFModel.TextMatrix(VSFModel.Rows - 1, 0) = Trim(VSFModel.Rows - 1)
    VSFModel.TextMatrix(VSFModel.Rows - 1, 1) = "Fill The Required Fields"
    ResetForm
    
End Sub

Private Sub cmdDeleteRow_Click()
Dim Sql As String
Dim Rs As ADODB.Recordset
   
    If Trim(txtModelDesc) = "" Then
        MsgBox "No Model Is Selected"
    End If
  
    If MsgBox(UCase("Do You Want To Delete?"), vbYesNo + vbInformation) = vbYes Then
  
        Sql = "Select * from Model_Set where ModelName='" & Trim(txtModelName) & "'"
        Set Rs = New ADODB.Recordset
        Rs.Open Sql, Con, adOpenForwardOnly, adLockOptimistic
        If Rs.EOF = True Then Exit Sub
        Rs.Delete
        Rs.Update
        
        DeleteCSV Trim$(txtModelName) & "-FORCE"
        DeleteCSV Trim$(txtModelName) & "-TRAVEL"
    End If


    ResetForm
    FillModelGrid

End Sub

Private Sub cmdReset_Click()
    If MsgBox(UCase("Reset the form?"), vbYesNo) = vbYes Then
       FillModelGrid
       ResetForm
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmmenu.Show
End Sub

Private Sub CmdSave_Click()
On Error GoTo Error
Dim Sql As String
Dim Rs As ADODB.Recordset

    If CheckValidEntry = False Then Exit Sub
    
    Sql = "Select * from Model_Set where ModelName = '" & Trim(txtModelName.Text) & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic

    If Rs.EOF = True Then
        MsgBox "Creating New Record", vbyesOnly
        Rs.AddNew
    ElseIf Rs.EOF = False Then
         MsgBox "Record with this Model Name Exist, Updating the record", vbyesOnly
    End If
        
    Rs("ModelName") = Trim(txtModelName.Text)
    Rs("ModelDesc") = Trim(txtModelDesc.Text)
    Rs.Update
        
    MsgBox UCase("Saved Successfully")
    
    FillModelGrid
    ResetForm
    
Exit Sub
Error:
'MsgBox Error, vbInformation
ErrorLog Err.Number, Err.Description, Erl, Me.Name, "Save Model Setting"
Resume Next
End Sub

Private Sub Form_Load()
On Error GoTo Error

'Settings
Me.WindowState = 2
''Me.BackColor = &H80000010
''Picture1.BorderStyle = 1
''Picture1.Appearance = 0
''Picture1.BackColor = vbButtonFace
''Picture1.Left = (Screen.Width - Picture1.Width) / 2
''Picture1.Top = (Screen.Height - Picture1.Height) / 2 - 400

txtModelName.Locked = True
ExtraSetting Load

FillModelGrid
'LoadGrid

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
        txtlinespeed.Text = Rs("cycletime")
    ElseIf Action = 2 Then
        Rs("cycletime") = txtlinespeed.Text
        Rs.Update
    End If

Exit Sub
Error:
MsgBox "Error in ComPort Setting" & vbNewLine & Error, vbInformation
End Sub

Private Sub LoadData()
On Error GoTo Error
Dim Rs As ADODB.Recordset
Dim Sql As String
    
    Sql = "Select * from Model_Set where ModelName ='" & Trim(txtModelName.Text) & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    
    txtModelName.Text = Trim(Rs("ModelName"))
    'txtModelDesc.Text = Trim(Rs("ModelDesc"))
    txtsaveBatch.Text = Rs("batchcounter")
    txtSaveCoupler.Text = Rs("CouplerCounter")
    txtProductionCounter.Text = Rs("ProductionCounter")
    txtokcount.Text = Val(GetSetting(App.Title, ModelName, "OkCounter", 0))
    txtngcount.Text = Val(GetSetting(App.Title, ModelName, "NgCounter", 0))
    txtSetCoupler.Text = Val(GetSetting(App.Title, ModelName, "CouplerCounter", 0))
    txtSetbatch.Text = Val(GetSetting(App.Title, ModelName, "BatchCounter", 0))
    cmdenable
Exit Sub
Error:
ErrorLog Err.Number, Err.Description, Erl, Me.Name, "LoadData"
Resume Next
End Sub
Private Sub cmdenable()
Command1.Enabled = True
Command2.Enabled = True
cmdsaveBatch.Enabled = True
cmdsaveCoupler.Enabled = True
cmdsetBatch.Enabled = True
cmdsetCoupler.Enabled = True
cmdresetBatch.Enabled = True
cmdresetCoupler.Enabled = True
cmdsetOK.Enabled = True
cmdresetOK.Enabled = True
cmdsetNG.Enabled = True
cmdResetNG.Enabled = True
cmdresetTarget.Enabled = True

End Sub

Private Function CheckValidEntry() As Boolean
    
    If ValidLen(3, 30, txtModelName) = False Then Exit Function
For Each txt In Me
    If TypeOf txt Is TextBox Then
    
        txt.Text = ""
    End If

    If TypeOf txt Is CheckBox Then
        txt.Value = 0
    End If

    If TypeOf txt Is ComboBox Then
        txt.ListIndex = 0
    End If
Next

'    If ValidLen(1, 30, txtModelDesc) = False Then Exit Function


   
CheckValidEntry = True
End Function

Private Function ValidEntryGrd(Grid As VSFlexGrid, Row, Col As Integer, Min, Max As String) As Boolean

    If IsNumeric(Grid.TextMatrix(Row, Col)) = False Or _
        Val(Grid.TextMatrix(Row, Col)) < Val(Min) Or _
        Val(Grid.TextMatrix(Row, Col)) > Val(Max) Then
        MsgBox ("Kindly Enter Between " & Min & " To " & Max), vbCritical
        Grid.Select Row, Col
        Grid.EditCell
        Grid.Cell(flexcpBackColor, Row, Col, Row, Col) = vbRed
        ValidEntryGrd = False
    Else
        Grid.Cell(flexcpBackColor, Row, Col, Row, Col) = vbWhite
        ValidEntryGrd = True
    End If

End Function

Private Function ValidEntry(Min, Max As Double, Text As TextBox) As Boolean

    If IsNumeric(Text) = False Or (Val(Text) < Min Or Val(Text) > Max) Then
        MsgBox ("Kindly Enter Between " & Min & " To " & Max), vbInformation
        Text.SetFocus
        Text.BackColor = vbRed
        ValidEntry = False
    Else
        Text.BackColor = vbWhite
        ValidEntry = True
    End If

End Function

Private Function ValidLen(Min, Max As Long, Text As TextBox) As Boolean

    If Trim(Text) = "" Or (Len(Text) < Min Or Len(Text) > Max) Then
        MsgBox ("Kindly Enter Between " & Min & " To " & Max & " Characters"), vbCritical
        Text.SetFocus
        Text.BackColor = vbRed
        ValidLen = False
    Else
        Text.BackColor = vbWhite
        ValidLen = True
    End If

End Function

Private Sub ResetForm()
Dim txt As Control

For Each txt In Me
    If TypeOf txt Is TextBox Then
        txt.Text = ""
    End If

    If TypeOf txt Is CheckBox Then
        txt.Value = 0
    End If

    If TypeOf txt Is ComboBox Then
        txt.ListIndex = 0
    End If
Next

End Sub

