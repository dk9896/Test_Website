VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmsettings 
   Caption         =   "Setting Test Parameters"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   900
   ClientWidth     =   13260
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmsettings.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   13260
   Begin VB.PictureBox Picture1 
      Height          =   7455
      Left            =   120
      ScaleHeight     =   7395
      ScaleWidth      =   13035
      TabIndex        =   0
      Top             =   120
      Width           =   13095
      Begin MSComDlg.CommonDialog CD1 
         Left            =   19440
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame11 
         Caption         =   "Bypasses"
         ForeColor       =   &H000040C0&
         Height          =   3375
         Left            =   4200
         TabIndex        =   27
         Top             =   2400
         Width           =   3135
         Begin VB.CheckBox chkbypass 
            Caption         =   "Bypass - 3"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   6
            Left            =   120
            TabIndex        =   37
            Top             =   2520
            Width           =   2175
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "Bypass - 4"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   7
            Left            =   120
            TabIndex        =   36
            Top             =   2880
            Width           =   2655
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "Bypass - 1"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   4
            Left            =   120
            TabIndex        =   33
            Top             =   1800
            Width           =   2655
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "Bypass - 2"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   5
            Left            =   120
            TabIndex        =   32
            Top             =   2160
            Width           =   2655
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "Scanner Bypass"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   3
            Left            =   120
            TabIndex        =   31
            Top             =   1440
            Width           =   2175
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "Printer Bypass"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   2
            Left            =   120
            TabIndex        =   30
            Top             =   1080
            Width           =   2775
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "Start Test Bypass"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   29
            Top             =   720
            Width           =   2775
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "Reverse Polarity Bypass"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame FrameDM 
         Height          =   4935
         Left            =   120
         TabIndex        =   18
         Top             =   0
         Width           =   3975
         Begin VB.TextBox txtCurrentOffset 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   2040
            TabIndex        =   67
            Text            =   "0.000"
            Top             =   4320
            Width           =   735
         End
         Begin VB.TextBox txtVoltageOffset 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   2040
            TabIndex        =   65
            Text            =   "0.000"
            Top             =   3840
            Width           =   735
         End
         Begin VB.TextBox txtOutputCurrentMax 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   2880
            TabIndex        =   63
            Text            =   "0.000"
            Top             =   3360
            Width           =   735
         End
         Begin VB.TextBox txtOutputCurrentMin 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   1920
            TabIndex        =   62
            Text            =   "0.000"
            Top             =   3360
            Width           =   735
         End
         Begin VB.TextBox txtOutputVoltageMax 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   2880
            TabIndex        =   61
            Text            =   "0.000"
            Top             =   2880
            Width           =   735
         End
         Begin VB.TextBox txtOutputVoltMin 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   1920
            TabIndex        =   60
            Text            =   "0.000"
            Top             =   2880
            Width           =   735
         End
         Begin VB.TextBox txtEfficiency 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   2160
            TabIndex        =   25
            Text            =   "00"
            Top             =   2040
            Width           =   735
         End
         Begin VB.TextBox txtSTCheckTime 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   2160
            TabIndex        =   24
            Text            =   "00"
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtRPCheckTime 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   2160
            TabIndex        =   22
            Text            =   "00"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox txtTestVoltage 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   2160
            TabIndex        =   21
            Text            =   "0.000"
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   210
            Index           =   13
            Left            =   3000
            TabIndex        =   71
            Top             =   4440
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   210
            Index           =   12
            Left            =   3720
            TabIndex        =   70
            Top             =   3480
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "V"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   210
            Index           =   11
            Left            =   3000
            TabIndex        =   69
            Top             =   3960
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "V"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   210
            Index           =   10
            Left            =   3720
            TabIndex        =   68
            Top             =   3000
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Offset"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   9
            Left            =   240
            TabIndex        =   66
            Top             =   4440
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Voltage Offset"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   7
            Left            =   240
            TabIndex        =   64
            Top             =   3960
            Width           =   1230
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   6
            Left            =   3000
            TabIndex        =   59
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   5
            Left            =   2160
            TabIndex        =   58
            Top             =   2520
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Output Current"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   4
            Left            =   240
            TabIndex        =   57
            Top             =   3480
            Width           =   1275
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Output Voltage"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   3
            Left            =   240
            TabIndex        =   56
            Top             =   3000
            Width           =   1305
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "sec"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   210
            Index           =   2
            Left            =   3120
            TabIndex        =   55
            Top             =   1680
            Width           =   270
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "V"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   210
            Index           =   55
            Left            =   3240
            TabIndex        =   35
            Top             =   480
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "sec"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   210
            Index           =   54
            Left            =   3120
            TabIndex        =   34
            Top             =   960
            Width           =   270
         End
         Begin VB.Label lblvoltageoffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Efficiency"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   26
            Top             =   2160
            Width           =   1230
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Short Test Check Time"
            ForeColor       =   &H000040C0&
            Height          =   480
            Index           =   8
            Left            =   240
            TabIndex        =   23
            Top             =   1440
            Width           =   1200
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reverse Polarity Check Time"
            ForeColor       =   &H000040C0&
            Height          =   480
            Index           =   1
            Left            =   240
            TabIndex        =   20
            Top             =   840
            Width           =   1545
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Test Voltage"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2250
         Left            =   7440
         TabIndex        =   11
         Top             =   0
         Width           =   5415
         Begin VB.CommandButton cmdImage 
            Caption         =   "...."
            Height          =   240
            Left            =   4800
            TabIndex        =   52
            Top             =   1800
            Width           =   375
         End
         Begin VB.TextBox txtImagePath 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1440
            TabIndex        =   50
            Top             =   1680
            Width           =   3225
         End
         Begin VB.TextBox txtModelNo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   3840
            TabIndex        =   16
            Top             =   1200
            Width           =   1305
         End
         Begin VB.TextBox txtModelDesc 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1440
            TabIndex        =   13
            Top             =   720
            Width           =   3705
         End
         Begin VB.TextBox txtModelName 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1440
            TabIndex        =   12
            Top             =   240
            Width           =   3705
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Image Path"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   51
            Top             =   1680
            Width           =   1875
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model No"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   1200
            Width           =   1875
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model Desc"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   8
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   1875
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model Name"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   7
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1335
            WordWrap        =   -1  'True
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
         Height          =   4935
         Left            =   7440
         TabIndex        =   7
         Top             =   2280
         Width           =   5385
         Begin VSFlex7Ctl.VSFlexGrid VSFModel 
            Height          =   4125
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   5115
            _cx             =   9022
            _cy             =   7276
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
            FormatString    =   $"frmsettings.frx":116A
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
            TabIndex        =   10
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
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   9
            Left            =   600
            TabIndex        =   9
            Top             =   4560
            Width           =   3915
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
         Left            =   120
         TabIndex        =   1
         Top             =   6000
         Width           =   7215
         Begin VB.CommandButton CmdClose 
            Caption         =   "&Close"
            Height          =   810
            Left            =   5520
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmsettings.frx":11D9
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Close Screen"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   1215
         End
         Begin VB.CommandButton cmdReset 
            Caption         =   "&Reset"
            Height          =   810
            Left            =   480
            MaskColor       =   &H00404040&
            Picture         =   "frmsettings.frx":1E1B
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Reset All"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   1095
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   810
            Left            =   2160
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmsettings.frx":317D
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   975
         End
         Begin VB.CommandButton cmdAddRow 
            Caption         =   "&Add Row"
            Height          =   810
            Left            =   480
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmsettings.frx":3DBF
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Add new Line"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   1095
         End
         Begin VB.CommandButton cmdDeleteRow 
            Caption         =   "&Delete Row"
            Height          =   810
            Left            =   3720
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmsettings.frx":4A01
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Delete Record"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame Frame13 
         Height          =   855
         Left            =   120
         TabIndex        =   38
         Top             =   4920
         Width           =   3975
         Begin VB.TextBox txtMarkTime 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   2040
            TabIndex        =   39
            Text            =   "00"
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dot Mark Time"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   74
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   1275
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Printer Detail"
         ForeColor       =   &H000040C0&
         Height          =   2175
         Left            =   4200
         TabIndex        =   41
         Top             =   120
         Width           =   3135
         Begin VB.TextBox txtVandorId 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            TabIndex        =   54
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox txtPartNo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   45
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtBarcodeLength 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2400
            TabIndex        =   44
            Text            =   "0"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtSerialNo 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2040
            TabIndex        =   43
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txtHardwareVersion 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            TabIndex        =   42
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor ID"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   102
            Left            =   120
            TabIndex        =   53
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cust Part No"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   79
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Barcode Length"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   78
            Left            =   120
            TabIndex        =   48
            Top             =   720
            Width           =   1365
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Serial Starting Text"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   77
            Left            =   120
            TabIndex        =   47
            Top             =   1080
            Width           =   1665
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Part Revision No"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   75
            Left            =   120
            TabIndex        =   46
            Top             =   1440
            Width           =   1440
         End
      End
   End
End
Attribute VB_Name = "frmsettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Row As Long
Dim Col As Long

Private Sub CboSensorType_Click()

Select Case CboSensorType.ListIndex
    Case 0
        VSFChannel.Cell(flexcpBackColor, 3, 3, 4, 4) = vbWhite
        VSFChannel.Cell(flexcpBackColor, 6, 3, 6, 4) = vbWhite
'        VSFChannel.Cell(flexcpBackColor, 10, 3, 10, 4) = &H404040
    Case 1
        VSFChannel.Cell(flexcpBackColor, 3, 3, 4, 4) = &H404040
        VSFChannel.Cell(flexcpBackColor, 6, 3, 6, 4) = &H404040
'        VSFChannel.Cell(flexcpBackColor, 10, 3, 10, 4) = vbWhite
End Select

End Sub
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

Private Sub cmdImage_Click()
With CD1
    .DialogTitle = "Select File"
    .Filter = "(*.bmp; *.jpg;)"
    .ShowOpen
    txtImagePath.Text = .FileName
End With
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
'''1
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
Dim O, P As String
    If CheckValidEntry = False Then Exit Sub
    
    Sql = "Select * from Model_Set where ModelName = '" & Trim(txtModelName.Text) & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    If Rs.EOF = True Then
        MsgBox "Creating New Record", vbOKOnly
        Rs.AddNew
    ElseIf Rs.EOF = False Then
         MsgBox "Record with this Model Name Exist, Updating the record", vbOKOnly
    End If
    Rs("ModelName") = Trim(txtModelName.Text)
    Rs("ModelDesc") = Trim(txtModelDesc.Text)
    Rs("testVoltage") = Format(Val(txtTestVoltage.Text), "0.00")
    Rs("RPCheckTime") = Format(Val(txtRPCheckTime.Text), "0.00")
    Rs("STCheckTime") = Format(Val(txtSTCheckTime.Text), "0.00")
    Rs("Efficiency") = Format(Val(txtEfficiency.Text), "00")
    Rs("OutputVoltMin") = Format(Val(txtOutputVoltMin.Text), "0.00")
    Rs("OutputVoltMax") = Format(Val(txtOutputVoltageMax.Text), "0.00")
    Rs("OutputCurrentMin") = Format(Val(txtOutputCurrentMin.Text), "0.00")
    Rs("OutputCurrentMax") = Format(Val(txtOutputCurrentMax.Text), "0.00")
    Rs("VoltageOffset") = Format(Val(txtVoltageOffset.Text), "0.00")
    Rs("CurrentOffset") = Format(Val(txtCurrentOffset.Text), "0.00")
    
    
    Rs("PrintPartNo") = txtPartNo.Text
    Rs("PrintBarcodeLength") = txtBarcodeLength.Text
    Rs("HardwareNo") = txtHardwareVersion.Text
    Rs("SerialStartingtxt") = txtSerialNo.Text
    Rs("VandorId") = txtVandorId.Text
    
    Rs("DotMarkingTime") = Format(txtMarkTime.Text, "0.0")
    
    Rs("ModelNo") = txtModelNo.Text
    Rs("PartImage") = txtImagePath.Text
    Rs("PrinterBypass") = Val(chkbypass(2).Value)
    For i = 0 To 7
     Rs("Bypass" & i + 1) = Val(chkbypass(i).Value)
    Next
    'Rs("BatchCounter").Text
    'Rs("CouplerCounter") = .Text
    'Rs("productioncounter") =
    'Rs("CameraBypass") = Val(chkbypass(0).Value)
    'Rs("LSBypass") = Val(chkbypass(1).Value)
    'Rs("WLCBypass") = Val(chkbypass(2).Value)
    'Rs("BSBypass") = Val(chkbypass(3).Value)
    'Rs("ICBypass") = Val(chkbypass(5).Value)
    'Rs("ScannerBypass") = Val(chkbypass(6).Value)
    'Rs("PIDByPass") = Val(chkbypass(7).Value)
    'Rs("PressureGuageByPass") = Val(chkbypass(8).Value)
    'Rs("UpperCoverByPass") = Val(chkbypass(9).Value)
    
    Rs.Update
'    WriteCSV VSFData1, Trim$(txtModelName)
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
Me.BackColor = &H80000010
Picture1.BorderStyle = 1
Picture1.Appearance = 0
Picture1.BackColor = vbButtonFace
Picture1.Left = (Screen.Width - Picture1.Width) / 2
Picture1.Top = (Screen.Height - Picture1.Height) / 2 - 400

FillModelGrid





Exit Sub
Error:
MsgBox Error, vbInformation
End Sub

Private Sub LoadData()
On Error GoTo Error
Dim Rs As ADODB.Recordset
Dim Sql As String
    
    Sql = "Select * from Model_Set where ModelName ='" & Trim(txtModelName.Text) & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    
    'txtModelName.Text = Trim(Rs("ModelName"))
    txtModelDesc.Text = Trim(Rs("ModelDesc"))
    txtTestVoltage.Text = Rs("testVoltage")
    txtRPCheckTime.Text = Rs("RPCheckTime")
    txtSTCheckTime.Text = Rs("STCheckTime")
    txtEfficiency.Text = Rs("Efficiency")
    txtOutputVoltMin.Text = Rs("OutputVoltMin")
    txtOutputVoltageMax.Text = Rs("OutputVoltMax")
    txtOutputCurrentMin.Text = Rs("OutputCurrentMin")
    txtOutputCurrentMax.Text = Rs("OutputCurrentMax")
    txtVoltageOffset.Text = Rs("VoltageOffset")
    txtCurrentOffset.Text = Rs("CurrentOffset")
    
    txtPartNo.Text = Rs("PrintPartNo")
    txtBarcodeLength.Text = Rs("BarcodeLength")
    txtHardwareVersion.Text = Rs("HardwareNo")
    txtSerialNo.Text = Rs("SerialStartingtxt")
    txtVandorId.Text = Rs("VandorId")
    
    txtMarkTime.Text = Val(Rs("DotMarkingTime"))
    
    txtModelNo.Text = Val(Rs("ModelNo"))
    txtImagePath.Text = Rs("PartImage")
    For i = 0 To 7
      chkbypass(i).Value = Val(Rs("Bypass" & i + 1))
    Next

'    For i = 0 To 8
'      txtCurrentOffset(i).Text = Rs("CurrentOffset" & i + 1)
'      txtVoltageOffset(i).Text = Rs("VoltageOffset" & i + 1)
'    Next
'    txtICMin.Text = Rs("ICMin")
'    txtICMax.Text = Rs("ICMax")
'    'txtICMinRH.Text = Rs("ICMinRH")
'    'txtICMaxRH.Text = Rs("ICMaxRH")
'    'txtWirevoltageMin.Text = Rs("WVMin")
'    'txtWireVoltageMax.Text = Rs("WVMax")
'
'    txtPartNo.Text = Rs("PrintPartNo")
'    txtBarcodeLength.Text = Rs("PrintBarcodeLength")
'    txtBarcodeLength.Text = Rs("BarcodeLength")
'    txtHardwareVersion.Text = Rs("HardwareNo")
'    txtSerialNo.Text = Rs("SerialStartingtxt")
'    txtDebounceTime.Text = Rs("DebounceTime")
'    txtHoldTime.Text = Rs("HoldTime")
'    txtCheckTime.Text = Rs("CheckTime")
'    txtMarkTime.Text = Rs("DotMarkingTime")
'    txtModelNo.Text = Rs("ModelNo")
'    'Rs("BatchCounter").Text
'    'Rs("CouplerCounter") = .Text
'    txtImagePath.Text = Rs("PartImage")
'    'Rs("productioncounter") =
'    For i = 0 To 9
'      chkbypass(i).Value = Val(Rs("Bypass" & i + 1))
'    Next
'    'chkbypass(1).Value = Val(Rs("LSBypass"))
'    'chkbypass(2).Value = Val(Rs("WLCBypass"))
'    'chkbypass(3).Value = Val(Rs("BSBypass"))
'    'chkbypass(4).Value = Val(Rs("PrinterBypass"))
'    'chkbypass(5).Value = Val(Rs("ICBypass"))
'    'chkbypass(6).Value = Val(Rs("ScannerBypass"))
'    'chkbypass(7).Value = Val(Rs("PIDByPass"))
'    'chkbypass(8).Value = Val(Rs("PressureGuageByPass"))
'    'chkbypass(9).Value = Val(Rs("UpperCoverByPass"))
'
    Exit Sub
Error:
ErrorLog Err.Number, Err.Description, Erl, Me.Name, "LoadData"
Resume Next
End Sub

Private Function CheckValidEntry() As Boolean
    
    If ValidLen(3, 30, txtModelName) = False Then Exit Function
    If ValidLen(1, 40, txtModelDesc) = False Then Exit Function
    'If ValidLen(4, 4, txtvendorCode) = False Then Exit Function
    'If ValidLen(1, 1, txtlinecode) = False Then Exit Function
    'If ValidLen(11, 11, txtPartNo) = False Then Exit Function
    'If ValidLen(5, 5, txtLastPartno) = False Then Exit Function
    
'    If ValidEntry(0, 320, txtDataMin3) = False Then Exit Function
'    If ValidEntry(0, 320, txtDataMax3) = False Then Exit Function
'
'    If ValidLen(10, 10, txtDataMin4) = False Then Exit Function
'    If ValidLen(8, 8, txtDataMax4) = False Then Exit Function
'
'
'
'    If ValidEntry(0, 180, txtServoFastSpeed) = False Then Exit Function
'    If ValidEntry(0, 90, txtServoFastDegree) = False Then Exit Function
'    If ValidEntry(0, 90, txtServoSlowSpeed) = False Then Exit Function
'    If ValidEntry(0, 320, txtClampingTime) = False Then Exit Function
'
'    If ValidEntry(1, 90, txtTestCycle) = False Then Exit Function
'    If ValidEntry(0, 30000, txtCameraJob) = False Then Exit Function
'    If ValidEntryGrd(VSFChannel, 1, 1, 0, 300) = False Then Exit Function
'    If ValidEntryGrd(VSFChannel, 1, 2, 0, 300) = False Then Exit Function
'    If ValidEntryGrd(VSFChannel, 1, 3, 0, 300) = False Then Exit Function
'    If ValidEntryGrd(VSFChannel, 1, 4, 0, 300) = False Then Exit Function

   
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



'LoadGrid

End Sub

