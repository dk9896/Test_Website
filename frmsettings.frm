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
   ForeColor       =   &H000040C0&
   Icon            =   "frmsettings.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   Begin VB.PictureBox Picture1 
      Height          =   9735
      Left            =   480
      ScaleHeight     =   9675
      ScaleWidth      =   13875
      TabIndex        =   0
      Top             =   240
      Width           =   13935
      Begin VB.Frame Frame4 
         Height          =   1695
         Left            =   120
         TabIndex        =   63
         Top             =   7800
         Width           =   3975
         Begin VB.TextBox txtEfficiencyoffset 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   2640
            TabIndex        =   79
            Text            =   "00"
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox txtVoltageOffset 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   2640
            TabIndex        =   65
            Text            =   "0.000"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtCurrentOffset 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   2640
            TabIndex        =   64
            Text            =   "0.000"
            Top             =   720
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
            Index           =   24
            Left            =   3600
            TabIndex        =   81
            Top             =   1320
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Efficiency Offset"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   22
            Left            =   120
            TabIndex        =   80
            Top             =   1320
            Width           =   1410
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Voltage Offset"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   7
            Left            =   120
            TabIndex        =   69
            Top             =   360
            Width           =   1230
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Offset"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   9
            Left            =   120
            TabIndex        =   68
            Top             =   840
            Width           =   1200
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
            Left            =   3600
            TabIndex        =   67
            Top             =   360
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
            Index           =   13
            Left            =   3600
            TabIndex        =   66
            Top             =   840
            Width           =   120
         End
      End
      Begin VB.Frame FrameDM 
         Caption         =   "With Load"
         ForeColor       =   &H000040C0&
         Height          =   4095
         Left            =   4200
         TabIndex        =   18
         Top             =   1560
         Width           =   3975
         Begin VB.OptionButton txtCurrOption3 
            Alignment       =   1  'Right Justify
            Caption         =   "3Amp"
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   1680
            TabIndex        =   74
            Top             =   1320
            Width           =   975
         End
         Begin VB.OptionButton txtCurrOption2 
            Alignment       =   1  'Right Justify
            Caption         =   "1.5Amp"
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   2880
            TabIndex        =   73
            Top             =   960
            Width           =   975
         End
         Begin VB.OptionButton txtCurrOption1 
            Alignment       =   1  'Right Justify
            Caption         =   "2Amp"
            ForeColor       =   &H000040C0&
            Height          =   255
            Left            =   1680
            TabIndex        =   72
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtEfficiencyMax 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   2880
            TabIndex        =   70
            Text            =   "00"
            Top             =   3480
            Width           =   735
         End
         Begin VB.TextBox txtInputCurrentMin 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   1800
            TabIndex        =   61
            Text            =   "00.00"
            Top             =   2040
            Width           =   735
         End
         Begin VB.TextBox txtInputCurrentMax 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   2880
            TabIndex        =   60
            Text            =   "00.00"
            Top             =   2040
            Width           =   735
         End
         Begin VB.TextBox txtOutputCurMax 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   2880
            TabIndex        =   55
            Text            =   "00.00"
            Top             =   3000
            Width           =   735
         End
         Begin VB.TextBox txtOutputCurMin 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   1800
            TabIndex        =   54
            Text            =   "00.00"
            Top             =   3000
            Width           =   735
         End
         Begin VB.TextBox txtOutputVoltMax 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   2880
            TabIndex        =   53
            Text            =   "00.00"
            Top             =   2520
            Width           =   735
         End
         Begin VB.TextBox txtOutputVoltMin 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   1800
            TabIndex        =   52
            Text            =   "00.00"
            Top             =   2520
            Width           =   735
         End
         Begin VB.TextBox txtEfficiencyMin 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   1800
            TabIndex        =   21
            Text            =   "00"
            Top             =   3480
            Width           =   735
         End
         Begin VB.TextBox txtTestVolt 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   1800
            TabIndex        =   20
            Text            =   "0.000"
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Test Current"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   1
            Left            =   240
            TabIndex        =   71
            Top             =   1080
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Input Current"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   25
            Left            =   240
            TabIndex        =   62
            Top             =   2040
            Width           =   1110
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
            TabIndex        =   57
            Top             =   2640
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
            TabIndex        =   56
            Top             =   2160
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   6
            Left            =   3000
            TabIndex        =   51
            Top             =   1680
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   5
            Left            =   1920
            TabIndex        =   50
            Top             =   1680
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
            TabIndex        =   49
            Top             =   3000
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
            TabIndex        =   48
            Top             =   2520
            Width           =   1305
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
            Left            =   2760
            TabIndex        =   30
            Top             =   480
            Width           =   120
         End
         Begin VB.Label lblvoltageoffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Efficiency"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   22
            Top             =   3480
            Width           =   1230
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
      Begin VB.Frame Frame2 
         Caption         =   "Without Load"
         ForeColor       =   &H000040C0&
         Height          =   6615
         Left            =   120
         TabIndex        =   58
         Top             =   120
         Width           =   3975
         Begin VB.Frame Frame9 
            BorderStyle     =   0  'None
            Caption         =   "Frame9"
            Height          =   1095
            Left            =   120
            TabIndex        =   109
            Top             =   1320
            Width           =   3735
            Begin VB.CheckBox ChkBypass 
               Height          =   255
               Index           =   9
               Left            =   3120
               TabIndex        =   113
               Top             =   0
               Width           =   255
            End
            Begin VB.TextBox txtCutoffVolt 
               Alignment       =   2  'Center
               Height          =   360
               Left            =   2040
               TabIndex        =   112
               Text            =   "00.00"
               Top             =   0
               Width           =   735
            End
            Begin VB.TextBox txtCutoffVoltMin 
               Alignment       =   2  'Center
               Height          =   360
               Left            =   2040
               TabIndex        =   111
               Text            =   "00.00"
               Top             =   600
               Width           =   735
            End
            Begin VB.TextBox txtCutoffVoltMax 
               Alignment       =   2  'Center
               Height          =   360
               Left            =   2880
               TabIndex        =   110
               Text            =   "00.00"
               Top             =   600
               Width           =   735
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CutOff Voltage"
               ForeColor       =   &H000040C0&
               Height          =   240
               Index           =   27
               Left            =   0
               TabIndex        =   117
               Top             =   0
               Width           =   1320
               WordWrap        =   -1  'True
            End
            Begin VB.Label lblvoltageoffset 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Max"
               ForeColor       =   &H000040C0&
               Height          =   240
               Index           =   3
               Left            =   3000
               TabIndex        =   116
               Top             =   360
               Width           =   495
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CutOff Voltage Limit"
               ForeColor       =   &H000040C0&
               Height          =   240
               Index           =   14
               Left            =   0
               TabIndex        =   115
               Top             =   720
               Width           =   1740
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Min"
               ForeColor       =   &H000040C0&
               Height          =   240
               Index           =   18
               Left            =   2160
               TabIndex        =   114
               Top             =   360
               Width           =   315
            End
         End
         Begin VB.Frame Frame8 
            BorderStyle     =   0  'None
            Caption         =   "Frame8"
            Height          =   1215
            Left            =   120
            TabIndex        =   100
            Top             =   4920
            Width           =   3735
            Begin VB.TextBox txtOutputVolt3 
               Alignment       =   2  'Center
               Height          =   360
               Left            =   2040
               TabIndex        =   104
               Text            =   "00.00"
               Top             =   120
               Width           =   735
            End
            Begin VB.CheckBox ChkBypass 
               Height          =   255
               Index           =   12
               Left            =   3120
               TabIndex        =   103
               Top             =   120
               Width           =   255
            End
            Begin VB.TextBox txtOutputVolt3Min 
               Alignment       =   2  'Center
               Height          =   360
               Left            =   2040
               TabIndex        =   102
               Text            =   "00.00"
               Top             =   720
               Width           =   735
            End
            Begin VB.TextBox txtOutputVolt3Max 
               Alignment       =   2  'Center
               Height          =   360
               Left            =   2880
               TabIndex        =   101
               Text            =   "00.00"
               Top             =   720
               Width           =   735
            End
            Begin VB.Label lblvoltageoffset 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Test Voltage-3"
               ForeColor       =   &H000040C0&
               Height          =   240
               Index           =   2
               Left            =   0
               TabIndex        =   108
               Top             =   240
               Width           =   1350
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Output Voltage-3 Limit"
               ForeColor       =   &H000040C0&
               Height          =   240
               Index           =   17
               Left            =   0
               TabIndex        =   107
               Top             =   840
               Width           =   1950
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Min"
               ForeColor       =   &H000040C0&
               Height          =   240
               Index           =   20
               Left            =   2280
               TabIndex        =   106
               Top             =   480
               Width           =   315
            End
            Begin VB.Label lblvoltageoffset 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Max"
               ForeColor       =   &H000040C0&
               Height          =   240
               Index           =   5
               Left            =   3120
               TabIndex        =   105
               Top             =   480
               Width           =   495
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame Frame7 
            BorderStyle     =   0  'None
            Caption         =   "Frame7"
            Height          =   1335
            Left            =   120
            TabIndex        =   91
            Top             =   3600
            Width           =   3735
            Begin VB.TextBox txtoutputvolt2 
               Alignment       =   2  'Center
               Height          =   360
               Left            =   2040
               TabIndex        =   95
               Text            =   "00.00"
               Top             =   240
               Width           =   735
            End
            Begin VB.CheckBox ChkBypass 
               Height          =   255
               Index           =   11
               Left            =   3120
               TabIndex        =   94
               Top             =   240
               Width           =   255
            End
            Begin VB.TextBox txtOutputVolt2Min 
               Alignment       =   2  'Center
               Height          =   360
               Left            =   2040
               TabIndex        =   93
               Text            =   "00.00"
               Top             =   840
               Width           =   735
            End
            Begin VB.TextBox txtOutputVolt2Max 
               Alignment       =   2  'Center
               Height          =   360
               Left            =   2880
               TabIndex        =   92
               Text            =   "00.00"
               Top             =   840
               Width           =   735
            End
            Begin VB.Label lblvoltageoffset 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Test Voltage-2"
               ForeColor       =   &H000040C0&
               Height          =   240
               Index           =   1
               Left            =   0
               TabIndex        =   99
               Top             =   360
               Width           =   1350
               WordWrap        =   -1  'True
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Output Voltage-2 Limit"
               ForeColor       =   &H000040C0&
               Height          =   240
               Index           =   16
               Left            =   0
               TabIndex        =   98
               Top             =   960
               Width           =   1950
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Min"
               ForeColor       =   &H000040C0&
               Height          =   240
               Index           =   21
               Left            =   2160
               TabIndex        =   97
               Top             =   600
               Width           =   315
            End
            Begin VB.Label lblvoltageoffset 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Max"
               ForeColor       =   &H000040C0&
               Height          =   240
               Index           =   6
               Left            =   3000
               TabIndex        =   96
               Top             =   600
               Width           =   495
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame Frame6 
            BorderStyle     =   0  'None
            Height          =   1095
            Left            =   120
            TabIndex        =   82
            Top             =   2520
            Width           =   3615
            Begin VB.TextBox txtOutputVolt1 
               Alignment       =   2  'Center
               Height          =   360
               Left            =   2040
               TabIndex        =   86
               Text            =   "00.00"
               Top             =   0
               Width           =   735
            End
            Begin VB.CheckBox ChkBypass 
               Height          =   255
               Index           =   10
               Left            =   3120
               TabIndex        =   85
               Top             =   0
               Width           =   255
            End
            Begin VB.TextBox txtOutputVolt1Min 
               Alignment       =   2  'Center
               Height          =   360
               Left            =   2040
               TabIndex        =   84
               Text            =   "00.00"
               Top             =   600
               Width           =   735
            End
            Begin VB.TextBox txtOutputVolt1Max 
               Alignment       =   2  'Center
               Height          =   360
               Left            =   2880
               TabIndex        =   83
               Text            =   "00.00"
               Top             =   600
               Width           =   735
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Test Voltage-1"
               ForeColor       =   &H000040C0&
               Height          =   240
               Index           =   23
               Left            =   0
               TabIndex        =   90
               Top             =   0
               Width           =   1260
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Output Voltage-1 Limit"
               ForeColor       =   &H000040C0&
               Height          =   240
               Index           =   15
               Left            =   0
               TabIndex        =   89
               Top             =   720
               Width           =   1950
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Min"
               ForeColor       =   &H000040C0&
               Height          =   240
               Index           =   19
               Left            =   2160
               TabIndex        =   88
               Top             =   360
               Width           =   315
            End
            Begin VB.Label lblvoltageoffset 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Max"
               ForeColor       =   &H000040C0&
               Height          =   240
               Index           =   4
               Left            =   3000
               TabIndex        =   87
               Top             =   360
               Width           =   495
               WordWrap        =   -1  'True
            End
         End
         Begin VB.CheckBox ChkBypass 
            Height          =   255
            Index           =   8
            Left            =   3240
            TabIndex        =   78
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox ChkBypass 
            Height          =   255
            Index           =   13
            Left            =   3240
            TabIndex        =   77
            Top             =   6240
            Width           =   255
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Output Voltage Short Test"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   8
            Left            =   120
            TabIndex        =   76
            Top             =   6240
            Width           =   2265
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bypass"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   2
            Left            =   3120
            TabIndex        =   75
            Top             =   240
            Width           =   705
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reverse Polarity"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   28
            Left            =   120
            TabIndex        =   59
            Top             =   720
            Width           =   1545
            WordWrap        =   -1  'True
         End
      End
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
         Height          =   3615
         Left            =   4200
         TabIndex        =   23
         Top             =   5880
         Width           =   3975
         Begin VB.CheckBox ChkBypass 
            Caption         =   "PID Bypass"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   6
            Left            =   120
            TabIndex        =   32
            Top             =   2520
            Width           =   1815
         End
         Begin VB.CheckBox ChkBypass 
            Caption         =   "Bypass - 1"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   7
            Left            =   120
            TabIndex        =   31
            Top             =   2880
            Width           =   2895
         End
         Begin VB.CheckBox ChkBypass 
            Caption         =   "Pressure SW."
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   4
            Left            =   120
            TabIndex        =   29
            Top             =   1800
            Width           =   2295
         End
         Begin VB.CheckBox ChkBypass 
            Caption         =   "Rejection Bin Bypass"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   5
            Left            =   120
            TabIndex        =   28
            Top             =   2160
            Width           =   2295
         End
         Begin VB.CheckBox ChkBypass 
            Caption         =   "Laser Bypass"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   3
            Left            =   120
            TabIndex        =   27
            Top             =   1440
            Width           =   1815
         End
         Begin VB.CheckBox ChkBypass 
            Caption         =   "Printer Bypass"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   2
            Left            =   120
            TabIndex        =   26
            Top             =   1080
            Width           =   2415
         End
         Begin VB.CheckBox ChkBypass 
            Caption         =   "Safety Guard Bypass"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   2415
         End
         Begin VB.CheckBox ChkBypass 
            Caption         =   "With Load Testing Bypass"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2250
         Left            =   8280
         TabIndex        =   11
         Top             =   0
         Width           =   5415
         Begin VB.CommandButton cmdImage 
            Caption         =   "...."
            Height          =   240
            Left            =   4800
            TabIndex        =   45
            Top             =   1800
            Width           =   375
         End
         Begin VB.TextBox txtImagePath 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1440
            TabIndex        =   43
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
            TabIndex        =   44
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
         Height          =   6135
         Left            =   8280
         TabIndex        =   7
         Top             =   2280
         Width           =   5385
         Begin VSFlex7Ctl.VSFlexGrid VSFModel 
            Height          =   5325
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   5115
            _cx             =   9022
            _cy             =   9393
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
            Top             =   5760
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
         Left            =   8280
         TabIndex        =   1
         Top             =   8400
         Width           =   5415
         Begin VB.CommandButton CmdClose 
            Caption         =   "&Close"
            Height          =   810
            Left            =   4080
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
            Left            =   120
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
            Left            =   1440
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
            Left            =   120
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
            Left            =   2640
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
         Height          =   1215
         Left            =   120
         TabIndex        =   33
         Top             =   6600
         Width           =   3975
         Begin VB.TextBox txtScanDelayTime 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   2640
            TabIndex        =   118
            Text            =   "00"
            Top             =   720
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtMarkTime 
            Alignment       =   2  'Center
            Height          =   360
            Left            =   2640
            TabIndex        =   34
            Text            =   "00"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Scan Delay Time"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   26
            Left            =   120
            TabIndex        =   119
            Top             =   720
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dot Mark Time"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   74
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   1275
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Printer Detail"
         ForeColor       =   &H000040C0&
         Height          =   1335
         Left            =   4200
         TabIndex        =   36
         Top             =   120
         Width           =   3975
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
            Left            =   1560
            TabIndex        =   47
            Top             =   1800
            Visible         =   0   'False
            Width           =   2175
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
            Left            =   1560
            TabIndex        =   39
            Top             =   360
            Width           =   2175
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
            Left            =   2280
            TabIndex        =   38
            Top             =   1440
            Visible         =   0   'False
            Width           =   1455
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
            Left            =   1560
            TabIndex        =   37
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vendor ID"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   102
            Left            =   120
            TabIndex        =   46
            Top             =   1800
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CLPL Part No"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   79
            Left            =   120
            TabIndex        =   42
            Top             =   360
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Serial Starting Text"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   77
            Left            =   120
            TabIndex        =   41
            Top             =   1440
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Minda Part NO."
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   75
            Left            =   120
            TabIndex        =   40
            Top             =   840
            Width           =   1350
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
Private Sub cmdClose_Click()
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
Dim rs As ADODB.Recordset
Dim Row As Integer
    
    VSFModel.Rows = 1
    
    Sql = "Select * from Model_Set order by ModelName"
    Set rs = New ADODB.Recordset
    rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    
    Do While rs.EOF = False
        VSFModel.Rows = VSFModel.Rows + 1
        Row = VSFModel.Rows - 1
        VSFModel.TextMatrix(Row, 0) = Trim(Row)
        VSFModel.TextMatrix(Row, 1) = Trim(rs("ModelName"))
        rs.MoveNext
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
Dim rs As ADODB.Recordset
   
    If Trim(txtModelDesc) = "" Then
        MsgBox "No Model Is Selected"
    End If
  
    If MsgBox(UCase("Do You Want To Delete?"), vbYesNo + vbInformation) = vbYes Then
  
        Sql = "Select * from Model_Set where ModelName='" & Trim(txtModelName) & "'"
        Set rs = New ADODB.Recordset
        rs.Open Sql, Con, adOpenForwardOnly, adLockOptimistic
        If rs.EOF = True Then Exit Sub
        rs.Delete
        rs.Update
        
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
Dim rs As ADODB.Recordset
Dim O, P As String
    If CheckValidEntry = False Then Exit Sub
    
    Sql = "Select * from Model_Set where ModelName = '" & Trim(txtModelName.Text) & "'"
    Set rs = New ADODB.Recordset
    rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    If rs.EOF = True Then
        MsgBox "Creating New Record", vbOKOnly
        rs.AddNew
    ElseIf rs.EOF = False Then
         MsgBox "Record with this Model Name Exist, Updating the record", vbOKOnly
    End If
    rs("ModelName") = Trim(txtModelName.Text)
    rs("ModelDesc") = Trim(txtModelDesc.Text)
    
    rs("CutoffVolt") = Format(Val(txtCutoffVolt.Text), "0.00")
    rs("OutputVolt1") = Format(Val(txtOutputVolt1.Text), "0.00")
    rs("OutputVolt2") = Format(Val(txtoutputvolt2.Text), "0.00")
    rs("OutputVolt3") = Format(Val(txtOutputVolt3.Text), "0.00")
    'rs("ScanDelayTime") = Format(Val(txtScanDelayTime.Text), "0.0")
    rs("EfficiencyOffset") = Format(Val(txtEfficiencyoffset.Text), "00")
    rs("CutoffVoltMin") = Format(Val(txtCutoffVoltMin.Text), "0.00")
    rs("OutputVolt1Min") = Format(Val(txtOutputVolt1Min.Text), "0.00")
    rs("OutputVolt2Min") = Format(Val(txtOutputVolt2Min.Text), "0.00")
    rs("OutputVolt3Min") = Format(Val(txtOutputVolt3Min.Text), "0.00")
    rs("CutoffVoltMax") = Format(Val(txtCutoffVoltMax.Text), "0.00")
    rs("OutputVolt1Max") = Format(Val(txtOutputVolt1Max.Text), "0.00")
    rs("OutputVolt2Max") = Format(Val(txtOutputVolt2Max.Text), "0.00")
    rs("OutputVolt3Max") = Format(Val(txtOutputVolt3Max.Text), "0.00")
    
    
    rs("testVoltage") = Format(Val(txtTestVolt.Text), "00.00")
    If txtCurrOption1.Value = True Then
        rs("testCurrent") = 1
    ElseIf txtCurrOption2.Value = True Then
        rs("testCurrent") = 2
    ElseIf txtCurrOption3.Value = True Then
        rs("testCurrent") = 3
    Else
        rs("testCurrent") = 0
    End If
    rs("EfficiencyMin") = Format(Val(txtEfficiencyMin.Text), "00")
    rs("EfficiencyMax") = Format(Val(txtEfficiencyMax.Text), "00")
    rs("InputCurrentMin") = Format(Val(txtInputCurrentMin.Text), "0.000")
    rs("InputCurrentMax") = Format(Val(txtInputCurrentMax.Text), "0.000")
    rs("OutputVoltMin") = Format(Val(txtOutputVoltMin.Text), "00.00")
    rs("OutputVoltMax") = Format(Val(txtOutputVoltMax.Text), "00.00")
    rs("OutputCurrentMin") = Format(Val(txtOutputCurMin.Text), "0.000")
    rs("OutputCurrentMax") = Format(Val(txtOutputCurMax.Text), "0.000")
    rs("VoltageOffset") = Format(Val(txtVoltageOffset.Text), "00.00")
    rs("CurrentOffset") = Format(Val(txtCurrentOffset.Text), "0.000")
    
    
    rs("PrintPartNo") = txtPartNo.Text
    rs("HardwareNo") = txtHardwareVersion.Text
    'rs("SerialStartingtxt") = txtSerialNo.Text
    'rs("VandorId") = txtVandorId.Text
    
    rs("DotMarkingTime") = Format(txtMarkTime.Text, "0.0")
    
    rs("ModelNo") = txtModelNo.Text
    rs("PartImage") = txtImagePath.Text
    rs("PrinterBypass") = Val(ChkBypass(2).Value)
    For i = 0 To 13
     rs("Bypass" & i + 1) = Val(ChkBypass(i).Value)
    Next
    
    rs.Update
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



UserAccess

Exit Sub
Error:
MsgBox Error, vbInformation
End Sub

Private Sub LoadData()
On Error GoTo Error
Dim rs As ADODB.Recordset
Dim Sql As String
    
    Sql = "Select * from Model_Set where ModelName ='" & Trim(txtModelName.Text) & "'"
    Set rs = New ADODB.Recordset
    rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    
    'txtModelName.Text = Trim(Rs("ModelName"))
    txtModelDesc.Text = Trim(rs("ModelDesc"))
    txtCutoffVolt.Text = Format(Val(rs("CutoffVolt")), "0.00")
    txtOutputVolt1.Text = Format(Val(rs("OutputVolt1")), "0.00")
    txtoutputvolt2.Text = Format(Val(rs("OutputVolt2")), "0.00")
    txtOutputVolt3.Text = Format(Val(rs("OutputVolt3")), "0.00")
    
    txtTestVolt.Text = Format(Val(rs("testVoltage")), "00.00")
    If rs("testCurrent") = 1 Then
        txtCurrOption1.Value = True
    ElseIf rs("testCurrent") = 2 Then
        txtCurrOption2.Value = True
    ElseIf rs("testCurrent") = 3 Then
        txtCurrOption3.Value = True
    End If
    'txtScanDelayTime.Text = Format(Val(rs("ScanDelayTime")), "0.0")
    txtEfficiencyMin.Text = Format(Val(rs("EfficiencyMin")), "00")
    txtEfficiencyMax.Text = Format(Val(rs("EfficiencyMax")), "00")
    txtEfficiencyoffset.Text = Format(Val(rs("EfficiencyOffset")), "00")
    txtInputCurrentMin.Text = Format(Val(rs("InputCurrentMin")), "0.000")
    txtInputCurrentMax.Text = Format(Val(rs("InputCurrentMax")), "0.000")
    txtOutputVoltMin.Text = Format(Val(rs("OutputVoltMin")), "00.00")
    txtOutputVoltMax.Text = Format(Val(rs("OutputVoltMax")), "00.00")
    txtOutputCurMin.Text = Format(Val(rs("OutputCurrentMin")), "0.000")
    txtOutputCurMax.Text = Format(Val(rs("OutputCurrentMax")), "0.000")
    txtVoltageOffset.Text = Format(Val(rs("VoltageOffset")), "00.00")
    txtCurrentOffset.Text = Format(Val(rs("CurrentOffset")), "0.000")
    
         txtCutoffVoltMin.Text = Format(Val(rs("CutoffVoltMin")), "0.00")
         txtOutputVolt1Min.Text = Format(Val(rs("OutputVolt1Min")), "0.00")
         txtOutputVolt2Min.Text = Format(Val(rs("OutputVolt2Min")), "0.00")
         txtOutputVolt3Min.Text = Format(Val(rs("OutputVolt3Min")), "0.00")
         txtCutoffVoltMax.Text = Format(Val(rs("CutoffVoltMax")), "0.00")
         txtOutputVolt1Max.Text = Format(Val(rs("OutputVolt1Max")), "0.00")
         txtOutputVolt2Max.Text = Format(Val(rs("OutputVolt2Max")), "0.00")
         txtOutputVolt3Max.Text = Format(Val(rs("OutputVolt3Max")), "0.00")
    
    txtPartNo.Text = rs("PrintPartNo")
    txtHardwareVersion.Text = rs("HardwareNo")
    'txtSerialNo.Text = rs("SerialStartingtxt")
    'txtVandorId.Text = rs("VandorId")
    
    txtMarkTime.Text = Format(rs("DotMarkingTime"), "0.0")
    
    txtModelNo.Text = rs("ModelNo")
    txtImagePath.Text = rs("PartImage")
    For i = 0 To 13
      ChkBypass(i).Value = Val(rs("Bypass" & i + 1))
    Next
    
    If AccessType <> 2 And ChkBypass(9).Value = 1 Then
    Frame9.Visible = False
    End If
    
    If AccessType <> 2 And ChkBypass(10).Value = 1 Then
    Frame6.Visible = False
    End If
    If AccessType <> 2 And ChkBypass(11).Value = 1 Then
    Frame7.Visible = False
    End If
    If AccessType <> 2 And ChkBypass(12).Value = 1 Then
    Frame8.Visible = False
    End If
    Exit Sub
Error:
ErrorLog Err.Number, Err.Description, Erl, Me.Name, "LoadData"
Resume Next
End Sub
Private Sub UserAccess()

If AccessType = "0" Then 'Disable or Hide For Operators
   
   ChkBypass(8).Visible = False
   ChkBypass(9).Visible = False
   ChkBypass(10).Visible = False
   ChkBypass(11).Visible = False
   ChkBypass(12).Visible = False
   ChkBypass(13).Visible = False
   Label1(2).Visible = False
   Label1(8).Visible = False
   Label1(28).Visible = False
   Frame11.Visible = False
   Frame4.Visible = False
    
ElseIf AccessType = "1" Then 'Disable or Hide for AccessType 1

   ChkBypass(8).Visible = False
   ChkBypass(9).Visible = False
   ChkBypass(10).Visible = False
   ChkBypass(11).Visible = False
   ChkBypass(12).Visible = False
   ChkBypass(13).Visible = False
   Label1(2).Visible = False
   Label1(8).Visible = False
   Label1(28).Visible = False
   
   Frame11.Visible = False
   Frame4.Visible = False

ElseIf AccessType = "2" Then 'Show All Which Will Disable or Hide For One

End If

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

