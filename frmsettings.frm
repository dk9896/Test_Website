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
      Height          =   8775
      Left            =   120
      ScaleHeight     =   8715
      ScaleWidth      =   19755
      TabIndex        =   0
      Top             =   120
      Width           =   19815
      Begin VB.CheckBox ChkSRMBypass 
         Caption         =   "Set/Reset Module Bypass"
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   8280
         TabIndex        =   268
         Top             =   6000
         Width           =   2655
      End
      Begin VB.Frame FrameSRM 
         Height          =   2415
         Left            =   8280
         TabIndex        =   250
         Top             =   6240
         Width           =   2655
         Begin VB.TextBox txtMVDMin 
            Alignment       =   2  'Center
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
            Index           =   14
            Left            =   840
            TabIndex        =   257
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtMVDMax 
            Alignment       =   2  'Center
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
            Index           =   14
            Left            =   1440
            TabIndex        =   256
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtCurrMin 
            Alignment       =   2  'Center
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
            Index           =   14
            Left            =   840
            TabIndex        =   255
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtCurrMax 
            Alignment       =   2  'Center
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
            Index           =   14
            Left            =   1440
            TabIndex        =   254
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtTestCycle 
            Alignment       =   2  'Center
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
            Index           =   8
            Left            =   240
            TabIndex        =   253
            Text            =   "00"
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txtVoltageOffset 
            Alignment       =   2  'Center
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
            Index           =   8
            Left            =   1920
            TabIndex        =   252
            Text            =   "0.000"
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtCurrentOffset 
            Alignment       =   2  'Center
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
            Index           =   8
            Left            =   1920
            TabIndex        =   251
            Text            =   "0.000"
            Top             =   1560
            Width           =   615
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
            Index           =   100
            Left            =   2040
            TabIndex        =   267
            Top             =   840
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
            Index           =   99
            Left            =   2040
            TabIndex        =   266
            Top             =   1200
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MVD"
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
            Index           =   98
            Left            =   120
            TabIndex        =   265
            Top             =   720
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current"
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
            Index           =   97
            Left            =   120
            TabIndex        =   264
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ON"
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
            Index           =   96
            Left            =   1200
            TabIndex        =   263
            Top             =   240
            Width           =   225
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min"
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
            Index           =   95
            Left            =   960
            TabIndex        =   262
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max"
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
            Index           =   91
            Left            =   1560
            TabIndex        =   261
            Top             =   480
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Test Cycle"
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
            Index           =   87
            Left            =   120
            TabIndex        =   260
            Top             =   1680
            Width           =   765
         End
         Begin VB.Label lblcurentoffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Offset"
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
            Index           =   8
            Left            =   840
            TabIndex        =   259
            Top             =   1560
            Width           =   1050
         End
         Begin VB.Label lblvoltageoffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Voltage Offset"
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
            Index           =   8
            Left            =   840
            TabIndex        =   258
            Top             =   1920
            Width           =   1065
         End
      End
      Begin VB.CheckBox ChkCUMBypass 
         Caption         =   "Custom Module Bypass"
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   5520
         TabIndex        =   249
         Top             =   6000
         Width           =   2655
      End
      Begin VB.CheckBox chkHOMBypass 
         Caption         =   "Horn Module Bypass"
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   120
         TabIndex        =   248
         Top             =   6000
         Width           =   2655
      End
      Begin VB.Frame FrameCRM 
         Height          =   2415
         Left            =   2760
         TabIndex        =   209
         Top             =   6240
         Width           =   2655
         Begin VB.TextBox txtCurrentOffset 
            Alignment       =   2  'Center
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
            Index           =   6
            Left            =   1920
            TabIndex        =   216
            Text            =   "0.000"
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox txtVoltageOffset 
            Alignment       =   2  'Center
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
            Index           =   6
            Left            =   1920
            TabIndex        =   215
            Text            =   "0.000"
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtTestCycle 
            Alignment       =   2  'Center
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
            Index           =   6
            Left            =   240
            TabIndex        =   214
            Text            =   "00"
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txtCurrMax 
            Alignment       =   2  'Center
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
            Index           =   12
            Left            =   1440
            TabIndex        =   213
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtCurrMin 
            Alignment       =   2  'Center
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
            Index           =   12
            Left            =   840
            TabIndex        =   212
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtMVDMax 
            Alignment       =   2  'Center
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
            Index           =   12
            Left            =   1440
            TabIndex        =   211
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtMVDMin 
            Alignment       =   2  'Center
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
            Index           =   12
            Left            =   840
            TabIndex        =   210
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lblvoltageoffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Voltage Offset"
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
            Index           =   7
            Left            =   840
            TabIndex        =   226
            Top             =   1920
            Width           =   1065
         End
         Begin VB.Label lblcurentoffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Offset"
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
            Index           =   7
            Left            =   840
            TabIndex        =   225
            Top             =   1560
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Test Cycle"
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
            Index           =   84
            Left            =   120
            TabIndex        =   224
            Top             =   1680
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max"
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
            Index           =   83
            Left            =   1560
            TabIndex        =   223
            Top             =   480
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min"
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
            Index           =   82
            Left            =   960
            TabIndex        =   222
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ON"
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
            Index           =   80
            Left            =   1200
            TabIndex        =   221
            Top             =   240
            Width           =   225
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current"
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
            Index           =   73
            Left            =   120
            TabIndex        =   220
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MVD"
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
            Index           =   65
            Left            =   120
            TabIndex        =   219
            Top             =   720
            Width           =   345
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
            Index           =   64
            Left            =   2040
            TabIndex        =   218
            Top             =   1200
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
            Index           =   44
            Left            =   2040
            TabIndex        =   217
            Top             =   840
            Width           =   120
         End
      End
      Begin VB.Frame FrameCUM 
         Height          =   2415
         Left            =   5520
         TabIndex        =   191
         Top             =   6240
         Width           =   2655
         Begin VB.TextBox txtCurrentOffset 
            Alignment       =   2  'Center
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
            Index           =   7
            Left            =   1920
            TabIndex        =   198
            Text            =   "0.000"
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox txtVoltageOffset 
            Alignment       =   2  'Center
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
            Index           =   7
            Left            =   1920
            TabIndex        =   197
            Text            =   "0.000"
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtTestCycle 
            Alignment       =   2  'Center
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
            Index           =   7
            Left            =   240
            TabIndex        =   196
            Text            =   "00"
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txtCurrMax 
            Alignment       =   2  'Center
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
            Index           =   13
            Left            =   1440
            TabIndex        =   195
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtCurrMin 
            Alignment       =   2  'Center
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
            Index           =   13
            Left            =   840
            TabIndex        =   194
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtMVDMax 
            Alignment       =   2  'Center
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
            Index           =   13
            Left            =   1440
            TabIndex        =   193
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtMVDMin 
            Alignment       =   2  'Center
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
            Index           =   13
            Left            =   840
            TabIndex        =   192
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lblvoltageoffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Voltage Offset"
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
            Index           =   5
            Left            =   840
            TabIndex        =   208
            Top             =   1920
            Width           =   1065
         End
         Begin VB.Label lblcurentoffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Offset"
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
            Index           =   5
            Left            =   840
            TabIndex        =   207
            Top             =   1560
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Test Cycle"
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
            Index           =   43
            Left            =   120
            TabIndex        =   206
            Top             =   1680
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max"
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
            Index           =   42
            Left            =   1560
            TabIndex        =   205
            Top             =   480
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min"
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
            Index           =   41
            Left            =   960
            TabIndex        =   204
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ON"
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
            Index           =   40
            Left            =   1200
            TabIndex        =   203
            Top             =   240
            Width           =   225
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current"
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
            Index           =   39
            Left            =   120
            TabIndex        =   202
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MVD"
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
            Index           =   20
            Left            =   120
            TabIndex        =   201
            Top             =   720
            Width           =   345
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
            Index           =   14
            Left            =   2040
            TabIndex        =   200
            Top             =   1200
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
            Left            =   2040
            TabIndex        =   199
            Top             =   840
            Width           =   120
         End
      End
      Begin VB.CheckBox chkNMBypass 
         Caption         =   "Navigation Module Bypass"
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   120
         TabIndex        =   189
         Top             =   3000
         Width           =   3135
      End
      Begin VB.CheckBox chkBMBypass 
         Caption         =   "Blinker Module Bypass"
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   6240
         TabIndex        =   188
         Top             =   3000
         Width           =   2535
      End
      Begin VB.CheckBox ChkHAMBypass 
         Caption         =   "Hazard Module Bypass"
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   7320
         TabIndex        =   187
         Top             =   120
         Width           =   2535
      End
      Begin VB.CheckBox chkCRMBypass 
         Caption         =   "Cruise Module Bypass"
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   2760
         TabIndex        =   186
         Top             =   6000
         Width           =   2655
      End
      Begin VB.CheckBox chkPMBypass 
         Caption         =   "Pass Module Bypass"
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   4080
         TabIndex        =   185
         Top             =   120
         Width           =   2535
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   19440
         Top             =   840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame12 
         ForeColor       =   &H000040C0&
         Height          =   975
         Left            =   11040
         TabIndex        =   158
         Top             =   0
         Width           =   3135
         Begin VB.TextBox txtICMin 
            Alignment       =   2  'Center
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
            Left            =   1440
            TabIndex        =   160
            Text            =   "0.000"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtICMax 
            Alignment       =   2  'Center
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
            Left            =   2160
            TabIndex        =   159
            Text            =   "0.000"
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ILLumination Curr."
            ForeColor       =   &H000040C0&
            Height          =   480
            Index           =   101
            Left            =   120
            TabIndex        =   269
            Top             =   360
            Width           =   1155
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min"
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
            Index           =   72
            Left            =   1560
            TabIndex        =   163
            Top             =   240
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max"
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
            Index           =   71
            Left            =   2280
            TabIndex        =   162
            Top             =   240
            Width           =   300
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
            Index           =   69
            Left            =   2880
            TabIndex        =   161
            Top             =   600
            Width           =   120
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Bypasses"
         ForeColor       =   &H000040C0&
         Height          =   3855
         Left            =   11040
         TabIndex        =   136
         Top             =   4800
         Width           =   3135
         Begin VB.CheckBox chkbypass 
            Caption         =   "Upper Cover Case ByPass"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   9
            Left            =   120
            TabIndex        =   190
            Top             =   3480
            Width           =   2895
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "Pressure Guage ByPass"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   8
            Left            =   120
            TabIndex        =   157
            Top             =   3120
            Width           =   2895
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "Scanner ByPass"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   6
            Left            =   120
            TabIndex        =   156
            Top             =   2400
            Width           =   2175
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "PID ByPass"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   7
            Left            =   120
            TabIndex        =   155
            Top             =   2760
            Width           =   2655
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "ILLumination Current ByPass"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   4
            Left            =   120
            TabIndex        =   142
            Top             =   2040
            Width           =   2895
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "Printer ByPass"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   5
            Left            =   120
            TabIndex        =   141
            Top             =   1680
            Width           =   2655
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "Body Short ByPass"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   3
            Left            =   120
            TabIndex        =   140
            Top             =   1320
            Width           =   2175
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "Wire Length Check ByPass"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   2
            Left            =   120
            TabIndex        =   139
            Top             =   960
            Width           =   2775
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "Blinker Limit Switch Bypass"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   138
            Top             =   600
            Width           =   2775
         End
         Begin VB.CheckBox chkbypass 
            Caption         =   "Camera Bypass"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   137
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame FrameNM 
         Height          =   2415
         Left            =   120
         TabIndex        =   113
         Top             =   3240
         Width           =   6135
         Begin VB.TextBox txtCurrMax 
            Alignment       =   2  'Center
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
            Index           =   7
            Left            =   5280
            TabIndex        =   244
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtCurrMin 
            Alignment       =   2  'Center
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
            Index           =   7
            Left            =   4680
            TabIndex        =   243
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtMVDMax 
            Alignment       =   2  'Center
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
            Index           =   7
            Left            =   5280
            TabIndex        =   242
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtMVDMin 
            Alignment       =   2  'Center
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
            Index           =   7
            Left            =   4680
            TabIndex        =   241
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtCurrMax 
            Alignment       =   2  'Center
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
            Index           =   6
            Left            =   3960
            TabIndex        =   237
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtCurrMin 
            Alignment       =   2  'Center
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
            Index           =   6
            Left            =   3360
            TabIndex        =   236
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtMVDMax 
            Alignment       =   2  'Center
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
            Index           =   6
            Left            =   3960
            TabIndex        =   235
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtMVDMin 
            Alignment       =   2  'Center
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
            Index           =   6
            Left            =   3360
            TabIndex        =   234
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtMVDMin 
            Alignment       =   2  'Center
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
            Index           =   4
            Left            =   720
            TabIndex        =   124
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtMVDMax 
            Alignment       =   2  'Center
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
            Index           =   4
            Left            =   1320
            TabIndex        =   123
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtMVDMin 
            Alignment       =   2  'Center
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
            Index           =   5
            Left            =   2040
            TabIndex        =   122
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtMVDMax 
            Alignment       =   2  'Center
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
            Index           =   5
            Left            =   2640
            TabIndex        =   121
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtCurrMin 
            Alignment       =   2  'Center
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
            Index           =   4
            Left            =   720
            TabIndex        =   120
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtCurrMax 
            Alignment       =   2  'Center
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
            Index           =   4
            Left            =   1320
            TabIndex        =   119
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtCurrMin 
            Alignment       =   2  'Center
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
            Index           =   5
            Left            =   2040
            TabIndex        =   118
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtCurrMax 
            Alignment       =   2  'Center
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
            Index           =   5
            Left            =   2640
            TabIndex        =   117
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtTestCycle 
            Alignment       =   2  'Center
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
            Index           =   3
            Left            =   1080
            TabIndex        =   116
            Text            =   "00"
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtVoltageOffset 
            Alignment       =   2  'Center
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
            Index           =   3
            Left            =   3240
            TabIndex        =   115
            Text            =   "0.000"
            Top             =   1680
            Width           =   615
         End
         Begin VB.TextBox txtCurrentOffset 
            Alignment       =   2  'Center
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
            Index           =   3
            Left            =   5280
            TabIndex        =   114
            Text            =   "0.000"
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max"
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
            Index           =   94
            Left            =   5400
            TabIndex        =   247
            Top             =   480
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min"
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
            Index           =   93
            Left            =   4800
            TabIndex        =   246
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BACK"
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
            Index           =   92
            Left            =   5040
            TabIndex        =   245
            Top             =   240
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max"
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
            Index           =   90
            Left            =   4080
            TabIndex        =   240
            Top             =   480
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min"
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
            Index           =   89
            Left            =   3480
            TabIndex        =   239
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ENTER"
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
            Index           =   88
            Left            =   3720
            TabIndex        =   238
            Top             =   240
            Width           =   480
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
            Index           =   67
            Left            =   5880
            TabIndex        =   154
            Top             =   840
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
            Index           =   66
            Left            =   5880
            TabIndex        =   153
            Top             =   1200
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MVD"
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
            Index           =   53
            Left            =   120
            TabIndex        =   135
            Top             =   720
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current"
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
            Index           =   52
            Left            =   120
            TabIndex        =   134
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UP"
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
            Index           =   51
            Left            =   1080
            TabIndex        =   133
            Top             =   240
            Width           =   195
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DOWN"
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
            Index           =   50
            Left            =   2400
            TabIndex        =   132
            Top             =   240
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min"
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
            Index           =   49
            Left            =   840
            TabIndex        =   131
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max"
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
            Index           =   48
            Left            =   1440
            TabIndex        =   130
            Top             =   480
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min"
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
            Index           =   47
            Left            =   2160
            TabIndex        =   129
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max"
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
            Index           =   46
            Left            =   2760
            TabIndex        =   128
            Top             =   480
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Test Cycle"
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
            Index           =   45
            Left            =   120
            TabIndex        =   127
            Top             =   1560
            Width           =   765
         End
         Begin VB.Label lblcurentoffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Offset"
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
            Index           =   6
            Left            =   4080
            TabIndex        =   126
            Top             =   1680
            Width           =   1050
         End
         Begin VB.Label lblvoltageoffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Voltage Offset"
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
            Index           =   6
            Left            =   2040
            TabIndex        =   125
            Top             =   1680
            Width           =   1065
         End
      End
      Begin VB.Frame FrameHAM 
         Height          =   2415
         Left            =   7320
         TabIndex        =   97
         Top             =   360
         Width           =   2895
         Begin VB.TextBox txtCurrentOffset 
            Alignment       =   2  'Center
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
            Index           =   2
            Left            =   2160
            TabIndex        =   104
            Text            =   "0.000"
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox txtVoltageOffset 
            Alignment       =   2  'Center
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
            Index           =   2
            Left            =   2160
            TabIndex        =   103
            Text            =   "0.000"
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtTestCycle 
            Alignment       =   2  'Center
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
            Index           =   2
            Left            =   240
            TabIndex        =   102
            Text            =   "00"
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txtCurrMax 
            Alignment       =   2  'Center
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
            Index           =   3
            Left            =   1440
            TabIndex        =   101
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtCurrMin 
            Alignment       =   2  'Center
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
            Index           =   3
            Left            =   840
            TabIndex        =   100
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtMVDMax 
            Alignment       =   2  'Center
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
            Index           =   3
            Left            =   1440
            TabIndex        =   99
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtMVDMin 
            Alignment       =   2  'Center
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
            Index           =   3
            Left            =   840
            TabIndex        =   98
            Text            =   "0.000"
            Top             =   720
            Width           =   495
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
            Index           =   63
            Left            =   2040
            TabIndex        =   152
            Top             =   840
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
            Index           =   62
            Left            =   2040
            TabIndex        =   151
            Top             =   1200
            Width           =   120
         End
         Begin VB.Label lblvoltageoffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Voltage Offset"
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
            Index           =   4
            Left            =   960
            TabIndex        =   112
            Top             =   1920
            Width           =   1065
         End
         Begin VB.Label lblcurentoffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Offset"
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
            Index           =   4
            Left            =   960
            TabIndex        =   111
            Top             =   1560
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Test Cycle"
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
            Index           =   38
            Left            =   120
            TabIndex        =   110
            Top             =   1560
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max"
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
            Index           =   37
            Left            =   1560
            TabIndex        =   109
            Top             =   480
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min"
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
            Index           =   36
            Left            =   960
            TabIndex        =   108
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ON"
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
            Index           =   35
            Left            =   1200
            TabIndex        =   107
            Top             =   240
            Width           =   225
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current"
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
            Index           =   34
            Left            =   120
            TabIndex        =   106
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MVD"
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
            Index           =   33
            Left            =   120
            TabIndex        =   105
            Top             =   720
            Width           =   345
         End
      End
      Begin VB.Frame FrameHoM 
         Height          =   2415
         Left            =   120
         TabIndex        =   81
         Top             =   6240
         Width           =   2655
         Begin VB.TextBox txtCurrentOffset 
            Alignment       =   2  'Center
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
            Index           =   5
            Left            =   1920
            TabIndex        =   88
            Text            =   "0.000"
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox txtVoltageOffset 
            Alignment       =   2  'Center
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
            Index           =   5
            Left            =   1920
            TabIndex        =   87
            Text            =   "0.000"
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtTestCycle 
            Alignment       =   2  'Center
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
            Index           =   5
            Left            =   240
            TabIndex        =   86
            Text            =   "00"
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txtCurrMax 
            Alignment       =   2  'Center
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
            Index           =   11
            Left            =   1440
            TabIndex        =   85
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtCurrMin 
            Alignment       =   2  'Center
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
            Index           =   11
            Left            =   840
            TabIndex        =   84
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtMVDMax 
            Alignment       =   2  'Center
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
            Index           =   11
            Left            =   1440
            TabIndex        =   83
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtMVDMin 
            Alignment       =   2  'Center
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
            Index           =   11
            Left            =   840
            TabIndex        =   82
            Text            =   "0.000"
            Top             =   720
            Width           =   495
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
            Index           =   61
            Left            =   2040
            TabIndex        =   150
            Top             =   840
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
            Index           =   60
            Left            =   2040
            TabIndex        =   149
            Top             =   1200
            Width           =   120
         End
         Begin VB.Label lblvoltageoffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Voltage Offset"
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
            Index           =   3
            Left            =   840
            TabIndex        =   96
            Top             =   1920
            Width           =   1065
         End
         Begin VB.Label lblcurentoffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Offset"
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
            Index           =   3
            Left            =   840
            TabIndex        =   95
            Top             =   1560
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Test Cycle"
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
            Index           =   32
            Left            =   120
            TabIndex        =   94
            Top             =   1680
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max"
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
            Index           =   31
            Left            =   1560
            TabIndex        =   93
            Top             =   480
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min"
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
            Index           =   30
            Left            =   960
            TabIndex        =   92
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ON"
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
            Index           =   29
            Left            =   1200
            TabIndex        =   91
            Top             =   240
            Width           =   225
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current"
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
            Index           =   28
            Left            =   120
            TabIndex        =   90
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MVD"
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
            Index           =   27
            Left            =   120
            TabIndex        =   89
            Top             =   720
            Width           =   345
         End
      End
      Begin VB.Frame FrameBM 
         Height          =   2415
         Left            =   6240
         TabIndex        =   58
         Top             =   3240
         Width           =   4815
         Begin VB.TextBox txtCurrMax 
            Alignment       =   2  'Center
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
            Index           =   10
            Left            =   3960
            TabIndex        =   230
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtCurrMin 
            Alignment       =   2  'Center
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
            Index           =   10
            Left            =   3360
            TabIndex        =   229
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtMVDMax 
            Alignment       =   2  'Center
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
            Index           =   10
            Left            =   3960
            TabIndex        =   228
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtMVDMin 
            Alignment       =   2  'Center
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
            Index           =   10
            Left            =   3360
            TabIndex        =   227
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtMVDMin 
            Alignment       =   2  'Center
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
            Index           =   9
            Left            =   2040
            TabIndex        =   77
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtMVDMax 
            Alignment       =   2  'Center
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
            Index           =   9
            Left            =   2640
            TabIndex        =   76
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtCurrMin 
            Alignment       =   2  'Center
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
            Index           =   9
            Left            =   2040
            TabIndex        =   75
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtCurrMax 
            Alignment       =   2  'Center
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
            Index           =   9
            Left            =   2640
            TabIndex        =   74
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtMVDMin 
            Alignment       =   2  'Center
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
            Index           =   8
            Left            =   720
            TabIndex        =   65
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtMVDMax 
            Alignment       =   2  'Center
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
            Index           =   8
            Left            =   1320
            TabIndex        =   64
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtCurrMin 
            Alignment       =   2  'Center
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
            Index           =   8
            Left            =   720
            TabIndex        =   63
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtCurrMax 
            Alignment       =   2  'Center
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
            Index           =   8
            Left            =   1320
            TabIndex        =   62
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtTestCycle 
            Alignment       =   2  'Center
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
            Index           =   4
            Left            =   360
            TabIndex        =   61
            Text            =   "00"
            Top             =   1800
            Width           =   495
         End
         Begin VB.TextBox txtVoltageOffset 
            Alignment       =   2  'Center
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
            Index           =   4
            Left            =   2760
            TabIndex        =   60
            Text            =   "0.000"
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtCurrentOffset 
            Alignment       =   2  'Center
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
            Index           =   4
            Left            =   2760
            TabIndex        =   59
            Text            =   "0.000"
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max"
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
            Index           =   86
            Left            =   4080
            TabIndex        =   233
            Top             =   480
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min"
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
            Index           =   85
            Left            =   3480
            TabIndex        =   232
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RIGHT"
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
            Index           =   81
            Left            =   3720
            TabIndex        =   231
            Top             =   240
            Width           =   450
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
            Index           =   59
            Left            =   4560
            TabIndex        =   148
            Top             =   840
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
            Index           =   58
            Left            =   4560
            TabIndex        =   147
            Top             =   1200
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OFF"
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
            Index           =   26
            Left            =   2400
            TabIndex        =   80
            Top             =   240
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min"
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
            Index           =   25
            Left            =   2160
            TabIndex        =   79
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max"
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
            Left            =   2760
            TabIndex        =   78
            Top             =   480
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MVD"
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
            Index           =   23
            Left            =   120
            TabIndex        =   73
            Top             =   720
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current"
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
            Index           =   22
            Left            =   120
            TabIndex        =   72
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LEFT"
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
            Index           =   21
            Left            =   1080
            TabIndex        =   71
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min"
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
            Index           =   19
            Left            =   840
            TabIndex        =   70
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max"
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
            Index           =   18
            Left            =   1440
            TabIndex        =   69
            Top             =   480
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Test Cycle"
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
            Left            =   240
            TabIndex        =   68
            Top             =   1560
            Width           =   765
         End
         Begin VB.Label lblcurentoffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Offset"
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
            Left            =   1560
            TabIndex        =   67
            Top             =   1560
            Width           =   1050
         End
         Begin VB.Label lblvoltageoffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Voltage Offset"
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
            Left            =   1560
            TabIndex        =   66
            Top             =   1920
            Width           =   1065
         End
      End
      Begin VB.Frame FramePM 
         Height          =   2415
         Left            =   4080
         TabIndex        =   42
         Top             =   360
         Width           =   2895
         Begin VB.TextBox txtMVDMin 
            Alignment       =   2  'Center
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
            Index           =   2
            Left            =   840
            TabIndex        =   49
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtMVDMax 
            Alignment       =   2  'Center
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
            Index           =   2
            Left            =   1440
            TabIndex        =   48
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtCurrMin 
            Alignment       =   2  'Center
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
            Index           =   2
            Left            =   840
            TabIndex        =   47
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtCurrMax 
            Alignment       =   2  'Center
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
            Index           =   2
            Left            =   1440
            TabIndex        =   46
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtTestCycle 
            Alignment       =   2  'Center
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
            Index           =   1
            Left            =   240
            TabIndex        =   45
            Text            =   "00"
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox txtVoltageOffset 
            Alignment       =   2  'Center
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
            Index           =   1
            Left            =   2160
            TabIndex        =   44
            Text            =   "0.000"
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtCurrentOffset 
            Alignment       =   2  'Center
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
            Index           =   1
            Left            =   2160
            TabIndex        =   43
            Text            =   "0.000"
            Top             =   1560
            Width           =   615
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
            Index           =   57
            Left            =   2040
            TabIndex        =   146
            Top             =   840
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
            Index           =   56
            Left            =   2040
            TabIndex        =   145
            Top             =   1200
            Width           =   120
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MVD"
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
            Index           =   17
            Left            =   120
            TabIndex        =   57
            Top             =   720
            Width           =   345
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current"
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
            Index           =   16
            Left            =   120
            TabIndex        =   56
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pass"
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
            Index           =   15
            Left            =   1200
            TabIndex        =   55
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min"
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
            Left            =   960
            TabIndex        =   54
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max"
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
            Left            =   1560
            TabIndex        =   53
            Top             =   480
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Test Cycle"
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
            Index           =   9
            Left            =   120
            TabIndex        =   52
            Top             =   1560
            Width           =   765
         End
         Begin VB.Label lblcurentoffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Offset"
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
            Index           =   1
            Left            =   960
            TabIndex        =   51
            Top             =   1560
            Width           =   1050
         End
         Begin VB.Label lblvoltageoffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Voltage Offset"
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
            Index           =   1
            Left            =   960
            TabIndex        =   50
            Top             =   1920
            Width           =   1065
         End
      End
      Begin VB.CheckBox chkDMBypass 
         Caption         =   "Dipper Module Bypass"
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   2655
      End
      Begin VB.Frame FrameDM 
         Height          =   2415
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   3615
         Begin VB.TextBox txtCurrentOffset 
            Alignment       =   2  'Center
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
            Index           =   0
            Left            =   2760
            TabIndex        =   41
            Text            =   "0.000"
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox txtVoltageOffset 
            Alignment       =   2  'Center
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
            Index           =   0
            Left            =   2760
            TabIndex        =   39
            Text            =   "0.000"
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtTestCycle 
            Alignment       =   2  'Center
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
            Index           =   0
            Left            =   360
            TabIndex        =   37
            Text            =   "00"
            Top             =   1800
            Width           =   615
         End
         Begin VB.TextBox txtCurrMax 
            Alignment       =   2  'Center
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
            Index           =   1
            Left            =   2760
            TabIndex        =   35
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtCurrMin 
            Alignment       =   2  'Center
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
            Index           =   1
            Left            =   2160
            TabIndex        =   34
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtCurrMax 
            Alignment       =   2  'Center
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
            Index           =   0
            Left            =   1440
            TabIndex        =   33
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtCurrMin 
            Alignment       =   2  'Center
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
            Index           =   0
            Left            =   840
            TabIndex        =   32
            Text            =   "0.000"
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtMVDMax 
            Alignment       =   2  'Center
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
            Index           =   1
            Left            =   2760
            TabIndex        =   31
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtMVDMin 
            Alignment       =   2  'Center
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
            Index           =   1
            Left            =   2160
            TabIndex        =   30
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtMVDMax 
            Alignment       =   2  'Center
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
            Index           =   0
            Left            =   1440
            TabIndex        =   29
            Text            =   "0.000"
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtMVDMin 
            Alignment       =   2  'Center
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
            Index           =   0
            Left            =   840
            TabIndex        =   28
            Text            =   "0.000"
            Top             =   720
            Width           =   495
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
            Left            =   3360
            TabIndex        =   144
            Top             =   840
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
            Index           =   54
            Left            =   3360
            TabIndex        =   143
            Top             =   1200
            Width           =   120
         End
         Begin VB.Label lblvoltageoffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Voltage Offset"
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
            Index           =   0
            Left            =   1560
            TabIndex        =   40
            Top             =   1920
            Width           =   1065
         End
         Begin VB.Label lblcurentoffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Offset"
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
            Index           =   0
            Left            =   1560
            TabIndex        =   38
            Top             =   1560
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Test Cycle"
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
            Index           =   8
            Left            =   240
            TabIndex        =   36
            Top             =   1560
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max"
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
            Index           =   7
            Left            =   2880
            TabIndex        =   27
            Top             =   480
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min"
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
            Index           =   6
            Left            =   2280
            TabIndex        =   26
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max"
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
            Index           =   5
            Left            =   1560
            TabIndex        =   25
            Top             =   480
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Min"
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
            Index           =   4
            Left            =   960
            TabIndex        =   24
            Top             =   480
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HIGH"
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
            Index           =   3
            Left            =   2520
            TabIndex        =   23
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LOW"
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
            Left            =   1200
            TabIndex        =   22
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current"
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
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MVD"
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
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   720
            Width           =   345
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2250
         Left            =   14280
         TabIndex        =   11
         Top             =   0
         Width           =   5415
         Begin VB.CommandButton cmdImage 
            Caption         =   "...."
            Height          =   240
            Left            =   4800
            TabIndex        =   184
            Top             =   1800
            Width           =   375
         End
         Begin VB.TextBox txtImagePath 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   1440
            TabIndex        =   182
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
            TabIndex        =   183
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
         Height          =   5175
         Left            =   14280
         TabIndex        =   7
         Top             =   2280
         Width           =   5385
         Begin VSFlex7Ctl.VSFlexGrid VSFModel 
            Height          =   4365
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   5115
            _cx             =   9022
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
            Top             =   4800
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
         Left            =   14280
         TabIndex        =   1
         Top             =   7440
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
         Height          =   1695
         Left            =   11040
         TabIndex        =   164
         Top             =   960
         Width           =   3135
         Begin VB.TextBox txtMarkTime 
            Alignment       =   2  'Center
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
            TabIndex        =   171
            Text            =   "0.000"
            Top             =   1320
            Width           =   615
         End
         Begin VB.TextBox txtCheckTime 
            Alignment       =   2  'Center
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
            TabIndex        =   169
            Text            =   "0.000"
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtHoldTime 
            Alignment       =   2  'Center
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
            TabIndex        =   166
            Text            =   "0.000"
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox txtDebounceTime 
            Alignment       =   2  'Center
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
            TabIndex        =   165
            Text            =   "0.000"
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dot Mark Time"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   74
            Left            =   120
            TabIndex        =   172
            Top             =   1320
            Width           =   1275
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Curr/MVD Check Time"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   70
            Left            =   120
            TabIndex        =   170
            Top             =   960
            Width           =   2070
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hold Time"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   76
            Left            =   120
            TabIndex        =   168
            Top             =   600
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Debounce Time"
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   68
            Left            =   120
            TabIndex        =   167
            Top             =   240
            Width           =   1350
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Printer Detail"
         ForeColor       =   &H000040C0&
         Height          =   2175
         Left            =   11040
         TabIndex        =   173
         Top             =   2640
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
            TabIndex        =   271
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
            TabIndex        =   177
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
            TabIndex        =   176
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
            TabIndex        =   175
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
            TabIndex        =   174
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
            TabIndex        =   270
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
            TabIndex        =   181
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
            TabIndex        =   180
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
            TabIndex        =   179
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
            TabIndex        =   178
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

Private Sub Check1_Click()

End Sub

Private Sub chkDMBypass_Click()
If chkDMBypass.Value = 1 Then
FrameDM.Visible = False
Else
FrameDM.Visible = True
End If
End Sub
Private Sub chkPMBypass_Click()
If chkPMBypass.Value = 1 Then
FramePM.Visible = False
Else
FramePM.Visible = True
End If
End Sub
Private Sub chkBMBypass_Click()
If chkBMBypass.Value = 1 Then
FrameBM.Visible = False
Else
FrameBM.Visible = True
End If
End Sub
Private Sub chkhaMBypass_Click()
If ChkHAMBypass.Value = 1 Then
FrameHAM.Visible = False
Else
FrameHAM.Visible = True
End If
End Sub
Private Sub chkhoMBypass_Click()
If chkHOMBypass.Value = 1 Then
FrameHoM.Visible = False
Else
FrameHoM.Visible = True
End If
End Sub
Private Sub chknMBypass_Click()
If chkNMBypass.Value = 1 Then
FrameNM.Visible = False
Else
FrameNM.Visible = True
End If
End Sub
Private Sub chkcumBypass_Click()
If ChkCUMBypass.Value = 1 Then
FrameCUM.Visible = False
Else
FrameCUM.Visible = True
End If
End Sub

Private Sub chkcrmBypass_Click()
If chkCRMBypass.Value = 1 Then
FrameCRM.Visible = False
Else
FrameCRM.Visible = True
End If
End Sub

Private Sub chksrmBypass_Click()
If ChkSRMBypass.Value = 1 Then
FrameSRM.Visible = False
Else
FrameSRM.Visible = True
End If
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




Private Sub Combo1_Click()
If Combo1.ListIndex = 0 Then
 Combo2.Visible = True
 Combo3.Visible = False
 Else
   Combo2.Visible = False
   Combo3.Visible = True
End If
End Sub

Private Sub Combo9_Click()
If Combo9.ListIndex = 0 Then
 Combo8.Visible = True
 Combo7.Visible = False
 Else
   Combo8.Visible = False
   Combo7.Visible = True
End If
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
    Rs("DMBypass") = chkDMBypass.Value
    Rs("PMBypass") = chkPMBypass.Value
    Rs("HAMBypass") = ChkHAMBypass.Value
    Rs("NMBypass") = chkNMBypass.Value
    Rs("BMBypass") = chkBMBypass.Value
    Rs("HOMBypass") = chkHOMBypass.Value
    Rs("CUMBypass") = ChkCUMBypass.Value
    Rs("SRMBypass") = ChkSRMBypass.Value
    Rs("CRMBypass") = chkCRMBypass.Value
    
    Rs("DM1CurMin") = Format(txtCurrMin(0), "0.000")
    Rs("DM2CurMin") = Format(txtCurrMin(1), "0.000")
    Rs("PM1CurMin") = Format(txtCurrMin(2), "0.000")
    Rs("HAM1CurMin") = Format(txtCurrMin(3), "0.000")
    Rs("NM1CurMin") = Format(txtCurrMin(4), "0.000")
    Rs("NM2CurMin") = Format(txtCurrMin(5), "0.000")
    Rs("NM3CurMin") = Format(txtCurrMin(6), "0.000")
    Rs("NM4CurMin") = Format(txtCurrMin(7), "0.000")
    Rs("BM1CurMin") = Format(txtCurrMin(8), "0.000")
    Rs("BM2CurMin") = Format(txtCurrMin(9), "0.000")
    Rs("BM3CurMin") = Format(txtCurrMin(10), "0.000")
    Rs("HOM1CurMin") = Format(txtCurrMin(11), "0.000")
    Rs("CRM1CurMin") = Format(txtCurrMin(12), "0.000")
    Rs("CUM1CurMin") = Format(txtCurrMin(13), "0.000")
    Rs("SRM1CurMin") = Format(txtCurrMin(14), "0.000")
    
    Rs("DM1CurMax") = Format(txtCurrMax(0), "0.000")
    Rs("DM2CurMax") = Format(txtCurrMax(1), "0.000")
    Rs("PM1CurMax") = Format(txtCurrMax(2), "0.000")
    Rs("HAM1CurMax") = Format(txtCurrMax(3), "0.000")
    Rs("NM1CurMax") = Format(txtCurrMax(4), "0.000")
    Rs("NM2CurMax") = Format(txtCurrMax(5), "0.000")
    Rs("NM3CurMax") = Format(txtCurrMax(6), "0.000")
    Rs("NM4CurMax") = Format(txtCurrMax(7), "0.000")
    Rs("BM1CurMax") = Format(txtCurrMax(8), "0.000")
    Rs("BM2CurMax") = Format(txtCurrMax(9), "0.000")
    Rs("BM3CurMax") = Format(txtCurrMax(10), "0.000")
    Rs("HOM1CurMax") = Format(txtCurrMax(11), "0.000")
    Rs("CRM1CurMax") = Format(txtCurrMax(12), "0.000")
    Rs("CUM1CurMax") = Format(txtCurrMax(13), "0.000")
    Rs("SRM1CurMax") = Format(txtCurrMax(14), "0.000")
    
    Rs("DM1VoltMin") = Format(txtMVDMin(0), "0.000")
    Rs("DM2VoltMin") = Format(txtMVDMin(1), "0.000")
    Rs("PM1VoltMin") = Format(txtMVDMin(2), "0.000")
    Rs("HAM1VoltMin") = Format(txtMVDMin(3), "0.000")
    Rs("NM1VoltMin") = Format(txtMVDMin(4), "0.000")
    Rs("NM2VoltMin") = Format(txtMVDMin(5), "0.000")
    Rs("NM3VoltMin") = Format(txtMVDMin(6), "0.000")
    Rs("NM4VoltMin") = Format(txtMVDMin(7), "0.000")
    Rs("BM1VoltMin") = Format(txtMVDMin(8), "0.000")
    Rs("BM2VoltMin") = Format(txtMVDMin(9), "0.000")
    Rs("BM3VoltMin") = Format(txtMVDMin(10), "0.000")
    Rs("HOM1VoltMin") = Format(txtMVDMin(11), "0.000")
    Rs("CRM1VoltMin") = Format(txtMVDMin(12), "0.000")
    Rs("CUM1VoltMin") = Format(txtMVDMin(13), "0.000")
    Rs("SRM1VoltMin") = Format(txtMVDMin(14), "0.000")
    
    Rs("DM1VoltMax") = Format(txtMVDMax(0), "0.000")
    Rs("DM2VoltMax") = Format(txtMVDMax(1), "0.000")
    Rs("PM1VoltMax") = Format(txtMVDMax(2), "0.000")
    Rs("HAM1VoltMax") = Format(txtMVDMax(3), "0.000")
    Rs("NM1VoltMax") = Format(txtMVDMax(4), "0.000")
    Rs("NM2VoltMax") = Format(txtMVDMax(5), "0.000")
    Rs("NM3VoltMax") = Format(txtMVDMax(6), "0.000")
    Rs("NM4VoltMax") = Format(txtMVDMax(7), "0.000")
    Rs("BM1VoltMax") = Format(txtMVDMax(8), "0.000")
    Rs("BM2VoltMax") = Format(txtMVDMax(9), "0.000")
    Rs("BM3VoltMax") = Format(txtMVDMax(10), "0.000")
    Rs("HOM1VoltMax") = Format(txtMVDMax(11), "0.000")
    Rs("CUM1VoltMax") = Format(txtMVDMax(13), "0.000")
    Rs("SRM1VoltMax") = Format(txtMVDMax(14), "0.000")
    Rs("CRM1VoltMax") = Format(txtMVDMax(12), "0.000")
    
    Rs("DMTestCycle") = Format(txtTestCycle(0), "0")
    Rs("PMTestCycle") = Format(txtTestCycle(1), "0")
    Rs("HAMTestCycle") = Format(txtTestCycle(2), "0")
    Rs("NMTestCycle") = Format(txtTestCycle(3), "0")
    Rs("BMTestCycle") = Format(txtTestCycle(4), "0")
    Rs("HOMTestCycle") = Format(txtTestCycle(5), "0")
    Rs("CRMTestCycle") = Format(txtTestCycle(6), "0")
    Rs("CUMTestCycle") = Format(txtTestCycle(7), "0")
    Rs("SRMTestCycle") = Format(txtTestCycle(8), "0")
    
    For i = 0 To 8
      Rs("CurrentOffset" & i + 1) = Format(txtCurrentOffset(i).Text, "0.000")
      Rs("VoltageOffset" & i + 1) = Format(txtVoltageOffset(i).Text, "0.000")
    Next
    
    'Rs("WVMin") = Format(txtWirevoltageMin.Text, "0.00")
    'Rs("WVMax") = Format(txtWireVoltageMax.Text, "0.00")
    
    Rs("ICMin") = Format(txtICMin.Text, "0.000")
    Rs("ICMax") = Format(txtICMax.Text, "0.000")
    'Rs("ICMinRh") = Format(txtICMinRH.Text, "0.00")
    'Rs("ICMaxRh") = Format(txtICMaxRH.Text, "0.00")
    Rs("PrintPartNo") = txtPartNo.Text
    Rs("PrintBarcodeLength") = txtBarcodeLength.Text
    Rs("BarcodeLength") = txtBarcodeLength.Text
    Rs("HardwareNo") = txtHardwareVersion.Text
    Rs("SerialStartingtxt") = txtSerialNo.Text
    
    Rs("DebounceTime") = Format(txtDebounceTime.Text, "0.000")
    Rs("HoldTime") = Format(txtHoldTime.Text, "0.000")
    Rs("CheckTime") = Format(txtCheckTime.Text, "0.000")
    Rs("DotMarkingTime") = Format(txtMarkTime.Text, "0.0")
    Rs("ModelNo") = txtModelNo.Text
    'Rs("BatchCounter").Text
    'Rs("CouplerCounter") = .Text
    Rs("PartImage") = txtImagePath.Text
    'Rs("productioncounter") =
    'Rs("CameraBypass") = Val(chkbypass(0).Value)
    'Rs("LSBypass") = Val(chkbypass(1).Value)
    'Rs("WLCBypass") = Val(chkbypass(2).Value)
    'Rs("BSBypass") = Val(chkbypass(3).Value)
    Rs("PrinterBypass") = Val(chkbypass(4).Value)
    'Rs("ICBypass") = Val(chkbypass(5).Value)
    'Rs("ScannerBypass") = Val(chkbypass(6).Value)
    'Rs("PIDByPass") = Val(chkbypass(7).Value)
    'Rs("PressureGuageByPass") = Val(chkbypass(8).Value)
    'Rs("UpperCoverByPass") = Val(chkbypass(9).Value)
    For i = 0 To 9
     Rs("Bypass" & i + 1) = Val(chkbypass(i).Value)
    Next
    
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
    chkDMBypass.Value = Val(Rs("DMBypass"))
    txtCurrMin(0) = Rs("DM1CurMin")
    txtCurrMax(0) = Rs("DM1CurMax")
    txtMVDMin(0) = Rs("DM1VoltMin")
    txtMVDMax(0) = Rs("DM1VoltMax")
    txtCurrMin(1) = Rs("DM2CurMin")
    txtCurrMax(1) = Rs("DM2CurMax")
    txtMVDMin(1) = Rs("DM2VoltMin")
    txtMVDMax(1) = Rs("DM2VoltMax")
    txtTestCycle(0) = Rs("DMTestCycle")
    
    chkPMBypass.Value = Val(Rs("PMBypass"))
    txtCurrMin(2) = Rs("PM1CurMin")
    txtCurrMax(2) = Rs("PM1CurMax")
    txtMVDMin(2) = Rs("PM1VoltMin")
    txtMVDMax(2) = Rs("PM1VoltMax")
    txtTestCycle(1) = Rs("PMTestCycle")
    
    ChkHAMBypass.Value = Val(Rs("HAMBypass"))
    txtCurrMin(3) = Rs("HAM1CurMin")
    txtCurrMax(3) = Rs("HAM1CurMax")
    txtMVDMin(3) = Rs("HAM1VoltMin")
    txtMVDMax(3) = Rs("HAM1VoltMax")
    txtTestCycle(2) = Rs("HAMTestCycle")
    
    chkNMBypass.Value = Val(Rs("NMBypass"))
    txtCurrMin(4) = Rs("NM1CurMin")
    txtCurrMax(4) = Rs("NM1CurMax")
    txtMVDMin(4) = Rs("NM1VoltMin")
    txtMVDMax(4) = Rs("NM1VoltMax")
    txtCurrMin(5) = Rs("NM2CurMin")
    txtCurrMax(5) = Rs("NM2CurMax")
    txtMVDMin(5) = Rs("NM2VoltMin")
    txtMVDMax(5) = Rs("NM2VoltMax")
    txtCurrMin(6) = Rs("NM3CurMin")
    txtCurrMax(6) = Rs("NM3CurMax")
    txtMVDMin(6) = Rs("NM3VoltMin")
    txtMVDMax(6) = Rs("NM3VoltMax")
    txtCurrMin(7) = Rs("NM4CurMin")
    txtCurrMax(7) = Rs("NM4CurMax")
    txtMVDMin(7) = Rs("NM4VoltMin")
    txtMVDMax(7) = Rs("NM4VoltMax")
    txtTestCycle(3) = Rs("NMTestCycle")
    
    chkBMBypass.Value = Val(Rs("BMBypass"))
    txtCurrMin(8) = Rs("BM1CurMin")
    txtCurrMax(8) = Rs("BM1CurMax")
    txtMVDMin(8) = Rs("BM1VoltMin")
    txtMVDMax(8) = Rs("BM1VoltMax")
    txtCurrMin(9) = Rs("BM2CurMin")
    txtCurrMax(9) = Rs("BM2CurMax")
    txtMVDMin(9) = Rs("BM2VoltMin")
    txtMVDMax(9) = Rs("BM2VoltMax")
    txtCurrMin(10) = Rs("BM3CurMin")
    txtCurrMax(10) = Rs("BM3CurMax")
    txtMVDMin(10) = Rs("BM3VoltMin")
    txtMVDMax(10) = Rs("BM3VoltMax")
    txtTestCycle(4) = Rs("BMTestCycle")
    
    chkHOMBypass.Value = Val(Rs("HOMBypass"))
    txtCurrMin(11) = Rs("HOM1CurMin")
    txtCurrMax(11) = Rs("HOM1CurMax")
    txtMVDMin(11) = Rs("HOM1VoltMin")
    txtMVDMax(11) = Rs("HOM1VoltMax")
    txtTestCycle(5) = Rs("HOMTestCycle")
    
    chkCRMBypass.Value = Val(Rs("CRMBypass"))
    txtCurrMin(12) = Rs("CRM1CurMin")
    txtCurrMax(12) = Rs("CRM1CurMax")
    txtMVDMin(12) = Rs("CRM1VoltMin")
    txtMVDMax(12) = Rs("CRM1VoltMax")
    txtTestCycle(6) = Rs("CRMTestCycle")
    
    ChkCUMBypass.Value = Val(Rs("CUMBypass"))
    txtCurrMin(13) = Rs("CUM1CurMin")
    txtCurrMax(13) = Rs("CUM1CurMax")
    txtMVDMin(13) = Rs("CUM1VoltMin")
    txtMVDMax(13) = Rs("CUM1VoltMax")
    txtTestCycle(7) = Rs("CUMTestCycle")
    
    ChkSRMBypass.Value = Val(Rs("SRMBypass"))
    txtCurrMin(14) = Rs("SRM1CurMin")
    txtCurrMax(14) = Rs("SRM1CurMax")
    txtMVDMin(14) = Rs("SRM1VoltMin")
    txtMVDMax(14) = Rs("SRM1VoltMax")
    txtTestCycle(8) = Rs("SRMTestCycle")
    
    For i = 0 To 8
      txtCurrentOffset(i).Text = Rs("CurrentOffset" & i + 1)
      txtVoltageOffset(i).Text = Rs("VoltageOffset" & i + 1)
    Next
    txtICMin.Text = Rs("ICMin")
    txtICMax.Text = Rs("ICMax")
    'txtICMinRH.Text = Rs("ICMinRH")
    'txtICMaxRH.Text = Rs("ICMaxRH")
    'txtWirevoltageMin.Text = Rs("WVMin")
    'txtWireVoltageMax.Text = Rs("WVMax")
    
    txtPartNo.Text = Rs("PrintPartNo")
    txtBarcodeLength.Text = Rs("PrintBarcodeLength")
    txtBarcodeLength.Text = Rs("BarcodeLength")
    txtHardwareVersion.Text = Rs("HardwareNo")
    txtSerialNo.Text = Rs("SerialStartingtxt")
    txtDebounceTime.Text = Rs("DebounceTime")
    txtHoldTime.Text = Rs("HoldTime")
    txtCheckTime.Text = Rs("CheckTime")
    txtMarkTime.Text = Rs("DotMarkingTime")
    txtModelNo.Text = Rs("ModelNo")
    'Rs("BatchCounter").Text
    'Rs("CouplerCounter") = .Text
    txtImagePath.Text = Rs("PartImage")
    'Rs("productioncounter") =
    For i = 0 To 9
      chkbypass(i).Value = Val(Rs("Bypass" & i + 1))
    Next
    'chkbypass(1).Value = Val(Rs("LSBypass"))
    'chkbypass(2).Value = Val(Rs("WLCBypass"))
    'chkbypass(3).Value = Val(Rs("BSBypass"))
    'chkbypass(4).Value = Val(Rs("PrinterBypass"))
    'chkbypass(5).Value = Val(Rs("ICBypass"))
    'chkbypass(6).Value = Val(Rs("ScannerBypass"))
    'chkbypass(7).Value = Val(Rs("PIDByPass"))
    'chkbypass(8).Value = Val(Rs("PressureGuageByPass"))
    'chkbypass(9).Value = Val(Rs("UpperCoverByPass"))
       
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

