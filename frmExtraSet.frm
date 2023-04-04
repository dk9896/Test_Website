VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmExtraSet 
   Caption         =   "Extra Settings"
   ClientHeight    =   8205
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14085
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
   ScaleHeight     =   8205
   ScaleWidth      =   14085
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000010&
      Height          =   8055
      Left            =   120
      ScaleHeight     =   7995
      ScaleWidth      =   13080
      TabIndex        =   0
      Top             =   120
      Width           =   13140
      Begin VB.Frame Frame6 
         Caption         =   "Send Mail"
         Height          =   5295
         Left            =   6720
         TabIndex        =   18
         Top             =   120
         Width           =   6255
         Begin VB.CheckBox chkTomailbypass 
            BackColor       =   &H00808080&
            Caption         =   "Bypass"
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   0
            Left            =   2040
            TabIndex        =   45
            Top             =   2280
            Width           =   975
         End
         Begin VB.CheckBox chkTomailbypass 
            BackColor       =   &H00808080&
            Caption         =   "Bypass"
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   5
            Left            =   2040
            TabIndex        =   44
            Top             =   4680
            Width           =   975
         End
         Begin VB.CheckBox chkTomailbypass 
            BackColor       =   &H00808080&
            Caption         =   "Bypass"
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   4
            Left            =   2040
            TabIndex        =   43
            Top             =   4200
            Width           =   975
         End
         Begin VB.CheckBox chkTomailbypass 
            BackColor       =   &H00808080&
            Caption         =   "Bypass"
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   3
            Left            =   2040
            TabIndex        =   42
            Top             =   3720
            Width           =   975
         End
         Begin VB.CheckBox chkTomailbypass 
            BackColor       =   &H00808080&
            Caption         =   "Bypass"
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   2
            Left            =   2040
            TabIndex        =   41
            Top             =   3240
            Width           =   975
         End
         Begin VB.CheckBox chkTomailbypass 
            BackColor       =   &H00808080&
            Caption         =   "Bypass"
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   1
            Left            =   2040
            TabIndex        =   40
            Top             =   2760
            Width           =   975
         End
         Begin VB.TextBox txtTomail 
            Height          =   360
            Index           =   6
            Left            =   3120
            TabIndex        =   38
            Top             =   4680
            Width           =   3015
         End
         Begin VB.TextBox txtTomail 
            Height          =   360
            Index           =   5
            Left            =   3120
            TabIndex        =   36
            Top             =   4200
            Width           =   3015
         End
         Begin VB.TextBox txtTomail 
            Height          =   360
            Index           =   4
            Left            =   3120
            TabIndex        =   34
            Top             =   3720
            Width           =   3015
         End
         Begin VB.TextBox txtTomail 
            Height          =   360
            Index           =   3
            Left            =   3120
            TabIndex        =   30
            Top             =   3240
            Width           =   3015
         End
         Begin VB.TextBox txtTomail 
            Height          =   360
            Index           =   2
            Left            =   3120
            TabIndex        =   29
            Top             =   2760
            Width           =   3015
         End
         Begin VB.TextBox txtTomail 
            Height          =   360
            Index           =   1
            Left            =   3120
            TabIndex        =   28
            Top             =   2280
            Width           =   3015
         End
         Begin VB.TextBox txtTomail 
            Height          =   360
            Index           =   0
            Left            =   3120
            TabIndex        =   26
            Text            =   "ABC@GMAIL"
            Top             =   1800
            Width           =   3015
         End
         Begin VB.TextBox txtfromPassword 
            Height          =   360
            Left            =   3120
            TabIndex        =   22
            Text            =   "PASSWORD"
            Top             =   1320
            Width           =   3015
         End
         Begin VB.TextBox txtfromMail 
            Height          =   360
            Left            =   3120
            TabIndex        =   21
            Text            =   "mAIL"
            Top             =   840
            Width           =   3015
         End
         Begin VB.TextBox txtapilink 
            Height          =   360
            Left            =   3120
            TabIndex        =   20
            Text            =   "www.MindaRikaSendEmail.com"
            Top             =   360
            Width           =   3015
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Bypass"
            Height          =   255
            Left            =   1320
            TabIndex        =   19
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "TO MAIL 7"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   17
            Left            =   120
            TabIndex        =   39
            Top             =   4680
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "TO MAIL 6"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   16
            Left            =   120
            TabIndex        =   37
            Top             =   4200
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "TO MAIL 5"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   15
            Left            =   120
            TabIndex        =   35
            Top             =   3720
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "TO MAIL 2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   14
            Left            =   120
            TabIndex        =   33
            Top             =   2280
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "TO MAIL 3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   13
            Left            =   120
            TabIndex        =   32
            Top             =   2760
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "TO MAIL 4"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   12
            Left            =   120
            TabIndex        =   31
            Top             =   3240
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "TO MAIL 1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   11
            Left            =   120
            TabIndex        =   27
            Top             =   1800
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "Web Api Link"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   10
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "FROM MAIL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   9
            Left            =   120
            TabIndex        =   24
            Top             =   840
            Width           =   2895
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "FROM PASSWORD"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   7
            Left            =   120
            TabIndex        =   23
            Top             =   1320
            Width           =   2895
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1815
         Left            =   120
         TabIndex        =   10
         Top             =   4560
         Width           =   5775
         Begin VB.CheckBox ChkComPortBP2 
            BackColor       =   &H00808080&
            Caption         =   "ByPass"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   360
            Left            =   2040
            TabIndex        =   16
            Top             =   1320
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.ComboBox CboCom2 
            Height          =   360
            Left            =   3480
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1320
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txtPLC_Port 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2520
            TabIndex        =   13
            Text            =   "1234"
            Top             =   840
            Width           =   2655
         End
         Begin VB.TextBox txtPLC_IP 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2520
            TabIndex        =   12
            Text            =   "192.168.1.12"
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "Melexis ComPort"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   360
            Index           =   6
            Left            =   120
            TabIndex        =   17
            Top             =   1320
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "PLC Port"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   360
            Index           =   4
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   2295
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "PLC IP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   360
            Index           =   3
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Other"
         Height          =   975
         Left            =   6000
         TabIndex        =   7
         Top             =   5400
         Width           =   6255
         Begin VB.TextBox txtPrinterName1 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   2400
            TabIndex        =   8
            Text            =   "ZDesigner ZD230-203dpi ZPL"
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            Caption         =   "Printer Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   360
            Index           =   5
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   1485
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000008&
         BorderStyle     =   0  'None
         Height          =   1440
         Left            =   120
         TabIndex        =   5
         Top             =   6480
         Width           =   12135
         Begin VB.CommandButton cmdSave 
            BackColor       =   &H8000000B&
            Caption         =   "Save"
            Height          =   1005
            Left            =   4080
            Picture         =   "frmExtraSet.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   83
            ToolTipText     =   "Close Screen"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdClose 
            BackColor       =   &H8000000B&
            Height          =   1005
            Left            =   7320
            Picture         =   "frmExtraSet.frx":066A
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Close Screen"
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Time Shift"
         Height          =   4335
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6495
         Begin VB.CheckBox chkBreakEnable 
            BackColor       =   &H00808080&
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   1
            Left            =   1440
            TabIndex        =   88
            Top             =   2280
            Width           =   210
         End
         Begin VB.CheckBox chkBreakEnable 
            BackColor       =   &H00808080&
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   2
            Left            =   1440
            TabIndex        =   87
            Top             =   2760
            Width           =   255
         End
         Begin VB.CheckBox chkBreakEnable 
            BackColor       =   &H00808080&
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   3
            Left            =   1440
            TabIndex        =   86
            Top             =   3240
            Width           =   255
         End
         Begin VB.CheckBox chkBreakEnable 
            BackColor       =   &H00808080&
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   4
            Left            =   1440
            TabIndex        =   85
            Top             =   3720
            Width           =   255
         End
         Begin VB.CheckBox chkBreakEnable 
            BackColor       =   &H00808080&
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Index           =   0
            Left            =   1440
            TabIndex        =   84
            Top             =   1800
            Width           =   255
         End
         Begin MSComCtl2.DTPicker DTShift1Start 
            Height          =   375
            Left            =   2400
            TabIndex        =   46
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "HH:MM"
            Format          =   110231554
            CurrentDate     =   44970
         End
         Begin MSComCtl2.DTPicker DTShift1End 
            Height          =   375
            Left            =   4800
            TabIndex        =   47
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "HH:MM"
            Format          =   110231554
            CurrentDate     =   44970
         End
         Begin MSComCtl2.DTPicker DTShift2Start 
            Height          =   375
            Left            =   2400
            TabIndex        =   48
            Top             =   840
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "HH:MM"
            Format          =   110231554
            CurrentDate     =   44970
         End
         Begin MSComCtl2.DTPicker DTShift2End 
            Height          =   375
            Left            =   4800
            TabIndex        =   49
            Top             =   840
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "HH:MM"
            Format          =   110231554
            CurrentDate     =   44970
         End
         Begin MSComCtl2.DTPicker DTShift3Start 
            Height          =   375
            Left            =   2400
            TabIndex        =   50
            Top             =   1320
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "HH:MM"
            Format          =   110231554
            CurrentDate     =   44970
         End
         Begin MSComCtl2.DTPicker DTShift3End 
            Height          =   375
            Left            =   4800
            TabIndex        =   51
            Top             =   1320
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "HH:MM"
            Format          =   110231554
            CurrentDate     =   44970
         End
         Begin MSComCtl2.DTPicker DTBreak1Start 
            Height          =   375
            Left            =   2400
            TabIndex        =   53
            Top             =   1800
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "HH:MM"
            Format          =   110231554
            CurrentDate     =   44970
         End
         Begin MSComCtl2.DTPicker DTBreak1End 
            Height          =   375
            Left            =   4800
            TabIndex        =   54
            Top             =   1800
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "HH:MM"
            Format          =   110231554
            CurrentDate     =   44970
         End
         Begin MSComCtl2.DTPicker DTBreak2Start 
            Height          =   375
            Left            =   2400
            TabIndex        =   56
            Top             =   2280
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "HH:MM"
            Format          =   110231554
            CurrentDate     =   44970
         End
         Begin MSComCtl2.DTPicker DTBreak2End 
            Height          =   375
            Left            =   4800
            TabIndex        =   57
            Top             =   2280
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "HH:MM"
            Format          =   110231554
            CurrentDate     =   44970
         End
         Begin MSComCtl2.DTPicker DTBreak3Start 
            Height          =   375
            Left            =   2400
            TabIndex        =   59
            Top             =   2760
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "HH:MM"
            Format          =   110231554
            CurrentDate     =   44970
         End
         Begin MSComCtl2.DTPicker DTBreak3End 
            Height          =   375
            Left            =   4800
            TabIndex        =   60
            Top             =   2760
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "HH:MM"
            Format          =   110231554
            CurrentDate     =   44970
         End
         Begin MSComCtl2.DTPicker DTBreak4Start 
            Height          =   375
            Left            =   2400
            TabIndex        =   62
            Top             =   3240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "HH:MM"
            Format          =   110231554
            CurrentDate     =   44970
         End
         Begin MSComCtl2.DTPicker DTBreak4End 
            Height          =   375
            Left            =   4800
            TabIndex        =   63
            Top             =   3240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "HH:MM"
            Format          =   110231554
            CurrentDate     =   44970
         End
         Begin MSComCtl2.DTPicker DTBreak5Start 
            Height          =   375
            Left            =   2400
            TabIndex        =   65
            Top             =   3720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "HH:MM"
            Format          =   110231554
            CurrentDate     =   44970
         End
         Begin MSComCtl2.DTPicker DTBreak5End 
            Height          =   375
            Left            =   4800
            TabIndex        =   66
            Top             =   3720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "HH:MM"
            Format          =   110231554
            CurrentDate     =   44970
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "End"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   36
            Left            =   4200
            TabIndex        =   82
            Top             =   3720
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "End"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   35
            Left            =   4200
            TabIndex        =   81
            Top             =   3240
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "End"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   34
            Left            =   4200
            TabIndex        =   80
            Top             =   2760
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "End"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   33
            Left            =   4200
            TabIndex        =   79
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "End"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   32
            Left            =   4200
            TabIndex        =   78
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "End"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   31
            Left            =   4200
            TabIndex        =   77
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "End"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   30
            Left            =   4200
            TabIndex        =   76
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "End"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   29
            Left            =   4200
            TabIndex        =   75
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   28
            Left            =   1800
            TabIndex        =   74
            Top             =   3720
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   27
            Left            =   1800
            TabIndex        =   73
            Top             =   3240
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   26
            Left            =   1800
            TabIndex        =   72
            Top             =   2760
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   25
            Left            =   1800
            TabIndex        =   71
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   24
            Left            =   1800
            TabIndex        =   70
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   23
            Left            =   1800
            TabIndex        =   69
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   22
            Left            =   1800
            TabIndex        =   68
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   21
            Left            =   1800
            TabIndex        =   67
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "BreakTime 5"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   20
            Left            =   120
            TabIndex        =   64
            Top             =   3720
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "BreakTime 4"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   19
            Left            =   120
            TabIndex        =   61
            Top             =   3240
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "BreakTime 3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   18
            Left            =   120
            TabIndex        =   58
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "BreakTime 2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   8
            Left            =   120
            TabIndex        =   55
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "BreakTime 1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   52
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "Shift 3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   4
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "Shift 2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackColor       =   &H00808080&
            Caption         =   "Shift 1 "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   375
            Index           =   57
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frmExtraSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Enum BasicAction
    Load = 1
    Save = 2
End Enum

Private Sub cmdClose_Click()
    frmmenu.Show
    Unload Me
End Sub

Private Sub CmdSave_Click()
On Error GoTo Error
'Dim Sql As String
'Dim Rs As ADODB.Recordset

    'If ChkShift4.Value = 0 Then
    'If getShiftValid(txtShift1, txtShift2, txtShift3) = False Then Exit Sub
    'End If
    
    'If ValidLen(1, 3, txtMachineNo) = False Then Exit Sub
    

    'If ValidLen(1, 255, txtNetworkDB) = False Then Exit Sub

    ExtraSetting Save
    
    ResetForm

    MsgBox "Saved Successfully", vbInformation

    ExtraSetting Load
    
Exit Sub
Error:
MsgBox Error, vbInformation
End Sub

''Private Sub Command1_Click()
''    txtMachineNo = getShift
''
''    'PrintLabel
''End Sub

Private Sub Form_Load()
On Error GoTo Error

'Advance
Me.WindowState = 2
Me.BackColor = &H80000010
Picture1.BorderStyle = 1
Picture1.Appearance = 0
Picture1.BackColor = vbButtonFace
Picture1.Left = (Screen.Width - Picture1.Width) / 2
Picture1.Top = (Screen.Height - Picture1.Height) / 2 - 400

    'My cboCom Load
    
    ExtraSetting Load


Exit Sub
Error:
MsgBox Error, vbInformation
End Sub

Private Sub ExtraSetting(Action As BasicAction)
On Error GoTo Error
Dim Sql As String
Dim rs As ADODB.Recordset

    Sql = "Select * from Common_Set where SetType ='CommonSet'" 'SetType = Settings Type
    Set rs = New ADODB.Recordset
    rs.Open Sql, Con, adOpenDynamic, adLockOptimistic

    If Action = 1 Then
        txtPrinterName1 = rs("PrinterName1")
        txtPLC_IP = rs("PLC_IP")
        txtPLC_Port = rs("PLC_Port")
        txtapilink.Text = rs("WebApiLink")
        txtfromMail.Text = rs("SenderEmail")
        txtfromPassword.Text = rs("SenderPassword")
        txtTomail(0).Text = rs("ToEmail1")
        txtTomail(1).Text = rs("ToEmail2")
        txtTomail(2).Text = rs("ToEmail3")
        txtTomail(3).Text = rs("ToEmail4")
        txtTomail(4).Text = rs("ToEmail5")
        txtTomail(5).Text = rs("ToEmail6")
        txtTomail(6).Text = rs("ToEmail7")
        Check1.Value = Val(rs("EmailBypass"))
        For i = 0 To 5
         chkTomailbypass(i).Value = Val(rs("EmailBypass" & i + 1))
        Next
        
        chkBreakEnable(0).Value = Val(rs("Break1Enable"))
        chkBreakEnable(1).Value = Val(rs("Break2Enable"))
        chkBreakEnable(2).Value = Val(rs("Break3Enable"))
        chkBreakEnable(3).Value = Val(rs("Break4Enable"))
        chkBreakEnable(4).Value = Val(rs("Break5Enable"))
        DTBreak1Start.Value = Format(rs("Break1Start"), "HH:MM")
        DTBreak1End.Value = TimeValue(rs("Break1End"))
        DTBreak2Start.Value = TimeValue(rs("Break2Start"))
        DTBreak2End.Value = TimeValue(rs("Break2End"))
        DTBreak3Start.Value = TimeValue(rs("Break3Start"))
        DTBreak3End.Value = TimeValue(rs("Break3End"))
        DTBreak4Start.Value = TimeValue(rs("Break4Start"))
        DTBreak4End.Value = TimeValue(rs("Break4End"))
        DTBreak5Start.Value = TimeValue(rs("Break5Start"))
        DTBreak5End.Value = TimeValue(rs("Break5End"))
        DTShift1Start.Value = TimeValue(rs("Shift1Start"))
        DTShift1End.Value = TimeValue(rs("Shift1End"))
        DTShift2Start.Value = TimeValue(rs("Shift2Start"))
        DTShift2End.Value = TimeValue(rs("Shift2End"))
        DTShift3Start.Value = TimeValue(rs("Shift3Start"))
        DTShift3End.Value = TimeValue(rs("Shift3End"))
    ElseIf Action = 2 Then
        rs("PrinterName1") = Trim$(txtPrinterName1)
        rs("PLC_IP") = Trim$(txtPLC_IP.Text)
        rs("PLC_Port") = Trim$(txtPLC_Port.Text)
        rs("EmailBypass") = Check1.Value
        rs("WebApiLink") = txtapilink.Text
        rs("SenderEmail") = txtfromMail.Text
        rs("SenderPassword") = txtfromPassword.Text
        rs("ToEmail1") = txtTomail(0).Text
        rs("ToEmail2") = txtTomail(1).Text
        rs("ToEmail3") = txtTomail(2).Text
        rs("ToEmail4") = txtTomail(3).Text
        rs("ToEmail5") = txtTomail(4).Text
        rs("ToEmail6") = txtTomail(5).Text
        rs("ToEmail7") = txtTomail(6).Text
        For i = 0 To 5
           rs("EmailBypass" & i + 1) = chkTomailbypass(i).Value
        Next
        rs("Break1Enable") = chkBreakEnable(0).Value
        rs("Break2Enable") = chkBreakEnable(1).Value
        rs("Break3Enable") = chkBreakEnable(2).Value
        rs("Break4Enable") = chkBreakEnable(3).Value
        rs("Break5Enable") = chkBreakEnable(4).Value
        rs("Break1Start") = DTBreak1Start.Value
        rs("Break1End") = DTBreak1End.Value
        rs("Break2Start") = DTBreak2Start.Value
        rs("Break2End") = DTBreak2End.Value
        rs("Break3Start") = DTBreak3Start.Value
        rs("Break3End") = DTBreak3End.Value
        rs("Break4Start") = DTBreak4Start.Value
        rs("Break4End") = DTBreak4End.Value
        rs("Break5Start") = DTBreak5Start.Value
        rs("Break5End") = DTBreak5End.Value
        rs("Shift1Start") = DTShift1Start.Value
        rs("Shift1End") = DTShift1End.Value
        rs("Shift2Start") = DTShift2Start.Value
        rs("Shift2End") = DTShift2End.Value
        rs("Shift3Start") = DTShift3Start.Value
        rs("Shift3End") = DTShift3End.Value
        rs.Update
    End If

Exit Sub
Error:
MsgBox "Error in ComPort Setting" & vbNewLine & Error, vbInformation
End Sub

Private Function ValidLen(ByVal Min, Max As Long, Text As TextBox) As Boolean

    If Trim(Text) = "" Or (Len(Text) < Min Or Len(Text) > Max) Then
        MsgBox ("Kindly Enter Between " & Min & " To " & Max & " Characters"), vbCritical
        Text.SetFocus
        Text.SelStart = 0
        Text.SelLength = Len(Text)
        Text.BackColor = vbRed
        ValidLen = False
    Else
        Text.BackColor = vbWhite
        ValidLen = True
    End If

End Function

Private Function ValidEntry(ByVal Min, Max As Double, Text As TextBox) As Boolean

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

Private Function getShiftValid(sTime1, sTime2, sTime3 As String) As Boolean
On Error GoTo Error
Dim AmPm(2) As String
Dim TmpLan(2) As String
'Const MaxTime = "11:59 PM"

    TmpLan(0) = Len(sTime1)
    TmpLan(1) = Len(sTime2)
    TmpLan(2) = Len(sTime3)
    AmPm(0) = Right(sTime1, 2)
    AmPm(1) = Right(sTime2, 2)
    AmPm(2) = Right(sTime3, 2)

    For i = 0 To 2

        If (TmpLan(i) > "8" Or TmpLan(i) < "8") Then
            MsgBox "Please Enter a Valid Time", vbInformation
            Exit Function

        End If

        If (AmPm(i) = "AM" Or AmPm(i) = "PM") Then
        Else
            MsgBox "Enter a Valid Time With AM/PM", vbInformation
            Exit Function

        End If

    Next

    If TimeValue(sTime2) < TimeValue(sTime1) Then
        MsgBox "Shift 2 Time Should Be Greater Then Shift 1", vbInformation
        Exit Function

    End If

    If TimeValue(sTime3) < TimeValue(sTime2) Then
        MsgBox "Shift 3 Time Should Be Greater Then Shift 2", vbInformation
        Exit Function

    End If

    getShiftValid = True

Exit Function
Error:
MsgBox "Error Found in Time Shift Validation" & vbNewLine & Error, vbInformation
End Function

Private Sub ResetForm()
Dim txt As Control

For Each txt In Me
    If TypeOf txt Is TextBox Then
        txt.Text = ""
    End If
Next

For Each txt In Me
    If TypeOf txt Is CheckBox Then
        txt.Value = 0
    End If
Next

End Sub

