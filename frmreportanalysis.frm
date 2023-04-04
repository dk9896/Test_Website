VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{84E5CF37-E467-4AC2-89C4-C6002FFB5055}#25.1#0"; "ChartViewer.ocx"
Begin VB.Form frmreportanalysis 
   Caption         =   "Report Analysis"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   8775
      Left            =   120
      ScaleHeight     =   8715
      ScaleWidth      =   19155
      TabIndex        =   0
      Top             =   120
      Width           =   19215
      Begin CDChartViewer.ChartViewer CD1 
         Height          =   5415
         Left            =   10080
         Top             =   1200
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   9551
      End
      Begin VB.Frame Frame4 
         Height          =   975
         Left            =   240
         TabIndex        =   47
         Top             =   1320
         Width           =   8415
         Begin VB.OptionButton opt4 
            Caption         =   "ALL Day"
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
            Left            =   6480
            TabIndex        =   51
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton opt3 
            Caption         =   "Shift C"
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
            Left            =   4320
            TabIndex        =   50
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton opt2 
            Caption         =   "Shift B"
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
            Left            =   2280
            TabIndex        =   49
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton opt1 
            Caption         =   "Shift A"
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
            Left            =   240
            TabIndex        =   48
            Top             =   480
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.CommandButton CmdClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   17640
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmreportanalysis.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   7080
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
      Begin VB.Frame Frame3 
         Height          =   1575
         Left            =   240
         TabIndex        =   39
         Top             =   6960
         Width           =   8415
         Begin VB.TextBox txtTargetSpeed 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6240
            Locked          =   -1  'True
            TabIndex        =   41
            Text            =   "0"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtLineSpeed 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   40
            Text            =   "0"
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "units/hr"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   24
            Left            =   7320
            TabIndex        =   45
            Top             =   720
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "units/hr"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   17
            Left            =   2880
            TabIndex        =   44
            Top             =   720
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Target Speed"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   14
            Left            =   4680
            TabIndex        =   43
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Line Speed"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   12
            Left            =   240
            TabIndex        =   42
            Top             =   720
            Width           =   1200
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2175
         Left            =   240
         TabIndex        =   26
         Top             =   4800
         Width           =   8415
         Begin VB.TextBox txtAF 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   30
            Text            =   "0"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtPF 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   29
            Text            =   "0"
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtOEE 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3720
            Locked          =   -1  'True
            TabIndex        =   28
            Text            =   "0"
            Top             =   1560
            Width           =   975
         End
         Begin VB.TextBox txtQF 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   27
            Text            =   "0"
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "AF"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   23
            Left            =   240
            TabIndex        =   38
            Top             =   720
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "PF"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   22
            Left            =   3240
            TabIndex        =   37
            Top             =   720
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "QF"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   21
            Left            =   6000
            TabIndex        =   36
            Top             =   720
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   20
            Left            =   2040
            TabIndex        =   35
            Top             =   720
            Width           =   210
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   19
            Left            =   4680
            TabIndex        =   34
            Top             =   720
            Width           =   210
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   18
            Left            =   4800
            TabIndex        =   33
            Top             =   1680
            Width           =   210
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "OEE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   15
            Left            =   3000
            TabIndex        =   32
            Top             =   1680
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   13
            Left            =   7800
            TabIndex        =   31
            Top             =   720
            Width           =   210
         End
      End
      Begin VB.ComboBox cbomodelname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         TabIndex        =   24
         Text            =   "ALL"
         Top             =   480
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Frame Frame1 
         Height          =   2535
         Left            =   240
         TabIndex        =   3
         Top             =   2280
         Width           =   8415
         Begin VB.TextBox txtTotalTarget 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6240
            Locked          =   -1  'True
            TabIndex        =   14
            Text            =   "0"
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtTotalOK 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6240
            Locked          =   -1  'True
            TabIndex        =   13
            Text            =   "0"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtshiftCOK 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   12
            Text            =   "0"
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox txtshiftbOK 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   11
            Text            =   "0"
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox txtshiftAok 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   10
            Text            =   "0"
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox txtTotalShiftC 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   9
            Text            =   "0"
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox txtTotalShiftB 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   8
            Text            =   "0"
            Top             =   1320
            Width           =   975
         End
         Begin VB.TextBox txtTotalshiftA 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   7
            Text            =   "0"
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "nos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   11
            Left            =   7320
            TabIndex        =   23
            Top             =   1800
            Width           =   390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "nos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   10
            Left            =   7200
            TabIndex        =   22
            Top             =   1080
            Width           =   390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total Counts"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   9
            Left            =   4680
            TabIndex        =   21
            Top             =   1800
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total OK Parts"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   8
            Left            =   4680
            TabIndex        =   20
            Top             =   1080
            Width           =   1515
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "OK Parts"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   7
            Left            =   3120
            TabIndex        =   19
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total Parts"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   6
            Left            =   1080
            TabIndex        =   18
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "nos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   2160
            TabIndex        =   17
            Top             =   2040
            Width           =   390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "nos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   2160
            TabIndex        =   16
            Top             =   1440
            Width           =   390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "nos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   2160
            TabIndex        =   15
            Top             =   840
            Width           =   390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Shift C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   240
            TabIndex        =   6
            Top             =   2040
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Shift B"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   240
            TabIndex        =   5
            Top             =   1440
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Shift A"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   4
            Top             =   840
            Width           =   675
         End
      End
      Begin MSComCtl2.DTPicker DTFrom 
         Height          =   405
         Left            =   6120
         TabIndex        =   25
         Top             =   480
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   -2147483646
         CalendarForeColor=   16711680
         CalendarTitleForeColor=   49344
         Format          =   112066561
         CurrentDate     =   39022
      End
      Begin VB.Label Label4 
         Caption         =   "Model "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmreportanalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TotalShiftTime As Double
Dim TotalBreaktime As Double
Dim TotalBreakdownTime As Double
Private Sub cmdClose_Click()
  frmmenu.Show
  Unload Me
End Sub
Public Sub createChart(viewer As Object)

    Dim cd As New ChartDirector.API

    ' The data for the bar chart
    Dim data()
    data = Array(Val(txtAF.Text), Val(txtPF.Text), Val(txtQF.Text), Val(txtOEE.Text))

    ' The labels for the bar chart
    Dim labels()
    labels = Array("AF", "PF", "QF", "OEE")

    ' Create a XYChart object of size 600 x 360 pixels
    Dim c As XYChart
    Set c = cd.XYChart(600, 360)

    ' Set the plotarea at (70, 20) and of size 500 x 300 pixels, with transparent background and
    ' border and light grey (0xcccccc) horizontal grid lines
    Call c.setPlotArea(70, 20, 500, 300, cd.Transparent, -1, cd.Transparent, &HCCCCCC)

    ' Set the x and y axis stems to transparent and the label font to 12pt Arial
    Call c.xAxis().setColors(cd.Transparent)
    Call c.yAxis().setColors(cd.Transparent)
    Call c.xAxis().setLabelStyle("arial.ttf", 12)
    Call c.yAxis().setLabelStyle("arial.ttf", 12)

    ' Add a blue (0x6699bb) bar chart layer using the given data
    Dim layer As BarLayer
    Set layer = c.addBarLayer(data, &H6699BB)

    ' Use bar gradient lighting with the light intensity from 0.8 to 1.3
    Call layer.setBorderColor(cd.Transparent, cd.barLighting(0.8, 1.3))

    ' Set rounded corners for bars
    'Call layer.setRoundedCorners

    ' Display labela on top of bars using 12pt Arial font
    Call layer.setAggregateLabelStyle("Arial", 12)

    ' Set the labels on the x axis.
    Call c.xAxis().setLabels(labels)

    ' For the automatic y-axis labels, set the minimum spacing to 40 pixels.
    Call c.yAxis().setTickDensity(40)

    ' Add a title to the y axis using dark grey (0x555555) 14pt Arial Bold font
    Call c.yAxis().setTitle("Calculated Value", "arialbd.ttf", 14, &H555555)

    ' Output the chart
    Set viewer.Picture = c.makePicture()

    'include tool tip for the chart
    viewer.ImageMap = c.getHTMLImageMap("clickable", "", "title='{xLabel}: ${value}M'")

End Sub

Private Sub ShiftCalculation()
    'Dim starttime As Date
    'Dim endtime As Date
    Dim rs As ADODB.Recordset
    'Dim TotalShiftTime As Double
    'Dim TotalBreaktime As Double
    Dim ShiftTimeA As Double
    Dim ShiftTimeB As Double
    Dim ShiftTimeC As Double
    Dim BreakTimeA As Double
    Dim BreakTimeB As Double
    Dim BreakTimeC As Double
    
    Sql = "Select * from Common_Set where SetType ='CommonSet'" 'SetType = Settings Type
    Set rs = New ADODB.Recordset
    rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    txtTargetSpeed.Text = rs("cycletime")
    ShiftTimeA = DateDiff("n", rs("Shift1Start"), rs("Shift1End"))
    ShiftTimeB = DateDiff("n", rs("Shift2Start"), rs("Shift2End"))
    ShiftTimeC = DateDiff("n", rs("Shift3Start"), rs("Shift3End"))
    If ShiftTimeA < 0 Then
        ShiftTimeA = 1440 + ShiftTimeA
    End If
    If ShiftTimeB < 0 Then
        ShiftTimeB = 1440 + ShiftTimeB
    End If
    
    If ShiftTimeC < 0 Then
        ShiftTimeC = 1440 + ShiftTimeC
    End If
    '- TimeValue(rs("Shift1Start"))
    'ShiftTimeB = TimeValue(rs("Shift2End")) - TimeValue(rs("Shift2Start"))
    'ShiftTimeC = TimeValue(rs("Shift3End")) - TimeValue(rs("Shift3Start"))
    For i = 0 To 4
      If Val(rs("Break" & i + 1 & "Enable")) = 1 Then
        If TimeValue(rs("Shift1Start")) <= TimeValue(rs("Break" & i + 1 & "Start")) And TimeValue(rs("Shift1End")) >= TimeValue(rs("Break" & i + 1 & "Start")) Then
            If TimeValue(rs("Shift1End")) >= TimeValue(rs("Break" & i + 1 & "End")) Then
                BreakTimeA = BreakTimeA + DateDiff("n", rs("Break" & i + 1 & "Start"), rs("Break" & i + 1 & "End"))
            Else
                BreakTimeA = BreakTimeA + DateDiff("n", rs("Break" & i + 1 & "Start"), rs("Shift1End"))
            End If
        ElseIf TimeValue(rs("Shift2Start")) <= TimeValue(rs("Break" & i + 1 & "Start")) And TimeValue(rs("Shift2End")) >= TimeValue(rs("Break" & i + 1 & "Start")) Then
            If TimeValue(rs("Shift2End")) >= TimeValue(rs("Break" & i + 1 & "End")) Then
                BreakTimeB = BreakTimeB + DateDiff("n", rs("Break" & i + 1 & "Start"), rs("Break" & i + 1 & "End"))
            Else
                BreakTimeB = BreakTimeB + DateDiff("n", rs("Break" & i + 1 & "Start"), rs("Shift2End"))
            End If
        Else 'If TimeValue(rs("Shift3Start")) <= TimeValue(rs("Break" & i + 1 & "Start")) And TimeValue(rs("Shift3End")) >= TimeValue(rs("Break" & i + 1 & "Start")) Then
            If TimeValue(rs("Shift3End")) >= TimeValue(rs("Break" & i + 1 & "End")) Then
                BreakTimeC = BreakTimeC + DateDiff("n", rs("Break" & i + 1 & "Start"), rs("Break" & i + 1 & "End"))
            Else
                BreakTimeC = BreakTimeC + DateDiff("n", rs("Break" & i + 1 & "Start"), rs("Shift3End"))
            End If
        End If
      End If
    Next
    Dim endtime As Date
    
    If opt4.Value = True Then
     TotalShiftTime = ShiftTimeA + ShiftTimeB + ShiftTimeC
     TotalBreaktime = BreakTimeA + BreakTimeB + BreakTimeC
     TotalBreakdownTime = BreakdownTimeCalculation(Format(DTFrom.Value, "dd-mm-yyyy") & " " & Format(rs("Shift1Start"), "HH:MM"), DateAdd("d", 1, Format(DTFrom.Value, "dd-mm-yyyy")) & " " & Format(rs("Shift1Start"), "HH:MM"))
    ElseIf opt1.Value = True Then
     TotalShiftTime = ShiftTimeA
     TotalBreaktime = BreakTimeA
     endtime = Format(DTFrom.Value, "dd-mm-yyyy") & " " & Format(rs("shift1end"), "HH:MM")
     If TimeValue(rs("Shift1Start")) >= TimeValue(rs("shift1end")) Then
        endtime = DateAdd("d", 1, Format(DTFrom.Value, "dd-mm-yyyy")) & " " & Format(rs("Shift1end"), "HH:MM")
     End If
     TotalBreakdownTime = BreakdownTimeCalculation(Format(DTFrom.Value, "dd-mm-yyyy") & " " & Format(rs("Shift1Start"), "HH:MM"), endtime)
    ElseIf opt2.Value = True Then
     TotalShiftTime = ShiftTimeB
     TotalBreaktime = BreakTimeB
     endtime = Format(rs("shift2end"), "HH:MM")
     If TimeValue(rs("Shift2Start")) >= TimeValue(rs("shift2end")) Then
        endtime = DateAdd("d", 1, Format(DTFrom.Value, "dd-mm-yyyy")) & " " & Format(rs("Shift2end"), "HH:MM")
     End If
     TotalBreakdownTime = BreakdownTimeCalculation(Format(DTFrom.Value, "dd-mm-yyyy") & " " & Format(rs("Shift2Start"), "HH:MM"), endtime)
    ElseIf opt3.Value = True Then
     TotalShiftTime = ShiftTimeC
     TotalBreaktime = BreakTimeC
     endtime = Format(rs("shift3end"), "HH:MM")
     If TimeValue(rs("Shift3Start")) >= TimeValue(rs("shift3end")) Then
        endtime = DateAdd("d", 1, Format(DTFrom.Value, "dd-mm-yyyy")) & " " & Format(rs("Shift3end"), "HH:MM")
     End If
     TotalBreakdownTime = BreakdownTimeCalculation(Format(DTFrom.Value, "dd-mm-yyyy") & " " & Format(rs("Shift3Start"), "HH:MM"), endtime)
    End If
    
End Sub



Private Sub DTFrom_Change()
Calculation
End Sub

Private Sub Form_Load()
  Me.WindowState = 2
 Picture1.Left = (Screen.Width - Picture1.Width) / 2
 Picture1.Top = ((Screen.Height - Picture1.Height) / 2) + 100
DTFrom.Value = Format(Now, "dd/mm/yyyy")

'A = BreakdownTimeCalculation("08-07-2022 02:30:00 PM", "08-07-2022 10:00:00 PM")
Calculation
End Sub
Private Sub Calculation()
    'Dim TotalShiftTime As Double
    'Dim TotalBreakdownTime As Double
    'Dim TotalBreaktime As Double
    Dim TotalProduction As Double
    Dim TotalOkParts As Double
    txtAF.Text = 0
    txtQF.Text = 0
    txtPF.Text = 0
    txtOEE.Text = 0
    ShiftCalculation
    Productioncalculation
    TotalProduction = Val(txtTotalTarget.Text)
    TotalOkParts = Val(txtTotalOK.Text)
        
    Dim PlannedTime As Double
    PlannedTime = TotalShiftTime - TotalBreaktime

    Dim RunTime As Double
    RunTime = PlannedTime - TotalBreakdownTime
    If PlannedTime <> 0 Then
    txtAF.Text = Format(RunTime / PlannedTime, "0.0000")
    End If
    If TotalProduction <> 0 Then
    txtQF.Text = Format(TotalOkParts / TotalProduction, "0.0000")
    End If
    Dim IdealCycleTime As Double
    If Val(txtTargetSpeed.Text) = 0 Then
    IdealCycleTime = 0
    Else
    IdealCycleTime = 3600 / txtTargetSpeed.Text
    End If
    If RunTime <> 0 Then
    txtPF.Text = Format((IdealCycleTime * TotalProduction) / (RunTime * 60), "0.0000")
    End If
    
    txtOEE.Text = Format(Val(txtAF.Text) * Val(txtQF.Text) * Val(txtPF.Text), "0.0000")
    
    createChart CD1
End Sub
Private Sub Productioncalculation()
    txtTotalshiftA.Text = 0
    txtTotalShiftB.Text = 0
    txtTotalShiftC.Text = 0
    txtshiftAok.Text = 0
    txtshiftbOK.Text = 0
    txtshiftCOK.Text = 0
    txtTotalOK.Text = 0
    txtTotalTarget = 0
    Dim rs As ADODB.Recordset
    Dim okcount As Integer
    Dim ngcount As Integer
    Sql = "Select * from Model_report_counter where datetime = #" & Format(DTFrom.Value, "mm-dd-yyyy") & "# "
    If opt1.Value = True Then
        Sql = Sql & "and shifttime = '1'"
    ElseIf opt2.Value = True Then
        Sql = Sql & "and shifttime = '2'"
    
    ElseIf opt3.Value = True Then
        Sql = Sql & "and shifttime = '3'"
    
    ElseIf opt4.Value = True Then
    
    End If
    Set rs = New ADODB.Recordset
    rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    If rs.EOF = False Then
        Do While rs.EOF = False
            If rs("Shifttime") = "1" Then
                txtTotalshiftA.Text = Val(txtTotalshiftA) + Val(rs("OKcounter")) + Val(rs("ngcounter"))
                txtshiftAok.Text = Val(txtshiftAok) + Val(rs("okcounter"))
            ElseIf rs("Shifttime") = "2" Then
                txtTotalShiftB.Text = Val(txtTotalShiftB) + Val(rs("OKcounter")) + Val(rs("ngcounter"))
                txtshiftbOK.Text = Val(txtshiftbOK) + Val(rs("okcounter"))
            ElseIf rs("Shifttime") = "3" Then
                txtTotalShiftC.Text = Val(txtTotalShiftC) + Val(rs("OKcounter")) + Val(rs("ngcounter"))
                txtshiftCOK.Text = Val(txtshiftCOK) + Val(rs("okcounter"))
            End If
            rs.MoveNext
        Loop
        txtTotalOK.Text = Val(txtshiftAok.Text) + Val(txtshiftbOK.Text) + Val(txtshiftCOK.Text)
        txtTotalTarget.Text = Val(txtTotalshiftA.Text) + Val(txtTotalShiftB.Text) + Val(txtTotalShiftC.Text)
    End If
    
End Sub
Private Function BreakdownTimeCalculation(starttime As Date, endtime As Date) As Double
Dim rs As ADODB.Recordset
    Sql = "Select * from Model_Report_breakdown where (starttime BETWEEN  #" & Format(starttime, "mm-dd-yyyy HH:MM:SS") & "# And #" & Format(endtime, "mm-dd-yyyy HH:MM:SS") & "# ) OR (endtime BETWEEN  #" & Format(starttime, "mm-dd-yyyy HH:MM:SS") & "# And #" & Format(endtime, "mm-dd-yyyy HH:MM:SS") & "# )"   ' where SetType ='CommonSet'" 'SetType = Settings Type
    Set rs = New ADODB.Recordset
    rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    If rs.EOF = False Then
    rs.MoveFirst
        Do While rs.EOF = False
         If rs("starttime") >= starttime And rs("endtime") <= endtime Then
            BreakdownTimeCalculation = BreakdownTimeCalculation + DateDiff("s", rs("starttime"), rs("endtime"))
         ElseIf rs("starttime") <= starttime And rs("endtime") <= endtime Then
            BreakdownTimeCalculation = BreakdownTimeCalculation + DateDiff("s", starttime, rs("endtime"))
         ElseIf rs("starttime") >= starttime And rs("endtime") >= endtime Then
            BreakdownTimeCalculation = BreakdownTimeCalculation + DateDiff("s", rs("starttime"), endtime)
         End If
        
        rs.MoveNext
        Loop
    Else
        BreakdownTimeCalculation = 0
    End If
    BreakdownTimeCalculation = BreakdownTimeCalculation / 60
End Function

Private Sub opt1_Click()
Calculation
End Sub

Private Sub opt2_Click()
Calculation
End Sub

Private Sub opt3_Click()
Calculation
End Sub

Private Sub opt4_Click()
Calculation
End Sub
