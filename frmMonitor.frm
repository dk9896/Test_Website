VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{97C0E9D8-AD04-4920-9B7A-4B99616579F9}#2.0#0"; "TextPrinter.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMonitor 
   Caption         =   "MI_7646_USB_Charger"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   15630
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8625
   ScaleWidth      =   15630
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   17895
      Begin VB.PictureBox PictureBreakdown 
         BackColor       =   &H80000010&
         Height          =   6015
         Left            =   4560
         ScaleHeight     =   5955
         ScaleWidth      =   8595
         TabIndex        =   43
         Top             =   2160
         Visible         =   0   'False
         Width           =   8655
         Begin VB.CommandButton cmdclosebreakdownscreen 
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
            Left            =   7200
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMonitor.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   4680
            UseMaskColor    =   -1  'True
            Width           =   1275
         End
         Begin VB.TextBox txtbreakdownsummary 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Left            =   2280
            TabIndex        =   47
            Top             =   4440
            Width           =   4575
         End
         Begin VB.CommandButton cmdgolive 
            BackColor       =   &H0000FF00&
            Caption         =   "Go Live"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   3240
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   2760
            Width           =   1935
         End
         Begin VB.CommandButton cmdfullbreakdown 
            BackColor       =   &H000000FF&
            Caption         =   "Full Breakdown"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   5280
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   840
            Width           =   1815
         End
         Begin VB.CommandButton cmdrunningbreakdown 
            BackColor       =   &H000080FF&
            Caption         =   "Running Breakdown"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   1560
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "BreakDown Summary"
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
            Left            =   120
            TabIndex        =   49
            Top             =   4800
            Width           =   2295
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   55
         Top             =   6840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   5760
         TabIndex        =   54
         Text            =   "Text3"
         Top             =   480
         Width           =   6135
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   615
         Left            =   4320
         TabIndex        =   53
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtproductioncounter 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2160
         TabIndex        =   52
         Top             =   7440
         Width           =   2490
      End
      Begin VB.TextBox txtTargetProduction 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2160
         TabIndex        =   51
         Top             =   6840
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Breakdown"
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
         Left            =   2520
         TabIndex        =   42
         Top             =   360
         Width           =   1575
      End
      Begin VB.PictureBox Picture1 
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1035
         ScaleWidth      =   2235
         TabIndex        =   41
         Top             =   240
         Width           =   2295
         Begin VB.Image Image1 
            Height          =   735
            Left            =   0
            Picture         =   "frmMonitor.frx":0C42
            Stretch         =   -1  'True
            Top             =   0
            Width           =   2295
         End
      End
      Begin VB.TextBox txtBarcode 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1320
         TabIndex        =   34
         Top             =   6240
         Width           =   4575
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   16440
         TabIndex        =   26
         Top             =   7320
         Width           =   1335
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
            Left            =   -120
            MaskColor       =   &H00FFFFFF&
            Picture         =   "frmMonitor.frx":37C8
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   1275
         End
      End
      Begin VB.Frame FrmResult 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   15360
         TabIndex        =   23
         Top             =   5400
         Width           =   2295
         Begin VB.Label lblGo 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   72
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   1425
            Left            =   0
            TabIndex        =   25
            Top             =   120
            Visible         =   0   'False
            Width           =   2265
         End
         Begin VB.Label lblNg 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NG"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   72
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1665
            Left            =   60
            TabIndex        =   24
            Top             =   120
            Visible         =   0   'False
            Width           =   2175
         End
      End
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   15240
         TabIndex        =   18
         Top             =   3480
         Width           =   2415
         Begin VB.TextBox txtNGCounter 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   32
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   1560
            Width           =   990
         End
         Begin VB.TextBox txtOKCounter 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   1080
            Width           =   990
         End
         Begin VB.TextBox txtBatchCounter 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   120
            Width           =   1005
         End
         Begin VB.TextBox txtCouplerCounter 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Text            =   "0"
            Top             =   600
            Width           =   990
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "NG Count"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "OK Count"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Coupler Count"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Batch Count"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
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
         Left            =   15480
         TabIndex        =   15
         Top             =   240
         Width           =   2415
         Begin VB.TextBox txtCycleTime 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   29
            Top             =   360
            Width           =   720
         End
         Begin VB.Label Label3 
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
            Height          =   255
            Left            =   2160
            TabIndex        =   35
            Top             =   480
            Width           =   375
         End
         Begin VB.Shape shapeInternet 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            Height          =   255
            Left            =   1320
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Cycle Time"
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
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Width           =   1575
         End
         Begin VB.Shape ShapePLCState 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            Height          =   255
            Left            =   1320
            Top             =   0
            Width           =   855
         End
         Begin VB.Label Label17 
            Caption         =   "PLC Comm"
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
            Left            =   120
            TabIndex        =   17
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00808080&
            BackStyle       =   0  'Transparent
            Caption         =   "Internet Con"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   840
            Width           =   1050
         End
      End
      Begin VB.TextBox txtCommandLine 
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
         Height          =   630
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   14
         Text            =   "frmMonitor.frx":440A
         Top             =   8280
         Width           =   17535
      End
      Begin VB.Frame Frame13 
         Caption         =   "Frame13"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   12240
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   3375
         Begin VB.Timer Timer7 
            Left            =   360
            Top             =   1320
         End
         Begin VB.Timer Timer13 
            Enabled         =   0   'False
            Left            =   0
            Top             =   0
         End
         Begin VB.Timer Timer12 
            Enabled         =   0   'False
            Left            =   0
            Top             =   0
         End
         Begin VB.Timer Timer11 
            Enabled         =   0   'False
            Left            =   0
            Top             =   0
         End
         Begin VB.Timer Timer6 
            Left            =   0
            Top             =   0
         End
         Begin VB.Timer Timer3 
            Left            =   840
            Top             =   960
         End
         Begin VB.Timer Timer1 
            Interval        =   1
            Left            =   120
            Top             =   960
         End
         Begin VB.Timer Timer2 
            Left            =   480
            Top             =   960
         End
         Begin VB.Timer Timer4 
            Left            =   1320
            Top             =   960
         End
         Begin VB.TextBox txtServoSpeedSet 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   360
            Width           =   1440
         End
         Begin VB.Timer Timer5 
            Left            =   2640
            Top             =   1080
         End
         Begin MSWinsockLib.Winsock WinSock1 
            Left            =   1920
            Top             =   960
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   393216
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
            TabIndex        =   13
            Top             =   240
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
         End
      End
      Begin VB.TextBox txtModelDesc 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1140
         Left            =   2400
         TabIndex        =   10
         Text            =   "MODEL DESC"
         Top             =   240
         Width           =   12975
      End
      Begin VB.Frame Frame10 
         Caption         =   "Frame10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   1
         Top             =   4680
         Visible         =   0   'False
         Width           =   5775
         Begin VB.Frame Frame5 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   5415
            Begin VB.TextBox txtIP 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   6
               Text            =   "127.0.0.1"
               Top             =   480
               Width           =   1095
            End
            Begin VB.TextBox txtPort 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2400
               Locked          =   -1  'True
               TabIndex        =   5
               Text            =   "1232"
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox txtIP_Host 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   4
               Text            =   "127.0.0.1"
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   3120
               Locked          =   -1  'True
               TabIndex        =   3
               Top             =   480
               Width           =   2175
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
               TabIndex        =   9
               Top             =   240
               Width           =   495
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
               TabIndex        =   8
               Top             =   240
               Width           =   495
            End
            Begin VB.Label Label15 
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
               TabIndex        =   7
               Top             =   240
               Width           =   615
            End
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid Grid1 
         Height          =   1845
         Left            =   240
         TabIndex        =   56
         Top             =   1920
         Width           =   8475
         _cx             =   14949
         _cy             =   3254
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
         FormatString    =   $"frmMonitor.frx":441C
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
      Begin VSFlex7Ctl.VSFlexGrid Grid2 
         Height          =   1845
         Left            =   240
         TabIndex        =   57
         Top             =   4080
         Width           =   8475
         _cx             =   14949
         _cy             =   3254
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
         FormatString    =   $"frmMonitor.frx":448B
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
      Begin VB.Image ImgPart 
         Height          =   1695
         Left            =   14520
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   3135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Production Counter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   43
         Left            =   240
         TabIndex        =   40
         Top             =   7440
         Width           =   1665
      End
      Begin VB.Label t 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Target Production"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   40
         Left            =   240
         TabIndex        =   39
         Top             =   6840
         Width           =   1530
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   240
         TabIndex        =   36
         Top             =   6240
         Width           =   720
      End
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   -1800
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   495
      Left            =   9480
      TabIndex        =   48
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ILLumination Curr. LH "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   34
      Left            =   7920
      TabIndex        =   38
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Barcode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   16
      Left            =   7920
      TabIndex        =   37
      Top             =   7320
      Width           =   375
   End
End
Attribute VB_Name = "frmMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim MsgCode As Integer
Dim Pulse As Boolean
Dim pulse1 As Boolean
Dim pulse2 As Boolean
Dim pulse3 As Boolean
Dim pulse4 As Boolean
Dim PulseScan As Boolean
Dim pulseBreakdown As Boolean
Dim PulseReset As Boolean
Dim pulsePrinterBypass As Boolean
Dim FSO As New FileSystemObject
Dim ExcelFileName As String
Dim Row As Long
Dim Col As Long
Dim setCouplerCounter As Integer
Dim setBatchCounter As Integer
'----------------
Dim PLC_Communication_Error As Boolean
Dim MsgText() As String
Dim MsgColor() As Integer
Dim MsgCount As Integer
Dim CloseScreen As Boolean
Dim runningreportdate As Date
Dim runningreportshift As String
Dim ModelNo As Integer
Private Declare Function InternetGetConnectedState Lib _
    "wininet" (ByRef dwflags As Long, ByVal dwReserved As _
    Long) As Long

Private Sub CmdClose_Click()
CloseScreen = True
CloseMe
End Sub

Private Sub CloseMe()

If MSComm1.PortOpen = True Then MSComm1.PortOpen = False

frmmenu.Show
Unload Me

End Sub

Private Sub CmdNgCounter_Click()
  If MsgBox("Are you Sure You Want To Reset NG Counter", vbInformation + vbYesNo) = vbYes Then
    txtNGCounter.Text = 0
    SaveCounterValue
  End If
End Sub

Private Sub CmdOKCounter_Click()
If MsgBox("Are you Sure You Want To Reset OK Counter", vbInformation + vbYesNo) = vbYes Then
    txtOKCounter.Text = 0
    SaveCounterValue
  End If
End Sub

Private Sub cmdclosebreakdownscreen_Click()
    PictureBreakdown.Visible = False
    Command2.Enabled = True
End Sub

Private Sub cmdfullbreakdown_Click()
    cmdrunningbreakdown.Enabled = False
    cmdfullbreakdown.Enabled = False
    cmdgolive.Enabled = True
    cmdclosebreakdownscreen.Enabled = False
    SaveBreakDown 3, 1
    PLcdata(348) = 3
End Sub

Private Sub cmdgolive_Click()
    cmdrunningbreakdown.Enabled = True
    cmdfullbreakdown.Enabled = True
    cmdgolive.Enabled = False
    cmdclosebreakdownscreen.Enabled = True
    SaveBreakDown 1, 0
    PLcdata(348) = 1
End Sub

Private Sub cmdrunningbreakdown_Click()
    cmdrunningbreakdown.Enabled = False
    cmdfullbreakdown.Enabled = False
    cmdgolive.Enabled = True
    cmdclosebreakdownscreen.Enabled = False
    SaveBreakDown 2, 1
    PLcdata(348) = 2
End Sub

Private Sub Command1_Click()
  If Val(txtTargetProduction.Text) > 0 Then
      Command1.Visible = False
      txtTargetProduction.Enabled = False
      txtTargetProduction.BackColor = vbWhite
      runningreportshift = getShift
      runningreportdate = TempReportDate
      SaveSetting App.Title, ModelName, "TargetProduction", txtTargetProduction.Text
      GetCounterValue
      PLcdata(349) = 0
  Else
    txtTargetProduction.BackColor = vbRed
  End If
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    PictureBreakdown.Visible = True
End Sub


Private Sub Command3_Click()
'PLcdata(109) = Val(Text3.Text)
'AssignPLCdata
sendEmail
End Sub

Private Sub Command7_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
If CloseScreen = False Then
    CloseMe
Else
    CloseScreen = False
End If
End Sub

Public Sub ConnectToPLC()
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
''On Error GoTo Error
Me.WindowState = 2
UserAccess
Frame1.Top = ((Screen.Height - Frame1.Height) / 2) - 100
Frame1.Left = ((Screen.Width - Frame1.Width) / 2)
LoadSettingsData
Call Load_Message_File
runningreportshift = GetSetting(App.Title, ModelName, "saveshift", 0)
runningreportdate = GetSetting(App.Title, ModelName, "savedate", 0)
PLcdata(340) = 1
GetCounterValue
ConnectToPLC
LoadGrid
Timer1.Enabled = True
Timer1.Interval = 1000
Timer2.Enabled = True
Timer2.Interval = 1000
Timer3.Interval = 500
Timer3.Enabled = True
'txtDate.Text = Date
'txttime.Text = Format(Time(), "hh:mm:ss")
'txtOperName.Text = LoginUser

Pulse = False
Exit Sub
End Sub

Private Sub UserAccess()
   If AccessType = "0" Then 'Disable or Hide For Operator
      'CmdOKCounter.Visible = False
      'CmdNgCounter.Visible = False
      'Command1.Visible = False
   ElseIf AccessType = "1" Then 'Disable or Hide for AccessType 1
      'CmdOKCounter.Visible = False
      'CmdNgCounter.Visible = False
      'Command1.Visible = False
   ElseIf AccessType = "2" Then 'Show All Which Will Disable or Hide For One
      'CmdOKCounter.Visible = True
      'CmdNgCounter.Visible = True
   End If
End Sub
Private Sub LoadGrid()
With Grid1
    .CellAlignment = flexAlignLeftCenter
    .RowHeight(0) = 1000
    .ColWidthMin = 1100
    .ColWidthMax = 1200
    .Cols = 6
    .Rows = 2
    .TextMatrix(0, 0) = "Reverse" & vbNewLine & "Polarity"
    .TextMatrix(0, 1) = "Cut-Off" & vbNewLine & "Voltage" & vbNewLine & "(AT<12.8V)"
    .TextMatrix(0, 2) = "O/P" & vbNewLine & "Voltage" & vbNewLine & "AT 12.8V"
    .TextMatrix(0, 3) = "O/P" & vbNewLine & "Voltage" & vbNewLine & "AT 14.3V"
    .TextMatrix(0, 4) = "O/P" & vbNewLine & "Voltage" & vbNewLine & "AT 18V"
    .TextMatrix(0, 5) = "O/P" & vbNewLine & "Voltage" & vbNewLine & "Short Test"
End With
With Grid2
    .CellAlignment = flexAlignLeftCenter
    .RowHeight(0) = 1000
    .ColWidthMin = 1100
    .ColWidthMax = 1200
    .Cols = 6
    .Rows = 2
    .TextMatrix(0, 0) = "Test" & vbNewLine & "Voltage"
    .TextMatrix(0, 1) = "Input" & vbNewLine & "Current"
    .TextMatrix(0, 2) = "O/P" & vbNewLine & "Voltage"
    .TextMatrix(0, 3) = "O/P" & vbNewLine & "Current"
    .TextMatrix(0, 4) = "Efficiency" & vbNewLine & "(>75%)"
    .TextMatrix(0, 5) = "Result"
End With

End Sub

Private Function AssignPLCdata()
On Error GoTo Error
   MsgCode = PLcdata(108)
   GridColorfunction Grid1, 1, 0, PLcdata(100), &H1, &H2
   GridColorfunction Grid1, 1, 1, PLcdata(100), &H4, &H8
   GridColorfunction Grid1, 1, 2, PLcdata(100), &H10, &H20
   GridColorfunction Grid1, 1, 3, PLcdata(100), &H40, &H80
   GridColorfunction Grid1, 1, 4, PLcdata(100), &H100, &H200
   GridColorfunction Grid1, 1, 5, PLcdata(100), &H400, &H800
   
   GridColorfunction Grid2, 1, 0, PLcdata(101), &H1, &H2
   GridColorfunction Grid2, 1, 1, PLcdata(101), &H4, &H8
   GridColorfunction Grid2, 1, 2, PLcdata(101), &H10, &H20
   GridColorfunction Grid2, 1, 3, PLcdata(101), &H40, &H80
   GridColorfunction Grid2, 1, 4, PLcdata(101), &H100, &H200
   GridColorfunction Grid2, 1, 5, PLcdata(101), &H400, &H800

   GridTextFunction Grid1, 1, 1, PLcdata(110), 100, "0.00"
   GridTextFunction Grid1, 1, 2, PLcdata(111), 100, "0.00"
   GridTextFunction Grid1, 1, 3, PLcdata(112), 100, "0.00"
   GridTextFunction Grid1, 1, 4, PLcdata(113), 100, "0.00"

   GridTextFunction Grid2, 1, 0, PLcdata(120), 100, "0.00"
   GridTextFunction Grid2, 1, 1, PLcdata(121), 100, "0.00"
   GridTextFunction Grid2, 1, 2, PLcdata(122), 100, "0.00"
   GridTextFunction Grid2, 1, 3, PLcdata(123), 100, "0.00"
   GridTextFunction Grid2, 1, 4, PLcdata(124), 100, "0.00"

   
   txtCycleTime.Text = Format(PLcdata(107) / 10, "0.0")


   If PLcdata(165) = 0 And pulseBreakdown = True Then
      pulseBreakdown = False
      'PictureBreakdown.Visible = False
   ElseIf PLcdata(165) = 1 And pulseBreakdown = False Then
      pulseBreakdown = True
      PictureBreakdown.Visible = True
      cmdrunningbreakdown.Enabled = False
      cmdfullbreakdown.Enabled = False
      cmdgolive.Enabled = True
      cmdclosebreakdownscreen.Enabled = False
      
   ElseIf PLcdata(165) = 2 And pulseBreakdown = False Then
      pulseBreakdown = True
      PictureBreakdown.Visible = True
       cmdrunningbreakdown.Enabled = False
      cmdfullbreakdown.Enabled = False
      cmdgolive.Enabled = True
      cmdclosebreakdownscreen.Enabled = False
   End If
   
   If PLcdata(170) = 0 And PulseScan = False Then
      PulseScan = True
      txtBarcode.Locked = False
      txtBarcode.BackColor = vbWhite
      txtBarcode.Locked = True
      PLcdata(350) = 0
   ElseIf PLcdata(170) = 1 And PulseScan = True Then
      PulseScan = False
      txtBarcode.Locked = False
      txtBarcode.BackColor = vbYellow
      txtBarcode.SetFocus
      
   End If
   If PLcdata(109) = 0 And pulse1 = False Then
      pulse1 = True
      lblGo.Visible = False
      lblNg.Visible = False
   ElseIf PLcdata(109) = 1 And pulse1 = True Then
      pulse1 = False
      lblGo.Visible = True
      GetCounterValue
      txtproductioncounter.Text = Val(txtproductioncounter.Text) + 1
      txtOKCounter.Text = Val(txtOKCounter.Text) + 1
      txtBatchCounter.Text = Val(txtBatchCounter.Text) + 1
      txtTargetProduction.Text = Val(txtTargetProduction.Text) - 1
      txtCouplerCounter.Text = Val(txtCouplerCounter.Text) + 1
      If pulsePrinterBypass = False Then
        PrintLabel JustPrinter1
      End If
      SaveProductioncounter
      SaveReport 1
      SaveCounter
      SaveCounterValue
   ElseIf PLcdata(109) = 2 And pulse1 = True Then
      pulse1 = False
      GetCounterValue
      lblNg.Visible = True
      txtNGCounter.Text = Val(txtNGCounter.Text) + 1
      SaveReport 2
      SaveCounter
      SaveCounterValue
   End If
      
Exit Function
Error:
   ErrorLog Err.Number, Err.Description & "---", Erl, Me.Name, "Assign PLC Data"
   Resume Next
End Function

Private Sub GridTextFunction(Grid As VSFlexGrid, Row As Integer, Col As Integer, Data As Integer, Devision As Integer, formatstring As String)
Grid.TextMatrix(Row, Col) = Format(Data / Devision, formatstring)
End Sub
Private Sub GridColorfunction(Grid As VSFlexGrid, Row As Integer, Col As Integer, Data As Integer, reg1 As Integer, reg2 As Integer)
    If (Data And reg1) Then
        If (Data And reg2) Then
           Grid.Cell(flexcpBackColor, Row, Col, Row, Col) = vbYellow
        Else
           Grid.Cell(flexcpBackColor, Row, Col, Row, Col) = vbGreen
        End If
    ElseIf (Data And reg2) Then
          Grid.Cell(flexcpBackColor, Row, Col, Row, Col) = vbRed
    Else
          Grid.Cell(flexcpBackColor, Row, Col, Row, Col) = vbWhite
    
    End If
End Sub

Private Sub ShapeColorsinglefunction(Data As Integer, reg1 As Integer, ctrl As Object)
    If (Data And reg1) <> 0 Then
          ctrl.BackColor = vbYellow
    Else
          ctrl.BackColor = vbWhite
    End If
End Sub
Private Sub ShapeColorsingleifunction(Data As Integer, reg1 As Integer, ctrl As Object)
    If (Data And reg1) <> 0 Then
          ctrl.BackColor = vbGreen
    Else
          ctrl.BackColor = vbWhite
    End If
End Sub

Private Sub Timer2_Timer()
'On Error Resume Next

'    txttime = Format(Time(), "Hh:Mm:Ss")

    Static TOGGLE As Boolean
    TOGGLE = Not (TOGGLE)
    Timer2.Interval = 400
    
    With txtCommandLine
        .BorderStyle = 1
        .Alignment = 2
        .FontBold = True
       
        .FontSize = 16
    End With
       
    If InternetGetConnectedState(0, 0) = 1 Then
        shapeInternet.BackColor = vbGreen
        'sendEmail
    Else
        shapeInternet.BackColor = vbRed
    End If
    
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
Public Function sendEmail()
'On Error GoTo Error
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim Sql As String


Sql = "Select * from Common_Set where SetType ='CommonSet'"
Set rs1 = New ADODB.Recordset
rs1.Open Sql, Con, adOpenDynamic, adLockOptimistic
If rs1("SenderEmail") <> "" And rs1("ToEmail1") <> "" Then
    'Sql = "select Top 1 * from model_report_counter where MailSent = false order by id desc"
    Sql = "select Top 1 * from model_report_counter order by id desc"
    Set rs2 = New ADODB.Recordset
    rs2.Open Sql, Con, adOpenDynamic, adLockOptimistic
    Do While rs2.EOF = False
        Dim Body As String
        Dim Subject As String
        Subject = "Production Report of Switch testing of " & rs2("ModelName") & "for date " & Format(rs2("DateTime"), "dd-mm-yyyy") & "and Shift " & rs2("ShiftTime")
        Body = "Dear Team," & vbNewLine
        Body = Body & "Below is the Production detail of Date '" & Format(rs2("DateTime"), "dd-mm-yyyy")
        Body = Body & "' and Shift '" & rs2("ShiftTime") & "' :" & vbNewLine
        Body = Body & "Model Name :- '" & rs2("ModelName") & "'" & vbNewLine
        Body = Body & "Total Ok Parts :- " & rs2("OKCounter") & vbNewLine
        Body = Body & "Total NG Parts :- " & rs2("NGCounter") & vbNewLine
        Body = Body & "Total Production Parts :- " & rs2("ProductionCounter") & vbNewLine
        If callSendEmailApi(rs1, Subject, Body) = True Then
         rs2("MailSent") = 1
         rs2.Update
        End If
        
        rs2.MoveNext
    Loop
End If
'End Function
'Error:
'ErrorLog Err.Number, Err.Description, Erl, Me.Name, "LoadData"
End Function
Private Function callSendEmailApi(rsGeneralset As ADODB.Recordset, Subject As String, Body As String) As Boolean
Dim ToEmail As String
    ToEmail = "&ToMailAddress%5b0%5d=" & rsGeneralset("ToEmail1")
    j = 0
    For i = 1 To 6
        If rsGeneralset("EmailBypass" & i) = False Then
            j = j + 1
            ReDim ToEmail1(j) As String
            ToEmail = ToEmail & "&ToMailAddress%5b" & j & "%5d=" & rsGeneralset("ToEmail" & i + 1)
        End If
    Next
    Dim URL As String
    Dim response As String
    
    URL = "http://" & rsGeneralset("WebApiLink") & "/SendMail?"
    URL = URL & "FromMailAddress=" & rsGeneralset("SenderEmail")
    URL = URL & "&FromMailPassword=" & rsGeneralset("SenderPassword")
    URL = URL & ToEmail
    URL = URL & "&subject=" & Subject
    URL = URL & "&body=" & Body
    
    Dim res As WinHttp.WinHttpRequest
    Set res = New WinHttp.WinHttpRequest
    With res
    
      ErrorLog 100, "API Initialise With URL - " & URL, "", "callsendEmailApi", ""
     .Open "Get", URL, False
     .Send
     .WaitForResponse
     response = .ResponseText
     ErrorLog 100, "API Response Recieved - " & response, "", "callsendEmailApi", ""
     If response = "SENT" Then
     callSendEmailApi = True
     Else
     callSendEmailApi = False
     
     End If
     
    
    End With
End Function
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

Private Sub LoadData()

On Error GoTo Error
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim strByPass(14) As Integer
Dim j As Integer

    Sql = "Select * from Model_Set where ModelName='" & ModelName & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    PLcdata(240) = 1
    PLcdata(210) = Val(Rs("testVoltage")) * 100
    PLcdata(211) = Val(Rs("RPCheckTime")) * 100
    PLcdata(212) = Val(Rs("STCheckTime")) * 100
    PLcdata(213) = Val(Rs("Efficiency")) * 100
    PLcdata(214) = Val(Rs("OutputVoltMin")) * 100
    PLcdata(215) = Val(Rs("OutputVoltMax")) * 100
    PLcdata(216) = Val(Rs("OutputCurrentMin")) * 100
    PLcdata(217) = Val(Rs("OutputCurrentMax")) * 100
    PLcdata(218) = Val(Rs("VoltageOffset")) * 100
    PLcdata(219) = Val(Rs("CurrentOffset")) * 100

    txtModelDesc.Text = Trim(Rs("ModelDesc"))
    If Val(txtCouplerCounter.Text) >= setCouplerCounter Then
        PLcdata(235) = 1
    ElseIf Val(txtBatchCounter.Text) >= setBatchCounter Then
        PLcdata(235) = 2
    Else
        PLcdata(235) = 0
    End If
    PartNo = Rs("PrintPartNo")
    BarcodeLength = Rs("BarcodeLength")
    HardwareNo = Rs("HardwareNo")
    SerialStartingtxt = Rs("SerialStartingtxt")
    
 '   PLcdata(320) = Val(Rs("DebounceTime")) * 100
 '   PLcdata(321) = Val(Rs("HoldTime")) * 100
 '   PLcdata(322) = Val(Rs("CheckTime")) * 1000
    
 '   PLcdata(332) = Val(Rs("DotMarkingTime")) * 10
    
    ModelNo = Rs("ModelNo")
    PLcdata(231) = Rs("ModelNo")
    
    'Rs("BatchCounter").Text
    'Rs("CouplerCounter") = .Text
    'Rs ("PartImage")
    'Rs("productioncounter") =
    
    PLcdata(230) = 0
    PLcdata(230) = PLcdata(230) + &H1 * Val(Rs("Bypass1"))
    PLcdata(230) = PLcdata(230) + &H2 * Val(Rs("Bypass2"))
    PLcdata(230) = PLcdata(230) + &H4 * Val(Rs("Bypass3"))
    PLcdata(230) = PLcdata(230) + &H8 * Val(Rs("Bypass4"))
    PLcdata(230) = PLcdata(230) + &H10 * Val(Rs("Bypass5"))
    PLcdata(230) = PLcdata(230) + &H20 * Val(Rs("Bypass6"))
    PLcdata(230) = PLcdata(230) + &H40 * Val(Rs("Bypass7"))
    PLcdata(230) = PLcdata(230) + &H80 * Val(Rs("ByPass8"))
  '  PLcdata(330) = PLcdata(330) + &H100 * Val(Rs("ByPass9"))
    'PLcdata(330) = PLcdata(330) + &H200 * Val(Rs("ByPass10"))
    'PLcdata(330) = PLcdata(330) + &H400 * Val(Rs("ByPass11"))
    chkproductioncount
Exit Sub
Error:
ErrorLog Err.Number, Err.Description, Erl, Me.Name, "LoadData"
Resume Next
End Sub
Private Sub chkproductioncount()
    tempgetshift = getShift
    'TempReportDate
       tempshift = GetSetting(App.Title, ModelName, "saveshift", 0)
       tempdate = GetSetting(App.Title, ModelName, "savedate", 0)
       If Val(txtTargetProduction.Text) > 0 And txtTargetProduction.BackColor <> vbYellow Then
        If TempReportDate <> DateValue(tempdate) Then
            txtTargetProduction.Enabled = True
            txtTargetProduction.Text = ""
            txtTargetProduction.SetFocus
            txtTargetProduction.BackColor = vbYellow
            Command1.Visible = True
            PLcdata(236) = 1
            Exit Sub
        Else
            If tempgetshift <> tempshift Then
                txtTargetProduction.Locked = False
                txtTargetProduction.Text = ""
                txtTargetProduction.SetFocus
                txtTargetProduction.BackColor = vbYellow
                Command1.Visible = True
                PLcdata(236) = 1
                Exit Sub
            End If
        End If
    ElseIf txtTargetProduction.BackColor <> vbYellow Then
        txtTargetProduction.Locked = False
        txtTargetProduction.Text = ""
        txtTargetProduction.SetFocus
        txtTargetProduction.BackColor = vbYellow
        Command1.Visible = True
        PLcdata(236) = 1
        
    End If
End Sub
Private Sub LoadSettingsData()
On Error GoTo Error
Dim Rs As ADODB.Recordset
Dim Sql As String

   Sql = "Select * from Model_Set where ModelName='" & ModelName & "'"
   Set Rs = New ADODB.Recordset
   Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
        
    txtModelDesc.Text = Rs("ModelDesc")
    PartNo = Rs("PrintPartNo")
    BarcodeLength = Rs("BarcodeLength")
    HardwareNo = Rs("HardwareNo")
    SerialStartingtxt = Rs("SerialStartingtxt")
    setBatchCounter = Rs("BatchCounter")
    setCouplerCounter = Rs("CouplerCounter")
    VendorId = Rs("VendorId")
    ImgPart.Picture = LoadPicture(Rs("PartImage"))
    txtproductioncounter.Text = Rs("productioncounter")
    If Val(Rs("PrinterBypass")) = 1 Then
        pulsePrinterBypass = True
    Else
        pulsePrinterBypass = False
    End If
Exit Sub
Error:
ErrorLog Err.Number, Err.Description, Erl, Me.Name, "LoadSettingsData"
Resume Next
End Sub
Private Function getresult(pic As PictureBox) As Integer
   If pic.BackColor = vbGreen Then
    getresult = 1
   ElseIf pic.BackColor = vbRed Then
    getresult = 2
   ElseIf pic.BackColor = vbWhite Then
    getresult = 0
   End If
End Function

Private Sub SaveReport(result As String)
'On Error GoTo Error
Dim Sql As String
Dim Rs As ADODB.Recordset
   Sql = "Select * from Model_Report"
   Set Rs = New ADODB.Recordset
   Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   Rs.AddNew
      Rs("ModelName") = ModelName
      Rs("OperatorName") = LoginUser
      Rs("Date") = Format(Now(), "mm/dd/yyyy")
      Rs("Time") = Format(Now(), "hh:mm:ss")
      Rs("Barcode") = barcode
      Rs("Result") = result
    Rs.Update
End Sub
Private Sub SaveCounter()
Dim Sql As String
Dim Rs As ADODB.Recordset
    Sql = "Select * from Model_Report_Counter where datetime = #" & runningreportdate & "# and shifttime = '" & runningreportshift & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    If Rs.EOF = True Then
      Rs.AddNew
      Rs("ModelName") = ModelName
      Rs("DateTime") = runningreportdate
      Rs("ShiftTime") = runningreportshift
      Rs("Mailsent") = 0
      Rs("ModelNo") = ModelNo
    End If
      Rs("ProductionCounter") = Val(txtproductioncounter.Text)
      Rs("OKCounter") = Val(txtOKCounter.Text)
      Rs("NGCounter") = Val(txtNGCounter.Text)
      Rs("CouplerCounter") = Val(txtCouplerCounter.Text)
      Rs("BatchCounter") = Val(txtBatchCounter.Text)
      If Val(txtTargetProduction.Text) > 0 Then
        Rs("TargetProduction") = Val(txtTargetProduction.Text)
      End If
      Rs.Update
End Sub
Private Sub SaveBreakDown(breakdownType As Integer, breakdownstatus As Integer)
Dim Sql As String
Dim Rs As ADODB.Recordset
   Sql = "Select Top 1 * from Model_Report_Breakdown "
   Set Rs = New ADODB.Recordset
   Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   If breakdownstatus = 1 Then
      Rs.AddNew
      Rs("StartTime") = Format(Now(), "mm/dd/yyyy hh:mm:ss")
      Rs("BreakdownType") = breakdownType
   Else
      Rs("Remarks") = txtbreakdownsummary.Text
      Rs("EndTime") = Format(Now(), "mm/dd/yyyy hh:mm:ss")
   End If
   Rs.Update
   Exit Sub
Error:
   ErrorLog Err.Number, Err.Description, Erl, Me.Name, "SaveReport"
   Resume Next
End Sub

Private Sub SaveCounterValue()
 Dim ProdDay As String
 SaveSetting App.Title, ModelName, "OkCounter", Val(txtOKCounter.Text)
 SaveSetting App.Title, ModelName, "NGCounter", Val(txtNGCounter.Text)
 SaveSetting App.Title, ModelName, "CouplerCounter", Val(txtCouplerCounter.Text)
 SaveSetting App.Title, ModelName, "BatchCounter", Val(txtBatchCounter.Text)
SaveSetting App.Title, ModelName, "TargetProduction", txtTargetProduction.Text
       
 'ProdDay = Format(Date, "ddmmyy")
 'SaveSetting App.Title, ModelName, "", Val(ProdDay)
 'SaveSetting App.Title, ModelName, "PrintCounter", txtprintcounter.Text
End Sub
Private Sub SaveProductioncounter()
Dim Rs As ADODB.Recordset
Dim Sql As String
    Sql = "Select * from Model_Set where ModelName ='" & Trim(txtModelName.Text) & "'"
    Set Rs = New ADODB.Recordset
    Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
    Rs("productioncounter") = Val(txtproductioncounter.Text)
    Rs.Update
    'txtSaveCoupler.Text = Rs("CouplerCounter")
End Sub
Private Sub GetCounterValue()
On Error Resume Next
Dim ProdDay As String
Dim Today As String
   txtOKCounter.Text = Val(GetSetting(App.Title, ModelName, "OkCounter", 0))
   txtNGCounter.Text = Val(GetSetting(App.Title, ModelName, "NgCounter", 0))
   txtCouplerCounter.Text = Val(GetSetting(App.Title, ModelName, "CouplerCounter", 0))
   txtBatchCounter.Text = Val(GetSetting(App.Title, ModelName, "BatchCounter", 0))
   txtTargetProduction.Text = GetSetting(App.Title, ModelName, "TargetProduction", 0)
         
   tempshift = GetSetting(App.Title, ModelName, "saveshift", 0)
   tempdate = GetSetting(App.Title, ModelName, "savedate", 0)
   If tempdate <> runningreportdate Or runningreportshift <> tempshift Then
      txtOKCounter.Text = 0
      txtNGCounter.Text = 0
      SaveSetting App.Title, ModelName, "saveshift", runningreportshift
      SaveSetting App.Title, ModelName, "savedate", runningreportdate
      'txtprintcounter.Text = 0
   End If
   SaveCounterValue
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

Private Sub txtBarcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtBarcode.Locked = True
   If txtBarcode.Text = barcode Then
     txtBarcode.BackColor = vbGreen
     PLcdata(350) = 1
   Else
     txtBarcode.BackColor = vbRed
     PLcdata(350) = 2
     'SaveReport "NG"
   End If
End If
End Sub

Private Function ValidateBarcode(barcode As String) As Boolean
Dim Rs As ADODB.Recordset
Dim Sql As String
   Sql = "Select * from Model_Report where barcode='" & barcode & "'"
   Set Rs = New ADODB.Recordset
   Rs.Open Sql, Con, adOpenDynamic, adLockOptimistic
   If Rs.EOF = False Then
      checkBarcoderepeat = True
   Else
      checkBarcoderepeat = False
   End If
End Function

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
