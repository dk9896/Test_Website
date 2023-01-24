VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{97C0E9D8-AD04-4920-9B7A-4B99616579F9}#2.0#0"; "TextPrinter.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMonitor 
   Caption         =   "Switch_Testing_With_Email_Job2"
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
   ScaleHeight     =   10935
   ScaleWidth      =   20250
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
      Height          =   10695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   19815
      Begin VB.PictureBox PictureBreakdown 
         BackColor       =   &H80000010&
         Height          =   6015
         Left            =   5160
         ScaleHeight     =   5955
         ScaleWidth      =   8595
         TabIndex        =   109
         Top             =   1560
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
            TabIndex        =   116
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
            TabIndex        =   113
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
            TabIndex        =   112
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
            TabIndex        =   111
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
            TabIndex        =   110
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
            TabIndex        =   115
            Top             =   4800
            Width           =   2295
         End
      End
      Begin VB.PictureBox PictureSRM 
         BackColor       =   &H80000010&
         Height          =   375
         Index           =   0
         Left            =   4920
         ScaleHeight     =   315
         ScaleWidth      =   2355
         TabIndex        =   166
         Top             =   6120
         Width           =   2415
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Set/Reset Module"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   49
            Left            =   240
            TabIndex        =   167
            Top             =   0
            Width           =   1995
         End
      End
      Begin VB.PictureBox PictureCUM 
         BackColor       =   &H80000010&
         Height          =   375
         Index           =   0
         Left            =   2520
         ScaleHeight     =   315
         ScaleWidth      =   2355
         TabIndex        =   164
         Top             =   6120
         Width           =   2415
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Custom Module"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   48
            Left            =   360
            TabIndex        =   165
            Top             =   0
            Width           =   1800
         End
      End
      Begin VB.PictureBox PictureCRM 
         BackColor       =   &H80000010&
         Height          =   375
         Index           =   0
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   2355
         TabIndex        =   162
         Top             =   6120
         Width           =   2415
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cruise Module"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   47
            Left            =   360
            TabIndex        =   163
            Top             =   0
            Width           =   1650
         End
      End
      Begin VB.PictureBox PictureSRM 
         BackColor       =   &H80000010&
         Height          =   3375
         Index           =   1
         Left            =   4920
         ScaleHeight     =   3315
         ScaleWidth      =   2355
         TabIndex        =   154
         Top             =   6480
         Width           =   2415
         Begin VB.Frame Frame19 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   120
            TabIndex        =   155
            Top             =   2280
            Width           =   2175
            Begin VB.TextBox txtVSRM 
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
               Locked          =   -1  'True
               TabIndex        =   157
               Text            =   "0.000"
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox txtCurSRM 
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
               Locked          =   -1  'True
               TabIndex        =   156
               Text            =   "0.000"
               Top             =   120
               Width           =   615
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured V"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   45
               Left            =   120
               TabIndex        =   159
               Top             =   480
               Width           =   1005
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured Amp"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   44
               Left            =   120
               TabIndex        =   158
               Top             =   120
               Width           =   1275
            End
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "OFF"
            Height          =   375
            Index           =   46
            Left            =   840
            TabIndex        =   161
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ON"
            Height          =   435
            Left            =   960
            TabIndex        =   160
            Top             =   1560
            Width           =   525
         End
         Begin VB.Shape ShpSRMInner 
            BackStyle       =   1  'Opaque
            Height          =   615
            Index           =   0
            Left            =   480
            Top             =   240
            Width           =   1455
         End
         Begin VB.Shape ShpSRMOuter 
            BackStyle       =   1  'Opaque
            Height          =   855
            Index           =   0
            Left            =   240
            Top             =   120
            Width           =   1935
         End
         Begin VB.Shape ShpSRMInner 
            BackStyle       =   1  'Opaque
            Height          =   615
            Index           =   1
            Left            =   480
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Shape ShpSRMOuter 
            BackStyle       =   1  'Opaque
            Height          =   855
            Index           =   1
            Left            =   240
            Top             =   1320
            Width           =   1935
         End
      End
      Begin VB.PictureBox PictureCUM 
         BackColor       =   &H80000010&
         Height          =   3375
         Index           =   1
         Left            =   2520
         ScaleHeight     =   3315
         ScaleWidth      =   2355
         TabIndex        =   146
         Top             =   6480
         Width           =   2415
         Begin VB.Frame Frame16 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   120
            TabIndex        =   147
            Top             =   2280
            Width           =   2175
            Begin VB.TextBox txtVCUM 
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
               Locked          =   -1  'True
               TabIndex        =   149
               Text            =   "0.000"
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox txtCurCUM 
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
               Locked          =   -1  'True
               TabIndex        =   148
               Text            =   "0.000"
               Top             =   120
               Width           =   615
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured V"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   41
               Left            =   120
               TabIndex        =   151
               Top             =   480
               Width           =   1005
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured Amp"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   37
               Left            =   120
               TabIndex        =   150
               Top             =   120
               Width           =   1275
            End
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "OFF"
            Height          =   375
            Index           =   42
            Left            =   840
            TabIndex        =   153
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ON"
            Height          =   435
            Left            =   960
            TabIndex        =   152
            Top             =   1560
            Width           =   525
         End
         Begin VB.Shape ShpCUMInner 
            BackStyle       =   1  'Opaque
            Height          =   615
            Index           =   0
            Left            =   480
            Top             =   240
            Width           =   1455
         End
         Begin VB.Shape ShpCUMOuter 
            BackStyle       =   1  'Opaque
            Height          =   855
            Index           =   0
            Left            =   240
            Top             =   120
            Width           =   1935
         End
         Begin VB.Shape ShpCUMInner 
            BackStyle       =   1  'Opaque
            Height          =   615
            Index           =   1
            Left            =   480
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Shape ShpCUMOuter 
            BackStyle       =   1  'Opaque
            Height          =   855
            Index           =   1
            Left            =   240
            Top             =   1320
            Width           =   1935
         End
      End
      Begin VB.PictureBox PictureCRM 
         BackColor       =   &H80000010&
         Height          =   3375
         Index           =   1
         Left            =   120
         ScaleHeight     =   3315
         ScaleWidth      =   2355
         TabIndex        =   138
         Top             =   6480
         Width           =   2415
         Begin VB.Frame Frame9 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   120
            TabIndex        =   139
            Top             =   2280
            Width           =   2175
            Begin VB.TextBox txtVCRM 
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
               Locked          =   -1  'True
               TabIndex        =   141
               Text            =   "0.000"
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox txtCurCRM 
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
               Locked          =   -1  'True
               TabIndex        =   140
               Text            =   "0.000"
               Top             =   120
               Width           =   615
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured V"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   24
               Left            =   120
               TabIndex        =   143
               Top             =   480
               Width           =   1005
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured Amp"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   23
               Left            =   120
               TabIndex        =   142
               Top             =   120
               Width           =   1275
            End
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "OFF"
            Height          =   375
            Index           =   25
            Left            =   840
            TabIndex        =   145
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ON"
            Height          =   435
            Left            =   960
            TabIndex        =   144
            Top             =   1560
            Width           =   525
         End
         Begin VB.Shape ShpCRMInner 
            BackStyle       =   1  'Opaque
            Height          =   615
            Index           =   1
            Left            =   480
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Shape ShpCRMInner 
            BackStyle       =   1  'Opaque
            Height          =   615
            Index           =   0
            Left            =   480
            Top             =   240
            Width           =   1455
         End
         Begin VB.Shape ShpCRMOuter 
            BackStyle       =   1  'Opaque
            Height          =   855
            Index           =   1
            Left            =   240
            Top             =   1320
            Width           =   1935
         End
         Begin VB.Shape ShpCRMOuter 
            BackStyle       =   1  'Opaque
            Height          =   855
            Index           =   0
            Left            =   240
            Top             =   120
            Width           =   1935
         End
      End
      Begin VB.PictureBox PictureNM 
         BackColor       =   &H80000010&
         Height          =   5655
         Index           =   1
         Left            =   7320
         ScaleHeight     =   5595
         ScaleWidth      =   7515
         TabIndex        =   122
         Top             =   1800
         Width           =   7575
         Begin VB.Frame Frame21 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   240
            TabIndex        =   173
            Top             =   2880
            Width           =   2175
            Begin VB.TextBox txtCurNM 
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
               Locked          =   -1  'True
               TabIndex        =   175
               Text            =   "0.000"
               Top             =   120
               Width           =   615
            End
            Begin VB.TextBox txtVNM 
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
               Locked          =   -1  'True
               TabIndex        =   174
               Text            =   "0.000"
               Top             =   480
               Width           =   615
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured Amp"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   53
               Left            =   120
               TabIndex        =   177
               Top             =   120
               Width           =   1275
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured V"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   52
               Left            =   120
               TabIndex        =   176
               Top             =   480
               Width           =   1005
            End
         End
         Begin VB.Frame Frame20 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   2640
            TabIndex        =   168
            Top             =   1080
            Width           =   2175
            Begin VB.TextBox txtVNM 
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
               Locked          =   -1  'True
               TabIndex        =   170
               Text            =   "0.000"
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox txtCurNM 
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
               Locked          =   -1  'True
               TabIndex        =   169
               Text            =   "0.000"
               Top             =   120
               Width           =   615
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured V"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   51
               Left            =   120
               TabIndex        =   172
               Top             =   480
               Width           =   1005
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured Amp"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   50
               Left            =   120
               TabIndex        =   171
               Top             =   120
               Width           =   1275
            End
         End
         Begin VB.Frame Frame17 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   2640
            TabIndex        =   128
            Top             =   4560
            Width           =   2175
            Begin VB.TextBox txtVNM 
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
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   130
               Text            =   "0.000"
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox txtCurNM 
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
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   129
               Text            =   "0.000"
               Top             =   120
               Width           =   615
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured V"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   26
               Left            =   120
               TabIndex        =   132
               Top             =   480
               Width           =   1005
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured Amp"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   27
               Left            =   120
               TabIndex        =   131
               Top             =   120
               Width           =   1275
            End
         End
         Begin VB.Frame Frame18 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   5160
            TabIndex        =   123
            Top             =   2880
            Width           =   2175
            Begin VB.TextBox txtCurNM 
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
               Locked          =   -1  'True
               TabIndex        =   125
               Text            =   "0.000"
               Top             =   120
               Width           =   615
            End
            Begin VB.TextBox txtVNM 
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
               Locked          =   -1  'True
               TabIndex        =   124
               Text            =   "0.000"
               Top             =   480
               Width           =   615
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured Amp"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   30
               Left            =   120
               TabIndex        =   127
               Top             =   120
               Width           =   1275
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured V"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   31
               Left            =   120
               TabIndex        =   126
               Top             =   480
               Width           =   1005
            End
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BACK"
            Height          =   435
            Left            =   720
            TabIndex        =   178
            Top             =   2160
            Width           =   1020
         End
         Begin VB.Shape ShpNMInner 
            BackStyle       =   1  'Opaque
            Height          =   615
            Index           =   4
            Left            =   600
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Shape ShpNMOuter 
            BackStyle       =   1  'Opaque
            Height          =   855
            Index           =   4
            Left            =   360
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DOWN"
            Height          =   435
            Left            =   3120
            TabIndex        =   135
            Top             =   3840
            Width           =   1125
         End
         Begin VB.Shape ShpNMInner 
            BackStyle       =   1  'Opaque
            Height          =   615
            Index           =   1
            Left            =   3000
            Top             =   3720
            Width           =   1455
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "UP"
            Height          =   375
            Index           =   29
            Left            =   3360
            TabIndex        =   134
            Top             =   360
            Width           =   855
         End
         Begin VB.Shape ShpNMInner 
            BackStyle       =   1  'Opaque
            Height          =   615
            Index           =   2
            Left            =   3000
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ENTER"
            Height          =   435
            Left            =   5640
            TabIndex        =   133
            Top             =   2160
            Width           =   1215
         End
         Begin VB.Shape ShpNMInner 
            BackStyle       =   1  'Opaque
            Height          =   615
            Index           =   3
            Left            =   5520
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Shape ShpNMOuter 
            BackStyle       =   1  'Opaque
            Height          =   855
            Index           =   3
            Left            =   5280
            Top             =   1920
            Width           =   1935
         End
         Begin VB.Shape ShpNMOuter 
            BackStyle       =   1  'Opaque
            Height          =   855
            Index           =   2
            Left            =   2760
            Top             =   120
            Width           =   1935
         End
         Begin VB.Shape ShpNMOuter 
            BackStyle       =   1  'Opaque
            Height          =   855
            Index           =   1
            Left            =   2760
            Top             =   3600
            Width           =   1935
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "OFF"
            Height          =   375
            Index           =   28
            Left            =   3360
            TabIndex        =   136
            Top             =   2520
            Width           =   855
         End
         Begin VB.Shape ShpNMInner 
            BackStyle       =   1  'Opaque
            Height          =   615
            Index           =   0
            Left            =   3000
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Shape ShpNMOuter 
            BackStyle       =   1  'Opaque
            Height          =   855
            Index           =   0
            Left            =   2760
            Top             =   2280
            Width           =   1935
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
         Left            =   10800
         TabIndex        =   121
         Top             =   8640
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
         TabIndex        =   120
         Text            =   "Text3"
         Top             =   480
         Visible         =   0   'False
         Width           =   6135
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   615
         Left            =   4320
         TabIndex        =   119
         Top             =   360
         Visible         =   0   'False
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
         Left            =   9600
         TabIndex        =   118
         Top             =   9120
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
         Left            =   9720
         TabIndex        =   117
         Top             =   8640
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
         TabIndex        =   108
         Top             =   360
         Width           =   1575
      End
      Begin VB.PictureBox Picture1 
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1035
         ScaleWidth      =   2235
         TabIndex        =   107
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
      Begin VB.TextBox txtILLH 
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
         Left            =   9720
         TabIndex        =   100
         Top             =   8160
         Width           =   975
      End
      Begin VB.PictureBox PictureNM 
         BackColor       =   &H80000010&
         Height          =   375
         Index           =   0
         Left            =   7320
         ScaleHeight     =   315
         ScaleWidth      =   7515
         TabIndex        =   97
         Top             =   1440
         Width           =   7575
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Navigation Module"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   38
            Left            =   2640
            TabIndex        =   98
            Top             =   0
            Width           =   2130
         End
      End
      Begin VB.PictureBox PictureHAM 
         BackColor       =   &H80000010&
         Height          =   375
         Index           =   0
         Left            =   4920
         ScaleHeight     =   315
         ScaleWidth      =   2355
         TabIndex        =   95
         Top             =   1440
         Width           =   2415
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hazard Module"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   36
            Left            =   240
            TabIndex        =   96
            Top             =   0
            Width           =   1725
         End
      End
      Begin VB.PictureBox PictureHOM 
         BackColor       =   &H80000010&
         Height          =   375
         Index           =   0
         Left            =   17280
         ScaleHeight     =   315
         ScaleWidth      =   2355
         TabIndex        =   93
         Top             =   1440
         Width           =   2415
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Horn Module"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   35
            Left            =   240
            TabIndex        =   94
            Top             =   0
            Width           =   1470
         End
      End
      Begin VB.PictureBox PictureBM 
         BackColor       =   &H80000010&
         Height          =   375
         Index           =   0
         Left            =   14880
         ScaleHeight     =   315
         ScaleWidth      =   2355
         TabIndex        =   91
         Top             =   1440
         Width           =   2415
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Blinker Module"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   33
            Left            =   240
            TabIndex        =   92
            Top             =   0
            Width           =   1710
         End
      End
      Begin VB.PictureBox PicturePM 
         BackColor       =   &H80000010&
         Height          =   375
         Index           =   0
         Left            =   2520
         ScaleHeight     =   315
         ScaleWidth      =   2355
         TabIndex        =   89
         Top             =   1440
         Width           =   2415
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pass Module"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   32
            Left            =   480
            TabIndex        =   90
            Top             =   0
            Width           =   1470
         End
      End
      Begin VB.PictureBox PictureHAM 
         BackColor       =   &H80000010&
         Height          =   4215
         Index           =   1
         Left            =   4920
         ScaleHeight     =   4155
         ScaleWidth      =   2355
         TabIndex        =   81
         Top             =   1800
         Width           =   2415
         Begin VB.Frame Frame15 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   120
            TabIndex        =   82
            Top             =   3120
            Width           =   2175
            Begin VB.TextBox txtCurHAM 
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
               Locked          =   -1  'True
               TabIndex        =   84
               Text            =   "0.000"
               Top             =   120
               Width           =   615
            End
            Begin VB.TextBox txtVHAM 
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
               Locked          =   -1  'True
               TabIndex        =   83
               Text            =   "0.000"
               Top             =   480
               Width           =   615
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured Amp"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   21
               Left            =   120
               TabIndex        =   86
               Top             =   120
               Width           =   1275
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured V"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   20
               Left            =   120
               TabIndex        =   85
               Top             =   480
               Width           =   1005
            End
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ON"
            Height          =   435
            Left            =   960
            TabIndex        =   88
            Top             =   2400
            Width           =   525
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "OFF"
            Height          =   375
            Index           =   22
            Left            =   840
            TabIndex        =   87
            Top             =   360
            Width           =   855
         End
         Begin VB.Shape ShpHAMInner 
            BackStyle       =   1  'Opaque
            Height          =   615
            Index           =   1
            Left            =   480
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Shape ShpHAMOuter 
            BackStyle       =   1  'Opaque
            Height          =   855
            Index           =   1
            Left            =   240
            Top             =   2160
            Width           =   1935
         End
         Begin VB.Shape ShpHAMInner 
            BackStyle       =   1  'Opaque
            Height          =   615
            Index           =   0
            Left            =   480
            Top             =   240
            Width           =   1455
         End
         Begin VB.Shape ShpHAMOuter 
            BackStyle       =   1  'Opaque
            Height          =   855
            Index           =   0
            Left            =   240
            Top             =   120
            Width           =   1935
         End
      End
      Begin VB.PictureBox PictureHOM 
         BackColor       =   &H80000010&
         Height          =   3735
         Index           =   1
         Left            =   17280
         ScaleHeight     =   3675
         ScaleWidth      =   2355
         TabIndex        =   73
         Top             =   1800
         Width           =   2415
         Begin VB.Frame Frame14 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   120
            TabIndex        =   74
            Top             =   2400
            Width           =   2175
            Begin VB.TextBox txtCurHOM 
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
               Locked          =   -1  'True
               TabIndex        =   76
               Text            =   "0.000"
               Top             =   120
               Width           =   615
            End
            Begin VB.TextBox txtVHOM 
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
               Locked          =   -1  'True
               TabIndex        =   75
               Text            =   "0.000"
               Top             =   480
               Width           =   615
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured Amp"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   18
               Left            =   120
               TabIndex        =   78
               Top             =   120
               Width           =   1275
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured V"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   17
               Left            =   120
               TabIndex        =   77
               Top             =   480
               Width           =   1005
            End
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ON"
            Height          =   435
            Left            =   840
            TabIndex        =   80
            Top             =   1680
            Width           =   525
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "OFF"
            Height          =   375
            Index           =   19
            Left            =   840
            TabIndex        =   79
            Top             =   360
            Width           =   855
         End
         Begin VB.Shape ShpHOMInner 
            BackStyle       =   1  'Opaque
            Height          =   615
            Index           =   0
            Left            =   480
            Top             =   240
            Width           =   1455
         End
         Begin VB.Shape ShpHOMOuter 
            BackStyle       =   1  'Opaque
            Height          =   855
            Index           =   0
            Left            =   240
            Top             =   120
            Width           =   1935
         End
         Begin VB.Shape ShpHOMInner 
            BackStyle       =   1  'Opaque
            Height          =   615
            Index           =   1
            Left            =   480
            Top             =   1560
            Width           =   1455
         End
         Begin VB.Shape ShpHOMOuter 
            BackStyle       =   1  'Opaque
            Height          =   855
            Index           =   1
            Left            =   240
            Top             =   1440
            Width           =   1935
         End
      End
      Begin VB.PictureBox PicturePM 
         BackColor       =   &H80000010&
         Height          =   4215
         Index           =   1
         Left            =   2520
         ScaleHeight     =   4155
         ScaleWidth      =   2355
         TabIndex        =   65
         Top             =   1800
         Width           =   2415
         Begin VB.Frame Frame7 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   120
            TabIndex        =   66
            Top             =   3120
            Width           =   2175
            Begin VB.TextBox txtCurPM 
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
               Locked          =   -1  'True
               TabIndex        =   68
               Text            =   "0.000"
               Top             =   120
               Width           =   615
            End
            Begin VB.TextBox txtVPM 
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
               Locked          =   -1  'True
               TabIndex        =   67
               Text            =   "0.000"
               Top             =   480
               Width           =   615
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured Amp"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   6
               Left            =   120
               TabIndex        =   70
               Top             =   120
               Width           =   1275
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured V"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   7
               Left            =   120
               TabIndex        =   69
               Top             =   480
               Width           =   1005
            End
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PASS"
            Height          =   435
            Left            =   720
            TabIndex        =   72
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "OFF"
            Height          =   375
            Index           =   10
            Left            =   840
            TabIndex        =   71
            Top             =   360
            Width           =   855
         End
         Begin VB.Shape ShpPMInner 
            BackStyle       =   1  'Opaque
            Height          =   615
            Index           =   0
            Left            =   480
            Top             =   240
            Width           =   1455
         End
         Begin VB.Shape ShpPMOuter 
            BackStyle       =   1  'Opaque
            Height          =   855
            Index           =   0
            Left            =   240
            Top             =   120
            Width           =   1935
         End
         Begin VB.Shape ShpPMInner 
            BackStyle       =   1  'Opaque
            Height          =   615
            Index           =   1
            Left            =   480
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Shape ShpPMOuter 
            BackStyle       =   1  'Opaque
            Height          =   855
            Index           =   1
            Left            =   240
            Top             =   2160
            Width           =   1935
         End
      End
      Begin VB.PictureBox PictureBM 
         BackColor       =   &H80000010&
         Height          =   5895
         Index           =   1
         Left            =   14880
         ScaleHeight     =   5835
         ScaleWidth      =   2355
         TabIndex        =   51
         Top             =   1800
         Width           =   2415
         Begin VB.Frame Frame22 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   120
            TabIndex        =   179
            Top             =   4920
            Width           =   2175
            Begin VB.TextBox txtCurBM 
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
               Locked          =   -1  'True
               TabIndex        =   181
               Text            =   "0.000"
               Top             =   120
               Width           =   615
            End
            Begin VB.TextBox txtVBM 
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
               Locked          =   -1  'True
               TabIndex        =   180
               Text            =   "0.000"
               Top             =   480
               Width           =   615
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured Amp"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   55
               Left            =   120
               TabIndex        =   183
               Top             =   120
               Width           =   1275
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured V"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   54
               Left            =   120
               TabIndex        =   182
               Top             =   480
               Width           =   1005
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   120
            TabIndex        =   57
            Top             =   3000
            Width           =   2175
            Begin VB.TextBox txtCurBM 
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
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   59
               Text            =   "0.000"
               Top             =   120
               Width           =   615
            End
            Begin VB.TextBox txtVBM 
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
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   58
               Text            =   "0.000"
               Top             =   480
               Width           =   615
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured Amp"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   14
               Left            =   120
               TabIndex        =   61
               Top             =   120
               Width           =   1275
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured V"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   13
               Left            =   120
               TabIndex        =   60
               Top             =   480
               Width           =   1005
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   120
            TabIndex        =   52
            Top             =   1080
            Width           =   2175
            Begin VB.TextBox txtVBM 
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
               Locked          =   -1  'True
               TabIndex        =   54
               Text            =   "0.000"
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox txtCurBM 
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
               Locked          =   -1  'True
               TabIndex        =   53
               Text            =   "0.000"
               Top             =   120
               Width           =   615
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured V"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   12
               Left            =   120
               TabIndex        =   56
               Top             =   480
               Width           =   1005
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured Amp"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   11
               Left            =   120
               TabIndex        =   55
               Top             =   120
               Width           =   1275
            End
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RIGHT"
            Height          =   435
            Left            =   600
            TabIndex        =   64
            Top             =   4200
            Width           =   1125
         End
         Begin VB.Shape ShpBMInner 
            BackStyle       =   1  'Opaque
            Height          =   615
            Index           =   2
            Left            =   480
            Top             =   4080
            Width           =   1455
         End
         Begin VB.Shape ShpBMOuter 
            BackStyle       =   1  'Opaque
            Height          =   855
            Index           =   2
            Left            =   240
            Top             =   3960
            Width           =   1935
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "OFF"
            Height          =   435
            Left            =   840
            TabIndex        =   63
            Top             =   2280
            Width           =   720
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "LEFT"
            Height          =   375
            Index           =   15
            Left            =   720
            TabIndex        =   62
            Top             =   360
            Width           =   975
         End
         Begin VB.Shape ShpBMInner 
            BackStyle       =   1  'Opaque
            Height          =   615
            Index           =   0
            Left            =   480
            Top             =   240
            Width           =   1455
         End
         Begin VB.Shape ShpBMOuter 
            BackStyle       =   1  'Opaque
            Height          =   855
            Index           =   0
            Left            =   240
            Top             =   120
            Width           =   1935
         End
         Begin VB.Shape ShpBMInner 
            BackStyle       =   1  'Opaque
            Height          =   615
            Index           =   1
            Left            =   480
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Shape ShpBMOuter 
            BackStyle       =   1  'Opaque
            Height          =   855
            Index           =   1
            Left            =   240
            Top             =   2040
            Width           =   1935
         End
      End
      Begin VB.PictureBox PictureDM 
         BackColor       =   &H80000010&
         Height          =   4215
         Index           =   1
         Left            =   120
         ScaleHeight     =   4155
         ScaleWidth      =   2355
         TabIndex        =   37
         Top             =   1800
         Width           =   2415
         Begin VB.Frame Frame4 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   120
            TabIndex        =   46
            Top             =   3120
            Width           =   2175
            Begin VB.TextBox txtCurDM 
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
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   48
               Text            =   "0.000"
               Top             =   120
               Width           =   615
            End
            Begin VB.TextBox txtVDM 
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
               Left            =   1440
               Locked          =   -1  'True
               TabIndex        =   47
               Text            =   "0.000"
               Top             =   480
               Width           =   615
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured Amp"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   5
               Left            =   120
               TabIndex        =   50
               Top             =   120
               Width           =   1275
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured V"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   4
               Left            =   120
               TabIndex        =   49
               Top             =   480
               Width           =   1005
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H8000000B&
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   120
            TabIndex        =   40
            Top             =   1080
            Width           =   2175
            Begin VB.TextBox txtVDM 
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
               Locked          =   -1  'True
               TabIndex        =   45
               Text            =   "0.000"
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox txtCurDM 
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
               Locked          =   -1  'True
               TabIndex        =   44
               Text            =   "0.000"
               Top             =   120
               Width           =   615
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured V"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   3
               Left            =   120
               TabIndex        =   43
               Top             =   480
               Width           =   1005
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Measured Amp"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   210
               Index           =   1
               Left            =   120
               TabIndex        =   41
               Top             =   120
               Width           =   1275
            End
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HIGH"
            Height          =   435
            Left            =   720
            TabIndex        =   39
            Top             =   2400
            Width           =   900
         End
         Begin VB.Shape ShpDMInner 
            BackStyle       =   1  'Opaque
            Height          =   615
            Index           =   1
            Left            =   480
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Shape ShpDMOuter 
            BackStyle       =   1  'Opaque
            Height          =   855
            Index           =   1
            Left            =   240
            Top             =   2160
            Width           =   1935
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "LOW"
            Height          =   375
            Index           =   0
            Left            =   720
            TabIndex        =   38
            Top             =   360
            Width           =   855
         End
         Begin VB.Shape ShpDMInner 
            BackStyle       =   1  'Opaque
            Height          =   615
            Index           =   0
            Left            =   480
            Top             =   240
            Width           =   1455
         End
         Begin VB.Shape ShpDMOuter 
            BackStyle       =   1  'Opaque
            Height          =   855
            Index           =   0
            Left            =   240
            Top             =   120
            Width           =   1935
         End
      End
      Begin VB.PictureBox PictureDM 
         BackColor       =   &H80000010&
         Height          =   375
         Index           =   0
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   2355
         TabIndex        =   36
         Top             =   1440
         Width           =   2415
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dipper Module"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   360
            TabIndex        =   42
            Top             =   0
            Width           =   1665
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
         Left            =   8760
         TabIndex        =   34
         Top             =   7560
         Width           =   5175
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
         Left            =   18360
         TabIndex        =   26
         Top             =   9720
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
            Left            =   0
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
         Left            =   17400
         TabIndex        =   23
         Top             =   7680
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
         Height          =   2295
         Left            =   17280
         TabIndex        =   18
         Top             =   5520
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
         Left            =   17160
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
         Top             =   9960
         Width           =   18015
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
         Width           =   14175
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
      Begin VB.Shape ShapeWLC 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   12840
         Top             =   8280
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Wire Length OK"
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
         Index           =   2
         Left            =   11280
         TabIndex        =   137
         Top             =   8280
         Width           =   1395
      End
      Begin VB.Image ImgPart 
         Height          =   1695
         Left            =   14040
         Stretch         =   -1  'True
         Top             =   7800
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
         Left            =   7680
         TabIndex        =   106
         Top             =   9240
         Width           =   1665
      End
      Begin VB.Label Label4 
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
         Left            =   7680
         TabIndex        =   105
         Top             =   8760
         Width           =   1530
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A"
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
         Index           =   39
         Left            =   10800
         TabIndex        =   104
         Top             =   8280
         Width           =   135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ILLumination Curr. "
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
         Index           =   9
         Left            =   7680
         TabIndex        =   101
         Top             =   8280
         Width           =   1635
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
         Left            =   7680
         TabIndex        =   99
         Top             =   7680
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
      TabIndex        =   114
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
      TabIndex        =   103
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
      TabIndex        =   102
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
AssignPLCdata
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
   WinSock1.Protocol = sckTCPProtocol
   txtIP.Text = WinSock1.LocalIP
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

Private Function AssignPLCdata()
On Error GoTo Error
   MsgCode = PLcdata(108)
   ShapeColorsinglefunction PLcdata(100), &H1, ShpDMOuter(0)
   ShapeColorsinglefunction PLcdata(100), &H2, ShpDMOuter(1)
   ShapeColorsinglefunction PLcdata(100), &H4, ShpPMOuter(0)
   ShapeColorsinglefunction PLcdata(100), &H8, ShpPMOuter(1)
   ShapeColorsinglefunction PLcdata(100), &H10, ShpHAMOuter(0)
   ShapeColorsinglefunction PLcdata(100), &H20, ShpHAMOuter(1)
   ShapeColorsinglefunction PLcdata(100), &H40, ShpNMOuter(0)
   ShapeColorsinglefunction PLcdata(100), &H80, ShpNMOuter(1)
   ShapeColorsinglefunction PLcdata(100), &H100, ShpNMOuter(2)
   ShapeColorsinglefunction PLcdata(100), &H200, ShpNMOuter(3)
   ShapeColorsinglefunction PLcdata(100), &H400, ShpNMOuter(4)
   ShapeColorsinglefunction PLcdata(101), &H1, ShpBMOuter(0)
   ShapeColorsinglefunction PLcdata(101), &H2, ShpBMOuter(1)
   ShapeColorsinglefunction PLcdata(101), &H4, ShpBMOuter(2)
   ShapeColorsinglefunction PLcdata(101), &H8, ShpHOMOuter(0)
   ShapeColorsinglefunction PLcdata(101), &H10, ShpHOMOuter(1)
   ShapeColorsinglefunction PLcdata(101), &H20, ShpCRMOuter(0)
   ShapeColorsinglefunction PLcdata(101), &H40, ShpCRMOuter(1)
   ShapeColorsinglefunction PLcdata(101), &H80, ShpCUMOuter(0)
   ShapeColorsinglefunction PLcdata(101), &H100, ShpCUMOuter(1)
   ShapeColorsinglefunction PLcdata(101), &H200, ShpSRMOuter(0)
   ShapeColorsinglefunction PLcdata(101), &H400, ShpSRMOuter(1)
   
   ShapeColorfunction1 PLcdata(102), &H1, &H2, ShpDMInner(0)
   ShapeColorfunction1 PLcdata(102), &H4, &H8, ShpDMInner(1)
   ShapeColorfunction1 PLcdata(102), &H10, &H20, ShpPMInner(0)
   ShapeColorfunction1 PLcdata(102), &H40, &H80, ShpPMInner(1)
   ShapeColorfunction1 PLcdata(102), &H100, &H200, ShpHAMInner(0)
   ShapeColorfunction1 PLcdata(102), &H400, &H800, ShpHAMInner(1)
   ShapeColorfunction1 PLcdata(103), &H1, &H2, ShpNMInner(0)
   ShapeColorfunction1 PLcdata(103), &H4, &H8, ShpNMInner(1)
   ShapeColorfunction1 PLcdata(103), &H10, &H20, ShpNMInner(2)
   ShapeColorfunction1 PLcdata(103), &H40, &H80, ShpNMInner(3)
   ShapeColorfunction1 PLcdata(103), &H100, &H200, ShpNMInner(4)
   ShapeColorfunction1 PLcdata(104), &H1, &H2, ShpBMInner(0)
   ShapeColorfunction1 PLcdata(104), &H4, &H8, ShpBMInner(1)
   ShapeColorfunction1 PLcdata(104), &H10, &H20, ShpBMInner(2)
   ShapeColorfunction1 PLcdata(104), &H40, &H80, ShpHOMInner(0)
   ShapeColorfunction1 PLcdata(104), &H100, &H200, ShpHOMInner(1)
   ShapeColorfunction1 PLcdata(105), &H1, &H2, ShpCRMInner(0)
   ShapeColorfunction1 PLcdata(105), &H4, &H8, ShpCRMInner(1)
   ShapeColorfunction1 PLcdata(105), &H10, &H20, ShpCUMInner(0)
   ShapeColorfunction1 PLcdata(105), &H40, &H80, ShpCUMInner(1)
   ShapeColorfunction1 PLcdata(105), &H100, &H200, ShpSRMInner(0)
   ShapeColorfunction1 PLcdata(105), &H400, &H800, ShpSRMInner(1)

   ShapeColorfunction PLcdata(160), &H1, &H2, PictureDM(0)
   ShapeColorfunction PLcdata(160), &H4, &H8, PicturePM(0)
   ShapeColorfunction PLcdata(160), &H10, &H20, PictureHAM(0)
   ShapeColorfunction PLcdata(160), &H40, &H80, PictureNM(0)
   ShapeColorfunction PLcdata(160), &H100, &H200, PictureBM(0)
   ShapeColorfunction PLcdata(160), &H400, &H800, PictureHOM(0)
   ShapeColorfunction PLcdata(160), &H1000, &H2000, PictureCRM(0)
   ShapeColorfunction PLcdata(161), &H1, &H2, PictureCUM(0)
   ShapeColorfunction PLcdata(161), &H4, &H8, PictureSRM(0)
   
   txtCycleTime.Text = Format(PLcdata(107) / 10, "0.0")


   txtVDM(0).Text = Format(PLcdata(110) / 1000, "0.000")
   txtCurDM(0).Text = Format(PLcdata(111) / 1000, "0.000")
   txtVDM(1).Text = Format(PLcdata(112) / 1000, "0.000")
   txtCurDM(1).Text = Format(PLcdata(113) / 1000, "0.000")
   txtVPM.Text = Format(PLcdata(114) / 1000, "0.000")
   txtCurPM.Text = Format(PLcdata(115) / 1000, "0.000")
   txtVHAM.Text = Format(PLcdata(116) / 1000, "0.000")
   txtCurHAM.Text = Format(PLcdata(117) / 1000, "0.000")
   txtVNM(0).Text = Format(PLcdata(118) / 1000, "0.000")
   txtCurNM(0).Text = Format(PLcdata(119) / 1000, "0.000")
   txtVNM(1).Text = Format(PLcdata(120) / 1000, "0.000")
   txtCurNM(1).Text = Format(PLcdata(121) / 1000, "0.000")
   txtVNM(2).Text = Format(PLcdata(122) / 1000, "0.000")
   txtCurNM(2).Text = Format(PLcdata(123) / 1000, "0.000")
   txtVNM(3).Text = Format(PLcdata(124) / 1000, "0.000")
   txtCurNM(3).Text = Format(PLcdata(125) / 1000, "0.000")
   txtVBM(0).Text = Format(PLcdata(126) / 1000, "0.000")
   txtCurBM(0).Text = Format(PLcdata(127) / 1000, "0.000")
   txtVBM(1).Text = Format(PLcdata(128) / 1000, "0.000")
   txtCurBM(1).Text = Format(PLcdata(129) / 1000, "0.000")
   txtVBM(2).Text = Format(PLcdata(130) / 1000, "0.000")
   txtCurBM(2).Text = Format(PLcdata(131) / 1000, "0.000")
   txtVHOM.Text = Format(PLcdata(132) / 1000, "0.000")
   txtCurHOM.Text = Format(PLcdata(133) / 1000, "0.000")
   txtVCRM.Text = Format(PLcdata(134) / 1000, "0.000")
   txtCurCRM.Text = Format(PLcdata(135) / 1000, "0.000")
   txtVCUM.Text = Format(PLcdata(136) / 1000, "0.000")
   txtCurCUM.Text = Format(PLcdata(137) / 1000, "0.000")
   txtVSRM.Text = Format(PLcdata(138) / 1000, "0.000")
   txtCurSRM.Text = Format(PLcdata(139) / 1000, "0.000")


   
   txtILLH.Text = Format(PLcdata(150) / 1000, "0.000")
   'txtILRH.Text = Format(PLcdata(151) / 100, "0.00")
   'txtWireVoltage.Text = Format(PLcdata(152) / 100, "0.00")
   'plcdata(185) = odcurrent
   'plcdata(186) = odil
   'plcdata(187) = odmvd
   
   If (PLcdata(155) And &H1) <> 0 Then
        ShapeWLC.BackColor = vbGreen
   ElseIf (PLcdata(155) And &H2) <> 0 Then
        ShapeWLC.BackColor = vbRed
   Else
        ShapeWLC.BackColor = vbWhite
   End If
   'If (PLcdata(156) And &H1) <> 0 Then
   '     ShapeRhBolt.BackColor = vbGreen
   'ElseIf (PLcdata(156) And &H2) <> 0 Then
   '     ShapeRhBolt.BackColor = vbRed
   'Else
   '     ShapeRhBolt.BackColor = vbWhite
   'End If
   'If (PLcdata(157) And &H1) <> 0 Then
   '     ShapeLHBolt.BackColor = vbGreen
   'ElseIf (PLcdata(157) And &H2) <> 0 Then
   '     ShapeLHBolt.BackColor = vbRed
   'Else
   '     ShapeLHBolt.BackColor = vbWhite
   'End If
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

Private Sub ShapeColorfunction(Data As Integer, reg1 As Integer, reg2 As Integer, ctrl As Object)
    If (Data And reg1) Then
        If (Data And reg2) Then
           ctrl.BackColor = vbRed
        Else
           ctrl.BackColor = vbYellow
         End If
    ElseIf (Data And reg2) Then
          ctrl.BackColor = vbGreen
    Else
          ctrl.BackColor = vbWhite
    End If
End Sub
Private Sub ShapeColorfunction1(Data As Integer, reg1 As Integer, reg2 As Integer, ctrl As Object)
    If (Data And reg1) Then
        If (Data And reg2) Then
           ctrl.BackColor = vbRed
        Else
           ctrl.BackColor = vbGreen
         End If
    ElseIf (Data And reg2) Then
          ctrl.BackColor = vbRed
    Else
          ctrl.BackColor = vbWhite
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
        sendEmail
    Else
        shapeInternet.BackColor = vbRed
    End If
    
    Text1.Text = WinsockStstus(WinSock1.State)


    If WinSock1.State = 7 Then
        ShapePLCState.BackColor = vbGreen
    Else
        ShapePLCState.BackColor = vbRed
    End If
    Dim Description As String
    
    Select Case WinSock1.State
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
    Sql = "select Top 1 * from model_report_counter where MailSent = false order by id desc"
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
    'txtModelDesc.Text = Trim(Rs("ModelDesc"))
    PLcdata(210) = Val(Rs("DMBypass"))
    If Val(txtCouplerCounter.Text) >= setCouplerCounter Then
        PLcdata(335) = 1
    ElseIf Val(txtBatchCounter.Text) >= setBatchCounter Then
        PLcdata(335) = 2
    Else
        PLcdata(335) = 0
    End If
   If Val(Rs("DMBypass")) = 0 Then
    PLcdata(213) = Val(Rs("DM1CurMin")) * 1000
    PLcdata(214) = Val(Rs("DM1CurMax")) * 1000
    PLcdata(217) = Val(Rs("DM2CurMin")) * 1000
    PLcdata(218) = Val(Rs("DM2CurMax")) * 1000
    PLcdata(211) = Val(Rs("DM1VoltMin")) * 1000
    PLcdata(212) = Val(Rs("DM1VoltMax")) * 1000
    PLcdata(215) = Val(Rs("DM2VoltMin")) * 1000
    PLcdata(216) = Val(Rs("DM2VoltMax")) * 1000
    PLcdata(219) = Val(Rs("DMTestCycle"))
  Else
    PictureDM(0).Visible = False
    PictureDM(1).Visible = False
    PLcdata(211) = 0
    PLcdata(212) = 0
    PLcdata(213) = 0
    PLcdata(214) = 0
    PLcdata(215) = 0
    PLcdata(216) = 0
    PLcdata(217) = 0
    PLcdata(218) = 0
    PLcdata(219) = 0
   End If
    PLcdata(220) = Val(Rs("PMBypass"))
   If Val(Rs("PMBypass")) = 0 Then
    PLcdata(223) = Val(Rs("PM1CurMin")) * 1000
    PLcdata(224) = Val(Rs("PM1CurMax")) * 1000
    PLcdata(221) = Val(Rs("PM1VoltMin")) * 1000
    PLcdata(222) = Val(Rs("PM1VoltMax")) * 1000
    PLcdata(225) = Val(Rs("PMTestCycle"))
   Else
    PicturePM(0).Visible = False
    PicturePM(1).Visible = False
    PLcdata(221) = 0
    PLcdata(222) = 0
    PLcdata(223) = 0
    PLcdata(224) = 0
    PLcdata(225) = 0
   End If
    
     PLcdata(230) = Val(Rs("HAMBypass"))
   If Val(Rs("HAMBypass")) = 0 Then
    PLcdata(233) = Val(Rs("HAM1CurMin")) * 1000
    PLcdata(234) = Val(Rs("HAM1CurMax")) * 1000
    PLcdata(231) = Val(Rs("HAM1VoltMin")) * 1000
    PLcdata(232) = Val(Rs("HAM1VoltMax")) * 1000
    PLcdata(235) = Val(Rs("HAMTestCycle"))
   Else
    PictureHAM(0).Visible = False
    PictureHAM(1).Visible = False
    PLcdata(231) = 0
    PLcdata(232) = 0
    PLcdata(233) = 0
    PLcdata(234) = 0
    PLcdata(235) = 0
   End If
   
    PLcdata(240) = Val(Rs("NMBypass"))
   If Val(Rs("NMBypass")) = 0 Then
    PLcdata(241) = Val(Rs("NM1VoltMin")) * 1000
    PLcdata(242) = Val(Rs("NM1VoltMax")) * 1000
    PLcdata(243) = Val(Rs("NM1CurMin")) * 1000
    PLcdata(244) = Val(Rs("NM1CurMax")) * 1000
    PLcdata(245) = Val(Rs("NM2VoltMin")) * 1000
    PLcdata(246) = Val(Rs("NM2VoltMax")) * 1000
    PLcdata(247) = Val(Rs("NM2CurMin")) * 1000
    PLcdata(248) = Val(Rs("NM2CurMax")) * 1000
    PLcdata(249) = Val(Rs("NM3VoltMin")) * 1000
    PLcdata(250) = Val(Rs("NM3VoltMax")) * 1000
    PLcdata(251) = Val(Rs("NM3CurMin")) * 1000
    PLcdata(252) = Val(Rs("NM3CurMax")) * 1000
    PLcdata(253) = Val(Rs("NM4VoltMin")) * 1000
    PLcdata(254) = Val(Rs("NM4VoltMax")) * 1000
    PLcdata(255) = Val(Rs("NM4CurMin")) * 1000
    PLcdata(256) = Val(Rs("NM4CurMax")) * 1000

    PLcdata(257) = Val(Rs("NMTestCycle"))
   Else
    PictureNM(0).Visible = False
    PictureNM(1).Visible = False
    PLcdata(241) = 0
    PLcdata(242) = 0
    PLcdata(243) = 0
    PLcdata(244) = 0
    PLcdata(245) = 0
    PLcdata(246) = 0
    PLcdata(247) = 0
    PLcdata(248) = 0
    PLcdata(249) = 0
    PLcdata(250) = 0
    PLcdata(251) = 0
    PLcdata(252) = 0
    PLcdata(253) = 0
    PLcdata(254) = 0
    PLcdata(255) = 0
    PLcdata(256) = 0
    PLcdata(257) = 0
   End If
   
    PLcdata(260) = Val(Rs("BMBypass"))
   If Val(Rs("BMBypass")) = 0 Then
    PLcdata(261) = Val(Rs("BM1VoltMin")) * 1000
    PLcdata(262) = Val(Rs("BM1VoltMax")) * 1000
    PLcdata(263) = Val(Rs("BM1CurMin")) * 1000
    PLcdata(264) = Val(Rs("BM1CurMax")) * 1000
    PLcdata(265) = Val(Rs("BM2VoltMin")) * 1000
    PLcdata(266) = Val(Rs("BM2VoltMax")) * 1000
    PLcdata(267) = Val(Rs("BM2CurMin")) * 1000
    PLcdata(268) = Val(Rs("BM2CurMax")) * 1000
    PLcdata(269) = Val(Rs("BM3VoltMin")) * 1000
    PLcdata(270) = Val(Rs("BM3VoltMax")) * 1000
    PLcdata(271) = Val(Rs("BM3CurMin")) * 1000
    PLcdata(272) = Val(Rs("BM3CurMax")) * 1000
    PLcdata(273) = Val(Rs("BMTestCycle"))
   Else
    PictureBM(0).Visible = False
    PictureBM(1).Visible = False
    PLcdata(261) = 0
    PLcdata(262) = 0
    PLcdata(263) = 0
    PLcdata(264) = 0
    PLcdata(265) = 0
    PLcdata(266) = 0
    PLcdata(267) = 0
    PLcdata(268) = 0
    PLcdata(269) = 0
    PLcdata(270) = 0
    PLcdata(271) = 0
    PLcdata(272) = 0
    PLcdata(273) = 0
   End If
   
    PLcdata(275) = Val(Rs("HOMBypass"))
   If Val(Rs("HOMBypass")) = 0 Then
    PLcdata(276) = Val(Rs("HOM1VoltMin")) * 1000
    PLcdata(277) = Val(Rs("HOM1VoltMax")) * 1000
    PLcdata(278) = Val(Rs("HOM1CurMin")) * 1000
    PLcdata(279) = Val(Rs("HOM1CurMax")) * 1000
    PLcdata(280) = Val(Rs("HOMTestCycle"))
   Else
    PictureHOM(0).Visible = False
    PictureHOM(1).Visible = False
    PLcdata(276) = 0
    PLcdata(277) = 0
    PLcdata(278) = 0
    PLcdata(279) = 0
    PLcdata(280) = 0
   End If
    PLcdata(285) = Val(Rs("CRMBypass"))
   If Val(Rs("CRMBypass")) = 0 Then
    PLcdata(286) = Val(Rs("CRM1VoltMin")) * 1000
    PLcdata(287) = Val(Rs("CRM1VoltMax")) * 1000
    PLcdata(288) = Val(Rs("CRM1CurMin")) * 1000
    PLcdata(289) = Val(Rs("CRM1CurMax")) * 1000
    PLcdata(290) = Val(Rs("CRMTestCycle"))
   Else
    PictureCRM(0).Visible = False
    PictureCRM(1).Visible = False
    PLcdata(286) = 0
    PLcdata(287) = 0
    PLcdata(288) = 0
    PLcdata(289) = 0
    PLcdata(290) = 0
   End If
    PLcdata(295) = Val(Rs("CUMBypass"))
   If Val(Rs("CUMBypass")) = 0 Then
    PLcdata(296) = Val(Rs("CUM1VoltMin")) * 1000
    PLcdata(297) = Val(Rs("CUM1VoltMax")) * 1000
    PLcdata(298) = Val(Rs("CUM1CurMin")) * 1000
    PLcdata(299) = Val(Rs("CUM1CurMax")) * 1000
    PLcdata(300) = Val(Rs("CUMTestCycle"))
   Else
    PictureCUM(0).Visible = False
    PictureCUM(1).Visible = False
    PLcdata(296) = 0
    PLcdata(297) = 0
    PLcdata(298) = 0
    PLcdata(299) = 0
    PLcdata(300) = 0
   End If
    PLcdata(305) = Val(Rs("SRMBypass"))
   If Val(Rs("SRMBypass")) = 0 Then
    PLcdata(306) = Val(Rs("SRM1VoltMin")) * 1000
    PLcdata(307) = Val(Rs("SRM1VoltMax")) * 1000
    PLcdata(308) = Val(Rs("SRM1CurMin")) * 1000
    PLcdata(309) = Val(Rs("SRM1CurMax")) * 1000
    PLcdata(310) = Val(Rs("SRMTestCycle"))
   Else
    PictureSRM(0).Visible = False
    PictureSRM(1).Visible = False
    PLcdata(306) = 0
    PLcdata(307) = 0
    PLcdata(308) = 0
    PLcdata(309) = 0
    PLcdata(310) = 0
   End If
   
   
    For i = 0 To 8
     PLcdata(351 + i) = Val(Rs("CurrentOffset" & i + 1)) * 1000
     PLcdata(361 + i) = Val(Rs("VoltageOffset" & i + 1)) * 1000
    Next
    
    PLcdata(325) = Val(Rs("ICMin")) * 1000
    PLcdata(326) = Val(Rs("ICMax")) * 1000
    'PLcdata(327) = Val(Rs("ICMinRH")) * 100
    'PLcdata(328) = Val(Rs("ICMaxRH")) * 100
    
    'PLcdata(318) = Val(Rs("WVMin")) * 100
    'PLcdata(319) = Val(Rs("WVMax")) * 100
    'PartNo = Rs("PrintPartNo")
    'BarcodeLength = Rs("BarcodeLength")
    'HardwareNo = Rs("HardwareNo")
    'SerialStartingtxt = Rs("SerialStartingtxt")
    
    PLcdata(320) = Val(Rs("DebounceTime")) * 100
    PLcdata(321) = Val(Rs("HoldTime")) * 100
    PLcdata(322) = Val(Rs("CheckTime")) * 1000
    
    PLcdata(332) = Val(Rs("DotMarkingTime")) * 10
    
    ModelNo = Rs("ModelNo")
    PLcdata(331) = Rs("ModelNo")
    
    'Rs("BatchCounter").Text
    'Rs("CouplerCounter") = .Text
    'Rs ("PartImage")
    'Rs("productioncounter") =
    
    PLcdata(330) = 0
    PLcdata(330) = PLcdata(330) + &H1 * Val(Rs("Bypass1"))
    PLcdata(330) = PLcdata(330) + &H2 * Val(Rs("Bypass2"))
    PLcdata(330) = PLcdata(330) + &H4 * Val(Rs("Bypass3"))
    PLcdata(330) = PLcdata(330) + &H8 * Val(Rs("Bypass4"))
    PLcdata(330) = PLcdata(330) + &H10 * Val(Rs("Bypass5"))
    PLcdata(330) = PLcdata(330) + &H20 * Val(Rs("Bypass6"))
    PLcdata(330) = PLcdata(330) + &H40 * Val(Rs("Bypass7"))
    PLcdata(330) = PLcdata(330) + &H80 * Val(Rs("ByPass8"))
    PLcdata(330) = PLcdata(330) + &H100 * Val(Rs("ByPass9"))
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
            PLcdata(349) = 1
            Exit Sub
        Else
            If tempgetshift <> tempshift Then
                txtTargetProduction.Locked = False
                txtTargetProduction.Text = ""
                txtTargetProduction.SetFocus
                txtTargetProduction.BackColor = vbYellow
                Command1.Visible = True
                PLcdata(349) = 1
                Exit Sub
            End If
        End If
    ElseIf txtTargetProduction.BackColor <> vbYellow Then
        txtTargetProduction.Locked = False
        txtTargetProduction.Text = ""
        txtTargetProduction.SetFocus
        txtTargetProduction.BackColor = vbYellow
        Command1.Visible = True
        PLcdata(349) = 1
        
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
    If Val(Rs("bypass3")) = 1 Then
      Label1(2).Visible = False
      ShapeWLC.Visible = False
    End If
    If Val(Rs("bypass6")) = 1 Then
        txtILLH.Visible = False
        Label4(9).Visible = False
        Label4(39).Visible = False
        
'        ShapeRhBolt.Visible = False
'        Label1(2).Visible = False
    End If
    If Val(Rs("bypass10")) = 1 Then
            
    End If
    
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
      Rs("DMResult") = getresult(PictureDM(0))
      Rs("DM1Cur") = txtCurDM(0).Text
      Rs("DM1Volt") = txtVDM(0).Text
      Rs("DM2Cur") = txtCurDM(1).Text
      Rs("DM2Volt") = txtVDM(1).Text
      Rs("PMResult") = getresult(PicturePM(0))
      Rs("PM1Cur") = txtCurPM.Text
      Rs("PM1Volt") = txtVPM.Text
      Rs("HAMResult") = getresult(PictureHAM(0))
      Rs("HAM1Cur") = txtCurHAM.Text
      Rs("HAM1Volt") = txtVHAM.Text
      Rs("NMResult") = getresult(PictureNM(0))
      Rs("NM1Cur") = txtCurNM(0).Text
      Rs("NM1Volt") = txtVNM(0).Text
      Rs("NM2Cur") = txtCurNM(1).Text
      Rs("NM2Volt") = txtVNM(1).Text
      Rs("NM3Cur") = txtCurNM(2).Text
      Rs("NM3Volt") = txtVNM(2).Text
      Rs("NM4Cur") = txtCurNM(2).Text
      Rs("NM4Volt") = txtVNM(2).Text
      Rs("BMResult") = getresult(PictureBM(0))
      Rs("BM1Cur") = txtCurBM(0).Text
      Rs("BM1Volt") = txtVBM(0).Text
      Rs("BM2Cur") = txtCurBM(1).Text
      Rs("BM2Volt") = txtVBM(1).Text
      Rs("BM3Cur") = txtCurBM(2).Text
      Rs("BM3Volt") = txtVBM(2).Text
      Rs("HOMResult") = getresult(PictureHOM(0))
      Rs("HOM1Cur") = txtCurHOM.Text
      Rs("HOM1Volt") = txtVHOM.Text
      Rs("CUMResult") = getresult(PictureCUM(0))
      Rs("CUM1Cur") = txtCurCUM.Text
      Rs("CUM1Volt") = txtVCUM.Text
      Rs("CRMResult") = getresult(PictureCRM(0))
      Rs("CRM1Cur") = txtCurCRM.Text
      Rs("CRM1Volt") = txtVCRM.Text
      Rs("SRMResult") = getresult(PictureSRM(0))
      Rs("SRM1Cur") = txtCurSRM.Text
      Rs("SRM1volt") = txtVSRM.Text
      'Rs("ICLH") = txtILLH.Text
      'Rs("ICRH") = txtILRH.Text
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
   WinSock1.Close
   WinSock1.RemoteHost = txtIP_Host.Text
   WinSock1.RemotePort = txtPort.Text
   WinSock1.Connect
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
   If (WinSock1.State = 7) And (CommandOn = False) Then
      Timer1.Enabled = False
      Select Case CommandType
         Case 1
            Call GetReadArray(StdReadStartAddress, StdReadCount, ReadArray)
            WinSock1.SendData ReadArray
            CVRead = CVRead + 1
            CommandOn = True
            Timer5.Interval = 800
            Timer5.Enabled = True
         Case 2
            Call GetWriteArray(StdWriteStartAddress, StdWriteCount, WriteArray)
            WinSock1.SendData WriteArray
            CommandOn = True
            Timer5.Interval = 800
            Timer5.Enabled = True
         Case 3
            Call GetReadArray((ExtendedReadStartAddress + (ExtendedReadCount * CVExtPktNo)), ExtendedReadCount, ReadArray)
            WinSock1.SendData ReadArray
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

   If (WinSock1.State <> 7) Then 'And (WinSock1.State <> 6) Then
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
   WinSock1.GetData SocketData
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
