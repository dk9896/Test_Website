VERSION 5.00
Begin VB.Form frmReportType 
   ClientHeight    =   6990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   5895
      Left            =   240
      ScaleHeight     =   5835
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   360
      Width           =   7935
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080C0FF&
         Caption         =   "BreakDown Summary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3960
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Daily Production"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2040
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Production Summary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2040
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Production Analysis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2040
         Width           =   1935
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H000040C0&
         Height          =   855
         Left            =   120
         ScaleHeight     =   795
         ScaleWidth      =   7515
         TabIndex        =   1
         Top             =   120
         Width           =   7575
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "SELECT REPORT TYPE"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   615
            Left            =   240
            TabIndex        =   2
            Top             =   0
            Width           =   7095
         End
      End
   End
End
Attribute VB_Name = "frmReportType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
 reporttype = 1
 frmReport.Show
 Unload Me
End Sub

Private Sub Command3_Click()
 reporttype = 3
 frmReport.Show
 Unload Me
End Sub

Private Sub Command4_Click()
 reporttype = 2
 frmReport.Show
 Unload Me
End Sub
