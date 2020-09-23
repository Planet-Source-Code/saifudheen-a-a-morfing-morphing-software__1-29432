VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Morfy"
   ClientHeight    =   2085
   ClientLeft      =   4335
   ClientTop       =   7440
   ClientWidth     =   4605
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4605
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Left            =   4140
      Top             =   1770
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   345
      Left            =   2040
      TabIndex        =   0
      Top             =   1740
      Width           =   1185
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C0C0C0&
      X1              =   0
      X2              =   6000
      Y1              =   690
      Y2              =   690
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   6030
      Y1              =   660
      Y2              =   660
   End
   Begin VB.Label lblVMemory 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2010
      TabIndex        =   9
      Top             =   1290
      Width           =   2595
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Available Virtual Memory :"
      Height          =   195
      Left            =   0
      TabIndex        =   8
      Top             =   1290
      Width           =   1815
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Available Physical Memory :"
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   1050
      Width           =   1965
   End
   Begin VB.Label Label7 
      Caption         =   "Global Memory Status :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1230
      TabIndex        =   6
      Top             =   780
      Width           =   2265
   End
   Begin VB.Label lblPMemory 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2010
      TabIndex        =   5
      Top             =   1050
      Width           =   2625
   End
   Begin VB.Label Label4 
      Caption         =   "Copy write (2001) Saifsoft Inc."
      Height          =   195
      Left            =   870
      TabIndex        =   4
      Top             =   330
      Width           =   3405
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   30
      X2              =   5760
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   30
      X2              =   5700
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label3 
      Caption         =   "keraleeyan@msn.com"
      Height          =   195
      Left            =   30
      TabIndex        =   3
      Top             =   1830
      Width           =   1965
   End
   Begin VB.Label Label2 
      Caption         =   "Morfy. ( Vertion 1.0 )"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   870
      TabIndex        =   2
      Top             =   90
      Width           =   3465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Author : A.A. Saifudheeen."
      Height          =   195
      Left            =   30
      TabIndex        =   1
      Top             =   1650
      Width           =   1890
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "frmAbout.frx":0442
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Dim MemStat As MEMORYSTATUS

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    GlobalMemoryStatus MemStat
    lblPMemory.Caption = MemStat.dwAvailPhys / 1024 & " KB /" & MemStat.dwTotalPhys / 1024 & " KB "
    lblVMemory.Caption = MemStat.dwAvailVirtual / 1024 & " KB /" & MemStat.dwTotalVirtual / 1024 & " KB "
End Sub

Private Sub lblMemory_Click()

End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    GlobalMemoryStatus MemStat
    lblPMemory.Caption = MemStat.dwAvailPhys / 1024 & " KB /" & MemStat.dwTotalPhys / 1024 & " KB "
    lblVMemory.Caption = MemStat.dwAvailVirtual / 1024 & " KB /" & MemStat.dwTotalVirtual / 1024 & " KB "

End Sub
