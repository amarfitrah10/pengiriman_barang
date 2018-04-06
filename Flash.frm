VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form3 
   BackColor       =   &H8000000D&
   Caption         =   "Form3"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12135
   ControlBox      =   0   'False
   FillColor       =   &H008080FF&
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form3"
   Picture         =   "Flash.frx":0000
   ScaleHeight     =   7410
   ScaleWidth      =   12135
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   735
      Left            =   1200
      TabIndex        =   0
      Top             =   3720
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   1296
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1200
      Top             =   2160
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      Height          =   1695
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12255
   End
   Begin VB.Label progress 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   9600
      TabIndex        =   2
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Selamat Datang Di Pengiriman Barang "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   1680
      Width           =   7455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 2
progress.Caption = ProgressBar1.Value & "%"
If ProgressBar1.Value = ProgressBar1.Max Then
Unload Me
Menuutama.Show
End If
End Sub
