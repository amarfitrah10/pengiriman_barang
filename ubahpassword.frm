VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H8000000D&
   Caption         =   "Form4"
   ClientHeight    =   6105
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12075
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   ScaleHeight     =   6105
   ScaleWidth      =   12075
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton save 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   8
      Top             =   4920
      Width           =   2055
   End
   Begin VB.TextBox konfirmasi 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   3720
      Width           =   3015
   End
   Begin VB.TextBox passwordbaru 
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   2760
      Width           =   3135
   End
   Begin VB.TextBox passwordlama 
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox kodeuser 
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "UBAH PASSWORD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      TabIndex        =   9
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label4 
      Caption         =   "Konfirmasi Password Baru"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Password Baru"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Password Lama"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Kode User"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
kodeuser.Text = Menuutama.StatusBar1.Panels(1).Text
kodeuser.Enabled = True
passwordbaru.Enabled = True
konfirmasi.Enabled = True
save.Enabled = True
passwordlama.SetFocus
End Sub
Private Sub passwordlama_KeyPress(KeyAscii As Integer)
Call koneksi
If KeyAscii = 13 Then
Rsuser.Open "select * from user where password='" & passwordlama.Text & "'", con
If Rsuser.EOF Then
MsgBox "Password Lama Salah !", vbOKOnly, "info"
passwordlama.Text = ""
Else
MsgBox "Masukan Password Baru Anda", vbOKOnly, "info"
passwordbaru.Enabled = True
konfirmasi.Enabled = True
passwordbaru.SetFocus
save.Enabled = True
End If
End If
End Sub

Private Sub save_Click()
Call koneksi
If KeyAscii = 13 Then
If konfirmasi.Text <> passwordbaru.Text Then
MsgBox "Maaf Konfirmasi Password Tidak Sama", vbOKOnly, "Info"
konfirmasi.Text = ""
passwordbaru.Text = ""
passwordbaru.SetFocus
Else
Rsuser.Open "Update user set password ='" & konfirmasi.Text & "' where kode_user='" & kodeuser.Text & "'", con
konfirmasi.Text = ""
passwordbaru.Text = ""
Unload Me
End If
End If
End Sub
