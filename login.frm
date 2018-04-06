VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H8000000A&
   Caption         =   "Form2"
   ClientHeight    =   4455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11265
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   Picture         =   "login.frx":0000
   ScaleHeight     =   4455
   ScaleWidth      =   11265
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   360
      Left            =   5040
      Top             =   240
   End
   Begin VB.TextBox username 
      BackColor       =   &H8000000D&
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
      Left            =   4560
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox password 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   4560
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label jam 
      Height          =   255
      Left            =   8160
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label tanggal 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mypic As PictureBox
Dim stdpic As New StdPicture
Sub RedrawScreen()
mypic.Width = Me.ScaleWidth
mypic.Height = Me.ScaleHeight
mypic.BorderStyle = 0
mypic.Visible = False
mypic.AutoRedraw = True
mypic.PaintPicture stdpic, 0, 0, mypic.Width, mypic.Height
Set Me.Picture = mypic.Image
End Sub

Private Sub form_Resize()
Call RedrawScreen
End Sub
Private Sub Form_Load()
Set mypic = Me.Controls.Add("VB.picturebox", "mypic")
Set stdpic = Me.Picture
Call RedrawScreen
End Sub
Private Sub Form_Activate()
    username.SetFocus
username.Enabled = True
password.Enabled = False
End Sub
Private Sub login_Click()
Call koneksi
Rsuser.Open "Select * from user where kode_user='" & username.Text & "' AND password='" & password.Text & "'", con
If Rsuser.EOF Then
MsgBox "Username Atau Password Mohon Diisi !", vbOKOnly, "Info"
username.Text = ""
password.Text = ""
username.SetFocus
password.Enabled = False
Else
Form3.Show
Menuutama.StatusBar1.Panels(1) = Rsuser!Kode_User
Menuutama.StatusBar1.Panels(2) = Rsuser!Nama_User
Unload Me
End If
End Sub
Private Sub password_KeyPress(KeyAscii As Integer)
Call koneksi
If KeyAscii = 13 Then
Rsuser.Open "select *from user where kode_user='" & username.Text & "' AND password='" & password.Text & "'", con
If Rsuser.EOF Then
    MsgBox "Password Anda Salah!"
    password.Text = ""
Else
Form3.Show
Menuutama.StatusBar1.Panels(1) = Rsuser!Kode_User
Menuutama.StatusBar1.Panels(2) = Rsuser!Nama_User
Unload Me
End If
End If
End Sub


Private Sub username_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase((Chr(KeyAscii))))
Call koneksi
password.Enabled = True
If KeyAscii = 13 Then
password.SetFocus
Rsuser.Open "select * from user where kode_user='" & username.Text & "'", con
If Not Rsuser.EOF Then
Else
MsgBox "username tidak terdaftar"
username.SetFocus
username = ""
End If
End If
End Sub

Private Sub Timer1_Timer()
   jam.Caption = Time
   Tanggal.Caption = Format(Date, "dd/mm/yyyy")
   
End Sub
