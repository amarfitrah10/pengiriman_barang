VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm Menuutama 
   BackColor       =   &H8000000D&
   Caption         =   "MENU"
   ClientHeight    =   6210
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   9645
   LinkTopic       =   "MDIForm"
   Picture         =   "menuutama.frx":0000
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   5595
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   1085
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu mnadmin 
         Caption         =   "Data User"
      End
      Begin VB.Menu mntarif 
         Caption         =   "Data Tarif"
      End
   End
   Begin VB.Menu Transaksi 
      Caption         =   "Transaksi"
      Begin VB.Menu mntransaksi 
         Caption         =   "Transaksi"
      End
      Begin VB.Menu mnretur 
         Caption         =   "Retur"
      End
   End
   Begin VB.Menu Laporan 
      Caption         =   "Laporan"
      Begin VB.Menu mntrans 
         Caption         =   "Laporan Transaksi"
      End
      Begin VB.Menu mnlap 
         Caption         =   "Laporan Pelanggan"
      End
      Begin VB.Menu lapretur 
         Caption         =   "Laporan Retur"
      End
   End
   Begin VB.Menu Utility 
      Caption         =   "Utility"
      Begin VB.Menu mnpass 
         Caption         =   "Ubah Password"
      End
   End
   Begin VB.Menu keluar 
      Caption         =   "Keluar"
      Begin VB.Menu mnkeluar 
         Caption         =   "Logout"
      End
      Begin VB.Menu mnexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Menuutama"
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

Private Sub lapretur_Click()
Form9.Show
End Sub

Private Sub MDIForm_Activate()
Call koneksi
Rsuser.Open "select * from user where Kode_user ='" & Menuutama.StatusBar1.Panels(1).Text & "'", con
If Not Rsuser.EOF Then
Call tampiluser
End If
End Sub

Private Sub tampiluser()
Menuutama.StatusBar1.Panels(2) = Rsuser!hak_akses
End Sub

Private Sub MDIForm_Load()
Menuutama.StatusBar1.Panels(1) = Form1.kodeuser.Text
Menuutama.StatusBar1.Panels(2) = Rsuser!hak_akses
If Menuutama.StatusBar1.Panels(2).Text = "ADMIN" Then
File.Visible = True
mntransaksi.Visible = False
Transaksi.Visible = False
mntransaksi.Enabled = False
Laporan.Enabled = True
Utility.Enabled = True
Utility.Visible = True
Else
If Menuutama.StatusBar1.Panels(2).Text = "KASIR" Then
File.Visible = False
Laporan.Visible = True
mntransaksi.Enabled = True
Utility.Enabled = False
Utility.Visible = False
End If
End If
End Sub

Private Sub mnadmin_Click()
Form1.Show
End Sub

Private Sub mnexit_Click()
Menuutama.Show
End Sub

Private Sub mnkeluar_Click()
Form2.Show
Unload Me
End Sub

Private Sub mnlap_Click()
Form11.Show
End Sub

Private Sub mnpass_Click()
Form4.Show
End Sub

Private Sub mnpelanggan_Click()
Form6.Show
End Sub

Private Sub mnretur_Click()
Form8.Show
End Sub

Private Sub mntarif_Click()
Form5.Show
End Sub

Private Sub mntrans_Click()
Form10.Show

End Sub

Private Sub mntransaksi_Click()
Form6.Show
End Sub



