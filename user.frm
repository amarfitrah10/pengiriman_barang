VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   "s"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12615
   ControlBox      =   0   'False
   Picture         =   "user.frx":0000
   ScaleHeight     =   7590
   ScaleWidth      =   12615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   9600
      TabIndex        =   28
      Top             =   600
      Width           =   2535
      Begin VB.TextBox tgl 
         Height          =   285
         Left            =   1200
         TabIndex        =   30
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Tanggal"
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
         TabIndex        =   29
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   120
      TabIndex        =   23
      Top             =   360
      Width           =   3015
      Begin VB.TextBox user 
         Height          =   285
         Left            =   1560
         TabIndex        =   25
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox kduser 
         Height          =   285
         Left            =   1560
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Nama User"
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
         TabIndex        =   27
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label8 
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
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.ComboBox hakakses 
      Height          =   315
      Left            =   2760
      TabIndex        =   21
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   9720
      TabIndex        =   19
      Top             =   2040
      Width           =   1815
      Begin VB.CommandButton cmdcari 
         Caption         =   "Cari User"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         Picture         =   "user.frx":6F8B
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSDataGridLib.DataGrid grid_user 
      Height          =   1215
      Left            =   0
      TabIndex        =   18
      Top             =   6240
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   2143
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Button"
      Height          =   3975
      Left            =   5520
      TabIndex        =   11
      Top             =   1920
      Width           =   3855
      Begin VB.CommandButton exit 
         Caption         =   "Keluar"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         Picture         =   "user.frx":798D
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton hapus 
         Caption         =   "Hapus"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         Picture         =   "user.frx":838F
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton edit 
         Caption         =   "Ubah"
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
         Left            =   2280
         TabIndex        =   15
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cancel 
         Caption         =   "Batal"
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
         Left            =   360
         TabIndex        =   14
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton save 
         Caption         =   "Simpan"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         Picture         =   "user.frx":8D91
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton baru 
         Caption         =   "Tambah"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         Picture         =   "user.frx":9793
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.TextBox pass 
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox nohp 
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox alamat 
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox namauser 
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox kodeuser 
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "User"
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
      Left            =   4800
      TabIndex        =   31
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000012&
      Caption         =   "Label7"
      Height          =   1695
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   12135
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000D&
      Caption         =   "Hak Akses"
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
      Left            =   240
      TabIndex        =   5
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000D&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000D&
      Caption         =   "No.Handphone"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Caption         =   "Alamat"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Caption         =   "Kode User"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Caption         =   "Nama User"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2640
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aktif()
    kodeuser.Enabled = True
    namauser.Enabled = True
    alamat.Enabled = True
    nohp.Enabled = True
    tgl.Text = Format(Date, "YYYY-MM-DD")
    pass.Enabled = True
    hakakses.Enabled = True
End Sub
Private Sub tidakaktif()
    kodeuser.Enabled = False
    namauser.Enabled = False
    alamat.Enabled = False
    tgl.Text = Format(Date, "YYYY-MM-DD")
    nohp.Enabled = False
    pass.Enabled = False
    hakakses.Enabled = False
    save.Enabled = False
    edit.Enabled = False
    hapus.Enabled = False
End Sub
Private Sub kosong()
    kodeuser.Text = ""
    namauser.Text = ""
    alamat.Text = ""
    nohp.Text = ""
    pass.Text = ""
    hakakses.Text = ""
End Sub
Private Sub tampil()
    kodeuser.Text = Rsuser!Kode_User
    namauser.Text = Rsuser!Nama_User
    alamat.Text = Rsuser!alamat
    nohp.Text = Rsuser!No_Handphone
    pass.Text = Rsuser!password
    hakakses.Text = Rsuser!hak_akses
End Sub
Private Sub cancel_Click()
 Call Form_Activate
 save.Enabled = True
End Sub
Private Sub Form_Load()
hakakses.AddItem "ADMIN"
hakakses.AddItem "KASIR"
End Sub

Private Sub hakakses_Click()
nomer = "select Kode_user from user where left (Kode_user, 3)='" & "ADM" & "'order by Kode_user desc"
Set Rsuser = con.Execute(nomer, , adCmdText)
Dim urutan As String * 10
If Not Rsuser.EOF Then
   hitung = Right(Rsuser!Kode_User, 2) + 1
 Select Case hakakses
 Case "ADMIN"
 kodeuser.Text = "ADM" + "0" & (Trim(Str(hitung)))
 namauser.Enabled = True
 namauser.SetFocus
 Exit Sub
 End Select
 End If
nomer = "select Kode_user from user where left (Kode_user, 3)='" & "KSR" & "'order by Kode_user desc"
Set Rsuser = con.Execute(nomer, , adCmdText)
If Not Rsuser.EOF Then
   hitung = Right(Rsuser!Kode_User, 2) + 1
 Select Case hakakses
 Case "KASIR"
 kodeuser.Text = "KSR" + "0" & (Trim(Str(hitung)))
 namauser.Enabled = True
 namauser.SetFocus
 End Select
 End If
 End Sub
Private Sub hapus_Click()
Dim hapus As String
    hapus = "DELETE FROM user WHERE kode_user = '" & kodeuser.Text & "'"
    con.Execute hapus
    MsgBox "Yakin Data Mau Dihapus !", vbYesNo, "Info"
    Call Form_Activate
    save.Enabled = False
    baru.Enabled = True
End Sub
Private Sub edit_Click()
Dim update As String
    update = "UPDATE user SET nama_user = '" & namauser.Text & "', alamat = '" & alamat.Text & "', no_handphone = '" & nohp.Text & "', hak_akses='" & hakakses.Text & "', password ='" & pass.Text & "' WHERE kode_user = '" & kodeuser.Text & "'"
    con.Execute update
    MsgBox "Data berhasil diubah !", vbOKOnly, "Info"
    Call Form_Activate
    save.Enabled = False
    baru.Enabled = True
End Sub
Private Sub exit_click()
Unload Me
End Sub
Private Sub Form_Activate()
    Call kosong
    Call tidakaktif
    Call koneksi
    kduser.Text = Menuutama.StatusBar1.Panels(1).Text
    user.Text = Menuutama.StatusBar1.Panels(2).Text
    baru.Enabled = True

    save.Enabled = False
    edit.Enabled = False
    hapus.Enabled = False
    cancel.Enabled = False
    Rsuser.Open "SELECT * FROM user", con
    Set grid_user.DataSource = Rsuser.DataSource
    
Call tampiluser
    baru.SetFocus
grid_user.Columns(2).Width = 3000

End Sub
Private Sub baru_Click()
    Call aktif
    Call kosong
    save.Enabled = False
    edit.Enabled = False
    hapus.Enabled = False
    hakakses.SetFocus
End Sub
Private Sub grid_user_Click()
    Call tampil
    Call aktif
    kodeuser.Enabled = False
    namauser.Enabled = False
    save.Enabled = False
    baru.Enabled = False
    edit.Enabled = True
    hapus.Enabled = True
End Sub
Private Sub pass_Change()
cancel.Enabled = True
baru.Enabled = False
save.Enabled = True
End Sub
Private Sub save_Click()
If kodeuser.Text = "" Or namauser.Text = "" Or nohp.Text = "" Or pass.Text = "" Or hakakses.Text = "" Then
    MsgBox "Isi data dengan lengkap", , "INFORMASI"
    save.Enabled = True
Else
AUF = True
    con.Execute "insert into user(Kode_User,Nama_User,Alamat,No_Handphone,Password,Hak_Akses) values ('" & kodeuser.Text & "','" & namauser.Text & "','" & alamat.Text & "','" & nohp.Text & "','" & pass.Text & "','" & hakakses.Text & "')"
    MsgBox "Data Sudah Tersimpan", , "SAVING...."
    Call Form_Activate
    baru.Enabled = True
    End If
End Sub
Private Sub cmdcari_Click()
a = InputBox("Masukan kode user yang akan dicari....!!!", "pencarian data")
B = "select * from user where Kode_user='" & a & "'"
Set Rsuser = con.Execute(B, , adCmdText)
    If Rsuser.EOF Then
        MsgBox "Kota tujuan yang Anda Cari Tidak Ditemukan", vbExclamation, ".::INFO::."
        cmdcari.SetFocus
        Else
        Call aktif
        baru.Enabled = False
        kodeuser.Text = Rsuser!Kode_User
        namauser.Text = Rsuser!Nama_User
        alamat.Text = Rsuser!alamat
        nohp.Text = Rsuser!No_Handphone
        pass.Text = Rsuser!password
        hakakses.Text = Rsuser!hak_akses
        save.Enabled = False
        edit.Enabled = True
        pass.Enabled = False
        hapus.Enabled = True
        cancel.Enabled = True
        cmdcari.Enabled = True
        End If
End Sub
Private Sub tampiluser()
kduser.Text = Rsuser!Kode_User
user.Text = Rsuser!Nama_User
End Sub
