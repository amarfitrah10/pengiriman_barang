VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form8 
   BackColor       =   &H8000000A&
   Caption         =   "Form8"
   ClientHeight    =   8745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14880
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   Picture         =   "Retur.frx":0000
   ScaleHeight     =   8745
   ScaleWidth      =   14880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Height          =   975
      Left            =   8760
      TabIndex        =   34
      Top             =   480
      Width           =   3255
      Begin VB.TextBox Tanggal 
         Height          =   285
         Left            =   1560
         TabIndex        =   35
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label13 
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
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1215
      Left            =   240
      TabIndex        =   29
      Top             =   360
      Width           =   3255
      Begin VB.TextBox user 
         Height          =   285
         Left            =   1680
         TabIndex        =   33
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox kduser 
         Height          =   285
         Left            =   1680
         TabIndex        =   31
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label12 
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
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label11 
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
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   9600
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSDataGridLib.DataGrid grid_retur 
      Height          =   1575
      Left            =   240
      TabIndex        =   25
      Top             =   7200
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   2778
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
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   5280
      TabIndex        =   20
      Top             =   4920
      Width           =   2415
      Begin VB.CommandButton cmdcari 
         Caption         =   "Cari Nomor Transaksi"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Picture         =   "Retur.frx":6F8B
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      Height          =   3015
      Left            =   5160
      TabIndex        =   13
      Top             =   1800
      Width           =   3855
      Begin VB.CommandButton cmdhapus 
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
         Left            =   480
         TabIndex        =   19
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "Edit"
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
         TabIndex        =   18
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdkeluar 
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
         Left            =   2040
         TabIndex        =   17
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "Cancel"
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
         Left            =   2040
         TabIndex        =   16
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdsimpan 
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
         Left            =   2040
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdtambah 
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
         TabIndex        =   14
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Height          =   5415
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   4815
      Begin VB.TextBox noresi 
         Height          =   375
         Left            =   2040
         TabIndex        =   28
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox notrans 
         Height          =   375
         Left            =   2040
         TabIndex        =   24
         Top             =   360
         Width           =   2535
      End
      Begin VB.ComboBox status 
         Height          =   315
         Left            =   2040
         TabIndex        =   12
         Top             =   4080
         Width           =   1455
      End
      Begin VB.TextBox biayaretur 
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   4800
         Width           =   2415
      End
      Begin VB.TextBox biayakirim 
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   3360
         Width           =   2415
      End
      Begin VB.TextBox brg 
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox nmpenerima 
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox idpelanggan 
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         Caption         =   "No Resi"
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
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         Caption         =   "No Transaksi"
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
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         Caption         =   "Biaya Retur"
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
         Left            =   120
         TabIndex        =   10
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         Caption         =   "Status"
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
         Left            =   120
         TabIndex        =   9
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         Caption         =   "Biaya Kirim"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         Caption         =   "Barang"
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
         Left            =   120
         TabIndex        =   5
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   "Nama Penerima"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   "Id Pelanggan"
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
         Left            =   120
         TabIndex        =   1
         Top             =   1440
         Width           =   1335
      End
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "Retur"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   37
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000012&
      Height          =   1695
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   12255
   End
   Begin VB.Label label1 
      BackColor       =   &H8000000E&
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
      Left            =   7680
      TabIndex        =   22
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub aktif()
Tanggal.Text = Format(Date, "YYYY-MM-DD")
idpelanggan.Enabled = True

nmpenerima.Enabled = True
brg.Enabled = True
biayakirim.Enabled = True
status.Enabled = True
biayaretur.Enabled = True
notrans.Enabled = True
noresi.Enabled = True
End Sub
Private Sub tidakaktif()
idpelanggan.Enabled = False
Tanggal.Text = Format(Date, "YYYY-MM-DD")
nmpenerima.Enabled = False

brg.Enabled = False
biayakirim.Enabled = False
status.Enabled = False
biayaretur.Enabled = False

notrans.Enabled = False
noresi.Enabled = False
End Sub
Private Sub kosong()
idpelanggan.Text = ""
nmpenerima.Text = ""
noresi.Text = ""
brg.Text = ""
biayakirim.Text = ""
notrans.Text = ""
status.Text = ""
biayaretur.Text = ""
End Sub

Private Sub biayaretur_Change()
cmdsimpan.Enabled = True
cmdcancel.Enabled = True
cmdtambah.Enabled = False
End Sub

Private Sub cmdcancel_Click()
Call Form_Activate
cmdsimpan.Enabled = False
End Sub
Private Sub cmdedit_Click()
Dim update As String
    update = "UPDATE retur SET Id_Pelanggan = '" & idpelanggan.Text & "',Nama_Penerima = '" & nmpenerima.Text & "',Barang ='" & brg.Text & "', Biaya_Kirim='" & biayakirim.Text & "',Status='" & status.Text & "',Biaya_Retur='" & biayaretur.Text & "' WHERE No_Trans = '" & notrans.Text & "'"
    con.Execute update
    MsgBox "Data berhasil diubah !", vbOKOnly, "Info"
    Call Form_Activate
    cmdtambah.Enabled = True
End Sub
Private Sub cmdhapus_Click()
Dim hapus As String
    hapus = "DELETE FROM retur WHERE No_Trans = '" & notrans.Text & "'"
    con.Execute hapus
    MsgBox "Data Berhasil Dihapus !", vbOKOnly, "Info"
    Call Form_Activate
    cmdtambah.Enabled = True
End Sub
Private Sub cmdsimpan_Click()
If noresi.Text = "" Or idpelanggan.Text = "" Or Tanggal.Text = "" Or nmpenerima.Text = "" Or brg.Text = "" Or biayakirim.Text = "" Or status.Text = "" Or biayaretur.Text = "" Then
    MsgBox "Isi data dengan lengkap", , "INFORMASI"
    cmdsimpan.Enabled = True
Else
AUF = True
    con.Execute "insert into retur(No_Trans,No_Resi,Id_Pelanggan,Tanggal,Nama_Penerima,Barang,Biaya_Kirim,Status,Biaya_Retur) values ('" & notrans.Text & "','" & noresi.Text & "','" & idpelanggan.Text & "','" & Tanggal.Text & "','" & nmpenerima.Text & "','" & brg.Text & "','" & biayakirim.Text & "','" & status.Text & "','" & biayaretur.Text & "')"
    MsgBox "Data Sudah Tersimpan", , "SAVING...."
    Call cetak
    Call Form_Activate
    cmdtambah.Enabled = True
    End If
End Sub
Private Sub cmdtambah_Click()
Call aktif
Call kosong
notrans.SetFocus
End Sub
Private Sub Form_Load()
Call aktif
status.AddItem "Hilang"
status.AddItem "Rusak Ringan"
status.AddItem "Rusak Berat"
cmdtambah.Enabled = True
End Sub
Private Sub tampil()
notrans.Text = Rsretur!No_Trans
noresi.Text = Rsretur!No_Resi
idpelanggan.Text = Rsretur!Id_Pelanggan
Tanggal.Text = Rsretur!Tanggal
nmpenerima.Text = Rsretur!Nama_Penerima
brg.Text = Rsretur!Barang
biayakirim.Text = Rsretur!Biaya_Kirim
status.Text = Rsretur!status
biayaretur.Text = Rsretur!biaya_retur
End Sub
Private Sub Form_Activate()
    Call kosong
    Call tidakaktif
    Call koneksi
    kduser.Text = Menuutama.StatusBar1.Panels(1).Text
    user.Text = Menuutama.StatusBar1.Panels(2).Text
    Rsretur.Open "SELECT * FROM retur", con
    Set grid_retur.DataSource = Rsretur.DataSource
        Rsuser.Open "select * from user where kode_user ='" & kduser.Text & "'", con
If Not Rsuser.EOF Then
Call tampiluser
    cmdtambah.Enabled = True
    cmdkeluar.Enabled = True
    cmdcari.Enabled = True
    cmdsimpan.Enabled = False
    cmdedit.Enabled = False
    cmdhapus.Enabled = False
    cmdcancel.Enabled = False
    End If
    End Sub
Private Sub grid_retur_Click()
Call tampil
Call aktif
End Sub
Private Sub notrans_KeyPress(KeyAscii As Integer)
Call koneksi
If KeyAscii = 13 Then
Rstransaksi.Open "SELECT No_Trans,Noresi,Id_Pelanggan,Barang,Total_Pengiriman,Nama_Penerima FROM transaksi where No_Trans = '" & notrans.Text & "'", con
If Not Rstransaksi.EOF Then
notrans.Text = Rstransaksi!No_Trans
noresi.Text = Rstransaksi!noresi
idpelanggan.Text = Rstransaksi!Id_Pelanggan
nmpenerima.Text = Rstransaksi!Nama_Penerima
brg.Text = Rstransaksi!Barang
biayakirim.Text = Rstransaksi!Total_Pengiriman
nmpenerima.Text = Rstransaksi!Nama_Penerima
Call tampiltrans
Else
MsgBox "Data tidak ditemukan !", vbOKOnly, "info"
idpelanggan.Text = ""
idpelanggan.SetFocus
End If
End If
End Sub
Private Sub status_Click()
If (status.Text = "Hilang") Then
biayaretur.Text = Val(biayakirim.Text) * 10
ElseIf (status.Text = "Rusak Berat") Then
biayaretur.Text = Val(biayakirim.Text) * 8
Else
biayaretur.Text = Val(biayakirim.Text) * 5
End If
End Sub
Private Sub cmdkeluar_Click()
Unload Me
End Sub
Private Sub cmdcari_Click()
a = InputBox("Masukan nomor transaksi yang akan dicari....!!!", "pencarian data")
B = "select * from transaksi where No_Trans='" & a & "'"
Set Rstransaksi = con.Execute(B, , adCmdText)
    If Rstransaksi.EOF Then
        MsgBox "Nomor transaksi yang Anda Cari Tidak Ditemukan", vbExclamation, ".::INFO::."
        cmdcari.SetFocus
        Else
        Call aktif
        cmdsimpan.Enabled = True
        cmdtambah.Enabled = False
        notrans.Text = Rstransaksi!No_Trans
        noresi.Text = Rstransaksi!noresi
        idpelanggan.Text = Rstransaksi!Id_Pelanggan
        nmpenerima.Text = Rstransaksi!Nama_Penerima
        brg.Text = Rstransaksi!Barang
        biayakirim.Text = Rstransaksi!Total_Pengiriman
        cmdedit.Enabled = True
        cmdhapus.Enabled = True
        cmdcancel.Enabled = True
        cmdcari.Enabled = True
        cmdsimpan.Enabled = False
        cmdtambah.Enabled = False
        End If
End Sub
Private Sub tampiltrans()
notrans.Text = Rstransaksi!No_Trans
noresi.Text = Rstransaksi!noresi
idpelanggan.Text = Rstransaksi!Id_Pelanggan
nmpenerima.Text = Rstransaksi!Nama_Penerima
brg.Text = Rstransaksi!Barang
biayakirim.Text = Rstransaksi!Total_Pengiriman
End Sub
Sub cetak()
Call koneksi
CrystalReport1.SelectionFormula = "{retur.no_trans}='" & notrans.Text & "'"
CrystalReport1.ReportFileName = App.Path & "\strukretur.rpt"
CrystalReport1.WindowState = crptMaximized
CrystalReport1.RetrieveDataFiles
CrystalReport1.Action = 1
End Sub
Private Sub tampiluser()
kduser.Text = Rsuser!Kode_User
user.Text = Rsuser!Nama_User
End Sub

