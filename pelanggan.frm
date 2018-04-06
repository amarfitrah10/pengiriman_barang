VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form6 
   BackColor       =   &H8000000D&
   Caption         =   "Form6"
   ClientHeight    =   9315
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16350
   ControlBox      =   0   'False
   FillColor       =   &H000080FF&
   LinkTopic       =   "Form6"
   Picture         =   "pelanggan.frx":0000
   ScaleHeight     =   9315
   ScaleWidth      =   16350
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Height          =   1455
      Left            =   0
      TabIndex        =   50
      Top             =   120
      Width           =   5535
      Begin VB.TextBox user 
         Height          =   285
         Left            =   2040
         TabIndex        =   54
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox kduser 
         Height          =   285
         Left            =   2040
         TabIndex        =   52
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label21 
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
         TabIndex        =   53
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label20 
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
         TabIndex        =   51
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1335
      Left            =   10920
      TabIndex        =   35
      Top             =   4920
      Width           =   3495
      Begin VB.CommandButton cmdcari 
         Caption         =   "Cari Id Pelanggan"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   600
         Picture         =   "pelanggan.frx":4AF2
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   240
         Width           =   1935
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   14640
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame4 
      Caption         =   "Penerima"
      Height          =   6135
      Left            =   5520
      TabIndex        =   24
      Top             =   1680
      Width           =   5295
      Begin VB.TextBox Provinsipenerima 
         Height          =   375
         Left            =   2640
         TabIndex        =   47
         Top             =   5400
         Width           =   2535
      End
      Begin VB.TextBox kecematanpenerima 
         Height          =   375
         Left            =   2640
         TabIndex        =   45
         Top             =   4800
         Width           =   2535
      End
      Begin VB.TextBox kelurahanpenerima 
         Height          =   375
         Left            =   2640
         TabIndex        =   44
         Top             =   4200
         Width           =   2535
      End
      Begin VB.TextBox kotapenerima 
         Height          =   375
         Left            =   2640
         TabIndex        =   42
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox altpenerima 
         Height          =   1095
         Left            =   2640
         TabIndex        =   32
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox pospenerima 
         Height          =   375
         Left            =   2640
         TabIndex        =   31
         Top             =   3600
         Width           =   2535
      End
      Begin VB.TextBox tlppenerima 
         Height          =   375
         Left            =   2640
         TabIndex        =   28
         Top             =   3000
         Width           =   2535
      End
      Begin VB.TextBox nmpenerima 
         Height          =   375
         Left            =   2640
         TabIndex        =   26
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label17 
         Caption         =   "Kecematan"
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
         TabIndex        =   48
         Top             =   4920
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Provinsi"
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
         TabIndex        =   46
         Top             =   5400
         Width           =   1335
      End
      Begin VB.Label Label16 
         Caption         =   "Kelurahan"
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
         TabIndex        =   43
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label15 
         Caption         =   "Kota/Kabupaten Penerima"
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
         TabIndex        =   41
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label Label12 
         Caption         =   "Kode Pos"
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
         TabIndex        =   30
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "Telpon Penerima"
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
         TabIndex        =   29
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Alamat Penerima"
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
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label9 
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
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   11160
      TabIndex        =   21
      Top             =   600
      Width           =   3135
      Begin VB.TextBox tgl 
         Height          =   285
         Left            =   1440
         TabIndex        =   23
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Tanggal"
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
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid grid_pelanggan 
      Height          =   2055
      Left            =   0
      TabIndex        =   20
      Top             =   7920
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   3625
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
   Begin VB.Frame Frame2 
      Caption         =   "Button"
      Height          =   3015
      Left            =   10920
      TabIndex        =   11
      Top             =   1800
      Width           =   3495
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
         Picture         =   "pelanggan.frx":54F4
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2160
         Width           =   1215
      End
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
         Left            =   240
         Picture         =   "pelanggan.frx":5EF6
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2160
         Width           =   1335
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
         Left            =   2040
         TabIndex        =   15
         Top             =   1320
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
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   1335
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
         Picture         =   "pelanggan.frx":68F8
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   240
         Picture         =   "pelanggan.frx":72FA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pengirim"
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   5535
      Begin VB.TextBox kecematan 
         Height          =   375
         Left            =   2760
         TabIndex        =   40
         Top             =   5040
         Width           =   2655
      End
      Begin VB.TextBox kelurahan 
         Height          =   375
         Left            =   2760
         TabIndex        =   38
         Top             =   4440
         Width           =   2655
      End
      Begin VB.TextBox idpelanggan 
         Height          =   375
         Left            =   2760
         TabIndex        =   33
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox provinsi 
         Height          =   375
         Left            =   2760
         TabIndex        =   19
         Top             =   5520
         Width           =   2655
      End
      Begin VB.TextBox altpengirim 
         Height          =   1215
         Left            =   2760
         TabIndex        =   10
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox kota 
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Top             =   3000
         Width           =   2655
      End
      Begin VB.TextBox kodepos 
         Height          =   375
         Left            =   2760
         TabIndex        =   7
         Top             =   3960
         Width           =   2655
      End
      Begin VB.TextBox telppengirim 
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   3480
         Width           =   2655
      End
      Begin VB.TextBox nmpengirim 
         Height          =   375
         Left            =   2760
         TabIndex        =   2
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label14 
         Caption         =   "Kecematan"
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
         TabIndex        =   39
         Top             =   5040
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "Kelurahan"
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
         LinkItem        =   "&H00C0C0FF&"
         TabIndex        =   37
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label1 
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
         TabIndex        =   34
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Provinsi"
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
         TabIndex        =   18
         Top             =   5520
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Kota/Kabupaten Pengirim"
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
         TabIndex        =   8
         Top             =   3000
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "Kode Pos"
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
         TabIndex        =   6
         Top             =   3960
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Telpon Pengirim"
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
         TabIndex        =   4
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Alamat Pengirim"
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
         TabIndex        =   3
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Nama Pengirim"
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
         TabIndex        =   1
         Top             =   960
         Width           =   2175
      End
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Caption         =   "Pelanggan"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   6120
      TabIndex        =   55
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label19 
      BackColor       =   &H80000012&
      Height          =   1695
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   15375
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aktif()
idpelanggan.Enabled = True
tgl.Text = Format(Date, "YYYY-MM-DD")
nmpengirim.Enabled = True
altpengirim.Enabled = True
kodepos.Enabled = True
provinsi.Enabled = True
telppengirim.Enabled = True
kota.Enabled = True
nmpenerima.Enabled = True
altpenerima.Enabled = True
tlppenerima.Enabled = True
pospenerima.Enabled = True
kelurahan.Enabled = True
kecematan.Enabled = True
provinsi.Enabled = True
kotapenerima.Enabled = True
kelurahanpenerima.Enabled = True
kecematanpenerima.Enabled = True
provinsipenerima.Enabled = True
cmdsimpan.Enabled = True
End Sub
Private Sub tidakaktif()
idpelanggan.Enabled = False
tgl.Text = Format(Date, "YYYY-MM-DD")
nmpengirim.Enabled = False
altpengirim.Enabled = False
kodepos.Enabled = False
provinsi.Enabled = False
telppengirim.Enabled = False
kota.Enabled = False
nmpenerima.Enabled = False
altpenerima.Enabled = False
tlppenerima.Enabled = False
pospenerima.Enabled = False
kelurahan.Enabled = False
kecematan.Enabled = False
provinsi.Enabled = False
kotapenerima.Enabled = False
kelurahanpenerima.Enabled = False
kecematanpenerima.Enabled = False
provinsipenerima.Enabled = False
cmdsimpan.Enabled = False
End Sub
Private Sub kosong()
nmpengirim.Text = ""
altpengirim.Text = ""
kodepos.Text = ""
provinsi.Text = ""
telppengirim.Text = ""
idpelanggan.Text = ""
kota.Text = ""
nmpenerima.Text = ""
altpenerima.Text = ""
tlppenerima.Text = ""
pospenerima.Text = ""
kelurahan.Text = ""
kecematan.Text = ""
kotapenerima.Text = ""
kecematanpenerima.Text = ""
provinsipenerima.Text = ""
kelurahanpenerima.Text = ""
kecematanpenerima.Text = ""
End Sub
Private Sub tampil()
idpelanggan.Text = Rspelanggan!Id_Pelanggan
nmpengirim.Text = Rspelanggan!Nama_Pengirim
altpengirim.Text = Rspelanggan!Alamat_Pengirim
kodepos.Text = Rspelanggan!Kode_Pos_Pengirim
telppengirim.Text = Rspelanggan!Telepon_Pengirim
kota.Text = Rspelanggan!Kota_Kabupaten_Pengirim
kelurahan.Text = Rspelanggan!Kelurahan_Pengirim
kecematan.Text = Rspelanggan!Kecematan_Pengirim
provinsi.Text = Rspelanggan!Provinsi_Pengirim
nmpenerima.Text = Rspelanggan!Nama_Penerima
altpenerima.Text = Rspelanggan!Alamat_Penerima
kotapenerima.Text = Rspelanggan!Kota_Kabupaten_Penerima
tlppenerima.Text = Rspelanggan!No_Telp_Penerima
pospenerima.Text = Rspelanggan!Kode_Pos_Penerima
kelurahanpenerima.Text = Rspelanggan!Kelurahan_Penerima
kecematanpenerima.Text = Rspelanggan!Kecematan_Penerima
provinsipenerima.Text = Rspelanggan!Provinsi_Penerima
tgl.Text = Rspelanggan!Tanggal
End Sub
Private Sub cmdcancel_Click()
 Call Form_Activate
 cmdsimpan.Enabled = True
End Sub
Private Sub cmdcari_Click()
a = InputBox("Masukan id pelanggan yang akan dicari....!!!", "pencarian data")
B = "select * from pelanggan where Id_Pelanggan='" & a & "'"
Set Rspelanggan = con.Execute(B, , adCmdText)
    If Rspelanggan.EOF Then
        MsgBox "Id pelanggan yang Anda Cari Tidak Ditemukan", vbExclamation, ".::INFO::."
        cmdcari.SetFocus
        Else
        Call aktif
        cmdsimpan.Enabled = True
        cmdtambah.Enabled = False
        idpelanggan.Text = Rspelanggan!Id_Pelanggan
        nmpengirim.Text = Rspelanggan!Nama_Pengirim
        altpengirim.Text = Rspelanggan!Alamat_Pengirim
        kota.Text = Rspelanggan!Kota_Kabupaten_Pengirim
        telppengirim.Text = Rspelanggan!Telepon_Pengirim
        kota.Text = Rspelanggan!Kota_Kabupaten_Pengirim
        kelurahan.Text = Rspelanggan!Kelurahan_Pengirim
        kecematan.Text = Rspelanggan!Kecematan_Pengirim
        kodepos.Text = Rspelanggan!Kode_Pos_Penerima
        provinsi.Text = Rspelanggan!Provinsi_Pengirim
        kotapenerima.Text = Rspelanggan!Kota_Kabupaten_Penerima
        nmpenerima.Text = Rspelanggan!Nama_Penerima
        altpenerima.Text = Rspelanggan!Alamat_Penerima
        tlppenerima.Text = Rspelanggan!No_Telp_Penerima
        pospenerima.Text = Rspelanggan!Kode_Pos_Penerima
        kelurahanpenerima.Text = Rspelanggan!Kelurahan_Penerima
        kecematanpenerima.Text = Rspelanggan!Kecematan_Penerima
        provinsipenerima.Text = Rspelanggan!Provinsi_Penerima
        cmdedit.Enabled = True
        cmdhapus.Enabled = True
        cmdcancel.Enabled = True
        cmdcari.Enabled = True
        cmdsimpan.Enabled = False
        End If

End Sub

Private Sub cmdhapus_Click()
Dim hapus As String
    hapus = "DELETE FROM pelanggan WHERE id_pelanggan = '" & idpelanggan.Text & "'"
    con.Execute hapus
    MsgBox "Data Berhasil Dihapus !", vbOKOnly, "Info"
    Call Form_Activate
    cmdtambah.Enabled = True
End Sub
Private Sub cmdedit_Click()
Dim update As String
    update = "UPDATE pelanggan SET Nama_Pengirim = '" & nmpengirim.Text & "', Alamat_Pengirim = '" & altpengirim.Text & "', Kode_Pos_Pengirim='" & kodepos.Text & "', Telepon_Pengirim ='" & telppengirim.Text & "', Kota_Kabupaten_Pengirim ='" & kota.Text & "',Kelurahan_Pengirim='" & kelurahan.Text & "',Kecematan_Pengirim='" & kecematan.Text & "',Provinsi_Pengirim ='" & provinsi.Text & "', Nama_Penerima='" & nmpenerima.Text & "', Alamat_Penerima='" & altpenerima.Text & "',Kota_Kabupaten_Penerima='" & kotapenerima.Text & "', No_Telp_Penerima='" & tlppenerima.Text & "', Kode_Pos_Penerima='" & pospenerima.Text & "',Kelurahan_Penerima='" & kelurahanpenerima.Text & "',Kecematan_Penerima='" & kecematanpenerima.Text & "',Provinsi_Penerima='" & provinsipenerima.Text & "' WHERE Id_Pelanggan = '" & idpelanggan.Text & "'"
    con.Execute update
    MsgBox "Data berhasil diubah !", vbOKOnly, "Info"
    Call Form_Activate
    cmdtambah.Enabled = True
End Sub
Private Sub cmdkeluar_Click()
Unload Me
End Sub
Private Sub Form_Activate()
    Call kosong
    Call tidakaktif
    Call koneksi
    kduser.Text = Menuutama.StatusBar1.Panels(1).Text
    user.Text = Menuutama.StatusBar1.Panels(2).Text
    cmdtambah.Enabled = True
    Rspelanggan.Open "SELECT * FROM pelanggan", con
    Set grid_pelanggan.DataSource = Rspelanggan.DataSource
    Rsuser.Open "select * from user where kode_user ='" & kduser.Text & "'", con
If Not Rsuser.EOF Then
Call tampiluser
    cmdtambah.SetFocus

    cmdsimpan.Enabled = False
    cmdedit.Enabled = False
    cmdkeluar.Enabled = True
    cmdcancel.Enabled = False
    cmdhapus.Enabled = False
grid_pelanggan.Columns(2).Width = 3000
End If
End Sub
Private Sub cmdtambah_Click()
Call aktif
Call kosong
Call auto
nmpengirim.SetFocus
cmdsimpan.Enabled = False
cmdhapus.Enabled = False
cmdedit.Enabled = False
cmdhapus.Enabled = False
cmdcancel.Enabled = False
End Sub

Private Sub Form_Load()
Form7.idpelanggan = Form6.idpelanggan
End Sub

Private Sub grid_pelanggan_Click()
    Call tampil
    Call aktif
    idpelanggan.Enabled = False
    nmpengirim.Enabled = False
    cmdsimpan.Enabled = False
    cmdtambah.Enabled = False
    cmdedit.Enabled = True
    cmdhapus.Enabled = True
  
End Sub
Private Sub cmdsimpan_Click()
 Call koneksi
If idpelanggan.Text = "" Or nmpenerima.Text = "" Or altpenerima.Text = "" Or kota.Text = "" Or telppengirim.Text = "" Or kodepos.Text = "" Or kelurahan.Text = "" Or kecematan.Text = "" Or provinsi.Text = "" Or nmpenerima.Text = "" Or altpenerima.Text = "" Or kotapenerima.Text = "" Or tlppenerima.Text = "" Or pospenerima.Text = "" Or kelurahanpenerima.Text = "" Or kecematanpenerima.Text = "" Or provinsipenerima.Text = "" Then
    MsgBox "Isi data dengan lengkap", , "INFORMASI"
    cmdsimpan.Enabled = True
Else
AUF = True
    con.Execute "insert into pelanggan(Id_Pelanggan,Nama_Pengirim,Alamat_Pengirim,Kode_Pos_Pengirim,Telepon_Pengirim,Kota_Kabupaten_Pengirim,Kelurahan_Pengirim,Kecematan_Pengirim,Provinsi_Pengirim,Nama_Penerima,Alamat_Penerima,Kota_Kabupaten_Penerima,No_Telp_Penerima,Kode_Pos_Penerima,Kelurahan_Penerima,Kecematan_Penerima,Provinsi_Penerima,Tanggal) values ('" & idpelanggan.Text & "','" & nmpengirim.Text & "','" & altpengirim.Text & "','" & kodepos.Text & "','" & telppengirim.Text & "','" & kota.Text & "','" & kelurahan.Text & "','" & kecematan.Text & "','" & provinsi.Text & "','" & nmpenerima.Text & "','" & altpenerima.Text & "','" & kotapenerima.Text & "','" & tlppenerima.Text & "','" & pospenerima.Text & "','" & kelurahanpenerima.Text & "','" & kecematanpenerima.Text & "','" & provinsipenerima.Text & "','" & tgl.Text & "')"
    MsgBox "Data Sudah Tersimpan", , "SAVING...."
 
'Form7.idpelanggan = Form6.idpelanggan.Text
Form7.Show
    Unload Me
End If

End Sub
Sub auto()
Call koneksi
Set Rspelanggan = con.Execute("select * from pelanggan order by Id_Pelanggan desc limit 1")
    With Rspelanggan
        If Rspelanggan.EOF Then
            idpelanggan.Text = "PL" & "001"
        Else
            idpelanggan.Text = "PL" & Right(Str(Val(Right(.Fields(0), 3)) + 1001), 3)
            
    End If
End With
End Sub

Private Sub nmpengirim_Change()
cmdtambah.Enabled = False
End Sub

Private Sub Provinsipenerima_Change()
cmdtambah.Enabled = False
cmdsimpan.Enabled = True
cmdedit.Enabled = False
cmdhapus.Enabled = False
cmdcancel.Enabled = True
cmdkeluar.Enabled = True
End Sub
Private Sub tampiluser()
kduser.Text = Rsuser!Kode_User
user.Text = Rsuser!Nama_User
End Sub
