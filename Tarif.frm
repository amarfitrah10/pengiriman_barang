VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form5 
   BackColor       =   &H8000000D&
   Caption         =   "Tarif"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12510
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   Picture         =   "Tarif.frx":0000
   ScaleHeight     =   8490
   ScaleWidth      =   12510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   9360
      TabIndex        =   30
      Top             =   960
      Width           =   2775
      Begin VB.TextBox tgl 
         Height          =   285
         Left            =   1200
         TabIndex        =   32
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label11 
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
         TabIndex        =   31
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   0
      TabIndex        =   25
      Top             =   480
      Width           =   3975
      Begin VB.TextBox user 
         Height          =   285
         Left            =   2040
         TabIndex        =   27
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox kduser 
         Height          =   285
         Left            =   2040
         TabIndex        =   26
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label10 
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
         TabIndex        =   29
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label9 
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
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Cari 
      Caption         =   "Cari"
      Height          =   855
      Left            =   6480
      TabIndex        =   22
      Top             =   4920
      Width           =   1695
      Begin VB.CommandButton cmdcari 
         Caption         =   "Cari"
         Height          =   495
         Left            =   360
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox tempat 
      Height          =   495
      Left            =   2520
      TabIndex        =   20
      Top             =   6120
      Width           =   2895
   End
   Begin VB.TextBox ttlvolume 
      Height          =   375
      Left            =   2520
      TabIndex        =   18
      Top             =   5520
      Width           =   2415
   End
   Begin VB.TextBox volume 
      Height          =   375
      Left            =   2520
      TabIndex        =   17
      Top             =   4920
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid grid_tarif 
      Height          =   2655
      Left            =   -360
      TabIndex        =   14
      Top             =   6840
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   4683
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
      Height          =   3015
      Left            =   5520
      TabIndex        =   8
      Top             =   1800
      Width           =   3255
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
         Height          =   495
         Left            =   360
         TabIndex        =   21
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton keluar 
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
         Height          =   495
         Left            =   1920
         TabIndex        =   13
         Top             =   2280
         Width           =   1095
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
         Height          =   495
         Left            =   1920
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
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
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   2280
         Width           =   1095
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
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   480
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
         Height          =   495
         Left            =   1920
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.TextBox txttarif 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox txtprovinsi 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox txtnama 
      Height          =   405
      Left            =   2520
      TabIndex        =   5
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox txtkode 
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Tarif"
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
      Left            =   4920
      TabIndex        =   33
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000012&
      Height          =   1695
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   12495
   End
   Begin VB.Label Label7 
      Caption         =   "Tempat Pengambilan"
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
      Left            =   0
      TabIndex        =   19
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Total Volume"
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
      TabIndex        =   16
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Volume"
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
      TabIndex        =   15
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Tarif"
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
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Provinsi"
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
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Nama Kota"
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
      TabIndex        =   1
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Kode Kota"
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
      TabIndex        =   0
      Top             =   1800
      Width           =   1455
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aktif()
    txtkode.Enabled = True
    tgl.Text = Format(Date, "YYYY-MM-DD")
    txtnama.Enabled = True
    txtprovinsi.Enabled = True
    txttarif.Enabled = True
    volume.Enabled = True
    ttlvolume.Enabled = True
    tempat.Enabled = True
   
    cmdsimpan.Enabled = True
    cmdhapus.Enabled = True
    cmdedit.Enabled = True
    keluar.Enabled = True
End Sub
 Private Sub tidakaktif()
 txtkode.Enabled = False
    txtnama.Enabled = False
    txtprovinsi.Enabled = False
    txttarif.Enabled = False
    volume.Enabled = False
    ttlvolume.Enabled = False
    tempat.Enabled = False
    HScroll1.Enabled = False
    cmdsimpan.Enabled = False
    cmdhapus.Enabled = False
    cmdedit.Enabled = False
    keluar.Enabled = False
 End Sub
Private Sub kosong()
    txtkode.Text = ""
    txtnama.Text = ""
    txtprovinsi.Text = ""
    txttarif.Text = ""
    volume.Text = ""
    ttlvolume.Text = ""
    tempat.Text = ""
End Sub
Private Sub tampil()
    txtkode.Text = Rstarif!Kode_Kota
    txtnama.Text = Rstarif!Nama_Kota
    txtprovinsi.Text = Rstarif!provinsi
    txttarif.Text = Rstarif!Tarif
    volume.Text = Rstarif!volume
    ttlvolume.Text = Rstarif!Total_Volume
    tempat.Text = Rstarif!pengambilan
End Sub
Private Sub cmdkeluar_Click()
 Call Form_Activate
 cmdsimpan.Enabled = True
End Sub

Private Sub cmdcancel_Click()
 Call Form_Activate
 cmdsimpan.Enabled = True
End Sub

Private Sub cmdcari_Click()
a = InputBox("Masukan Kode kota yang akan dicari....!!!", "pencarian data")
B = "select * from tarif where kode_kota='" & a & "'"
Set Rstarif = con.Execute(B, , adCmdText)
    If Rstarif.EOF Then
        MsgBox "Kode kota yang Anda Cari Tidak Ditemukan", vbExclamation, ".::INFO::."
        cmdcari.SetFocus
        Else
        Call aktif
        cmdsimpan.Enabled = True
        cmdtambah.Enabled = False
        txtkode.Text = Rstarif!Kode_Kota
        txtnama.Text = Rstarif!Nama_Kota
        txtprovinsi.Text = Rstarif!provinsi
        txttarif.Text = Rstarif!Tarif
        volume.Text = Rstarif!volume
        ttlvolume.Text = Rstarif!Total_Volume
        tempat.Text = Rstarif!pengambilan
End Sub

Private Sub cmdhapus_Click()
Dim hapus As String
    hapus = "DELETE FROM tarif WHERE kode_kota = '" & txtkode.Text & "'"
    con.Execute hapus
    MsgBox "Data Berhasil Dihapus !", vbOKOnly, "Info"
    Call Form_Activate
    cmdsimpan.Enabled = True
    cmdtambah.Enabled = True
End Sub
Private Sub cmdedit_Click()
Dim update As String
    update = "UPDATE tarif SET Nama_Kota = '" & txtnama.Text & "', provinsi = '" & txtprovinsi.Text & "', tarif = '" & txttarif.Text & "', Volume = '" & volume.Text & "', Total_Volume = '" & ttlvolume.Text & "', Pengambilan = '" & List1.Text & "' WHERE kode_kota = '" & txtkode.Text & "'"
    con.Execute update
    MsgBox "Data berhasil diubah !", vbOKOnly, "Info"
    Call Form_Activate
    cmdsimpan.Enabled = True
    cmdtambah.Enabled = True

End Sub



Private Sub keluar_click()
Unload Me
End Sub
Private Sub Form_Activate()
    Call kosong
    Call aktif
    Call koneksi
    kduser.Text = Menuutama.StatusBar1.Panels(1).Text
    user.Text = Menuutama.StatusBar1.Panels(2).Text
    cmdtambah.Enabled = True
    Rstarif.Open "SELECT * FROM tarif", con
    Set grid_tarif.DataSource = Rstarif.DataSource
     Rsuser.Open "select * from user where kode_user ='" & kduser.Text & "'", con
If Not Rsuser.EOF Then
Call tampiluser
    cmdtambah.SetFocus
    cmdsimpan.Enabled = False
    cmdcancel.Enabled = False
    cmdedit.Enabled = False
    keluar.Enabled = True
    cmdhapus.Enabled = False
    cmdcancel.Enabled = False
grid_tarif.Columns(0).Width = 3000
End If
End Sub
Private Sub cmdtambah_Click()
Call aktif
Call kosong
cmdsimpan.Enabled = False
cmdcancel.Enabled = False
cmdhapus.Enabled = False
cmdedit.Enabled = False
cmdtambah.Enabled = False
txtkode.SetFocus
End Sub
Private Sub grid_tarif_Click()
    Call tampil
    Call aktif
    txtkode.Enabled = False
    txtnama.Enabled = False
    cmdsimpan.Enabled = False
    cmdtambah.Enabled = False
    cmdedit.SetFocus
End Sub

Private Sub cmdsimpan_Click()
If txtkode.Text = "" Or txtnama.Text = "" Or txtprovinsi.Text = "" Or txttarif.Text = "" Or volume.Text = "" Or ttlvolume.Text = "" Or tempat.Text = "" Then
    MsgBox "Isi data dengan lengkap", , "INFORMASI"
    cmdsimpan.Enabled = True
Else
AUF = True
    con.Execute "insert into tarif(Kode_Kota,Nama_Kota,Provinsi,Tarif,Volume,Total_Volume,Pengambilan) values ('" & txtkode.Text & "','" & txtnama.Text & "','" & txtprovinsi.Text & "','" & txttarif.Text & "','" & volume.Text & "','" & ttlvolume.Text & "','" & tempat.Text & "')"
    MsgBox "Data Sudah Tersimpan", , "SAVING...."
    Call Form_Activate
  
End If
    
End Sub

Private Sub tampiluser()
kduser.Text = Rsuser!Kode_User
user.Text = Rsuser!Nama_User
End Sub

Private Sub tempat_Change()
cmdsimpan.Enabled = True
cmdcancel.Enabled = True
keluar.Enabled = True
cmdtambah.Enabled = False
End Sub
