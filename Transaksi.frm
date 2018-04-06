VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Form7 
   BackColor       =   &H8000000D&
   ClientHeight    =   11205
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   19455
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   LinkTopic       =   "Form7"
   Picture         =   "Transaksi.frx":0000
   ScaleHeight     =   11205
   ScaleWidth      =   19455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame11 
      Caption         =   "List Pelanggan"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   13800
      TabIndex        =   104
      Top             =   3240
      Width           =   6495
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1335
         Left            =   240
         TabIndex        =   105
         Top             =   480
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   2355
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
   End
   Begin VB.Frame Frame10 
      Caption         =   "Cari"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8760
      TabIndex        =   81
      Top             =   1800
      Width           =   4815
      Begin VB.CommandButton cmdcari 
         Caption         =   "Cek Tarif"
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
         Left            =   1680
         Picture         =   "Transaksi.frx":4AF2
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H8000000A&
      Caption         =   "Volume"
      Height          =   1455
      Left            =   0
      TabIndex        =   62
      Top             =   7800
      Width           =   8055
      Begin VB.TextBox hasilpembagian 
         Height          =   375
         Left            =   6840
         TabIndex        =   78
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox harga 
         Height          =   375
         Left            =   5520
         TabIndex        =   76
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox bagi 
         Height          =   375
         Left            =   4440
         TabIndex        =   74
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox hasil 
         Height          =   375
         Left            =   3600
         TabIndex        =   72
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox tinggi 
         Height          =   375
         Left            =   2400
         TabIndex        =   68
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox lebar 
         Height          =   375
         Left            =   1200
         TabIndex        =   67
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox panjang 
         Height          =   375
         Left            =   120
         TabIndex        =   66
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label33 
         BackColor       =   &H8000000A&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   77
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label32 
         BackColor       =   &H8000000A&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   75
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label31 
         BackColor       =   &H8000000A&
         Caption         =   "/"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   73
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label30 
         BackColor       =   &H8000000A&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   71
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label29 
         BackColor       =   &H8000000A&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   70
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label28 
         BackColor       =   &H8000000A&
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   69
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label27 
         BackColor       =   &H8000000A&
         Caption         =   "Tinggi"
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
         Left            =   2400
         TabIndex        =   65
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label26 
         BackColor       =   &H8000000A&
         Caption         =   "Lebar"
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
         Left            =   1320
         TabIndex        =   64
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label25 
         BackColor       =   &H8000000A&
         Caption         =   "Panjang"
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
         TabIndex        =   63
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H8000000D&
      Height          =   1455
      Left            =   3960
      TabIndex        =   57
      Top             =   1680
      Width           =   4695
      Begin VB.TextBox user 
         Height          =   285
         Left            =   1800
         TabIndex        =   80
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox kduser 
         Height          =   285
         Left            =   1800
         TabIndex        =   59
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label34 
         BackColor       =   &H8000000D&
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
         Left            =   240
         TabIndex        =   79
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
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
         Left            =   240
         TabIndex        =   58
         Top             =   360
         Width           =   855
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   18120
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSDataGridLib.DataGrid grid_transaksi 
      Height          =   1695
      Left            =   0
      TabIndex        =   49
      Top             =   10200
      Width           =   18735
      _ExtentX        =   33046
      _ExtentY        =   2990
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
   Begin VB.Frame Frame7 
      BackColor       =   &H8000000D&
      Caption         =   "Button"
      Height          =   1815
      Left            =   12960
      TabIndex        =   44
      Top             =   8160
      Width           =   3135
      Begin VB.CommandButton cmdkeluar 
         BackColor       =   &H8000000E&
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
         Left            =   1800
         MaskColor       =   &H000080FF&
         TabIndex        =   48
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdcancel 
         BackColor       =   &H8000000E&
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
         Left            =   1800
         MaskColor       =   &H000080FF&
         TabIndex        =   47
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmdsimpan 
         BackColor       =   &H8000000E&
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
         Left            =   240
         MaskColor       =   &H000080FF&
         TabIndex        =   46
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdtambah 
         BackColor       =   &H8000000E&
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
         MaskColor       =   &H000080FF&
         TabIndex        =   45
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H8000000A&
      Caption         =   "Pembayaran"
      Height          =   2895
      Left            =   12840
      TabIndex        =   37
      Top             =   5280
      Width           =   4935
      Begin VB.TextBox ukem 
         Height          =   375
         Left            =   2280
         TabIndex        =   43
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox ubay 
         Height          =   375
         Left            =   2280
         TabIndex        =   41
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox ttl 
         Height          =   375
         Left            =   2280
         TabIndex        =   39
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Pembayaran 
         BackColor       =   &H8000000A&
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
         Left            =   1680
         TabIndex        =   56
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label20 
         BackColor       =   &H8000000A&
         Caption         =   "Uang Kembali"
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
         TabIndex        =   42
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000A&
         Caption         =   "Uang Bayar"
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
         TabIndex        =   40
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000A&
         Caption         =   "Total Pengiriman"
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
         TabIndex        =   38
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H8000000D&
      Caption         =   "Tarif"
      Height          =   4815
      Left            =   8160
      TabIndex        =   28
      Top             =   5160
      Width           =   4695
      Begin VB.TextBox noresi 
         Height          =   285
         Left            =   2280
         TabIndex        =   107
         Top             =   4080
         Width           =   2055
      End
      Begin VB.TextBox pengambilan 
         Height          =   285
         Left            =   2280
         TabIndex        =   87
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox ktasal 
         Height          =   285
         Left            =   2280
         TabIndex        =   61
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox brt 
         Height          =   375
         Left            =   2280
         TabIndex        =   55
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox provinsi 
         Height          =   285
         Left            =   2280
         TabIndex        =   53
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox subtotal 
         Height          =   285
         Left            =   2280
         TabIndex        =   36
         Top             =   3600
         Width           =   2055
      End
      Begin VB.TextBox txttarif 
         Height          =   285
         Left            =   2280
         TabIndex        =   34
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox txtkota 
         Height          =   285
         Left            =   2280
         TabIndex        =   32
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox nmkota 
         Height          =   285
         Left            =   2280
         TabIndex        =   30
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label45 
         BackColor       =   &H8000000D&
         Caption         =   "Nomor Resi"
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
         TabIndex        =   106
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label36 
         BackColor       =   &H8000000D&
         Caption         =   "Pengambilan Barang"
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
         TabIndex        =   86
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label35 
         Caption         =   "Kg"
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
         Left            =   2880
         TabIndex        =   85
         Top             =   3240
         Width           =   375
      End
      Begin VB.Label Label24 
         BackColor       =   &H8000000D&
         Caption         =   "Kota Asal"
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
         TabIndex        =   60
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label23 
         BackColor       =   &H8000000D&
         Caption         =   "Berat"
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
         TabIndex        =   54
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label22 
         BackColor       =   &H8000000D&
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
         TabIndex        =   52
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label17 
         BackColor       =   &H8000000D&
         Caption         =   "Subtotal"
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
         TabIndex        =   35
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000D&
         Caption         =   "Tarif"
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
         TabIndex        =   33
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000D&
         Caption         =   "Kode Kota"
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
         TabIndex        =   31
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000D&
         Caption         =   "Kota Tujuan"
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
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000002&
      Caption         =   "Penerima"
      Height          =   2535
      Left            =   0
      TabIndex        =   19
      Top             =   5160
      Width           =   8175
      Begin VB.TextBox provinsipenerima 
         Height          =   285
         Left            =   5880
         TabIndex        =   103
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox kotapenerima 
         Height          =   285
         Left            =   3840
         TabIndex        =   101
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox kecematanpenerima 
         Height          =   285
         Left            =   5880
         TabIndex        =   99
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox kelurahanpenerima 
         Height          =   285
         Left            =   3840
         TabIndex        =   97
         Top             =   1320
         Width           =   1575
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   1800
         TabIndex        =   84
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox kodepospenerima 
         Height          =   285
         Left            =   5880
         TabIndex        =   27
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox telppenerima 
         Height          =   285
         Left            =   3840
         TabIndex        =   25
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox altpenerima 
         Height          =   1455
         Left            =   1800
         TabIndex        =   23
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox nmpenerima 
         Height          =   405
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label44 
         BackColor       =   &H80000002&
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
         Left            =   5880
         TabIndex        =   102
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label43 
         BackColor       =   &H80000002&
         Caption         =   "Kota/Kabupaten"
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
         Left            =   3840
         TabIndex        =   100
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label42 
         BackColor       =   &H80000002&
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
         Left            =   5880
         TabIndex        =   98
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label41 
         BackColor       =   &H80000002&
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
         Left            =   3840
         TabIndex        =   96
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000002&
         Caption         =   "Kode Pos Penerima"
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
         Left            =   5880
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000002&
         Caption         =   "No Telepon Penerima"
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
         Left            =   3840
         TabIndex        =   24
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000002&
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
         Left            =   1800
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000002&
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
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000D&
      Caption         =   "Pengirim"
      Height          =   2055
      Left            =   0
      TabIndex        =   8
      Top             =   3120
      Width           =   13695
      Begin VB.TextBox provinsipengirim 
         Height          =   375
         Left            =   11880
         TabIndex        =   95
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox kecematan 
         Height          =   375
         Left            =   10080
         TabIndex        =   93
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox kelurahan 
         Height          =   375
         Left            =   8160
         TabIndex        =   91
         Top             =   600
         Width           =   1815
      End
      Begin VB.ComboBox plat 
         Height          =   315
         Left            =   10080
         TabIndex        =   89
         Top             =   1440
         Width           =   1575
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   2160
         TabIndex        =   83
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox kodepospengirim 
         Height          =   375
         Left            =   5760
         TabIndex        =   51
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox brg 
         Height          =   375
         Left            =   8160
         TabIndex        =   18
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox kotapengirim 
         Height          =   375
         Left            =   5760
         TabIndex        =   16
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox telppengirim 
         Height          =   375
         Left            =   3840
         TabIndex        =   14
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox altpengirim 
         Height          =   1215
         Left            =   2160
         TabIndex        =   12
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox nmpengirim 
         Height          =   405
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label40 
         BackColor       =   &H8000000D&
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
         Height          =   255
         Left            =   12000
         TabIndex        =   94
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label39 
         BackColor       =   &H8000000D&
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
         Height          =   375
         Left            =   10200
         TabIndex        =   92
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label38 
         BackColor       =   &H8000000D&
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
         Left            =   8280
         TabIndex        =   90
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label37 
         BackColor       =   &H8000000D&
         Caption         =   "No Plat Bus"
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
         Left            =   10200
         TabIndex        =   88
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label21 
         BackColor       =   &H8000000D&
         Caption         =   "Kode Pos Pengirim"
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
         Left            =   5760
         TabIndex        =   50
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         Caption         =   "Isi Barang"
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
         Left            =   8280
         TabIndex        =   17
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
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
         Left            =   5760
         TabIndex        =   15
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         Caption         =   "No Telepon Pengirim"
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
         Left            =   3840
         TabIndex        =   13
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
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
         Left            =   2160
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
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
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000D&
      Height          =   1215
      Left            =   13680
      TabIndex        =   5
      Top             =   1800
      Width           =   3135
      Begin VB.TextBox tgl 
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
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
         TabIndex        =   6
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   3975
      Begin VB.TextBox idpelanggan 
         Height          =   375
         Left            =   2160
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox notrans 
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   240
         Width           =   1455
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
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
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
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Label Label47 
      Alignment       =   2  'Center
      Caption         =   "Transaksi"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4560
      TabIndex        =   109
      Top             =   240
      Width           =   5055
   End
   Begin VB.Label Label46 
      BackColor       =   &H80000007&
      Height          =   1695
      Left            =   0
      TabIndex        =   108
      Top             =   0
      Width           =   19455
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aktif()
notrans.Enabled = True
idpelanggan.Enabled = True
tgl.Text = Format(Date, "YYYY-MM-DD")
nmpengirim.Enabled = True
altpengirim.Enabled = True
telppengirim.Enabled = True
kotapengirim.Enabled = True
kodepospengirim.Enabled = True
kelurahan.Enabled = True
kecematan.Enabled = True
provinsipengirim.Enabled = True
kelurahanpenerima.Enabled = True
kecematanpenerima.Enabled = True
kotapenerima.Enabled = True
provinsipenerima.Enabled = True
brg.Enabled = True
brt.Enabled = True
noresi.Enabled = True
nmpengirim.Enabled = True
altpengirim.Enabled = True
kodepospengirim.Enabled = True
nmkota.Enabled = True
hasil.Enabled = True
panjang.Enabled = True
lebar.Enabled = True
tinggi.Enabled = True
hasilpembagian.Enabled = True
bagi.Enabled = True
harga.Enabled = True
txtkota.Enabled = True
txttarif.Enabled = True
ktasal.Enabled = True
subtotal.Enabled = True
pengambilan.Enabled = True
plat.Enabled = True
ttl.Enabled = True
ubay.Enabled = True
ukem.Enabled = True
End Sub
Private Sub tidakaktif()
notrans.Enabled = False
idpelanggan.Enabled = False
nmpengirim.Enabled = False
altpengirim.Enabled = False
telppengirim.Enabled = False
kotapengirim.Enabled = False
kodepospengirim.Enabled = False
ktasal.Enabled = False
noresi.Enabled = False
kelurahan.Enabled = False
kecematan.Enabled = False
tgl.Text = Format(Date, "YYYY-MM-DD")
provinsipengirim.Enabled = False
kelurahanpenerima.Enabled = False
kecematanpenerima.Enabled = False
kotapenerima.Enabled = False
provinsipenerima.Enabled = False
brg.Enabled = False
brt.Enabled = False
hasil.Enabled = False
panjang.Enabled = False
lebar.Enabled = False
tinggi.Enabled = False
hasilpembagian.Enabled = False
bagi.Enabled = False
harga.Enabled = False
nmpengirim.Enabled = False
altpenerima.Enabled = False
telppenerima.Enabled = False
kodepospengirim.Enabled = False
nmkota.Enabled = False
txtkota.Enabled = False
txttarif.Enabled = False
subtotal.Enabled = False
pengambilan.Enabled = False
plat.Enabled = False
ttl.Enabled = False
ubay.Enabled = False
ukem.Enabled = False
kduser.Enabled = False
user.Enabled = False

End Sub
Private Sub kosong()
notrans.Text = ""
idpelanggan.Text = ""
tgl.Text = ""
nmpengirim.Text = ""
altpengirim.Text = ""
ktasal.Text = ""
telppengirim.Text = ""
kotapengirim.Text = ""
kodepospengirim.Text = ""
kelurahan.Text = ""
kecematan.Text = ""
noresi.Text = ""
provinsipengirim.Text = ""
kelurahanpenerima.Text = ""
kecematanpenerima.Text = ""
kotapenerima.Text = ""
provinsipenerima.Text = ""
brg.Text = ""
hasil.Text = ""
panjang.Text = ""
lebar.Text = ""
tinggi.Text = ""
hasilpembagian.Text = ""
bagi.Text = ""
harga.Text = ""
brt.Text = ""
nmpenerima.Text = ""
altpenerima.Text = ""
telppenerima.Text = ""
kodepospenerima.Text = ""
txtkota.Text = ""
nmkota.Text = ""
txttarif.Text = ""
subtotal.Text = ""
pengambilan.Text = ""
plat.Text = ""
ttl.Text = ""
ubay.Text = ""
ukem.Text = ""

End Sub
Private Sub brt_KeyPress(KeyAscii As Integer)
Dim masuk As String
If KeyAscii = 13 Then
Call hitung
ttl.Text = Val(subtotal.Text) + Val(ttl.Text) + Val(hasilpembagian.Text)
Pembayaran.Caption = total
Pembayaran.Caption = "Rp.  " & Format(Pembayaran.Caption, "#,##0")
masuk = "insert into detailtransaksi(No_Trans,Id_Pelanggan,Kode_Kota,Subtotal) values ('" & notrans.Text & "','" & idpelanggan.Text & "','" & txtkota.Text & "','" & ttl.Text & "')"
con.Execute masuk

If provinsi.Text = txtkota.Text Then
     noresi.Text = txtkota.Text + subtotal.Text + idpelanggan.Text
  ElseIf provinsi.Text = "" Then
            noresi.Text = ""
        Else: noresi.Text = "JKT01" + Strings.Left(provinsi.Text, 3) + subtotal.Text
pesan = MsgBox("Tambah Data Lagi ??", vbYesNo, "Konfirmasi")
If pesan = vbYes Then
notrans.Text = ""
noresi.Text = ""
nmpenerima.Text = ""
altpenerima.Text = ""
telppenerima.Text = ""
kodepospenerima.Text = ""
txtkota.Text = ""
nmkota.Text = ""
nmkota.Text = ""
provinsi.Text = ""
brt.Text = ""
brg.Text = ""
txttarif.Text = ""
subtotal.Text = ""
Else
    brg.SetFocus
    End If
End If
End If
End Sub

Private Sub cmdhasil_Click()
Call koneksi
If KeyAscii = 13 Then
Rstarif.Open "select * from tarif where Volume ='" & bagi.Text & "'", con
If Not Rstarif.EOF Then
      Call tampil_tarif
Else
MsgBox "Data tidak ditemukan !", vbOKOnly, "info"
panjang.Text = ""
panjang.SetFocus
End If
End If

End Sub
Private Sub caritrans_Click()
a = InputBox("Masukan no transaksi yang akan dicari....!!!", "pencarian data")
B = "select * from transaksi where No_Trans='" & a & "'"
Set Rstransaksi = con.Execute(B, , adCmdText)
    If Rstransaksi.EOF Then
        MsgBox "no transaksi yang Anda Cari Tidak Ditemukan", vbExclamation, ".::INFO::."
        caritrans.SetFocus
        Else
        Call aktif
        cmdsimpan.Enabled = True
        cmdtambah.Enabled = False
        notrans.Text = Rstransaksi!No_Trans
        tgl.Text = Rstransaksi!tanggal
        kduser.Text = Rstransaksi!Kode_User
        idpelanggan.Text = Rstransaksi!Id_Pelanggan
        brg.Text = Rstransaksi!Barang
        brt.Text = Rstransaksi!Berat
        plat.Text = Rstransaksi!No_Plat_Bus
        noresi.Text = Rstransaksi!noresi
        txttarif.Text = Rstransaksi!Tarif
        ttl.Text = Rstransaksi!Total_Pengiriman
        
        cmdcancel.Enabled = True
        cmdcari.Enabled = True
        End If
End Sub
Private Sub cmdcari_Click()
a = InputBox("Masukan kota tujuan yang akan dicari....!!!", "pencarian data")
B = "select * from tarif where Nama_Kota='" & a & "'"
Set Rstarif = con.Execute(B, , adCmdText)
    If Rstarif.EOF Then
        MsgBox "Kota tujuan yang Anda Cari Tidak Ditemukan", vbExclamation, ".::INFO::."
        cmdcari.SetFocus
        Else
        Call aktif
        cmdsimpan.Enabled = False
        cmdtambah.Enabled = False
        txtkota.Text = Rstarif!Kode_Kota
        nmkota.Text = Rstarif!Nama_Kota
        provinsi.Text = Rstarif!provinsi
        txttarif.Text = Rstarif!Tarif
        harga.Text = Rstarif!volume
        hasilpembagian.Text = Rstarif!Total_Volume
        cmdedit.Enabled = True
        cmdhapus.Enabled = True
        cmdcancel.Enabled = True
        cmdcari.Enabled = True
        End If
End Sub
Private Sub cmdkeluar_Click()
Unload Me
End Sub
Private Sub cmdsimpan_Click()
If idpelanggan.Text = "" Or brg.Text = "" Or brt.Text = "" Or txttarif.Text = "" Or ttl.Text = "" Or noresi.Text = "" Or ubay.Text = "" Or ukem.Text = "" Then
    MsgBox "Isi data dengan lengkap", , "INFORMASI"
    cmdsimpan.Enabled = True
Else
AUF = True
   con.Execute "insert into transaksi(No_Trans,Tanggal,Kode_user,Id_Pelanggan,Nama_Penerima,Barang,Berat,No_Plat_Bus,Kode_Kota,Nama_Kota,Tarif,Total_Pengiriman,NoResi,Uang_Bayar,Uang_Kembali) values ('" & notrans.Text & "','" & tgl.Text & "','" & kduser.Text & "','" & idpelanggan.Text & "','" & nmpenerima.Text & "','" & brg.Text & "','" & brt.Text & "','" & plat.Text & "','" & txtkota.Text & "','" & nmkota.Text & "','" & txttarif.Text & "','" & ttl.Text & "','" & noresi.Text & "','" & ubay.Text & "','" & ukem.Text & "')"
    MsgBox "Data Sudah Tersimpan", , "SAVING...."
    Call cetak
    Call Form_Activate
    cmdtambah.Enabled = True
End If
End Sub
Private Sub cmdtambah_Click()
Call auto
Call aktif
cmdsimpan.Enabled = False
cmdcancel.Enabled = False
cmdkeluar.Enabled = False
brg.SetFocus
End Sub
Private Sub Command1_Click()
Form6.Show
End Sub
Private Sub Command2_Click()
Call koneksi

If provinsi.Text = txtkota.Text Then
     noresi.Text = txtkota.Text + subtotal.Text + idpelanggan.Text
  ElseIf provinsi.Text = "" Then
            noresi.Text = ""
        Else: noresi.Text = "JKT01" + Strings.Left(provinsi.Text, 3) + subtotal.Text
        

End If
End Sub

Private Sub DataGrid1_Click()
    DataGrid1.Columns(0).Locked = True
    DataGrid1.Columns(1).Locked = True
    DataGrid1.Columns(2).Locked = True
    DataGrid1.Columns(3).Locked = True
    DataGrid1.Columns(4).Locked = True
    DataGrid1.Columns(5).Locked = True
    DataGrid1.Columns(6).Locked = True
    DataGrid1.Columns(7).Locked = True
    DataGrid1.Columns(8).Locked = True
    DataGrid1.Columns(9).Locked = True
    DataGrid1.Columns(10).Locked = True
    DataGrid1.Columns(11).Locked = True
    DataGrid1.Columns(12).Locked = True
    DataGrid1.Columns(13).Locked = True
    DataGrid1.Columns(14).Locked = True
    DataGrid1.Columns(15).Locked = True
    DataGrid1.Columns(16).Locked = True
    Call tampil_pelanggan
    Call aktif
    idpelanggan.Enabled = False
    cmdsimpan.Enabled = False
    cmdtambah.Enabled = False
End Sub
Private Sub Form_Load()
Call aktif
plat.AddItem "B 7702 XT"
plat.AddItem "B 1515 CY"
plat.AddItem "B 3709 XY"
plat.AddItem "B 7705 XY"
plat.AddItem "B 7389 BW"
plat.AddItem "B 7709 TK"
plat.AddItem "B 7389 CY"
plat.AddItem "B 3989 CY"
plat.AddItem "B 3708 XT"
cmdcancel.Enabled = False
cmdkeluar.Enabled = False
cmdsimpan.Enabled = False
End Sub
Sub auto()
Call koneksi
Set Rstransaksi = con.Execute("select * from transaksi order by no_trans desc limit 1")
    With Rstransaksi
        If .EOF Then
            notrans.Text = "TR" & "001"
        Else
            notrans.Text = "TR" & Right(Str(Val(Right(.Fields(0), 3)) + 1001), 3)
    End If
End With
End Sub
Private Sub idpelanggan_KeyPress(KeyAscii As Integer)
Call koneksi
If KeyAscii = 13 Then
Rspelanggan.Open "select * from pelanggan where Id_pelanggan ='" & idpelanggan.Text & "'", con
If Not Rspelanggan.EOF Then
      Call tampil_pelanggan
      brg.SetFocus
Else
MsgBox "Data tidak ditemukan !", vbOKOnly, "info"
idpelanggan.Text = ""
idpelanggan.SetFocus
End If
End If
End Sub
Private Sub tampil_pelanggan()
idpelanggan.Text = Rspelanggan!Id_Pelanggan
nmpengirim.Text = Rspelanggan!Nama_Pengirim
altpengirim.Text = Rspelanggan!Alamat_Pengirim
kodepospengirim.Text = Rspelanggan!Kode_Pos_Pengirim
telppengirim.Text = Rspelanggan!Telepon_Pengirim
kotapengirim.Text = Rspelanggan!Kota_Kabupaten_Pengirim
kelurahan.Text = Rspelanggan!Kelurahan_Pengirim
kecematan.Text = Rspelanggan!Kecematan_Pengirim
provinsipengirim.Text = Rspelanggan!Provinsi_Pengirim
nmpenerima.Text = Rspelanggan!Nama_Penerima
altpenerima.Text = Rspelanggan!Alamat_Penerima
kotapenerima.Text = Rspelanggan!Kota_Kabupaten_Penerima
telppenerima.Text = Rspelanggan!No_Telp_Penerima
kodepospenerima.Text = Rspelanggan!Kode_Pos_Penerima
kelurahanpenerima.Text = Rspelanggan!Kelurahan_Penerima
kecematanpenerima.Text = Rspelanggan!Kecematan_Penerima
provinsipenerima.Text = Rspelanggan!Provinsi_Penerima
End Sub



Private Sub nmkota_KeyPress(KeyAscii As Integer)
Call koneksi
If KeyAscii = 13 Then
Rstarif.Open "select * from tarif where Nama_Kota ='" & nmkota.Text & "'", con
If Not Rstarif.EOF Then
      Call tampil_tarif
      MsgBox "Menghitung volume?", vbYesNo, "info"
Else
MsgBox "Data tidak ditemukan !", vbOKOnly, "info"
nmkota.Text = ""
nmkota.SetFocus
End If
End If
End Sub
Private Sub tampil_tarif()
txtkota.Text = Rstarif!Kode_Kota
nmkota.Text = Rstarif!Nama_Kota
txttarif.Text = Rstarif!Tarif
provinsi.Text = Rstarif!provinsi
bagi.Text = Rstarif!volume
harga.Text = Rstarif!Total_Volume
pengambilan.Text = Rstarif!pengambilan
End Sub
Sub hitung()
subtotal.Text = Val(txttarif.Text) * Val(brt.Text)
End Sub
Private Sub grid_transaksi_Click()
    Call tampil
    Call aktif
    notrans.Enabled = False
    idpelanggan.Enabled = False
    cmdsimpan.Enabled = False
    cmdtambah.Enabled = False
End Sub
Private Sub tampil()
notrans.Text = Rstransaksi!No_Trans
kduser.Text = Rstransaksi!Kode_User
idpelanggan.Text = Rstransaksi!Id_Pelanggan
nmpenerima.Text = Rstransaksi!Nama_Penerima
brg.Text = Rstransaksi!Barang
brt.Text = Rstransaksi!Berat
plat.Text = Rstransaksi!No_Plat_Bus
noresi.Text = Rstransaksi!No_Resi
txttarif.Text = Rstransaksi!Tarif
ttl.Text = Rstransaksi!Total_Pengiriman
End Sub
Private Sub Form_Activate()
    Call kosong
    Call tidakaktif
    Call koneksi
    kduser.Text = Menuutama.StatusBar1.Panels(1).Text
    user.Text = Menuutama.StatusBar1.Panels(2).Text
    idpelanggan.Text = Form6.idpelanggan.Text
    nmpenerima.Text = Form6.nmpenerima.Text
    nmpengirim.Text = Form6.nmpengirim.Text
    altpengirim.Text = Form6.altpengirim.Text
    telppengirim.Text = Form6.telppengirim.Text
    kotapengirim.Text = Form6.kota.Text
    kodepospengirim.Text = Form6.kodepos.Text
    kelurahan.Text = Form6.kelurahan.Text
    kecematan.Text = Form6.kecematan.Text
    provinsipengirim.Text = Form6.provinsi.Text
    kelurahanpenerima.Text = Form6.kelurahanpenerima.Text
    kecematanpenerima.Text = Form6.kecematanpenerima.Text
    kotapenerima.Text = Form6.kotapenerima.Text
    provinsipenerima.Text = Form6.provinsipenerima.Text
    cmdtambah.Enabled = True
    Rstransaksi.Open "SELECT * FROM transaksi", con
    Rspelanggan.Open "select * from pelanggan", con
    Rsdetailtransaksi.Open "Select * From detailtransaksi", con
    Set DataGrid1.DataSource = Rspelanggan.DataSource
    Set grid_transaksi.DataSource = Rstransaksi.DataSource
    'Rspelanggan.Open "select * from pelanggan Id_Pelanggan where  ='" & idpelanggan.Text & "'", con
If Not Rspelanggan.EOF Then
Call tampilform

Rsuser.Open "select * from user where kode_user ='" & kduser.Text & "'", con
If Not Rsuser.EOF Then
Call tampiluser
 
grid_transaksi.Columns(0).Width = 1100
grid_transaksi.Columns(2).Width = 1100
grid_transaksi.Columns(3).Width = 1400
grid_transaksi.Columns(4).Width = 1300
grid_transaksi.Columns(5).Width = 1300
grid_transaksi.Columns(6).Width = 1300
 cmdtambah.SetFocus

End If
End If
End Sub



Private Sub tinggi_KeyPress(KeyAscii As Integer)
Dim l, p, t  As Integer
p = Val(panjang.Text)
t = Val(tinggi.Text)
l = Val(lebar.Text)
hasil.Text = Val(panjang.Text) * Val(tinggi.Text) * Val(lebar.Text)
hasil.SetFocus
End Sub
Private Sub harga_KeyPress(KeyAscii As Integer)
hasilpembagian.Text = Val(hasil.Text) / Val(bagi.Text) * Val(harga.Text)
End Sub
Private Sub ubay_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(ubay.Text) < Val(ttl.Text) Then
    MsgBox "uang anda kurang", vbInformation, "info"
    ukem.Text = ""
    Else
        ukem.Text = Val(ubay.Text) - Val(ttl.Text)
        MsgBox "Uang Kembali=" + ukem.Text, vbInformation, "info"
End If
End If
cmdsimpan.Enabled = True
cmdkeluar.Enabled = True
cmdcancel.Enabled = True
End Sub
Private Sub cmdcancel_Click()
Call Form_Activate
cmdsimpan.Enabled = False
End Sub
Sub cetak()
Call koneksi
CrystalReport1.SelectionFormula = "{transaksi.no_trans}='" & notrans.Text & "'"
CrystalReport1.ReportFileName = App.Path & "\struk.rpt"
CrystalReport1.WindowState = crptMaximized
CrystalReport1.RetrieveDataFiles
CrystalReport1.Action = 1
End Sub
Private Sub tampiluser()
kduser.Text = Rsuser!Kode_User
user.Text = Rsuser!Nama_User
End Sub
Private Sub tampilform()
idpelanggan.Text = Rspelanggan!Id_Pelanggan
nmpengirim.Text = Rspelanggan!Nama_Pengirim
altpengirim.Text = Rspelanggan!Alamat_Pengirim
kodepospengirim.Text = Rspelanggan!Kode_Pos_Pengirim
telppengirim.Text = Rspelanggan!Telepon_Pengirim
kotapengirim.Text = Rspelanggan!Kota_Kabupaten_Pengirim
kelurahan.Text = Rspelanggan!Kelurahan_Pengirim
kecematan.Text = Rspelanggan!Kecematan_Pengirim
provinsipengirim.Text = Rspelanggan!Provinsi_Pengirim
nmpenerima.Text = Rspelanggan!Nama_Penerima
altpenerima.Text = Rspelanggan!Alamat_Penerima
kotapenerima.Text = Rspelanggan!Kota_Kabupaten_Penerima
telppenerima.Text = Rspelanggan!No_Telp_Penerima
kodepospenerima.Text = Rspelanggan!Kode_Pos_Penerima
kelurahanpenerima.Text = Rspelanggan!Kelurahan_Penerima
kecematanpenerima.Text = Rspelanggan!Kecematan_Penerima
provinsipenerima.Text = Rspelanggan!Provinsi_Penerima
End Sub

