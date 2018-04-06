VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form10 
   BackColor       =   &H8000000D&
   Caption         =   "Form10"
   ClientHeight    =   8835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16785
   ControlBox      =   0   'False
   LinkTopic       =   "Form10"
   ScaleHeight     =   8835
   ScaleWidth      =   16785
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdkeluar 
      Caption         =   "Keluar"
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdcetak 
      Caption         =   "Cetak"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   3120
      Width           =   1335
   End
   Begin Crystal.CrystalReport cr 
      Left            =   5400
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      Caption         =   "Laporan Transaksi"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   4335
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   134938625
         CurrentDate     =   42963
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   134938625
         CurrentDate     =   42963
      End
      Begin VB.Label Label2 
         Caption         =   "Tanggal Akhir"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Tanggal Awal"
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcetak_Click()
    cr.SelectionFormula = "{transaksi.tanggal}>=date(" & Year(DTPicker1.Value) & "," & Month(DTPicker1.Value) & "," & Day(DTPicker1.Value) & ") and {transaksi.tanggal}<=date (" & Year(DTPicker2.Value) & "," & Month(DTPicker2.Value) & "," & Day(DTPicker2.Value) & ")"
     'cr.ReportFileName = App.Path & "\REPORT_PERIODE.rpt"
    cr.ReportFileName = App.Path & "\LAPORAN TRANSAKSI.rpt"
    cr.RetrieveDataFiles
    cr.Action = 1
End Sub
Private Sub cmdkeluar_Click()
Unload Me
End Sub

Private Sub Form_Activate()
DTPicker1.Value = Format(Date, "DD/MM/YYYY")
DTPicker2.Value = Format(Date, "DD/MM/YYYY")
cmdcetak.SetFocus
End Sub

