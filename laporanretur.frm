VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form9 
   BackColor       =   &H8000000D&
   Caption         =   "Form9"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12105
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   ScaleHeight     =   5610
   ScaleWidth      =   12105
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Laporan Retur"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   4695
      Begin VB.CommandButton cmdkeluar 
         BackColor       =   &H8000000D&
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
         Left            =   2520
         TabIndex        =   6
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdcetak 
         Caption         =   "Cetak"
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
         TabIndex        =   5
         Top             =   2640
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   1920
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         _Version        =   393216
         Format          =   88473601
         CurrentDate     =   42937
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   840
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         _Version        =   393216
         Format          =   88473601
         CurrentDate     =   42937
      End
      Begin VB.Label Label2 
         Caption         =   "Tanggal Akhir"
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
         Left            =   480
         TabIndex        =   4
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Tanggal Awal"
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
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
   End
   Begin Crystal.CrystalReport cr 
      Left            =   5520
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcetak_Click()
    cr.SelectionFormula = "{retur.tanggal}>=date(" & Year(DTPicker1.Value) & "," & Month(DTPicker1.Value) & "," & Day(DTPicker1.Value) & ")and {retur.tanggal}<=date (" & Year(DTPicker2.Value) & "," & Month(DTPicker2.Value) & "," & Day(DTPicker2.Value) & ")"
     'cr.ReportFileName = App.Path & "\REPORT_PERIODE.rpt"
    cr.ReportFileName = App.Path & "\LaporanReturr.rpt"
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

