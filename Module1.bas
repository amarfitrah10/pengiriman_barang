Attribute VB_Name = "Module1"
Public con  As New ADODB.Connection
Public Rsuser As ADODB.Recordset
Public Rsbarang As ADODB.Recordset
Public Rstarif As ADODB.Recordset
Public Rstransaksi As ADODB.Recordset
Public Rspelanggan As ADODB.Recordset
Public Rsretur As ADODB.Recordset
Public Rsdetailtransaksi As ADODB.Recordset
Sub koneksi()
Set con = New ADODB.Connection
Set Rsuser = New ADODB.Recordset
Set Rsbarang = New ADODB.Recordset
Set Rstarif = New ADODB.Recordset
Set Rstransaksi = New ADODB.Recordset
Set Rspelanggan = New ADODB.Recordset
Set Rsretur = New ADODB.Recordset
Set Rsdetailtransaksi = New ADODB.Recordset
con.Open "DRIVER={MySQL ODBC 5.1 Driver};SERVER=localhost;DATABASE=pengirimanbarangobl;UID=root;PWD="
con.CursorLocation = adUseClient
End Sub






