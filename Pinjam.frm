VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Pinjam 
   Caption         =   "Peminjaman Buku"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9375
   LinkTopic       =   "Form3"
   ScaleHeight     =   5055
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4560
      Left            =   6720
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1800
      TabIndex        =   5
      Top             =   2760
      Width           =   800
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   960
      TabIndex        =   4
      Top             =   2760
      Width           =   800
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   800
   End
   Begin VB.TextBox TxtNomorAgt 
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   1250
   End
   Begin MSAdodcLib.Adodc DT 
      Height          =   375
      Left            =   2280
      Top             =   1680
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Transaksi"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DTCari 
      Height          =   375
      Left            =   2760
      Top             =   2760
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DTCari"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DG1 
      Bindings        =   "Pinjam.frx":0000
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   2990
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Nomor"
         Caption         =   "Nomor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Kode"
         Caption         =   "Kode"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Judul"
         Caption         =   "Judul"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Jumlah"
         Caption         =   "Jumlah"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   645,165
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3750,236
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   750,047
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DG2 
      Bindings        =   "Pinjam.frx":001A
      Height          =   1695
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2990
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   2
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "NOMORBK"
         Caption         =   "Kode"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "JUDUL"
         Caption         =   "Judul Buku"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "JUMLAHBK"
         Caption         =   "Jumlah"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4500,284
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnWidth     =   750,047
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Telah Pinjam"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2520
      TabIndex        =   16
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label LbltelahPjm 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3600
      TabIndex        =   15
      Top             =   120
      Width           =   540
   End
   Begin VB.Label LblNamaAgt 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2520
      TabIndex        =   13
      Top             =   480
      Width           =   3900
   End
   Begin VB.Label LblTotalPjm 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5640
      TabIndex        =   12
      Top             =   2760
      Width           =   750
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5040
      TabIndex        =   11
      Top             =   2760
      Width           =   555
   End
   Begin VB.Label LblTanggal 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4920
      TabIndex        =   10
      Top             =   120
      Width           =   1500
   End
   Begin VB.Label LblNomorPjm 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1200
      TabIndex        =   9
      Top             =   120
      Width           =   1250
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Anggota"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4200
      TabIndex        =   7
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " No Pinjam"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1005
   End
End
Attribute VB_Name = "Pinjam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
'hubungkan objek adodc ke database
DT.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOPustaka.mdb"
'hubungkan objek adodc ke tabel
DT.RecordSource = "Transaksi"
'sumber data untuk grid1 adalah data dalam objek
Set DG1.DataSource = DT
'grid di-refresh
DG1.Refresh
'panggil file database
Call BukaDB
'buka tabel buku dan tampilkan kode dan judlnya di list sebelah kanan
RSBuku.Open "SELECT * FROM BUKU WHERE STOK >0", Conn
List1.Clear
Do Until RSBuku.EOF
    List1.AddItem RSBuku!Judul & Space(50) & RSBuku!NomorBk
    RSBuku.MoveNext
Loop
'tampilkan nomor pinjam otomatis
Call AutoNomor
LblTanggal.Caption = Format(Date, "dd-mm-yyyy")
Call Tabel_Kosong
DT.Recordset.MoveFirst
DG1.Col = 1
End Sub

'cari nomor pinjaman terakhir
Private Sub AutoNomor()
Call BukaDB
RSPinjam.Open "select * from Pinjam Where NomorPjm In(Select Max(NomorPjm)From Pinjam)Order By NomorPjm Desc", Conn
RSPinjam.Requery
    Dim Urutan As String * 8
    Dim Hitung As Long
    With RSPinjam
        If .EOF Then
            Urutan = Format(Date, "yymmdd") + "01"
            LblNomorPjm = Urutan
        Else
            If Left(!NomorPjm, 6) <> Format(Date, "yymmdd") Then
                Urutan = Format(Date, "yymmdd") + "01"
            Else
                Hitung = (!NomorPjm) + 1
                Urutan = Format(Date, "yymmdd") + Right("00" & Hitung, 2)
            End If
        End If
        LblNomorPjm = Urutan
    End With
End Sub

Private Sub TxtNomorAgt_KeyPress(Keyascii As Integer)
TxtNomorAgt.MaxLength = 4
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 27 Then Unload Me
If Keyascii = 13 Then
    Call BukaDB
    'cari nomor anggota yang diketik
    RSAnggota.Open "Select * from anggota where nomoragt='" & TxtNomorAgt & "'", Conn
    'jika ditemukan
    If Not RSAnggota.EOF Then
        'tampilkan namanya
        LblNamaAgt.Caption = RSAnggota!Namaagt
        DG1.SetFocus
        DG1.Col = 1
    Else
        'jika tidak ditemukan, munculkan pesan
        MsgBox "Nomor anggota tidak terdaftar"
        TxtNomorAgt.SetFocus
        Exit Sub
    End If
       
    Call Pinjaman
    
    'batas-batas peminjaman
    If LbltelahPjm = 0 Or LbltelahPjm = "" Then
        Pesan = MsgBox(" " & LblNamaAgt & " Silahkan Pinjam Maksimal " & 4 & " Buku", 0, "Informasi Peminjaman Buku")
        DG1.SetFocus
        DG1.Col = 1
    ElseIf LbltelahPjm = 1 Then
        Pesan = MsgBox(" " & LblNamaAgt & " Boleh Meminjam " & 3 & " Buku Lagi", 0, "Informasi Peminjaman Buku")
        DG1.SetFocus
        DG1.Col = 1
        Exit Sub
    ElseIf LbltelahPjm = 2 Then
        Pesan = MsgBox(" " & LblNamaAgt & " Boleh Meminjam " & 2 & " Buku Lagi", 0, "Informasi Peminjaman Buku")
        DG1.SetFocus
        DG1.Col = 1
        Exit Sub
    ElseIf LbltelahPjm = 3 Then
        Pesan = MsgBox(" " & LblNamaAgt & " Boleh Meminjam " & 1 & " Buku Lagi", 0, "Informasi Peminjaman Buku")
        DG1.SetFocus
        DG1.Col = 1
        Exit Sub
    ElseIf LbltelahPjm >= 4 Then
        Pesan = MsgBox(" " & LblNamaAgt & "  Tidak Boleh Meminjam Lagi...!", 0, "Informasi Peminjaman")
        LbltelahPjm = ""
        LblNamaAgt = ""
        TxtNomorAgt.SetFocus
        Exit Sub
    End If
End If
End Sub

Sub Pinjaman()
DTCari.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOPustaka.mdb"
DTCari.RecordSource = "Select Buku.Nomorbk,Judul,Jumlahbk From Anggota,Pinjam,Buku,Detailpjm Where Buku.Nomorbk=Detailpjm.Nomorbk And Pinjam.Nomorpjm=Left(Detailpjm.Nomorpjm,8) And Anggota.Nomoragt=Pinjam.Nomoragt And Anggota.Nomoragt='" & TxtNomorAgt & "'"
DTCari.Refresh
DG2.Refresh
LbltelahPjm.Caption = DTCari.Recordset.RecordCount
End Sub

'transaksi peminjaman dlam grid1
Private Sub DG1_AfterColEdit(ByVal ColIndex As Integer)
If DG1.Col = 1 Then
    Call BukaDB
    Dim RS As New ADODB.Recordset
    RS.Open "Select * from transaksi where kode='" & DT.Recordset!KODE & "'", Conn
    If Not RS.EOF Then
        MsgBox "Kode buku sudah dientri sebelumnya"
        Exit Sub
    End If
    'cari kode buku
    RSBuku.Open "Select * from Buku where NomorBK='" & DT.Recordset!KODE & "'", Conn
    'jika tidak ditemukan, munculkan pesan
    If RSBuku.EOF Then
        Pesan = MsgBox("Kode Buku Tidak Terdaftar")
        DG1.Col = 1
        Exit Sub
    End If
    'jika ditemukan, tampilkan nomor dan judul buku
    DT.Recordset!KODE = RSBuku!NomorBk
    DT.Recordset!Judul = RSBuku!Judul
    'jumlah pinjam asumsinya 1 buku
    DT.Recordset!Jumlah = 1
    'pindah ke baris berikutnya
    Call Tambah_Baris
    DT.Recordset.MoveNext
    DG1.Col = 1
    DT.Recordset.MoveLast
    'tampilkan jumlah total pinjaman
    LblTotalPjm.Caption = DT.Recordset.RecordCount - 1
End If

If DG1.Col = 3 Then
    DT.Recordset!Jumlah = DT.Recordset!Jumlah
    DT.Recordset.Update
    DT.Recordset.MoveNext
    DG1.Refresh
    DG1.Col = 1
    LblTotalPjm.Caption = DT.Recordset.RecordCount - 1
End If

If Val(LbltelahPjm) + Val(LblTotalPjm) = 4 Then
    MsgBox "pinjaman sudah masimal"
    DG1.AllowAddNew = False
    DG1.AllowUpdate = False
    CmdSimpan.SetFocus
    Exit Sub
'jika jumlah telah pinjam dan pinjaman sekarang lebih dari 4,
'munculkan pesan bahwa pinjaman telah maksimal
ElseIf Val(LbltelahPjm) + Val(LblTotalPjm) > 4 Then
    MsgBox "pinjaman melebihi batas, edit jumlah pinjaman"
    DG1.AllowAddNew = True
    DG1.AllowUpdate = True
    DG1.SetFocus
    CmdSimpan.SetFocus
    Exit Sub
End If
End Sub

Private Sub cmdSimpan_Click()
'jika total pinjaman belum ada, tampilkan pesan
If LblTotalPjm.Caption = "" Then
    MsgBox "Tidak ada transaksi peminjaman"
    TxtNomorAgt.SetFocus
    Exit Sub
End If

'simpan ke tabel pinjam
Dim SQLInput1 As String
SQLInput1 = "Insert Into Pinjam(Nomorpjm,TanggalPjm,TotalPjm,Nomoragt)" & _
"values('" & LblNomorPjm.Caption & "','" & LblTanggal.Caption & "','" & LblTotalPjm.Caption & "','" & TxtNomorAgt & "')"
Conn.Execute (SQLInput1)

'simpan ke tabel detailpjm
DT.Recordset.MoveFirst
Do While Not DT.Recordset.EOF
    If DT.Recordset!KODE <> vbNullString Then
        Dim SQLInput2 As String
        SQLInput2 = "Insert Into DetailPjm(Nomorpjm,Nomorbk,Jumlahbk) " & _
        "values ('" & LblNomorPjm.Caption + DT.Recordset!Nomor & "','" & DT.Recordset!KODE & "','" & DT.Recordset!Jumlah & "')"
        Conn.Execute (SQLInput2)
    End If
DT.Recordset.MoveNext
Loop
    
'Pengurangan Jumlah buku
DT.Recordset.MoveFirst
Do While Not DT.Recordset.EOF
    If DT.Recordset!KODE <> vbNullString Then
        Call BukaDB
        RSBuku.Open "Select * from Buku where Nomorbk='" & DT.Recordset!KODE & "'", Conn
        If Not RSBuku.EOF Then
            Dim kurangi As String
            kurangi = "update buku set stok='" & RSBuku!Stok - DT.Recordset!Jumlah & "' where nomorbk='" & DT.Recordset!KODE & "'"
            Conn.Execute (kurangi)
        End If
    End If
DT.Recordset.MoveNext
Loop
Bersihkan
Form_Activate
cmdbatal_Click
End Sub

Sub Bersihkan()
TxtNomorAgt = ""
LblNamaAgt.Caption = ""
LblTotalPjm.Caption = ""
LbltelahPjm.Caption = ""
End Sub

Function Tabel_Kosong()
DT.Recordset.MoveFirst
Do While Not DT.Recordset.EOF
    DT.Recordset.Delete
    DT.Recordset.MoveNext
Loop
For i = 1 To 1
    DT.Recordset.AddNew
    DT.Recordset!Nomor = i
    DT.Recordset.Update
Next i
End Function

Private Sub cmdbatal_Click()
Form_Activate
TxtNomorAgt = ""
LblNamaAgt = ""
LblTotalPjm = ""
LbltelahPjm = ""
DG1.Enabled = True
Call Pinjaman
TxtNomorAgt.SetFocus
End Sub

Private Sub cmdtutup_Click()
Unload Me
End Sub

Function Tambah_Baris()
For i = DT.Recordset.RecordCount To DT.Recordset.RecordCount
    DT.Recordset.AddNew
    DT.Recordset!Nomor = i + 1
    DT.Recordset.Update
Next i
End Function

Function Kurangi_Baris()
For i = DT.Recordset.RecordCount To DT.Recordset.RecordCount
    DT.Recordset.Delete
    DT.Recordset.Update
Next i
End Function

'jika menekan ESC dalam grid transaksi
'data akan hilang (dibatalkan) dan baris berkurang
Private Sub DG1_Keypress(Keyascii As Integer)
On Error GoTo salah

Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 27 Then
    DT.Recordset!KODE = Null
    DT.Recordset!Judul = Null
    DT.Recordset!Jumlah = Null
    DT.Recordset.Update
    Call Kurangi_Baris
    LblTotalPjm.Caption = DT.Recordset.RecordCount - 1
End If
On Error GoTo 0
Exit Sub
salah:
cmdbatal_Click
End Sub

Private Sub List1_keyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        If DG1.SelText <> Right(List1, 4) Then
            DG1.SelText = Right(List1, 4)
            DT.Recordset.Update
            Call BukaDB
            
            'Call BukaDB
            Dim cari As New ADODB.Recordset
            cari.Open "select * from transaksi where KODE= '" & Right(List1, 4) & "'", Conn
            'cari.Open "select * from transaksi where KODE= '" & DT.Recordset!KODE & "'", Conn
            If Not cari.EOF Then
                MsgBox "data jangan dientri dua kali"
                Exit Sub
            Else
            '    Call SelectAllVisible
            'End If
    
                RSBuku.Open "Select * from Buku where nomorBk ='" & Right(List1, 4) & "'", Conn
                RSBuku.Requery
                If Not RSBuku.EOF Then
                    DT.Recordset!KODE = RSBuku!NomorBk
                    DT.Recordset!Judul = RSBuku!Judul
                    DT.Recordset!Jumlah = 1
                    Call Tambah_Baris
                    DT.Recordset.MoveNext
                    DG1.Col = 1
                    DT.Recordset.MoveLast
                    'LblTotalPjm.Caption = Format(TotalPjm, "##")
                    LblTotalPjm.Caption = DT.Recordset.RecordCount - 1
                    
                    If Val(LbltelahPjm) + Val(LblTotalPjm) = 4 Then
                        MsgBox "Pinjaman Sudah Maksimal"
                        DG1.AllowAddNew = False
                        DG1.AllowUpdate = False
                        CmdSimpan.SetFocus
                        Exit Sub
                    'jika jumlah telah pinjam dan pinjaman sekarang lebih dari 4,
                    'munculkan pesan bahwa pinjaman telah maksimal
                    ElseIf Val(LbltelahPjm) + Val(LblTotalPjm) > 4 Then
                        MsgBox "Pinjaman melebihi batas, edit jumlah pinjaman"
                        DG1.SetFocus
                        CmdSimpan.SetFocus
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
End Sub

