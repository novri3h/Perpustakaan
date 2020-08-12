VERSION 5.00
Begin VB.Form Buku 
   Caption         =   "Data Buku"
   ClientHeight    =   2850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4515
   BeginProperty Font 
      Name            =   "Century"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   2850
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtTahun 
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
      TabIndex        =   9
      Top             =   1560
      Width           =   3200
   End
   Begin VB.TextBox TxtStok 
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
      Top             =   1920
      Width           =   3200
   End
   Begin VB.TextBox TxtNomor 
      Height          =   350
      Left            =   1200
      TabIndex        =   5
      Top             =   120
      Width           =   3200
   End
   Begin VB.TextBox TxtJudul 
      Height          =   350
      Left            =   1200
      TabIndex        =   6
      Top             =   480
      Width           =   3200
   End
   Begin VB.TextBox TxtPengarang 
      Height          =   350
      Left            =   1200
      TabIndex        =   7
      Top             =   840
      Width           =   3200
   End
   Begin VB.TextBox TxtPenerbit 
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
      TabIndex        =   8
      Top             =   1200
      Width           =   3200
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "&Input"
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
      TabIndex        =   1
      Top             =   2400
      Width           =   1000
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
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
      TabIndex        =   2
      Top             =   2400
      Width           =   1000
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
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
      Left            =   2280
      TabIndex        =   3
      Top             =   2400
      Width           =   1000
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
      Left            =   3360
      TabIndex        =   4
      Top             =   2400
      Width           =   1000
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tahun"
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
      TabIndex        =   15
      Top             =   1560
      Width           =   1005
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Stok"
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
      TabIndex        =   14
      Top             =   1920
      Width           =   1005
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nomor"
      Height          =   345
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Judul"
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Pengarang"
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Penerbit"
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
      TabIndex        =   10
      Top             =   1200
      Width           =   1005
   End
End
Attribute VB_Name = "Buku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Form_Load()
    Call BukaDB
    TxtNomor.MaxLength = 4
    TxtJudul.MaxLength = 30
    TxtPengarang.MaxLength = 20
    TxtPenerbit.MaxLength = 30
    TxtTahun.MaxLength = 4
    TxtStok.MaxLength = 3
    KondisiAwal
End Sub

Function CariData()
    Call BukaDB
    RSBuku.Open "Select * From Buku where NomorBk='" & TxtNomor & "'", Conn
End Function

Private Sub KosongkanText()
    TxtNomor = ""
    TxtJudul = ""
    TxtPengarang = ""
    TxtPenerbit = ""
    TxtTahun = ""
    TxtStok = ""
End Sub

Private Sub SiapIsi()
    TxtNomor.Enabled = True
    TxtJudul.Enabled = True
    TxtPengarang.Enabled = True
    TxtPenerbit.Enabled = True
    TxtTahun.Enabled = True
    TxtStok.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    TxtNomor.Enabled = False
    TxtJudul.Enabled = False
    TxtPengarang.Enabled = False
    TxtPenerbit.Enabled = False
    TxtTahun.Enabled = False
    TxtStok.Enabled = False
End Sub

Private Sub KondisiAwal()
    KosongkanText
    TidakSiapIsi
    CmdInput.Caption = "&Input"
    CmdEdit.Caption = "&Edit"
    CmdHapus.Caption = "&Hapus"
    CmdTutup.Caption = "&Tutup"
    CmdInput.Enabled = True
    CmdEdit.Enabled = True
    CmdHapus.Enabled = True
End Sub

Private Sub TampilkanData()
    With RSBuku
        If Not RSBuku.EOF Then
            TxtJudul = RSBuku!Judul
            TxtPengarang = RSBuku!Pengarang
            TxtPenerbit = RSBuku!Penerbit
            TxtTahun = RSBuku!Tahun
            TxtStok = RSBuku!Stok
        End If
    End With
End Sub

Private Sub AutoNomor()
Call BukaDB
RSBuku.Open ("select * from Buku Where NomorBk In(Select Max(NomorBk)From Buku)Order By NomorBk Desc"), Conn
RSBuku.Requery
    Dim Urutan As String * 4
    Dim Hitung As Long
    With RSBuku
        If .EOF Then
            Urutan = "B" + "001"
            TxtNomor = Urutan
        Else
            Hitung = Right(!NomorBk, 3) + 1
            Urutan = "B" + Right("000" & Hitung, 3)
        End If
        TxtNomor = Urutan
    End With
End Sub

Private Sub CmdInput_Click()
    If CmdInput.Caption = "&Input" Then
        CmdInput.Caption = "&Simpan"
        CmdEdit.Enabled = False
        CmdHapus.Enabled = False
        CmdTutup.Caption = "&Batal"
        SiapIsi
        KosongkanText
        Call AutoNomor
        TxtNomor.Enabled = False
        TxtJudul.SetFocus
    Else
        If TxtNomor = "" Or TxtJudul = "" Or TxtPengarang = "" Or TxtPenerbit = "" Or TxtTahun = "" Or TxtStok = "" Then
            MsgBox "Data Belum Lengkap...!"
        Else
            Dim SQLTambah As String
            SQLTambah = "Insert Into Buku (NomorBk,Judul,Pengarang,Penerbit,tahun,Stok) values ('" & TxtNomor & "','" & TxtJudul & "','" & TxtPengarang & "','" & TxtPenerbit & "','" & TxtTahun & "','" & TxtStok & "')"
            Conn.Execute SQLTambah
            KondisiAwal
        End If
    End If
End Sub

Private Sub CmdEdit_Click()
    If CmdEdit.Caption = "&Edit" Then
        CmdInput.Enabled = False
        CmdEdit.Caption = "&Simpan"
        CmdHapus.Enabled = False
        CmdTutup.Caption = "&Batal"
        SiapIsi
        TxtNomor.SetFocus
    Else
        If TxtJudul = "" Or TxtPengarang = "" Or TxtPenerbit = "" Or TxtTahun = "" Or TxtStok = "" Then
            MsgBox "Masih Ada Data Yang Kosong"
        Else
            Dim SQLEdit As String
            SQLEdit = "Update Buku Set Judul= '" & TxtJudul & "', pengarang='" & TxtPengarang & "',penerbit='" & TxtPenerbit & "', tahun='" & TxtTahun & "',stok='" & TxtStok & "' where NomorBk='" & TxtNomor & "'"
            Conn.Execute SQLEdit
            KondisiAwal
        End If
    End If
End Sub

Private Sub CmdHapus_Click()
    If CmdHapus.Caption = "&Hapus" Then
        CmdInput.Enabled = False
        CmdEdit.Enabled = False
        CmdTutup.Caption = "&Batal"
        KosongkanText
        SiapIsi
        TxtNomor.SetFocus
    End If
End Sub

Private Sub cmdtutup_Click()
    Select Case CmdTutup.Caption
        Case "&Tutup"
            Unload Me
        Case "&Batal"
            TidakSiapIsi
            KondisiAwal
    End Select
End Sub

Private Sub TxtNomor_KeyPress(Keyascii As Integer)
Keyascii = Asc(UCase(Chr(Keyascii)))
If Keyascii = 13 Then
    If Len(TxtNomor) < 4 Then
        MsgBox "Kode Harus 4 Digit"
        TxtNomor.SetFocus
    Else
        TxtJudul.SetFocus
    End If

    If CmdInput.Caption = "&Simpan" Then
        Call CariData
            If Not RSBuku.EOF Then
                TampilkanData
                MsgBox "Kode Buku Sudah Ada"
                KosongkanText
                TxtNomor.SetFocus
            Else
                TxtJudul.SetFocus
            End If
    End If
    
    If CmdEdit.Caption = "&Simpan" Then
        Call CariData
            If Not RSBuku.EOF Then
                TampilkanData
                TxtNomor.Enabled = False
                TxtJudul.SetFocus
            Else
                MsgBox "Kode Buku Tidak Ada"
                TxtNomor = ""
                TxtNomor.SetFocus
            End If
    End If
    
    If CmdHapus.Enabled = True Then
        Call CariData
            If Not RSBuku.EOF Then
                TampilkanData
                Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
                If Pesan = vbYes Then
                    Dim SQLHapus As String
                    SQLHapus = "Delete From Buku where NomorBk= '" & TxtNomor & "'"
                    Conn.Execute SQLHapus
                    KondisiAwal
                Else
                    KondisiAwal
                    CmdHapus.SetFocus
                End If
            Else
                MsgBox "Data Tidak ditemukan"
                TxtNomor.SetFocus
            End If
    End If
End If
End Sub

Private Sub txtjudul_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then TxtPengarang.SetFocus
End Sub

Private Sub txtpengarang_keypress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then TxtPenerbit.SetFocus
End Sub

Private Sub txtpenerbit_keypress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then TxtTahun.SetFocus
End Sub

Private Sub txttahun_keypress(Keyascii As Integer)
    If Keyascii = 13 Then TxtStok.SetFocus
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

Private Sub txtstok_keypress(Keyascii As Integer)
    If Keyascii = 13 Then
        If CmdInput.Enabled = True Then
            CmdInput.SetFocus
        ElseIf CmdEdit.Enabled = True Then
            CmdEdit.SetFocus
        End If
    End If
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

