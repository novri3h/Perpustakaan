VERSION 5.00
Begin VB.Form Anggota 
   Caption         =   "Data Anggota"
   ClientHeight    =   2145
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdTutup 
      Caption         =   "&Tutup"
      Height          =   350
      Left            =   3480
      TabIndex        =   3
      Top             =   1680
      Width           =   850
   End
   Begin VB.CommandButton CmdHapus 
      Caption         =   "&Hapus"
      Height          =   350
      Left            =   1800
      TabIndex        =   2
      Top             =   1680
      Width           =   850
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   350
      Left            =   960
      TabIndex        =   1
      Top             =   1680
      Width           =   850
   End
   Begin VB.CommandButton CmdInput 
      Caption         =   "&Input"
      Height          =   350
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   850
   End
   Begin VB.TextBox TxtTelepon 
      Height          =   350
      Left            =   960
      TabIndex        =   7
      Top             =   1200
      Width           =   3400
   End
   Begin VB.TextBox TxtAlamat 
      Height          =   350
      Left            =   960
      TabIndex        =   6
      Top             =   840
      Width           =   3400
   End
   Begin VB.TextBox TxtNama 
      Height          =   350
      Left            =   960
      TabIndex        =   5
      Top             =   480
      Width           =   3400
   End
   Begin VB.TextBox TxtNomor 
      Height          =   350
      Left            =   960
      TabIndex        =   4
      Top             =   120
      Width           =   3400
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Telepon"
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   850
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Alamat"
      Height          =   345
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   850
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nama"
      Height          =   345
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Width           =   850
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nomor"
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   850
   End
End
Attribute VB_Name = "Anggota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Form_Load()
    Call BukaDB
    TxtNomor.MaxLength = 4
    TxtNama.MaxLength = 20
    TxtAlamat.MaxLength = 20
    TxtTelepon.MaxLength = 15
    KondisiAwal
End Sub

Function CariData()
    Call BukaDB
    RSAnggota.Open "Select * From Anggota where NomorAgt='" & TxtNomor & "'", Conn
End Function

Private Sub KosongkanText()
    TxtNomor = ""
    TxtNama = ""
    TxtAlamat = ""
    TxtTelepon = ""
End Sub

Private Sub SiapIsi()
    TxtNomor.Enabled = True
    TxtNama.Enabled = True
    TxtAlamat.Enabled = True
    TxtTelepon.Enabled = True
End Sub

Private Sub TidakSiapIsi()
    TxtNomor.Enabled = False
    TxtNama.Enabled = False
    TxtAlamat.Enabled = False
    TxtTelepon.Enabled = False
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
    With RSAnggota
        If Not RSAnggota.EOF Then
            TxtNama = RSAnggota!Namaagt
            TxtAlamat = RSAnggota!AlamatAgt
            TxtTelepon = RSAnggota!TeleponAgt
        End If
    End With
End Sub

Private Sub AutoNomor()
Call BukaDB
RSAnggota.Open ("select * from Anggota Where NomorAgt In(Select Max(NomorAgt)From Anggota)Order By NomorAgt Desc"), Conn
RSAnggota.Requery
    Dim Urutan As String * 4
    Dim Hitung As Long
    With RSAnggota
        If .EOF Then
            Urutan = "A" + "001"
            TxtNomor = Urutan
        Else
            Hitung = Right(!NomorAgt, 3) + 1
            Urutan = "A" + Right("000" & Hitung, 3)
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
        TxtNama.SetFocus
    Else
        If TxtNomor = "" Or TxtNama = "" Or TxtAlamat = "" Or TxtTelepon = "" Then
            MsgBox "Data Belum Lengkap...!"
        Else
            Dim SQLTambah As String
            SQLTambah = "Insert Into Anggota (NomorAgt,NamaAgt,AlamatAgt,TeleponAgt) values ('" & TxtNomor & "','" & TxtNama & "','" & TxtAlamat & "','" & TxtTelepon & "')"
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
        If TxtNama = "" Or TxtAlamat = "" Or TxtTelepon = "" Then
            MsgBox "Masih Ada Data Yang Kosong"
        Else
            Dim SQLEdit As String
            SQLEdit = "Update Anggota Set NamaAgt= '" & TxtNama & "', AlamatAgt='" & TxtAlamat & "',TeleponAgt='" & TxtTelepon & "' where NomorAgt='" & TxtNomor & "'"
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
        TxtNama.SetFocus
    End If

    If CmdInput.Caption = "&Simpan" Then
        Call CariData
            If Not RSAnggota.EOF Then
                TampilkanData
                MsgBox "Kode Anggota Sudah Ada"
                KosongkanText
                TxtNomor.SetFocus
            Else
                TxtNama.SetFocus
            End If
    End If
    
    If CmdEdit.Caption = "&Simpan" Then
        Call CariData
            If Not RSAnggota.EOF Then
                TampilkanData
                TxtNomor.Enabled = False
                TxtNama.SetFocus
            Else
                MsgBox "Kode Anggota Tidak Ada"
                TxtNomor = ""
                TxtNomor.SetFocus
            End If
    End If
    
    If CmdHapus.Enabled = True Then
        Call CariData
            If Not RSAnggota.EOF Then
                TampilkanData
                Pesan = MsgBox("Yakin akan dihapus", vbYesNo)
                If Pesan = vbYes Then
                    Dim SQLHapus As String
                    SQLHapus = "Delete From Anggota where NomorAgt= '" & TxtNomor & "'"
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

Private Sub TxtNama_KeyPress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then TxtAlamat.SetFocus
End Sub

Private Sub TxtAlamat_keypress(Keyascii As Integer)
    Keyascii = Asc(UCase(Chr(Keyascii)))
    If Keyascii = 13 Then TxtTelepon.SetFocus
End Sub

Private Sub TxtTelepon_keypress(Keyascii As Integer)
    If Keyascii = 13 Then
        If CmdInput.Enabled = True Then
            CmdInput.SetFocus
        ElseIf CmdEdit.Enabled = True Then
            CmdEdit.SetFocus
        End If
    End If
    If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9") Or Keyascii = vbKeyBack) Then Keyascii = 0
End Sub

