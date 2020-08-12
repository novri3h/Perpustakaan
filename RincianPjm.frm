VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form RincianPjm 
   Caption         =   "Data Rincian Peminjaman Buku"
   ClientHeight    =   2940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5715
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
   ScaleHeight     =   2940
   ScaleWidth      =   5715
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2640
      Top             =   2160
      Visible         =   0   'False
      Width           =   1845
      _ExtentX        =   3254
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
      Caption         =   "Adodc1"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "RincianPjm.frx":0000
      Height          =   1695
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   4335
      _ExtentX        =   7646
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
      ColumnCount     =   2
      BeginProperty Column00 
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
      BeginProperty Column01 
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
            ColumnWidth     =   3000,189
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   794,835
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List1 
      Height          =   1410
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1100
   End
   Begin VB.Label Jumlah 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   4800
      TabIndex        =   7
      Top             =   2400
      Width           =   750
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Jumlah"
      Height          =   315
      Left            =   4800
      TabIndex        =   6
      Top             =   2040
      Width           =   750
   End
   Begin VB.Label Anggota 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   2400
      Width           =   1245
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Anggota"
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   2040
      Width           =   1245
   End
   Begin VB.Label Tanggal 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1100
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tanggal"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1100
   End
End
Attribute VB_Name = "RincianPjm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error Resume Next
Call BukaDB
List1.Clear
RSPinjam.Open "Select Distinct NomorPjm from Pinjam where totalpjm<>0", Conn
Do Until RSPinjam.EOF
    List1.AddItem RSPinjam!NomorPjm
    RSPinjam.MoveNext
Loop
Conn.Close

End Sub

Private Sub list1_click()
Call BukaDB
Conn.CursorLocation = adUseClient
RSPinjam.Open "select * from Pinjam where NomorPjm='" & List1.Text & "'", Conn
RSPinjam.Requery

If Not RSPinjam.EOF Then Tanggal = RSPinjam!TanggalPjm

RSAnggota.Open "select * from Anggota where NomorAgt='" & RSPinjam!NomorAgt & "'", Conn
If Not RSAnggota.EOF Then Anggota = RSAnggota!Namaagt
Conn.Close

Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\ADOPustaka.mdb"
Adodc1.RecordSource = "select Judul, jumlahbk as Jumlah from Buku,detailpjm,Pinjam where DetailPjm.Nomorbk=Buku.Nomorbk and left(detailPjm.NomorPjm,8)=Pinjam.NomorPjm and Pinjam.NomorPjm='" & List1 & "'"
Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
DataGrid1.Refresh
Jumlah = Adodc1.Recordset.RecordCount
End Sub

Private Sub List1_keyPress(Keyascii As Integer)
If Keyascii = 27 Then Unload Me
End Sub

