VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Menu 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Menu Utama"
   ClientHeight    =   4440
   ClientLeft      =   195
   ClientTop       =   765
   ClientWidth     =   6225
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
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   4440
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   10
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "a"
            Object.ToolTipText     =   "Anggota"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "b"
            Object.ToolTipText     =   "Buku"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "c"
            Object.ToolTipText     =   "Peminjaman"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "d"
            Object.ToolTipText     =   "Pengembalian"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "e"
            Object.ToolTipText     =   "Laporan Data Anggota"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "f"
            Object.ToolTipText     =   "Laporan Data Buku"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "g"
            Object.ToolTipText     =   "Laporan Transaksi"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "h"
            Object.ToolTipText     =   "Rincian Peminjaman"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "i"
            Object.ToolTipText     =   "Rincian Pengembalian"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "k"
            Object.ToolTipText     =   "Keluar"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport CR 
      Left            =   2760
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":26876
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":26B90
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":26EAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":271C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":274DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":277F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":27B12
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":27E2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":28146
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Menu.frx":28460
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnfile 
      Caption         =   "File"
      Begin VB.Menu mnanggota 
         Caption         =   "Anggota"
      End
      Begin VB.Menu mnbuku 
         Caption         =   "Buku"
      End
   End
   Begin VB.Menu mntransaksi 
      Caption         =   "Transaksi"
      Begin VB.Menu mnpinjaman 
         Caption         =   "Pinjaman"
      End
      Begin VB.Menu mnkembali 
         Caption         =   "Pengembalian"
      End
   End
   Begin VB.Menu mnlaporan 
      Caption         =   "Laporan"
      Begin VB.Menu mnlapbuku 
         Caption         =   "Data Buku"
      End
      Begin VB.Menu mnlapanggota 
         Caption         =   "Data Anggota"
      End
      Begin VB.Menu mnlaptransaksi 
         Caption         =   "Laporan Transaksi"
      End
   End
   Begin VB.Menu mnrincian 
      Caption         =   "Rincian"
      Begin VB.Menu mnrincianpjm 
         Caption         =   "Pinjaman"
      End
      Begin VB.Menu mnrinciankbl 
         Caption         =   "Pengembalian"
      End
   End
   Begin VB.Menu mnkeluar 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then End
End Sub

Private Sub mnanggota_Click()
Anggota.Show
End Sub

Private Sub mnbuku_Click()
Buku.Show
End Sub

Private Sub mnkeluar_Click()
End
End Sub

Private Sub mnkembali_Click()
Kembali.Show
End Sub

Private Sub mnlapanggota_Click()
CR.ReportFileName = App.Path & "\Lap Anggota.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 0
End Sub

Private Sub mnlapbuku_Click()
CR.ReportFileName = App.Path & "\Lap Buku.rpt"
CR.WindowState = crptMaximized
CR.RetrieveDataFiles
CR.Action = 0
End Sub

Private Sub mnlaptransaksi_Click()
Laporan.Show
End Sub

Private Sub mnpinjaman_Click()
Pinjam.Show
End Sub

Private Sub mnrinciankbl_Click()
RincianKbl.Show
End Sub

Private Sub mnrincianpjm_Click()
RincianPjm.Show
End Sub

Private Sub mnuji_Click()
UjiSQL.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Key
    Case "a"
        Anggota.Show
    Case "b"
        Buku.Show
    Case "c"
        Pinjam.Show
    Case "d"
        Kembali.Show
    Case "e"
       CR.ReportFileName = App.Path & "\Lap anggota.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1
    Case "f"
       CR.ReportFileName = App.Path & "\Lap buku.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1
    Case "g"
        Laporan.Show
    Case "h"
        RincianPjm.Show
    Case "i"
        RincianKbl.Show
    Case "j"
        End
End Select

End Sub
