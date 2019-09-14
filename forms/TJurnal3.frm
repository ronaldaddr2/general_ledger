VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form TJurnal3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaksi Jurnal Umum"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14355
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   14355
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtcab 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "F1 : List"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "TJurnal3.frx":0000
      Left            =   9600
      List            =   "TJurnal3.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "TJurnal3.frx":0019
      Left            =   9600
      List            =   "TJurnal3.frx":0023
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Keterangan 
      Height          =   885
      Left            =   1800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1440
      Width           =   5415
   End
   Begin VB.TextBox Notrans 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   2295
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   12360
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3495
      Left            =   360
      TabIndex        =   4
      Top             =   2880
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   6165
      _Version        =   393216
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   0
      ForeColor       =   16711680
      HeadLines       =   1
      RowHeight       =   16
      TabAction       =   2
      WrapCellPointer =   -1  'True
      RowDividerStyle =   0
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin MSComCtl2.DTPicker Tgl 
      Height          =   285
      Left            =   4080
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   164167683
      CurrentDate     =   40197
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   14355
      _ExtentX        =   25321
      _ExtentY        =   1058
      ButtonWidth     =   1799
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Data Baru"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Rubah"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Simpan"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "List Data"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Hapus"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cetak"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Keluar"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   10560
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TJurnal3.frx":0032
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TJurnal3.frx":0A44
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TJurnal3.frx":1456
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TJurnal3.frx":1E68
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TJurnal3.frx":287A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TJurnal3.frx":328C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TJurnal3.frx":3C9E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label8 
      Caption         =   "Cabang"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lblcab 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3120
      TabIndex        =   16
      Top             =   2400
      Width           =   4455
   End
   Begin VB.Label total 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   9960
      TabIndex        =   12
      Top             =   6600
      Width           =   2895
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SELISIH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9120
      TabIndex        =   11
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEBET"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   6600
      Width           =   630
   End
   Begin VB.Label TTLDB 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Top             =   6600
      Width           =   2895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KREDIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4800
      TabIndex        =   8
      Top             =   6600
      Width           =   705
   End
   Begin VB.Label TTLCR 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5640
      TabIndex        =   7
      Top             =   6600
      Width           =   2895
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keterangan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No.Transaksi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   1080
      Width           =   1140
   End
End
Attribute VB_Name = "TJurnal3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##############################################3
'BY : RONALD
'PT : RIDIA GROUP
'EDIT : 9 OKTOBER 2018
'##############################################3

Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_HIDEDROPDOWN = &H14
Private M_OldNoTrx As String
Dim NumEdit As Boolean
Private Sub CmbJenis_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub




Private Sub Combo1_Click()
On Error Resume Next
DataGrid1.Columns(13) = Combo1
Combo1.Visible = False
DataGrid1.SelText = DataGrid1.Columns(13)
DataGrid1.SelStart = 0
DataGrid1.SelLength = Len(DataGrid1.SelText)

End Sub

Private Sub Combo1_DropDown()
Combo1.Visible = False
End Sub

Private Sub Combo2_Click()
On Error Resume Next
DataGrid1.Columns(14) = Combo2
Combo2.Visible = False
DataGrid1.SelText = DataGrid1.Columns(14)
DataGrid1.SelStart = 0
DataGrid1.SelLength = Len(DataGrid1.SelText)
End Sub

Private Sub Combo2_DropDown()
Combo2.Visible = False
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
On Error Resume Next
If DataGrid1.Col = 0 Then
    Set RsData7 = Nothing
    Set RsData7 = New ADODB.Recordset
    RsData7.Open LCase("SELECT * FROM V_ACCOUNT_JURNAL WHERE STS=1 AND KODE_REK='" & DataGrid1.Text & "'"), con, adOpenDynamic, adLockReadOnly, adCmdText
    If RsData7.RecordCount > 0 Then
    VIEWJENISAKUN DataGrid1.Columns("KODE_REK").Text, "NONE"
        DataGrid1.Columns("KODE_REK").Text = RsData7.Fields("KODE_REK").Value
        DataGrid1.Columns("NAMA_REK").Text = RsData7.Fields("NAMA_REK").Value
        DataGrid1.Columns("SUBSIDIARY_REKENING").Text = Empty
    Else
    VIEWJENISAKUN "", "NONE"
        DataGrid1.Columns("KODE_REK").Text = Empty
        DataGrid1.Columns("NAMA_REK").Text = Empty
        DataGrid1.Columns("SUBSIDIARY_REKENING").Text = Empty
    End If
    RsData7.Close
    Set RsData7 = Nothing
End If
If ColIndex = 9 Or ColIndex = 11 Then DataGrid1.Text = FormatNumber(DataGrid1.Text, 2)
If ColIndex = 10 Or ColIndex = 12 Then DataGrid1.Text = FormatNumber(DataGrid1.Text, 3)

End Sub

Private Sub DataGrid1_AfterDelete()
GETHITUNG
GETHITUNG2

End Sub

Private Sub DataGrid1_AfterUpdate()
GETHITUNG
GETHITUNG2

End Sub

Private Sub DataGrid1_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
If ColIndex = 1 Or ColIndex = 2 Or ColIndex = 4 Or ColIndex = 6 Or ColIndex = 7 Then Cancel = True

End Sub

Private Sub DataGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
If ColIndex = 8 Then SetTmpData.Fields("KREDIT").Value = 0
If ColIndex = 9 Then SetTmpData.Fields("DEBET").Value = 0

End Sub

Private Sub DataGrid1_ButtonClick(ByVal ColIndex As Integer)
On Error GoTo errgrid
    Dim mydisp1 As DisplayData
    Set mydisp1 = New DisplayData
    mydisp1.DISPL_COSTR
If ColIndex = 0 Then
    mydisp1.Col_Name = "kode_rek"
    mydisp1.MyData "select * from V_ACCOUNT_JURNAL WHERE STS=1"
    mydisp1.Show 1
    If Not mydisp1.kode = Empty Then
        DataGrid1.Columns("KODE_REK").Value = mydisp1.kode
        DataGrid1.Columns("NAMA_REK").Value = mydisp1.CariString1("SELECT KODE_REK,NAMA_REK FROM V_ACCOUNT_JURNAL WHERE STS=1 AND KODE_REK='" & mydisp1.kode & "' ORDER BY KODE_REK ASC", "nama_rek")
        DataGrid1.Columns("SUBSIDIARY_REKENING").Text = Empty
        DataGrid1.Columns("NOMOR_BUKTI").Text = Empty
    End If
    Unload mydisp1
    Set mydisp1 = Nothing
End If
If ColIndex = 2 Then
If MsgBox("YES = HUTANG/PIUTANG, NO=UANG MUKA PEMBELIAN/PENJUALAN", vbYesNo) = vbYes Then
    
    mydisp1.Col_Name = "kode_rek"
    mydisp1.MyData "select * from V_NAMA WHERE AKUN_AR='" & DataGrid1.Columns("KODE_REK").Value & "'"
    mydisp1.Show 1
    If Not mydisp1.kode = Empty Then
        DataGrid1.Columns("SUBSIDIARY_REKENING").Value = mydisp1.kode
    End If
    Unload mydisp1
    Set mydisp1 = Nothing
Else
    mydisp1.Col_Name = "kode_rek"
    mydisp1.MyData "select * from V_LISTUM WHERE KODE_UM='" & DataGrid1.Columns("KODE_REK").Value & "'"
    mydisp1.Show 1
    If Not mydisp1.kode = Empty Then
        DataGrid1.Columns("SUBSIDIARY_REKENING").Value = mydisp1.kode
    End If
    Unload mydisp1
    Set mydisp1 = Nothing

End If
End If

If ColIndex = 4 Then
Dim clsakun_ap As ClsNama
Set clsakun_ap = New ClsNama
    Dim MyBrowRek As FrmBrowDispRek
'If MsgBox("YES = INVOICE, NO=PO", vbYesNo) = vbYes Then

    '***********************************************
    'SETTING AKUN HUTANG / PIUTANG
    clsakun_ap.NAMA_Construct
    clsakun_ap.CariKode "select * from nama where kode_rek='" & DataGrid1.Columns("SUBSIDIARY_REKENING").Value & "'", "AKUN_AR"
    clsakun_ap.CariJenisKode "select tipe from nama where kode_rek='" & DataGrid1.Columns("SUBSIDIARY_REKENING").Value & "'", "TIPE"
    Set MyBrowRek = New FrmBrowDispRek
        MyBrowRek.Col_Name = "notrans"
        MyBrowRek.MyData "select * from v_listinvoice where kode_rek='" & DataGrid1.Columns("subsidiary_rekening").Value & "' and kd_cab='" & txtcab.Text & "' and jenis='" & clsakun_ap.JENIS_REK & "'"
        MyBrowRek.Show 1
        If Not MyBrowRek.kode = Empty Then
        DataGrid1.Columns("NOMOR_BUKTI").Value = MyBrowRek.kode
        End If
        Unload MyBrowRek
        Set MyBrowRek = Nothing
'Else
    '***********************************************
    'SETTING AKUN HUTANG / PIUTANG
'    clsakun_ap.NAMA_Construct
'    clsakun_ap.CariKode "select * from v_listum where kode_rek='" & DataGrid1.Columns("SUBSIDIARY_REKENING").Value & "'", "KODE_UM"
'    clsakun_ap.CariJenisKode "select tipe from v_listum where kode_rek='" & DataGrid1.Columns("SUBSIDIARY_REKENING").Value & "'", "TIPE"
'    Set MyBrowRek = New FrmBrowDispRek
'        MyBrowRek.Col_Name = "notrans"
'        MyBrowRek.MyData "select * from v_listpo where kode_rek='" & DataGrid1.Columns("subsidiary_rekening").Value & "' and kd_cab='" & txtcab.Text & "' and jenis='" & clsakun_ap.JENIS_REK & "'"
'        MyBrowRek.Show 1
'        If Not MyBrowRek.kode = Empty Then
'        DataGrid1.Columns("NOMOR_BUKTI").Value = MyBrowRek.kode
'        End If
'        Unload MyBrowRek
'        Set MyBrowRek = Nothing
'End If
End If

errgrid:
If Not Err.Number = 0 Then
    MsgBox Err.Description, vbCritical, Err.Number
End If
End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    If DataGrid1.Col = 0 Then GetNama "SELECT KODE_REK,NAMA_REK FROM V_ACCOUNT_JURNAL WHERE STS=1 ORDER BY KODE_REK ASC", 107
    If DataGrid1.Col = 6 Then GetNama "SELECT KODE_REK,NAMA FROM NAMA WHERE TIPE='SALES' ORDER BY KODE_REK ASC", 142
End If

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0
If KeyAscii = 13 Then SendKeys vbTab


End Sub

Private Sub Form_Load()
CheckFormStatus Me
'Me.Picture = LoadPicture(App.Path & "\bcgk.wmf")
blankforms
SETOBJMENU False
End Sub

Function blankforms()
OldNoTrx = Empty
NumEdit = False
Notrans.Text = Empty
Notrans.Locked = True
Tgl.Value = Now
Keterangan.Text = "-"
txtcab.Text = Empty
lblcab.Caption = Empty
TTLDB.Caption = Empty
TTLCR.Caption = Empty
total.Caption = Empty
Combo2.Clear
Set rsKategori = Nothing
Set rsKategori = New ADODB.Recordset
rsKategori.CursorLocation = adUseClient
rsKategori.Open LCase("SELECT * FROM KATEGORI_BARANG"), con, adOpenDynamic, adLockReadOnly, adCmdText
Do Until rsKategori.EOF
    Combo2.AddItem rsKategori!kategori
    rsKategori.MoveNext
Loop
Combo2.ListIndex = 0

Toolbar1.Buttons(1).Enabled = Param_1
Toolbar1.Buttons(2).Enabled = Param_2
Toolbar1.Buttons(3).Enabled = Param_3
Toolbar1.Buttons(5).Enabled = Param_4
Toolbar1.Buttons(6).Enabled = True
GetTmpData3
Set DataGrid1.DataSource = SetTmpData
DataGrid1.Columns("KODE_REK").Button = True
DataGrid1.Columns("SUBSIDIARY_REKENING").Button = True
DataGrid1.Columns("NOMOR_BUKTI").Button = True
DataGrid1.Columns("KODE_SALES").Button = True
DataGrid1.Columns("DEBET").NumberFormat = "#,##0.00;(#,##0.00)"
DataGrid1.Columns("KREDIT").NumberFormat = "#,##0.00;(#,##0.00)"
DataGrid1.Columns("KODE_REK").Width = 1500
DataGrid1.Columns("KETERANGAN").Width = 2700
DataGrid1.Columns("SUBSIDIARY_REKENING").Width = 1000
DataGrid1.Columns("NOMOR_BUKTI").Width = 300

DataGrid1.Columns("NOMOR_NOTA").Visible = False
DataGrid1.Columns("NOMOR_NOTA").Locked = True

DataGrid1.Columns("KODE_SALES").Visible = False
DataGrid1.Columns("KODE_SALES").Locked = True

DataGrid1.Columns("NAMA_SALES").Visible = False
DataGrid1.Columns("NAMA_SALES").Locked = True

DataGrid1.Columns("DEBET").Alignment = dbgRight
DataGrid1.Columns("KREDIT").Alignment = dbgRight
End Function

Function SETOBJMENU(ByVal SETOBJ As Boolean)
Notrans.Enabled = SETOBJ
Notrans.Locked = True
Tgl.Enabled = SETOBJ
Keterangan.Enabled = SETOBJ
txtcab.Enabled = SETOBJ
DataGrid1.Enabled = SETOBJ
End Function

Private Sub Keterangan_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Notrans_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
If NumEdit = True And Param_5 = True Then
    If MsgBox("Apakah No.Bukti akan Diganti", vbYesNo) = vbYes Then
        OldNoTrx = Notrans.Text
        Notrans.Locked = False
    End If
End If
End If
End Sub

Private Sub Notrans_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Tgl_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys vbTab
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
 Case Is = 1
    blankforms
    Toolbar1.Buttons(2).Enabled = False
    Toolbar1.Buttons(5).Enabled = False
    Toolbar1.Buttons(6).Enabled = False
    SETOBJMENU True
    '--------------------------
    'OTOMATIS NOMOR
    'Notrans.Text = nobukti(Me.Name)
    '--------------------------
    '--------------------------
    'MANUAL NOMOR
    Notrans.Enabled = True
    Notrans.Locked = False
    '--------------------------
    Notrans.SetFocus
 Case Is = 2
       Set rsPostFlag = Nothing
    Set rsPostFlag = New ADODB.Recordset
    rsPostFlag.Open LCase("SELECT * FROM JURNAL_UMUM_HD WHERE NO_JURNAL='" & Notrans.Text & "' AND POST='1'"), con, adOpenDynamic, adLockOptimistic, adCmdText
            If rsPostFlag.RecordCount > 0 Then
                MsgBox "DATA SUDAH DIPOSTING", vbCritical
            Else
                NumEdit = True
                Toolbar1.Buttons(5).Enabled = False
                Toolbar1.Buttons(6).Enabled = False
                 SETOBJMENU True
'                 Notrans.Enabled = False
                 Keterangan.SetFocus
            End If
     Set rsPostFlag = Nothing
 Case Is = 3
               If NumEdit = True Then
                   If SetTmpData.RecordCount > 0 And Not Notrans.Text = Empty Then
                   SetTmpData.MoveFirst
                   If CDbl(total.Caption) = 0 Then
                    EditData
                   Else
                    MsgBox "Jurnal belum seimbang...!", vbCritical
                    DataGrid1.SetFocus
                   End If
                   
                   End If
               End If
               
               If NumEdit = False Then
                   If SetTmpData.RecordCount > 0 And Not Notrans.Text = Empty Then
                   SetTmpData.MoveFirst
                   If CDbl(total.Caption) = 0 Then
                       SimpanData
                   Else
                    MsgBox "Jurnal belum seimbang...!", vbCritical
                    DataGrid1.SetFocus
                   End If
                   Else
                       MsgBox "Transaksi tersebut sudah ada", vbCritical
                   End If
               End If
 Case Is = 4
    blankforms
    SETOBJMENU False
    NumEdit = True
    GetNama "SELECT NO_JURNAL,TGL_JURNAL,KETERANGAN,KD_CAB,NAMA_CAB FROM V_JURNAL_HD WHERE kd_cab in(select kd_cab from v_list_user_cab where userid='" & idxuser & "')" & " AND JENIS='JURNAL UMUM' AND POST=0 ORDER BY NO_JURNAL,TGL_JURNAL ASC", 25
Case Is = 5

If MsgBox("Apakah Data akan Dihapus", vbYesNo) = vbYes Then
    If SetTmpData.RecordCount > 0 Then
                HapusData
                SETOBJMENU False
                blankforms
                MsgBox "Data telah terhapus", vbInformation
                Set RsData1 = Nothing
                Set RsData2 = Nothing
    End If
End If

Case Is = 6
    Set rsCetakBukti = Nothing
    Set rsCetakBukti = New ADODB.Recordset
    rsCetakBukti.Open LCase("SELECT * FROM V_JU WHERE NO_JURNAL='" & Notrans.Text & "' AND TGL_JURNAL='" & Format(Tgl.Value, "YYYY-MM-dd") & "'"), con, adOpenDynamic, adLockOptimistic, adCmdText
'CreateFieldDefFile rsCetakBukti, App.Path & "\Voucher1.ttx", 1
    With CrystalReport1
        .ReportTitle = "JURNAL VOUCHER"
        .ReportFileName = App.Path & "\Reports\" & idxSqlDb & "\JurnalVoucher5.rpt"
        .Formulas(0) = "TxtNama='JURNAL UMUM'"
        .Destination = crptToWindow
        .SetTablePrivateData 0, 3, rsCetakBukti
        .WindowState = crptMaximized
       .Action = 1
    End With

Case Is = 7
    Unload Me

End Select

End Sub

Private Sub txtkodeakun_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub




Function SimpanData()
On Error GoTo errJur3
'update nomor
'Notrans.Text = nobukti(Me.Name)

    SetTmpData.MoveLast
    SetTmpData.MoveFirst
    Set RsData1 = Nothing
    Set RsData1 = New ADODB.Recordset
    RsData1.Open LCase("SELECT * FROM JURNAL_UMUM_HD WHERE JENIS='JURNAL UMUM' AND NO_JURNAL='" & Notrans.Text & "'"), con, adOpenDynamic, adLockOptimistic, adCmdText
    If RsData1.EOF = True Then
    con.BeginTrans
        RsData1.AddNew
        RsData1.Fields("NO_JURNAL").Value = Notrans.Text
        RsData1.Fields("TGL_JURNAL").Value = DateValue(Tgl.Value)
        RsData1.Fields("KETERANGAN").Value = Keterangan.Text
        RsData1.Fields("KD_CAB").Value = txtcab.Text
        RsData1.Fields("JENIS").Value = "JURNAL UMUM"
        RsData1!iduser = idxuser
        RsData1.Update
    
    Set RsData2 = Nothing
    Set RsData2 = New ADODB.Recordset
    RsData2.Open LCase("SELECT * FROM JURNAL_UMUM_DT WHERE NO_JURNAL='" & Notrans.Text & "'"), con, adOpenDynamic, adLockOptimistic, adCmdText
    Do Until SetTmpData.EOF
       RsData2.AddNew
        RsData2.Fields("NO_JURNAL").Value = Notrans.Text
        RsData2.Fields("TGL_JURNAL").Value = DateValue(Tgl.Value)
        RsData2.Fields("KODE").Value = SetTmpData.Fields("KODE_REK").Value
        RsData2.Fields("NAMA").Value = SetTmpData.Fields("NAMA_REK").Value
        RsData2.Fields("KETERANGAN").Value = SetTmpData.Fields("Keterangan").Value
        RsData2.Fields("NOMOR_BUKTI").Value = SetTmpData.Fields("Nomor_Bukti").Value
        RsData2.Fields("NOMOR_NOTA").Value = SetTmpData.Fields("Nomor_Nota").Value
        RsData2.Fields("KODE_SALES").Value = SetTmpData.Fields("KODE_SALES").Value
        RsData2.Fields("NAMA_SALES").Value = SetTmpData.Fields("NAMA_SALES").Value
        RsData2.Fields("JENIS").Value = "JURNAL UMUM"
        If Len(SetTmpData.Fields("DEBET").Value) > 0 Then RsData2.Fields("DEBET_UANG").Value = CDbl(FormatNumber(SetTmpData.Fields("DEBET").Value))
        If Len(SetTmpData.Fields("KREDIT").Value) > 0 Then RsData2.Fields("KREDIT_UANG").Value = CDbl(FormatNumber(SetTmpData.Fields("KREDIT").Value))
        RsData2.Fields("SUBSIDIARY_REKENING").Value = SetTmpData.Fields("SUBSIDIARY_REKENING").Value
        RsData2.Update
       SetTmpData.MoveNext
    Loop
   'UpdateNoBukti Me.Name
   Dim myTRX As ClsUserAkses
   Set myTRX = New ClsUserAkses
   myTRX.WriteMyIP Notrans.Text, "UPDATE"
    
    con.CommitTrans
    Toolbar1.Buttons(1).Enabled = Param_1
    Toolbar1.Buttons(2).Enabled = Param_2
    Toolbar1.Buttons(3).Enabled = Param_3
    Toolbar1.Buttons(5).Enabled = Param_4
    Toolbar1.Buttons(6).Enabled = True
    SETOBJMENU False
        MsgBox "Data telah tersimpan", vbInformation
        RsData1.Close
        RsData2.Close
        Set RsData1 = Nothing
        Set RsData2 = Nothing
        SendMessages Notrans.Text & " - " & DateValue(Tgl.Value), "Transc Completed"
    Else
        MsgBox "Transaksi tersebut sudah ada", vbCritical
    End If
errJur3:
If Not Err.Number = 0 Then
con.RollbackTrans
MsgBox Err.Number & "-" & Err.Description, vbCritical, Err.Number
End If

End Function


Function HapusData()
    con.BeginTrans
    Set RsData1 = Nothing
    Set RsData1 = New ADODB.Recordset
    RsData1.Open LCase("DELETE FROM JURNAL_UMUM_DT WHERE JENIS='JURNAL UMUM' AND NO_JURNAL='" & Notrans.Text & "'"), con, adOpenDynamic, adLockOptimistic, adCmdText
    Set RsData2 = Nothing
    Set RsData2 = New ADODB.Recordset
    RsData2.Open LCase("DELETE FROM JURNAL_UMUM_HD WHERE JENIS='JURNAL UMUM' AND NO_JURNAL='" & Notrans.Text & "'"), con, adOpenDynamic, adLockOptimistic, adCmdText
   Dim myTRX As ClsUserAkses
   Set myTRX = New ClsUserAkses
   myTRX.WriteMyIP Notrans.Text, "DELETE"
    con.CommitTrans
    SendMessages Notrans.Text & " - " & DateValue(Tgl.Value), "Void Transc"

End Function



Function EditData()
    'HAPUS TERLEBIH DAHULU
    con.BeginTrans
    Set RsData1 = Nothing
    Set RsData1 = New ADODB.Recordset

If NumEdit = True And Param_5 = True And Not OldNoTrx = Empty Then
    RsData1.Open LCase("DELETE FROM JURNAL_UMUM_DT WHERE JENIS='JURNAL UMUM' AND NO_JURNAL='" & OldNoTrx & "'"), con, adOpenDynamic, adLockOptimistic, adCmdText
    'Set RsData2 = Nothing
    'Set RsData2 = New ADODB.Recordset
    RsData2.Open LCase("DELETE FROM JURNAL_UMUM_HD WHERE JENIS='JURNAL UMUM' AND NO_JURNAL='" & OldNoTrx & "'"), con, adOpenDynamic, adLockOptimistic, adCmdText
Else
    RsData1.Open LCase("DELETE FROM JURNAL_UMUM_DT WHERE JENIS='JURNAL UMUM' AND NO_JURNAL='" & Notrans.Text & "'"), con, adOpenDynamic, adLockOptimistic, adCmdText
    'Set RsData2 = Nothing
    'Set RsData2 = New ADODB.Recordset
    'RsData2.Open LCase("DELETE FROM JURNAL_UMUM_HD WHERE JENIS='JURNAL UMUM' AND NO_JURNAL='" & Notrans.Text & "'"), con, adOpenDynamic, adLockOptimistic, adCmdText
End If
    SetTmpData.MoveLast
    SetTmpData.MoveFirst
    Set RsData1 = Nothing
    Set RsData1 = New ADODB.Recordset
    RsData1.Open LCase("SELECT * FROM JURNAL_UMUM_HD WHERE JENIS='JURNAL UMUM' AND NO_JURNAL='" & Notrans.Text & "'"), con, adOpenDynamic, adLockOptimistic, adCmdText
If NumEdit = True And Param_5 = True And Not OldNoTrx = Empty Then
    If RsData1.EOF = True Then
        RsData1.AddNew
        RsData1.Fields("NO_JURNAL").Value = Notrans.Text
        RsData1.Fields("TGL_JURNAL").Value = DateValue(Tgl.Value)
        RsData1.Fields("KETERANGAN").Value = Keterangan.Text
        RsData1.Fields("KD_CAB").Value = txtcab.Text
        RsData1.Fields("JENIS").Value = "JURNAL UMUM"
        RsData1!iduser = idxuser
        RsData1.Update
        End If
Else
    If RsData1.EOF = False Then
        'RsData1.AddNew
        RsData1.Fields("NO_JURNAL").Value = Notrans.Text
        RsData1.Fields("TGL_JURNAL").Value = DateValue(Tgl.Value)
        RsData1.Fields("KETERANGAN").Value = Keterangan.Text
        RsData1.Fields("KD_CAB").Value = txtcab.Text
        RsData1.Fields("JENIS").Value = "JURNAL UMUM"
        RsData1!iduser = idxuser
        RsData1.Update
        End If

End If
    Set RsData2 = Nothing
    Set RsData2 = New ADODB.Recordset
    RsData2.Open LCase("SELECT * FROM JURNAL_UMUM_DT WHERE NO_JURNAL='" & Notrans.Text & "'"), con, adOpenDynamic, adLockOptimistic, adCmdText
    Do Until SetTmpData.EOF
       RsData2.AddNew
        RsData2.Fields("NO_JURNAL").Value = Notrans.Text
        RsData2.Fields("TGL_JURNAL").Value = DateValue(Tgl.Value)
        RsData2.Fields("KODE").Value = SetTmpData.Fields("KODE_REK").Value
        RsData2.Fields("NAMA").Value = SetTmpData.Fields("NAMA_REK").Value
        RsData2.Fields("KETERANGAN").Value = SetTmpData.Fields("Keterangan").Value
        RsData2.Fields("NOMOR_BUKTI").Value = SetTmpData.Fields("Nomor_Bukti").Value
        RsData2.Fields("NOMOR_NOTA").Value = SetTmpData.Fields("Nomor_Nota").Value
        RsData2.Fields("KODE_SALES").Value = SetTmpData.Fields("KODE_SALES").Value
        RsData2.Fields("NAMA_SALES").Value = SetTmpData.Fields("NAMA_SALES").Value
        RsData2.Fields("JENIS").Value = "JURNAL UMUM"
        If Len(SetTmpData.Fields("DEBET").Value) > 0 Then RsData2.Fields("DEBET_UANG").Value = CDbl(FormatNumber(SetTmpData.Fields("DEBET").Value))
        If Len(SetTmpData.Fields("KREDIT").Value) > 0 Then RsData2.Fields("KREDIT_UANG").Value = CDbl(FormatNumber(SetTmpData.Fields("KREDIT").Value))
        RsData2.Fields("SUBSIDIARY_REKENING").Value = SetTmpData.Fields("SUBSIDIARY_REKENING").Value
        RsData2.Update
       SetTmpData.MoveNext
    Loop
   Dim myTRX As ClsUserAkses
   Set myTRX = New ClsUserAkses
   myTRX.WriteMyIP Notrans.Text, "UPDATE"
    
    con.CommitTrans
    Toolbar1.Buttons(1).Enabled = Param_1
    Toolbar1.Buttons(2).Enabled = Param_2
    Toolbar1.Buttons(3).Enabled = Param_3
    Toolbar1.Buttons(5).Enabled = Param_4
    Toolbar1.Buttons(6).Enabled = True
    SETOBJMENU False
        MsgBox "Data telah tersimpan", vbInformation
        RsData1.Close
        RsData2.Close
        Set RsData1 = Nothing
        Set RsData2 = Nothing
        SendMessages Notrans.Text & " - " & DateValue(Tgl.Value), "Transc Completed"
'    Else
'        MsgBox "Transaksi tersebut sudah ada", vbCritical
'    End If

End Function


 Function GETHITUNG()
 Dim getjml As ClsSummary
 Set getjml = New ClsSummary
 total.Caption = FormatNumber(getjml.LetHitung2(SetTmpData, "DEBET", "KREDIT"), 2)
 TTLDB.Caption = FormatNumber(getjml.IDdb1, 2)
 TTLCR.Caption = FormatNumber(getjml.IDcr1, 2)
 End Function


Function GETHITUNG2()
 Dim getjml As ClsSummary
 Set getjml = New ClsSummary
 End Function



Private Sub txtcab_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    GetNama "SELECT KD_CAB,NAMA_CAB FROM V_LIST_USER_CAB WHERE USERID='" & idxuser & "'", 66
End If

End Sub

Public Property Get OldNoTrx() As String
OldNoTrx = M_OldNoTrx
End Property

Public Property Let OldNoTrx(ByVal vNewValue As String)
M_OldNoTrx = vNewValue
End Property

Private Sub txtcab_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab

End Sub
