VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form TJurnal1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaksi Penerimaan Kas/Bank"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14355
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   14355
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      Height          =   1965
      Left            =   7800
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1440
      Width           =   5415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   255
      Left            =   1800
      TabIndex        =   26
      ToolTipText     =   "Impor Data Pengajuan"
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "VALIDASI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      MaskColor       =   &H00800000&
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7200
      Width           =   14415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   7
      ToolTipText     =   "F1 : List"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox TxtKdAkun 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "F1 : List"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Notrans 
      Height          =   285
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   2295
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
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "TJurnal1.frx":0000
      Left            =   9600
      List            =   "TJurnal1.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "TJurnal1.frx":0019
      Left            =   9600
      List            =   "TJurnal1.frx":0023
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtcab 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "F1 : List"
      Top             =   2400
      Width           =   1215
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
      Height          =   2775
      Left            =   360
      TabIndex        =   10
      Top             =   3600
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   4895
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
      Left            =   4560
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   183762947
      CurrentDate     =   40197
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   16
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
            Picture         =   "TJurnal1.frx":0032
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TJurnal1.frx":0A44
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TJurnal1.frx":1456
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TJurnal1.frx":1E68
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TJurnal1.frx":287A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TJurnal1.frx":328C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TJurnal1.frx":3C9E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Catatan"
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
      Left            =   7800
      TabIndex        =   27
      Top             =   1080
      Width           =   675
   End
   Begin VB.Label Label11 
      Caption         =   "Petugas"
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
      TabIndex        =   24
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label10 
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
      TabIndex        =   8
      Top             =   3120
      Width           =   4455
   End
   Begin VB.Label Label9 
      Caption         =   "Kas / Bank"
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
      TabIndex        =   23
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label7 
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
      TabIndex        =   6
      Top             =   2760
      Width           =   4455
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
      TabIndex        =   22
      Top             =   1080
      Width           =   1140
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Terima dari"
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
      TabIndex        =   21
      Top             =   1440
      Width           =   960
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
      Left            =   1080
      TabIndex        =   12
      Top             =   6600
      Width           =   2895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KREDIT"
      Enabled         =   0   'False
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
      TabIndex        =   20
      Top             =   6600
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label TTLDB 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
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
      Left            =   2040
      TabIndex        =   11
      Top             =   7080
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JUMLAH"
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
      TabIndex        =   19
      Top             =   6600
      Width           =   750
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SELISIH"
      Enabled         =   0   'False
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
      TabIndex        =   18
      Top             =   6600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label total 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
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
      TabIndex        =   13
      Top             =   6600
      Visible         =   0   'False
      Width           =   2895
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
      TabIndex        =   4
      Top             =   2400
      Width           =   4455
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
End
Attribute VB_Name = "TJurnal1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const CB_SHOWDROPDOWN = &H14F
Private Const CB_HIDEDROPDOWN = &H14

Dim NumEdit As Boolean
Dim myClsAju As New ClsJurnal

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

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab

End Sub

Private Sub Command1_Click()
Dim mypostdata As TPostingKeu
Dim myTrxKasBank As ADODB.Recordset
Dim myTrxKasBank2 As ADODB.Recordset
Set mypostdata = New TPostingKeu
    'inisialisasi
    mypostdata.TRXKB
    
    Set myTrxKasBank = Nothing
    Set myTrxKasBank = New ADODB.Recordset
    myTrxKasBank.Open LCase("SELECT * FROM V_KASBANK WHERE NO_JURNAL='" & Notrans.Text & "'"), con, adOpenDynamic, adLockReadOnly, adCmdText
    
    Set myTrxKasBank2 = Nothing
    Set myTrxKasBank2 = New ADODB.Recordset
    myTrxKasBank2.Open LCase("SELECT SUM(KREDIT_UANG)AS TTLKB,KETERANGAN,CATATAN,NO_JURNAL,TGL_JURNAL,KODE_AKUN,NAMA_AKUN,SUBSIDIARY_REKENING,NOMOR_BUKTI,KD_CAB FROM V_KASBANK WHERE NO_JURNAL='" & Notrans.Text & "'"), con, adOpenDynamic, adLockReadOnly, adCmdText
    If myTrxKasBank.RecordCount > 0 Then
        mypostdata.JENIS_TRX = "PENERIMAAN"
        mypostdata.KASBANK = TxtKdAkun.Text
        mypostdata.NO_JURNAL = myTrxKasBank2.Fields("NO_JURNAL").Value
        mypostdata.TGL_JURNAL = myTrxKasBank2.Fields("TGL_JURNAL").Value
        mypostdata.KODE_CABANG = myTrxKasBank2.Fields("KD_CAB").Value
        mypostdata.TTLKB = myTrxKasBank2.Fields("TTLKB").Value
        mypostdata.KETKB = Text2.Text 'myTrxKasBank2.Fields("KETERANGAN").Value & vbCrLf & myTrxKasBank2.Fields("CATATAN").Value
        mypostdata.FromRsTmp = myTrxKasBank
        mypostdata.GenerateRows
        mypostdata.Show 1
    End If
    If mypostdata.STSPOST = True Then
        blankforms
        SETOBJMENU False
    End If
    
End Sub

Private Sub Command2_Click()
    Dim mydisp1 As DisplayData
    Set mydisp1 = Nothing
    Set mydisp1 = New DisplayData
    mydisp1.Col_Name = "no_jurnal"
    mydisp1.MyData "select NO_JURNAL,TGL_JURNAL,KETERANGAN,KD_CAB,KODE,NAMA,KODE_ADM,NAMA_ADM from v_ajukasbank_hd3 WHERE kd_cab in(select kd_cab from v_list_user_cab where userid='" & idxuser & "')" & " AND JENIS LIKE '%PENERIMAAN%'"
    mydisp1.Show 1
    If Not mydisp1.kode = Empty Then
    blankforms
    setGrid
        Notrans.Text = mydisp1.kode
        myClsAju.DisplayRecJU ("select * from AJUKASBANK_HD where no_jurnal='" & mydisp1.kode & "'")
        Tgl.Value = myClsAju.Tgl_Order
        TxtKdAkun.Text = myClsAju.Kode_Nama
        Label7.Caption = myClsAju.nama
        Keterangan.Text = myClsAju.ket1
        txtcab.Text = myClsAju.KODE_CABANG
        lblcab.Caption = myClsAju.CariString1("select nama_cab from cabang where kd_cab='" & myClsAju.KODE_CABANG & "'", "nama_cab")
        Text1.Text = myClsAju.Kode_Petugas
        Label10.Caption = myClsAju.CariString1("select nama from nama where tipe='PETUGAS' AND kode_rek='" & myClsAju.Kode_Petugas & "'", "nama")
    
        If myClsAju.RsTmpPO.RecordCount > 0 Then
            myClsAju.RsTmpPO.MoveFirst
        Do Until myClsAju.RsTmpPO.EOF
            SetTmpData.AddNew
            SetTmpData.Fields("KODE_REK").Value = myClsAju.RsTmpPO.Fields("KODE").Value
            SetTmpData.Fields("NAMA_REK").Value = myClsAju.RsTmpPO.Fields("NAMA").Value
            SetTmpData.Fields("NOMOR_BUKTI").Value = myClsAju.RsTmpPO.Fields("NOMOR_BUKTI").Value
            SetTmpData.Fields("SUBSIDIARY_REKENING").Value = myClsAju.RsTmpPO.Fields("SUBSIDIARY_REKENING").Value
            SetTmpData.Fields("KETERANGAN").Value = myClsAju.RsTmpPO.Fields("KETERANGAN").Value
            SetTmpData.Fields("JUMLAH").Value = myClsAju.RsTmpPO.Fields("KREDIT_UANG").Value
            SetTmpData.Update
        myClsAju.RsTmpPO.MoveNext
        Loop
        End If
    
    
    End If
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
        DataGrid1.Columns("NOMOR_BUKTI").Text = Empty
        
    Else
    VIEWJENISAKUN "", "NONE"
        DataGrid1.Columns("KODE_REK").Text = Empty
        DataGrid1.Columns("NAMA_REK").Text = Empty
        DataGrid1.Columns("SUBSIDIARY_REKENING").Text = Empty
        DataGrid1.Columns("NOMOR_BUKTI").Text = Empty
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

If ColIndex = 3 Then
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
    If DataGrid1.Col = 0 Then DataGrid1_ButtonClick (DataGrid1.Col)
    If DataGrid1.Col = 2 Then DataGrid1_ButtonClick (DataGrid1.Col)
    If DataGrid1.Col = 3 Then DataGrid1_ButtonClick (DataGrid1.Col)
End If

End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0
If KeyAscii = 13 Then SendKeys vbTab


End Sub

Private Sub Form_Load()
CheckFormStatus Me
Set myClsAju = New ClsJurnal
'Me.Picture = LoadPicture(App.Path & "\bcgk.wmf")
blankforms
SETOBJMENU False
End Sub

Function blankforms()
NumEdit = False
myClsAju.PO_Construct
Notrans.Text = Empty
Notrans.Locked = False
Tgl.Value = Now
Keterangan.Text = "-"
Text2.Text = "-"
txtcab.Text = Empty
TxtKdAkun.Text = Empty
Label7.Caption = Empty
lblcab.Caption = Empty
TTLDB.Caption = Empty
TTLCR.Caption = Empty
total.Caption = Empty
Text1.Text = Empty
Label10.Caption = Empty
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
Command1.Enabled = Param_6
setGrid
End Function

Function setGrid()
GetTmpData3KB1
Set DataGrid1.DataSource = SetTmpData
DataGrid1.Columns("KODE_REK").Button = True
DataGrid1.Columns("SUBSIDIARY_REKENING").Button = True
DataGrid1.Columns("NOMOR_BUKTI").Button = True
DataGrid1.Columns("KODE_REK").Width = 1500
DataGrid1.Columns("NAMA_REK").Width = 2000
DataGrid1.Columns("NAMA_REK").Locked = True
DataGrid1.Columns("SUBSIDIARY_REKENING").Width = 2000
DataGrid1.Columns("SUBSIDIARY_REKENING").Locked = True
DataGrid1.Columns("NOMOR_BUKTI").Width = 2000
DataGrid1.Columns("NOMOR_BUKTI").Locked = True
DataGrid1.Columns("JUMLAH").Width = 2500
DataGrid1.Columns("keterangan").Width = 3000
DataGrid1.Columns("jumlah").Alignment = dbgRight
DataGrid1.Columns("JUMLAH").NumberFormat = "#,##0.00;(#,##0.00)"

End Function
Function SETOBJMENU(ByVal SETOBJ As Boolean)
Command2.Enabled = False
Notrans.Enabled = SETOBJ
Tgl.Enabled = SETOBJ
Keterangan.Enabled = SETOBJ
Text2.Enabled = SETOBJ
txtcab.Enabled = SETOBJ
TxtKdAkun.Enabled = SETOBJ
DataGrid1.Enabled = SETOBJ
Text1.Enabled = SETOBJ
End Function

Private Sub Keterangan_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0

End Sub

Private Sub Notrans_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim myClsPO1 As New ClsInvoicing1
Set myClsPO1 = Nothing
Set myClsPO1 = New ClsInvoicing1
If KeyCode = vbKeyF1 Then
    Dim mydisp1 As DisplayData
    Set mydisp1 = New DisplayData
    mydisp1.Col_Name = "kode_rek"
    mydisp1.MyData "SELECT * FROM V_NAMA WHERE TIPE='PETUGAS'"
    mydisp1.Show 1
    If Not mydisp1.kode = Empty Then
        Text1.Text = mydisp1.kode
        Label10.Caption = myClsPO1.CariString1("select nama from v_nama where kode_rek='" & mydisp1.kode & "'", "nama")
    End If
    Unload mydisp1
    Set mydisp1 = Nothing
End If
Set myClsPO1 = Nothing
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0
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
    Command2.Enabled = True
    'Notrans.Text = nobukti(Me.Name)
    Notrans.SetFocus
 Case Is = 2
    Set rsPostFlag = Nothing
    Set rsPostFlag = New ADODB.Recordset
    rsPostFlag.Open LCase("SELECT * FROM KASBANK_HD WHERE NO_JURNAL='" & Notrans.Text & "' AND POST='1'"), con, adOpenDynamic, adLockOptimistic, adCmdText
            If rsPostFlag.RecordCount > 0 And Param_6 = False Then
                MsgBox "DATA TIDAK BISA DIRUBAH," & vbCrLf & "SUDAH DI VALIDASI OLEH ACCOUNTING", vbCritical
            Else
                NumEdit = True
                Toolbar1.Buttons(5).Enabled = False
                Toolbar1.Buttons(6).Enabled = False
                 SETOBJMENU True
                 Notrans.Enabled = False
                 'Tgl.Enabled = False
                 Keterangan.SetFocus
            End If
 Case Is = 3
               If NumEdit = True Then
    Set rsPostFlag = Nothing
    Set rsPostFlag = New ADODB.Recordset
    rsPostFlag.Open LCase("SELECT * FROM KASBANK_HD WHERE NO_JURNAL='" & Notrans.Text & "' AND POST='1'"), con, adOpenDynamic, adLockOptimistic, adCmdText
            If rsPostFlag.RecordCount > 0 And Param_6 = False Then
                MsgBox "DATA TIDAK BISA DIRUBAH," & vbCrLf & "SUDAH DI VALIDASI OLEH ACCOUNTING", vbCritical
            Else
                   If SetTmpData.RecordCount > 0 And Not Notrans.Text = Empty And Not txtcab.Text = Empty And Not TxtKdAkun.Text = Empty And Not Text1.Text = Empty Then
                   SetTmpData.MoveFirst
                   If CDbl(total.Caption) = 0 Then
                    EditData
                   Else
                    MsgBox "Jurnal belum seimbang...!", vbCritical
                    DataGrid1.SetFocus
                   End If
            End If
                   End If
               
               End If
               
               If NumEdit = False Then
                   If SetTmpData.RecordCount > 0 And Not Notrans.Text = Empty And Not txtcab.Text = Empty And Not TxtKdAkun.Text = Empty And Not Text1.Text = Empty Then
                   SetTmpData.MoveFirst
                   If CDbl(total.Caption) = 0 Then
                       SimpanData
                   Else
                    MsgBox "Jurnal belum seimbang...!", vbCritical
                    DataGrid1.SetFocus
                   End If
                   Else
                       MsgBox "DATA BELUM LENGKAP", vbCritical
                   End If
               End If
 Case Is = 4
    blankforms
    SETOBJMENU False
    NumEdit = True
    GetNama "SELECT NO_JURNAL,TGL_JURNAL,NO_AJU,KODE,NAMA,KODE_ADM,NAMA_ADM,KETERANGAN,CATATAN,KD_CAB,NAMA_CAB,POST FROM V_KASBANK_HD WHERE (JENIS='PENERIMAAN KAS' OR JENIS='PENERIMAAN BANK') AND kd_cab in(select kd_cab from v_list_user_cab where userid='" & idxuser & "')" & " ORDER BY NO_JURNAL,TGL_JURNAL ASC", 170
Case Is = 5

    Set rsPostFlag = Nothing
    Set rsPostFlag = New ADODB.Recordset
    rsPostFlag.Open LCase("SELECT * FROM KASBANK_HD WHERE NO_JURNAL='" & Notrans.Text & "' AND POST='1'"), con, adOpenDynamic, adLockOptimistic, adCmdText
            If rsPostFlag.RecordCount > 0 And Param_6 = False Then
                MsgBox "DATA TIDAK BISA DIHAPUS," & vbCrLf & "SUDAH DI VALIDASI OLEH ACCOUNTING", vbCritical
            Else

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
End If
Case Is = 6

'***********************************************
Dim clsakun_ap As ClsNama
Set clsakun_ap = New ClsNama
'pilih kas/bank
clsakun_ap.NAMA_Construct
clsakun_ap.CariKode "select jenis from v_account where kode_rek1='" & TxtKdAkun.Text & "'", "JENIS"
'***********************************************

    Set rsCetakBukti = Nothing
    Set rsCetakBukti = New ADODB.Recordset
    rsCetakBukti.Open LCase("SELECT * FROM V_KASBANK WHERE JENIS='PENERIMAAN " & UCase(clsakun_ap.Kode_rek) & "' AND NO_JURNAL='" & Notrans.Text & "' AND TGL_JURNAL='" & Format(Tgl.Value, "YYYY-MM-dd") & "'"), con, adOpenDynamic, adLockOptimistic, adCmdText
'CreateFieldDefFile rsCetakBukti, App.Path & "\ttkasbank.ttx", 1
    With CrystalReport1
        .ReportTitle = "JURNAL VOUCHER"
        .ReportFileName = App.Path & "\Reports\" & idxSqlDb & "\TTKASBANK.rpt"
        If rsCetakBukti.RecordCount > 0 Then .Formulas(1) = "TXTSAYS='" & UCase(Terbilang(CDbl(FormatNumber(TTLCR.Caption)))) & "'"
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

'***********************************************
Dim clsakun_ap As ClsNama
Set clsakun_ap = New ClsNama
'pilih kas/bank
clsakun_ap.NAMA_Construct
clsakun_ap.CariKode "select jenis from v_account where kode_rek1='" & TxtKdAkun.Text & "'", "JENIS"
'***********************************************
    
    SetTmpData.MoveLast
    SetTmpData.MoveFirst
    Set RsData1 = Nothing
    Set RsData1 = New ADODB.Recordset
    RsData1.Open LCase("SELECT * FROM KASBANK_HD WHERE JENIS='PENERIMAAN " & UCase(clsakun_ap.Kode_rek) & "' AND NO_JURNAL='" & Notrans.Text & "'"), con, adOpenDynamic, adLockOptimistic, adCmdText
    If RsData1.EOF = True Then
    con.BeginTrans
        RsData1.AddNew
        RsData1.Fields("NO_JURNAL").Value = Notrans.Text
        RsData1.Fields("NO_AJU").Value = myClsAju.No_RO
        RsData1.Fields("TGL_JURNAL").Value = DateValue(Tgl.Value)
        RsData1.Fields("KETERANGAN").Value = Keterangan.Text
        RsData1.Fields("CATATAN").Value = Text2.Text
        RsData1.Fields("KD_CAB").Value = txtcab.Text
        RsData1.Fields("JENIS").Value = "PENERIMAAN " & UCase(clsakun_ap.Kode_rek)
        RsData1!iduser = idxuser
        RsData1.Fields("KODE").Value = TxtKdAkun.Text
        RsData1.Fields("NAMA").Value = Label7.Caption
        RsData1.Fields("KODE_ADM").Value = Text1.Text
        RsData1.Fields("NAMA_ADM").Value = Label10.Caption
        RsData1.Update
    
    Set RsData2 = Nothing
    Set RsData2 = New ADODB.Recordset
    RsData2.Open LCase("SELECT * FROM KASBANK_DT WHERE NO_JURNAL='" & Notrans.Text & "'"), con, adOpenDynamic, adLockOptimistic, adCmdText
    Do Until SetTmpData.EOF
       RsData2.AddNew
        RsData2.Fields("NO_JURNAL").Value = Notrans.Text
        RsData2.Fields("TGL_JURNAL").Value = DateValue(Tgl.Value)
        RsData2.Fields("KODE").Value = SetTmpData.Fields("KODE_REK").Value
        RsData2.Fields("NAMA").Value = SetTmpData.Fields("NAMA_REK").Value
        RsData2.Fields("NOMOR_BUKTI").Value = SetTmpData.Fields("Nomor_Bukti").Value
        RsData2.Fields("SUBSIDIARY_REKENING").Value = SetTmpData.Fields("SUBSIDIARY_REKENING").Value
        RsData2.Fields("KETERANGAN").Value = SetTmpData.Fields("Keterangan").Value
        RsData2.Fields("JENIS").Value = "PENERIMAAN " & UCase(clsakun_ap.Kode_rek)
        If Len(SetTmpData.Fields("JUMLAH").Value) > 0 Then RsData2.Fields("KREDIT_UANG").Value = CDbl(FormatNumber(SetTmpData.Fields("JUMLAH").Value))
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
    
'***********************************************
Dim clsakun_ap As ClsNama
Set clsakun_ap = New ClsNama
'pilih kas/bank
clsakun_ap.NAMA_Construct
clsakun_ap.CariKode "select jenis from v_account where kode_rek1='" & TxtKdAkun.Text & "'", "JENIS"
'***********************************************
    
    Set RsData1 = Nothing
    Set RsData1 = New ADODB.Recordset
    RsData1.Open LCase("DELETE FROM KASBANK_DT WHERE NO_JURNAL='" & Notrans.Text & "'"), con, adOpenDynamic, adLockOptimistic, adCmdText
    Set RsData2 = Nothing
    Set RsData2 = New ADODB.Recordset
    RsData2.Open LCase("DELETE FROM KASBANK_HD WHERE NO_JURNAL='" & Notrans.Text & "'"), con, adOpenDynamic, adLockOptimistic, adCmdText
   Dim myTRX As ClsUserAkses
   Set myTRX = New ClsUserAkses
   myTRX.WriteMyIP Notrans.Text, "DELETE"
    con.CommitTrans
    SendMessages Notrans.Text & " - " & DateValue(Tgl.Value), "Void Transc"

End Function



Function EditData()
    'HAPUS TERLEBIH DAHULU
    con.BeginTrans
    
'***********************************************
Dim clsakun_ap As ClsNama
Set clsakun_ap = New ClsNama
'pilih kas/bank
clsakun_ap.NAMA_Construct
clsakun_ap.CariKode "select jenis from v_account where kode_rek1='" & TxtKdAkun.Text & "'", "JENIS"
'***********************************************
    
    Set RsData1 = Nothing
    Set RsData1 = New ADODB.Recordset
    RsData1.Open LCase("DELETE FROM KASBANK_DT WHERE NO_JURNAL='" & Notrans.Text & "'"), con, adOpenDynamic, adLockOptimistic, adCmdText
    
'    Set RsData2 = Nothing
'    Set RsData2 = New ADODB.Recordset
'    RsData2.Open LCase("DELETE FROM KASBANK_HD WHERE NO_JURNAL='" & Notrans.Text & "'"), con, adOpenDynamic, adLockOptimistic, adCmdText
    
    SetTmpData.MoveLast
    SetTmpData.MoveFirst
    Set RsData1 = Nothing
    Set RsData1 = New ADODB.Recordset
    RsData1.Open LCase("SELECT * FROM KASBANK_HD WHERE NO_JURNAL='" & Notrans.Text & "'"), con, adOpenDynamic, adLockOptimistic, adCmdText
    
    If RsData1.EOF = False Then
        'RsData1.AddNew
        'RsData1.Fields("NO_JURNAL").Value = Notrans.Text
        RsData1.Fields("TGL_JURNAL").Value = DateValue(Tgl.Value)
        RsData1.Fields("KETERANGAN").Value = Keterangan.Text
        RsData1.Fields("CATATAN").Value = Text2.Text
        RsData1.Fields("KD_CAB").Value = txtcab.Text
        RsData1.Fields("JENIS").Value = "PENERIMAAN " & UCase(clsakun_ap.Kode_rek)
        RsData1!iduser = idxuser
        RsData1.Fields("KODE").Value = TxtKdAkun.Text
        RsData1.Fields("NAMA").Value = Label7.Caption
        RsData1.Fields("KODE_ADM").Value = Text1.Text
        RsData1.Fields("NAMA_ADM").Value = Label10.Caption
        RsData1.Fields("POST").Value = 0
        RsData1.Update
    
    Set RsData2 = Nothing
    Set RsData2 = New ADODB.Recordset
    RsData2.Open LCase("SELECT * FROM KASBANK_DT WHERE NO_JURNAL='" & Notrans.Text & "'"), con, adOpenDynamic, adLockOptimistic, adCmdText
    Do Until SetTmpData.EOF
       RsData2.AddNew
        RsData2.Fields("NO_JURNAL").Value = Notrans.Text
        RsData2.Fields("TGL_JURNAL").Value = DateValue(Tgl.Value)
        RsData2.Fields("KODE").Value = SetTmpData.Fields("KODE_REK").Value
        RsData2.Fields("NAMA").Value = SetTmpData.Fields("NAMA_REK").Value
        RsData2.Fields("NOMOR_BUKTI").Value = SetTmpData.Fields("Nomor_Bukti").Value
        RsData2.Fields("SUBSIDIARY_REKENING").Value = SetTmpData.Fields("SUBSIDIARY_REKENING").Value
        RsData2.Fields("KETERANGAN").Value = SetTmpData.Fields("Keterangan").Value
        RsData2.Fields("JENIS").Value = "PENERIMAAN " & UCase(clsakun_ap.Kode_rek)
        If Len(SetTmpData.Fields("JUMLAH").Value) > 0 Then RsData2.Fields("KREDIT_UANG").Value = SetTmpData.Fields("JUMLAH").Value
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
    End If

End Function


 Function GETHITUNG()
 Dim getjml As ClsSummary
 Set getjml = New ClsSummary
 total.Caption = FormatNumber(getjml.LetHitung2(SetTmpData, "JUMLAH", "JUMLAH"), 2)
 TTLDB.Caption = FormatNumber(getjml.IDdb1, 2)
 TTLCR.Caption = FormatNumber(getjml.IDcr1, 2)
 End Function


Function GETHITUNG2()
 Dim getjml As ClsSummary
 Set getjml = New ClsSummary
 End Function



Private Sub txtcab_KeyDown(KeyCode As Integer, Shift As Integer)
Dim myClsPO1 As New ClsInvoicing1
Dim mydisp1 As DisplayData
Set myClsPO1 = Nothing
Set myClsPO1 = New ClsInvoicing1
If KeyCode = vbKeyF1 Then
    If SetTmpData.RecordCount > 0 Then
        If MsgBox("Detil Data akan Dihapus,bila ingin mengganti Kode Cabang", vbYesNo) = vbYes Then
            setGrid
            Set mydisp1 = New DisplayData
            mydisp1.Col_Name = "kd_cab"
            mydisp1.MyData "SELECT KD_CAB,NAMA_CAB FROM V_LIST_USER_CAB WHERE USERID='" & idxuser & "'"
            mydisp1.Show 1
            If Not mydisp1.kode = Empty Then
                txtcab.Text = mydisp1.kode
                lblcab.Caption = myClsPO1.CariString1("select nama_cab from cabang where kd_cab='" & mydisp1.kode & "'", "nama_cab")
            End If
            Unload mydisp1
            Set mydisp1 = Nothing
        End If
    Else
        Set mydisp1 = New DisplayData
        mydisp1.Col_Name = "kd_cab"
        mydisp1.MyData "SELECT KD_CAB,NAMA_CAB FROM V_LIST_USER_CAB WHERE USERID='" & idxuser & "'"
        mydisp1.Show 1
        If Not mydisp1.kode = Empty Then
            txtcab.Text = mydisp1.kode
            lblcab.Caption = myClsPO1.CariString1("select nama_cab from cabang where kd_cab='" & mydisp1.kode & "'", "nama_cab")
        End If
        Unload mydisp1
        Set mydisp1 = Nothing
    End If
End If
Set myClsPO1 = Nothing
End Sub

Private Sub txtcab_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub TxtKdAkun_KeyDown(KeyCode As Integer, Shift As Integer)
Dim myClsPO1 As New ClsInvoicing1
Set myClsPO1 = Nothing
Set myClsPO1 = New ClsInvoicing1
If KeyCode = vbKeyF1 Then
    Dim mydisp1 As DisplayData
    Set mydisp1 = New DisplayData
    mydisp1.Col_Name = "kode_rek1"
    mydisp1.MyData "SELECT KODE_REK1,NAMA_REK,JENIS FROM V_ACCOUNT WHERE STS=1 AND (JENIS='KAS' OR JENIS='BANK')"
    mydisp1.Show 1
    If Not mydisp1.kode = Empty Then
        TxtKdAkun.Text = mydisp1.kode
        Label7.Caption = myClsPO1.CariString1("select nama_rek from v_account where kode_rek1='" & mydisp1.kode & "'", "nama_rek")
    End If
    Unload mydisp1
    Set mydisp1 = Nothing
End If
Set myClsPO1 = Nothing

End Sub

Private Sub TxtKdAkun_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys vbTab

End Sub

