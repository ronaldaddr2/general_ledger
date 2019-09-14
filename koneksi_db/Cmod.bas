Attribute VB_Name = "Cmod"
    Option Explicit
Declare Function CreateFieldDefFile Lib "p2smon.dll" (lpUnk As Object, ByVal filename As String, ByVal bOverWriteExistingFile As Long) As Long

Public con As New ADODB.Connection
Public conCHAT As New ADODB.Connection

Public idxuser As String
Public idxpwd As String
Public idxservername As String
Public idxSqlDb As String
Public RsUser As New ADODB.Recordset
Public RsUserLevel As New ADODB.Recordset
Public RsList As New ADODB.Recordset
Public RsPeriodeAktif As New ADODB.Recordset

Public RightsMn As Boolean
Public setstatus As String
Public Param_1 As Boolean
Public Param_2 As Boolean
Public Param_3 As Boolean
Public Param_4 As Boolean
Public Param_5 As Boolean
Public Param_6 As Boolean
Public MYPORTDB As String





Function SendLogin(USER As String, Sandi As String, Server As String, SqlDB As String)
On Error GoTo errCon
Set con = Nothing
Set con = New ADODB.Connection
Dim xUsr, xUsrPwd As String
'XXXXXXXXXXXXXXXXXXXXXXXXXXX
' MYSQL USER
xUsr = "usrsys"
xUsrPwd = "usrsys123"
'XXXXXXXXXXXXXXXXXXXXXXXXXXX
con.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & LCase(Server) & ";PORT=" & "49000" & ";UID=" & LCase(USER) & ";PWD=" & LCase(Sandi) & ";DATABASE=" & LCase(SqlDB) & ";Socket=MySQL;Option=3;"
'con.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & Server & ";PORT=3306;UID=" & LCase(xUsr) & ";PWD=" & LCase(xUsrPwd) & ";DATABASE=" & SqlDB & ";Socket=MySQL;Option=3;"
'con.ConnectionString = "DRIVER={MySQL ODBC 5.3 ANSI Driver};SERVER=" & Server & ";PORT=3306;UID=" & LCase(USER) & ";PWD=" & LCase(Sandi) & ";DATABASE=" & SqlDB & ";Option=3;"

con.CursorLocation = adUseClient
idxuser = USER
idxpwd = Sandi
idxservername = Server
idxSqlDb = SqlDB
con.Open
Set RsUser = Nothing
Set RsUser = New ADODB.Recordset
RsUser.CursorLocation = adUseClient
RsUser.Open LCase("SELECT * FROM USER WHERE USER='" & USER & "' AND SANDI='" & Sandi & "'"), con, adOpenDynamic, adLockReadOnly
If Not RsUser.EOF Then
  With MenuUtama
     .StatusBar1.Panels(1).Text = "USER :  " & idxuser
     .StatusBar1.Panels(3).Text = "SERVER :  " & idxservername
     .StatusBar1.Panels(2).Text = "Bulan Aktif : " & MonthName(Mid(PeriodeAktif, 1, 2)) & "-" & (Mid(PeriodeAktif, 3, 4))
     .StatusBar1.Panels(4).Text = "DATABASE :  " & UCase(idxSqlDb)
     .MnMasuk.Enabled = False
     .MnKeluar.Enabled = True
     .Caption = UCase(idxSqlDb)
     .ConnectData
    'ID HOSTNAME MYSQL
     Dim myIPADDR As ClsUserAkses
     Set myIPADDR = New ClsUserAkses
     myIPADDR.ReadMyIP
     myIPADDR.SetTGLServer
     myIPADDR.WriteMyIP idxuser & "," & myIPADDR.TGL_SERVER, "LOGIN"
    .StatusBar1.Panels(5).Text = "IP Address : " & myIPADDR.myIP
    .MnGantiPwd.Visible = True
    .MnGantiPwd.Enabled = True
    .Toolbar1.Visible = True
  End With
    With FrmLogin
        Unload FrmLogin
    End With
    'FrmHome.Show
    'FrmHome.Left = 0
    'FrmHome.Width = 300
    

Else
  MsgBox "Akses Ditolak"
End If

errCon:
If Not Err.Number = 0 Then
    'Exit Function
    '& vbCrLf
    MsgBox Err.Description, vbCritical, Err.Number
    'MsgBox "Try Again,Databases not Respond", vbCritical, Err.Number
End If
End Function

Function StatusUser(TMENU As String)
On Error GoTo errmenu
If con.State = adStateOpen Then
RightsMn = False
Set RsUserLevel = Nothing
Set RsUserLevel = New ADODB.Recordset
RsUserLevel.Open LCase("SELECT * FROM VIEWUSERDATA1 WHERE USER='" & idxuser & "' AND FORMNAME='" & TMENU & "'"), con, adOpenDynamic, adLockReadOnly
If Not RsUserLevel.EOF Then
   RightsMn = True
End If
End If
errmenu:
If Not Err.Number = 0 Then
MsgBox Err.Description, vbCritical, Err.Number
End If
End Function


    Public Sub CheckFormStatus(Myform As Form)
    Dim objForm As Form
    Dim FlgLoaded As Boolean
    Dim FlgShown As Boolean
    FlgLoaded = False
    FlgShown = False
    For Each objForm In VB.Forms
    If (Trim(objForm.Name) = Trim(Myform.Name)) Then
    FlgLoaded = True
    If objForm.Visible Then
    FlgShown = True
    End If
    Exit For
    End If
    Next
    'MsgBox "Load Status: " & FlgLoaded & vbCrLf & "Show Status:" & FlgShown
      Param_1 = RsUserLevel.Fields("databaru").Value
      Param_2 = RsUserLevel.Fields("rubah").Value
      Param_3 = RsUserLevel.Fields("simpan").Value
      Param_4 = RsUserLevel.Fields("hapus").Value
      Param_5 = RsUserLevel.Fields("notrx").Value
      Param_6 = RsUserLevel.Fields("postdata").Value
    End Sub







Public Function SendMessages(Optional ByVal strmessages As String, Optional ByVal status As String)
With MenuUtama
If .Winsock1.State = 7 Then
    .StatusBar1.Panels(4).Text = Empty
    .Winsock1.SendData strmessages & "<" & status & ">" & "User:" & idxuser & " //Date: " & Now
    .StatusBar1.Panels(4).Text = "NO.TRANS :> " & strmessages & "<" & status & ">" & "User:" & idxuser
    '.lstmessage.AddItem "NO.TRANSAKSI :> " & strmessages
    strmessages = Empty
Else
    .StatusBar1.Panels(4).Text = .Winsock1.LocalIP & " :> Not currently connected to the server"
    'lbltransmissionstatus.Caption = "Incomplete Data Transmission"
End If
End With
End Function




Function PeriodeAktif() As String
Set RsPeriodeAktif = Nothing
Set RsPeriodeAktif = New ADODB.Recordset
RsPeriodeAktif.Open LCase("SELECT * FROM BULANAKTIF"), con, adOpenDynamic, adLockReadOnly
PeriodeAktif = RsPeriodeAktif.Fields("AKTIF").Value
End Function

Function PeriodeAktif1() As Date
Set RsPeriodeAktif = Nothing
Set RsPeriodeAktif = New ADODB.Recordset
RsPeriodeAktif.Open LCase("SELECT TGL_SALDO FROM BULANAKTIF"), con, adOpenDynamic, adLockReadOnly
PeriodeAktif1 = RsPeriodeAktif.Fields("TGL_SALDO").Value
End Function



Function tagihan()
Dim hutangAwal_gr As Currency
Dim jual_gr As Currency
Dim retur_gr As Currency
Dim retur_gr1 As Currency
Dim TitipanAwal_Rp As Currency
Dim Titipanptg_Rp As Currency
Dim putus_gr As Currency
Dim putus_gr1 As Currency
Dim putus_gr2 As Currency
Dim putus_gr3 As Currency
Dim putus_RP As Currency
Dim putus_RP1 As Currency
Dim bayar_RP As Currency
Dim bayar_RP1 As Currency
Dim col1 As Currency
Dim col1_1 As Currency
Dim col2 As Currency
Dim col2_1 As Currency
Dim col3 As Currency
Dim col3_1 As Currency
Dim col4 As Currency
Dim col4_1 As Currency
'NOTA PENJUALAN
'Set DataGrid1.DataSource = Nothing
TempDataSisaRp521
Set RsData1 = Nothing
Set RsData1 = New ADODB.Recordset
RsData1.CursorLocation = adUseClient
'ProgressBar1.Value = 0

'ProgressBar1.Max = RsData1.RecordCount
DoEvents
hutangAwal_gr = Val(0)
jual_gr = Val(0)
retur_gr = Val(0)
retur_gr1 = Val(0)
TitipanAwal_Rp = Val(0)
Titipanptg_Rp = Val(0)
putus_gr = Val(0)
putus_gr1 = Val(0)
putus_gr2 = Val(0)
putus_gr3 = Val(0)
putus_RP = Val(0)
putus_RP1 = Val(0)
bayar_RP = Val(0)
bayar_RP1 = Val(0)
col1 = Val(0)
col2 = Val(0)
col3 = Val(0)
col4 = Val(0)
col1_1 = Val(0)
col2_1 = Val(0)
col3_1 = Val(0)
col4_1 = Val(0)
'    ProgressBar1.Value = ProgressBar1.Value + 1
'======================
'insert data transaksi
'======================


'------------------
'SALDO AWAL PIUTANG
'------------------
 Set rsUP2 = Nothing
 Set rsUP2 = New ADODB.Recordset
 rsUP2.CursorLocation = adUseClient
 'rsUP2.Open "SELECT NOINV,TGLINV,IDSALES,IDSUPPLIER,JUMLAH FROM SALDOPIUTANG_HD,SALDOPIUTANG_DT WHERE SALDOPIUTANG_DT.STATUS='BELUM LUNAS' AND SALDOPIUTANG_HD.NOTRANS=SALDOPIUTANG_DT.NOTRANS AND SALDOPIUTANG_DT.TGL=SALDOPIUTANG_DT.TGL AND IDSUPPLIER='" & RsData6.Fields("NAMA").Value & "' AND IDSALES='" & RsData1.Fields("NAMA").Value & "'", con, 2, 1
 rsUP2.Open LCase("SELECT NOINV,TGLINV,IDSALES,IDSUPPLIER,STATUS,JUMLAH FROM SALDOPIUTANG_HD,SALDOPIUTANG_DT WHERE SALDOPIUTANG_DT.STATUS='BELUM LUNAS' AND SALDOPIUTANG_HD.NOTRANS=SALDOPIUTANG_DT.NOTRANS AND SALDOPIUTANG_DT.TGL=SALDOPIUTANG_DT.TGL"), con, 2, 1
 Do Until rsUP2.EOF
    hutangAwal_gr = rsUP2.Fields("jumlah").Value
    putus_gr = Val(0)
    putus_RP = Val(0)
    col1 = Val(0)
    col2 = Val(0)
    col3 = Val(0)
    col4 = Val(0)
    bayar_RP = Val(0)
        Set rsUP4 = Nothing
        Set rsUP4 = New ADODB.Recordset
        rsUP4.CursorLocation = adUseClient
        rsUP4.Open LCase("SELECT * FROM BAYARPIUTANG WHERE IDSUPPLIER='" & rsUP2.Fields("IDSUPPLIER").Value & "' AND IDSALES='" & rsUP2.Fields("IDSALES").Value & "' AND NOTRANS='" & rsUP2.Fields("NOINV").Value & "' AND TGL='" & Format(rsUP2.Fields("TGLINV").Value, "YYYY/MM/DD") & "'"), con, 2, 1
        Do Until rsUP4.EOF
            putus_gr = Val(putus_gr) + rsUP4.Fields("jumlah").Value
            putus_RP = Val(putus_RP) + rsUP4.Fields("jumlahrp").Value
            'bayar_RP = Val(0)
                    Set rsUP3 = Nothing
                    Set rsUP3 = New ADODB.Recordset
                    rsUP3.CursorLocation = adUseClient
                    rsUP3.Open LCase("SELECT * FROM BAYARPIUTANGRP WHERE IDSUPPLIER='" & rsUP2.Fields("IDSUPPLIER").Value & "' AND IDSALES='" & rsUP2.Fields("IDSALES").Value & "' AND NOTRANS='" & rsUP2.Fields("NOINV").Value & "' AND TGL='" & Format(rsUP2.Fields("TGLINV").Value, "YYYY/MM/DD") & "' AND IDBAYAR=" & rsUP4.Fields("NOURUT").Value), con, 2, 1
                    Do Until rsUP3.EOF
                        bayar_RP = Val(bayar_RP) + rsUP3.Fields("JUMLAHRP2").Value
                    rsUP3.MoveNext
                    Loop
        rsUP4.MoveNext
        Loop
        
    If Not (FormatNumber(((hutangAwal_gr) - (putus_gr)), 3)) = FormatNumber(0, 3) Then
                            rsTempLap7.AddNew
                            rsTempLap7!Notrans = rsUP2.Fields("noinv").Value
                            rsTempLap7!tanggal = rsUP2.Fields("tglinv").Value
                            rsTempLap7!toko = rsUP2.Fields("IDSUPPLIER").Value
                            rsTempLap7!sales = rsUP2.Fields("IDSALES").Value
                            rsTempLap7!SAPIUTANG_GR = hutangAwal_gr
                            rsTempLap7!jenis = "KREDIT"
                            rsTempLap7!status = rsUP2.Fields("STATUS").Value
                            rsTempLap7!jual_gr = Val(0)
                            rsTempLap7!retur_gr = Val(0)
                            rsTempLap7!bayar_Gr = putus_gr
                            rsTempLap7!sisa_GR = ((hutangAwal_gr) - (putus_gr))
                            rsTempLap7.Update
    End If
 
rsUP2.MoveNext
Loop

'-------------------------
'END OF SALDO AWAL PIUTANG
'-------------------------
    
    
'-------------------
'TRANSAKSI PENJUALAN
'-------------------

 Set rsUP2_1 = Nothing
 Set rsUP2_1 = New ADODB.Recordset
 rsUP2_1.CursorLocation = adUseClient
 'rsUP2_1.Open "SELECT JUAL_HD.NOTRANS,JUAL_HD.TGL,IDSALES,IDSUPPLIER,SUM((QTY*HARGA)-((QTY*JUAL_DT.DISC)/100))AS MURNI FROM JUAL_HD,JUAL_DT WHERE JUAL_HD.STATUS='BELUM LUNAS' AND JUAL_HD.NOTRANS=JUAL_DT.NOTRANS AND JUAL_HD.TGL=JUAL_DT.TGL AND IDSUPPLIER='" & RsData6.Fields("NAMA").Value & "' AND IDSALES='" & RsData1.Fields("NAMA").Value & "' GROUP BY NOTRANS,TGL", con, 2, 1
 'rsUP2_1.Open "SELECT JUAL_HD.NOTRANS,JUAL_HD.TGL,JENIS,STATUS,IDSALES,IDSUPPLIER,SUM(ROUND((QTY*HARGA),3)-(ROUND((QTY*JUAL_DT.DISC),3)/100))AS MURNI FROM JUAL_HD,JUAL_DT WHERE JUAL_HD.STATUS='BELUM LUNAS' AND JUAL_HD.NOTRANS=JUAL_DT.NOTRANS AND JUAL_HD.TGL=JUAL_DT.TGL GROUP BY NOTRANS,TGL", con, 2, 1
 rsUP2_1.Open LCase("SELECT JUAL_HD.NOTRANS,JUAL_HD.TGL,JENIS,STATUS,IDSALES,IDSUPPLIER,SUM(((QTY*HARGA))-(((QTY*JUAL_DT.DISC))/100))AS MURNI FROM JUAL_HD,JUAL_DT WHERE JUAL_HD.STATUS='BELUM LUNAS' AND JUAL_HD.NOTRANS=JUAL_DT.NOTRANS AND JUAL_HD.TGL=JUAL_DT.TGL GROUP BY NOTRANS,TGL"), con, 2, 1
 Do Until rsUP2_1.EOF
    jual_gr = rsUP2_1.Fields("murni").Value
    retur_gr = Val(0)
    putus_gr1 = Val(0)
    putus_RP1 = Val(0)
    col1_1 = Val(0)
    col2_1 = Val(0)
    col3_1 = Val(0)
    col4_1 = Val(0)
    bayar_RP1 = Val(0)
 '-----------------------------
        'trans. retur jual
'-----------------------------
         Set rsUP2_2 = Nothing
        Set rsUP2_2 = New ADODB.Recordset
        rsUP2_2.CursorLocation = adUseClient
        rsUP2_2.Open LCase("SELECT RETURJUAL_HD.NOTRANS,RETURJUAL_HD.TGL,IDSALES,IDSUPPLIER,SUM(((QTY*HARGA))-(((QTY*RETURJUAL_DT.DISC))/100))AS MURNI FROM RETURJUAL_HD,RETURJUAL_DT WHERE RETURJUAL_HD.NOTRANS=RETURJUAL_DT.NOTRANS AND RETURJUAL_HD.TGL=RETURJUAL_DT.TGL AND IDSUPPLIER='" & rsUP2_1.Fields("IDSUPPLIER").Value & "' AND IDSALES='" & rsUP2_1.Fields("IDSALES").Value & "' AND RETURJUAL_HD.NOTRANS='" & rsUP2_1.Fields("NOTRANS").Value & "' AND RETURJUAL_HD.TGL='" & Format(rsUP2_1.Fields("TGL").Value, "YYYY/MM/DD") & "' GROUP BY NOTRANS,TGL"), con, 2, 1
        Do Until rsUP2_2.EOF
        retur_gr = Val(retur_gr) + rsUP2_2.Fields("murni").Value
           rsUP2_2.MoveNext
        Loop
 
 '-----------------------------
        'trans. bayar gr & Rp
 '-----------------------------
        Set rsUP2_3 = Nothing
        Set rsUP2_3 = New ADODB.Recordset
        rsUP2_3.CursorLocation = adUseClient
        rsUP2_3.Open LCase("SELECT * FROM BAYARPIUTANG WHERE IDSUPPLIER='" & rsUP2_1.Fields("IDSUPPLIER").Value & "' AND IDSALES='" & rsUP2_1.Fields("IDSALES").Value & "' AND NOTRANS='" & rsUP2_1.Fields("NOTRANS").Value & "' AND TGL='" & Format(rsUP2_1.Fields("TGL").Value, "YYYY/MM/DD") & "'"), con, 2, 1
        Do Until rsUP2_3.EOF
            putus_gr1 = Val(putus_gr1) + rsUP2_3.Fields("jumlah").Value
            putus_RP1 = Val(putus_RP1) + rsUP2_3.Fields("jumlahrp").Value
            'bayar_RP1 = Val(0)
                    Set rsUP2_4 = Nothing
                    Set rsUP2_4 = New ADODB.Recordset
                    rsUP2_4.CursorLocation = adUseClient
                    rsUP2_4.Open LCase("SELECT * FROM BAYARPIUTANGRP WHERE IDSUPPLIER='" & rsUP2_1.Fields("IDSUPPLIER").Value & "' AND IDSALES='" & rsUP2_1.Fields("IDSALES").Value & "' AND NOTRANS='" & rsUP2_1.Fields("NOTRANS").Value & "' AND TGL='" & Format(rsUP2_1.Fields("TGL").Value, "YYYY/MM/DD") & "' AND IDBAYAR=" & rsUP2_3.Fields("NOURUT").Value), con, 2, 1
                    Do Until rsUP2_4.EOF
                        bayar_RP1 = Val(bayar_RP1) + rsUP2_4.Fields("JUMLAHRP2").Value
                    rsUP2_4.MoveNext
                    Loop
        rsUP2_3.MoveNext
        Loop
 
 

    If Not (((FormatNumber(jual_gr, 3)) - (FormatNumber(retur_gr, 3)) - (putus_gr1))) = FormatNumber(0, 3) Then
                            rsTempLap7.AddNew
                            rsTempLap7!Notrans = rsUP2_1.Fields("notrans").Value
                            rsTempLap7!tanggal = rsUP2_1.Fields("tgl").Value
                            rsTempLap7!toko = rsUP2_1.Fields("IDSUPPLIER").Value
                            rsTempLap7!sales = rsUP2_1.Fields("IDSALES").Value
                            rsTempLap7!SAPIUTANG_GR = (FormatNumber(jual_gr, 3) - FormatNumber(retur_gr, 3))
                            rsTempLap7!status = rsUP2_1.Fields("STATUS").Value
                            rsTempLap7!jenis = rsUP2_1.Fields("JENIS").Value
                            rsTempLap7!jual_gr = Val(0)
                            rsTempLap7!retur_gr = Val(0)
                            rsTempLap7!bayar_Gr = putus_gr1
                            rsTempLap7!sisa_GR = FormatNumber(jual_gr, 3) - FormatNumber(retur_gr, 3) - putus_gr1
                            rsTempLap7.Update
    End If
 
    rsUP2_1.MoveNext
 Loop

'--------------------------
'END OF TRANSAKSI PENJUALAN
'--------------------------

rsTempLap7.Filter = adFilterNone
'With FrmBrowUmurPiutang
'Set .DataGrid1.DataSource = rsTempLap7
'.DataGrid1.Columns("Umur_Nota").Visible = False
'.DataGrid1.Columns("Bayar_Gr").Visible = False
'.DataGrid1.Columns("Jual_Gr").Visible = False
'.DataGrid1.Columns("Retur_Gr").Visible = False
'.DataGrid1.Columns("Hutang_Rupiah").Visible = False
'.DataGrid1.Columns("Bayar_Rupiah").Visible = False
'.DataGrid1.Columns("Komisi").Visible = False
'.DataGrid1.Columns("col_1").Visible = False
'.DataGrid1.Columns("col_2").Visible = False
'.DataGrid1.Columns("col_3").Visible = False
'.DataGrid1.Columns("col_4").Visible = False
'.DataGrid1.Columns("bayar_rupiah2").Visible = False
'.DataGrid1.Columns("sisa_rupiah").Visible = False
'.DataGrid1.Columns("titipan_awal").Visible = False
'.DataGrid1.Columns("titipan_potong").Visible = False
'.DataGrid1.Columns("titipan_akhir").Visible = False
'.DataGrid1.Columns("kurs_bayar").Visible = False
'.DataGrid1.Columns("kurs_sekarang").Visible = False
'.DataGrid1.Columns("SAPIUTANG_GR").Visible = False
'.Show 1
'End With

End Function




Public Function nobukti(ByVal idfrm As String) As String
On Error GoTo Errnobukti
    Set rsNoBukti = Nothing
    Set rsNoBukti = New ADODB.Recordset
    rsNoBukti.Open LCase("SELECT * FROM NOMORBUKTI WHERE FRMID='" & idfrm & "'"), con, adOpenDynamic, adLockOptimistic, adCmdText
    If rsNoBukti.EOF = False Then
        nobukti = rsNoBukti.Fields("judul").Value & Format(rsNoBukti.Fields("var1").Value, "0000") & rsNoBukti.Fields("batas").Value & rsNoBukti.Fields("var2").Value & rsNoBukti.Fields("batas").Value & rsNoBukti.Fields("var3").Value
    End If
Errnobukti:
If Not Err.Number = 0 Then
    MsgBox Err.Description, vbCritical, Err.Number
End If

End Function

Public Function UpdateNoBukti(ByVal idfrm As String)
    Set rsNoBukti = Nothing
    Set rsNoBukti = New ADODB.Recordset
    rsNoBukti.Open LCase("SELECT * FROM NOMORBUKTI WHERE FRMID='" & idfrm & "' LOCK IN SHARE MODE"), con, adOpenDynamic, adLockOptimistic, adCmdText
    If rsNoBukti.EOF = False Then
        rsNoBukti.Fields("var1").Value = rsNoBukti.Fields("var1").Value + 1
        rsNoBukti.Update
    End If
End Function


Function KonekDBChat()
On Error Resume Next
Set conCHAT = Nothing
Set conCHAT = New ADODB.Connection
conCHAT.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & idxservername & ";PORT=3306;UID=" & LCase("root") & ";PWD=" & LCase("batuceper") & ";DATABASE=" & idxSqlDb & ";Option=3;"
conCHAT.CursorLocation = adUseClient
conCHAT.Open

End Function



