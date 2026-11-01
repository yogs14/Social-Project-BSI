Attribute VB_Name = "Modul_TransferPizza"
Option Explicit
'PF1 CSV exchange via Excel itself. Import both .bas modules into each .xlsm copy.
Private Function W(ByVal name As String) As Worksheet
    Set W = ThisWorkbook.Worksheets(name)
End Function
Private Sub Need(ByVal valid As Boolean, ByVal message As String)
    If Not valid Then Err.Raise vbObjectError + 801, "Transfer Pizza", message
End Sub
Private Function Loc() As String
    Loc = CStr(W("CONFIG").Range("B6").Value)
End Function
Private Function IsLoc(ByVal code As String) As Boolean
    IsLoc = (code = "MEDOKAN" Or code = "KARAH" Or code = "GUDANG")
End Function
Private Sub Token(ByVal text As String, Optional ByVal allowEmpty As Boolean = False)
    Dim i As Long, c As String
    If allowEmpty And Len(text) = 0 Then Exit Sub
    Need Len(text) > 0 And Len(text) <= 120, "Kode/ID kosong atau terlalu panjang."
    For i = 1 To Len(text)
        c = Mid$(text, i, 1)
        Need InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_-", c, vbBinaryCompare) > 0, "ID hanya huruf, angka, garis bawah atau tanda minus."
    Next i
End Sub
Private Function ISODate(ByVal value As String) As Date
    Need Len(value) = 10 And Mid$(value, 5, 1) = "-" And Mid$(value, 8, 1) = "-", "Tanggal CSV harus yyyy-mm-dd."
    ISODate = DateSerial(CInt(Left$(value, 4)), CInt(Mid$(value, 6, 2)), CInt(Right$(value, 2)))
    Need Format$(ISODate, "yyyy-mm-dd") = value And ISODate > 0 And ISODate <= Date, "Tanggal tidak valid atau di masa depan."
End Function
Private Function Q(ByVal text As String) As Double
    Dim i As Long
    Need Len(text) > 0, "Qty kosong."
    For i = 1 To Len(text)
        Need InStr(1, "0123456789.eE+-", Mid$(text, i, 1), vbBinaryCompare) > 0, "Qty memakai angka desimal titik, tanpa pemisah ribuan."
    Next i
    Q = CDbl(Replace(text, ".", Application.DecimalSeparator))
    Need Q > 0, "Qty harus positif."
End Function
Private Function Dot(ByVal v As Variant) As String
    Dot = Replace(CStr(v), Application.DecimalSeparator, ".")
End Function
Private Function Headers() As Variant
    Headers = Array("schema", "kind", "transfer_id", "shipped_on", "received_on", "origin", "destination", "material", "unit", "quantity_shipped", "quantity_received", "request_id", "note")
End Function
Private Function RowBy(ByVal tab As String, ByVal column As String, ByVal value As String, ByVal last As Long) As Long
    Dim r As Long, n As Long
    n = Application.CountIf(W(tab).Range(column & "7:" & column & last), value)
    Need n <= 1, "ID ganda pada " & tab & ": " & value
    If n = 0 Then Exit Function
    For r = 7 To last
        If CStr(W(tab).Range(column & r).Value) = value Then RowBy = r: Exit Function
    Next r
End Function
Private Function FreeRow(ByVal tab As String, ByVal last As Long, ByVal inputLast As String) As Long
    Dim r As Long
    For r = 7 To last
        If Len(CStr(W(tab).Cells(r, 1).Value)) = 0 Then
            Need Application.CountA(W(tab).Range("A" & r & ":" & inputLast & r)) = 0, "Ada baris parsial pada " & tab
            If tab = "MUTASI" Then Need Application.CountA(W(tab).Range("T" & r & ":V" & r)) = 0, "Metadata penerimaan belum kosong."
            FreeRow = r: Exit Function
        End If
    Next r
    Err.Raise vbObjectError + 802, , "Kapasitas tabel penuh."
End Function
Private Sub Put(ByVal tab As String, ByVal r As Long, ByVal values As Variant)
    Dim i As Long
    For i = LBound(values) To UBound(values)
        If Not IsEmpty(values(i)) Then
            If VarType(values(i)) = vbString Then W(tab).Cells(r, i + 1).NumberFormat = "@"
            W(tab).Cells(r, i + 1).Value = values(i)
        End If
    Next i
End Sub
Private Function ReadCSV(ByVal file As String) As Collection
    Dim imported As Workbook, t As Worksheet, fields(0 To 12) As Variant
    Dim alreadyOpen As Workbook
    Dim i As Long, j As Long, last As Long, h As Variant, data As Variant, a(0 To 12) As String
    Dim rows As Collection, prev As Variant, message As String
    On Error GoTo Failed
    For Each alreadyOpen In Application.Workbooks
        Need StrComp(alreadyOpen.FullName, file, vbTextCompare) <> 0, "Tutup CSV yang sudah terbuka sebelum impor agar perubahan pengguna tidak tertutup."
    Next alreadyOpen
    For i = 0 To 12: fields(i) = Array(i + 1, xlTextFormat): Next i
    Workbooks.OpenText Filename:=file, Origin:=65001, DataType:=xlDelimited, Comma:=True, TextQualifier:=xlTextQualifierDoubleQuote, FieldInfo:=fields, Local:=False
    Set imported = ActiveWorkbook: Need Not (imported Is ThisWorkbook), "Tidak dapat membuka CSV."
    Set t = imported.Worksheets(1): h = Headers()
    Need t.UsedRange.Columns.Count = 13, "CSV harus memiliki 13 kolom PF1."
    For i = 0 To 12: Need CStr(t.Cells(1, i + 1).Value) = h(i), "Header CSV PF1 tidak cocok.": Next i
    last = t.Cells(t.Rows.Count, 1).End(xlUp).Row
    Need last >= 2, "CSV tidak berisi transaksi."
    data = t.Range("A2:M" & last).Value2
    Set rows = New Collection
    For i = 1 To UBound(data, 1)
        For j = 0 To 12: a(j) = CStr(data(i, j + 1)): Next j
        For Each prev In rows: Need CStr(prev(2)) <> a(2), "ID transfer ganda dalam CSV.": Next prev
        rows.Add Array(a(0), a(1), a(2), a(3), a(4), a(5), a(6), a(7), a(8), a(9), a(10), a(11), a(12))
    Next i
    imported.Close SaveChanges:=False: Set imported = Nothing
    Set ReadCSV = rows: Exit Function
Failed:
    message = Err.Description
    On Error Resume Next
    If Not imported Is Nothing Then
        If Not (imported Is ThisWorkbook) Then imported.Close SaveChanges:=False
    End If
    On Error GoTo 0
    Err.Raise vbObjectError + 803, , message
End Function
Private Sub Validate(ByVal a As Variant, ByVal kind As String)
    Dim mr As Long, shipped As Date, received As Date, qty As Double
    Need a(0) = "PF1" And a(1) = kind, "Schema/jenis CSV berbeda."
    Token CStr(a(2)): Token CStr(a(7)): Token CStr(a(11)), True
    Need IsLoc(CStr(a(5))) And IsLoc(CStr(a(6))) And a(5) <> a(6), "Asal/tujuan tidak valid."
    mr = RowBy("BAHAN", "A", CStr(a(7)), 86): Need mr > 0, "Bahan CSV tidak dikenal."
    Need CStr(W("BAHAN").Cells(mr, 4).Value) = CStr(a(8)), "Satuan bahan berbeda; tidak ada konversi otomatis."
    shipped = ISODate(CStr(a(3))): qty = Q(CStr(a(9)))
    If kind = "RECEIPT" Then
        received = ISODate(CStr(a(4))): Need received >= shipped, "Tanggal terima sebelum dikirim."
        qty = Q(CStr(a(10)))
    Else
        Need Len(CStr(a(4))) = 0 And Len(CStr(a(10))) = 0, "CSV kirim tidak boleh berisi penerimaan."
    End If
End Sub
Private Sub ExportCSV(ByVal kind As String)
    Dim peer As String, rows As Collection, a As Variant, r As Long, mr As Long
    Dim expected As String, file As Variant, temp As Workbook, h As Variant, i As Long, j As Long, message As String
    On Error GoTo Failed
    Modul_PizzaFuzta.ValidasiPizza
    peer = Trim$(InputBox("Kode lokasi lawan (MEDOKAN/KARAH/GUDANG):", "Ekspor " & kind, IIf(Loc() = "GUDANG", "MEDOKAN", "GUDANG")))
    If Len(peer) = 0 Then Exit Sub
    Need IsLoc(peer) And peer <> Loc(), "Lokasi lawan salah."
    expected = IIf(kind = "SHIPMENT", "KIRIM", "TERIMA"): Set rows = New Collection
    For r = 7 To 4006
        If W("MUTASI").Cells(r, 3).Value = expected And W("MUTASI").Cells(r, 13).Value = "AKTIF" Then
            If CStr(W("MUTASI").Cells(r, IIf(kind = "SHIPMENT", 10, 9)).Value) = peer Then
                mr = RowBy("BAHAN", "A", CStr(W("MUTASI").Cells(r, 4).Value), 86)
                Need mr > 0, "Kode bahan tidak ada."
                a = Array("PF1", kind, CStr(W("MUTASI").Cells(r, 8).Value), "", "", CStr(W("MUTASI").Cells(r, 9).Value), CStr(W("MUTASI").Cells(r, 10).Value), CStr(W("MUTASI").Cells(r, 4).Value), CStr(W("BAHAN").Cells(mr, 4).Value), "", "", CStr(W("MUTASI").Cells(r, 11).Value), CStr(W("MUTASI").Cells(r, 14).Value))
                If kind = "SHIPMENT" Then
                    a(3) = Format$(W("MUTASI").Cells(r, 2).Value, "yyyy-mm-dd"): a(9) = Dot(W("MUTASI").Cells(r, 5).Value)
                    Need a(5) = Loc(), "Pengiriman bukan dari lokasi ini."
                Else
                    a(3) = Format$(W("MUTASI").Cells(r, 20).Value, "yyyy-mm-dd"): a(4) = Format$(W("MUTASI").Cells(r, 2).Value, "yyyy-mm-dd")
                    a(9) = Dot(W("MUTASI").Cells(r, 21).Value): a(10) = Dot(W("MUTASI").Cells(r, 5).Value)
                    Need a(6) = Loc(), "Penerimaan bukan untuk lokasi ini."
                End If
                Validate a, kind: rows.Add a
            End If
        End If
    Next r
    Need rows.Count > 0, "Tidak ada transfer aktif untuk lokasi lawan tersebut."
    file = Application.GetSaveAsFilename(InitialFileName:="PF1_" & kind & "_" & Loc() & "_" & peer & ".csv", FileFilter:="CSV (*.csv), *.csv")
    If VarType(file) = vbBoolean Then Exit Sub
    If LCase$(Right$(CStr(file), 4)) <> ".csv" Then file = CStr(file) & ".csv"
    Set temp = Workbooks.Add(xlWBATWorksheet): h = Headers()
    temp.Worksheets(1).Range("A:M").NumberFormat = "@"
    For j = 0 To 12: temp.Worksheets(1).Cells(1, j + 1).Value2 = h(j): Next j
    i = 2
    For Each a In rows
        For j = 0 To 12: temp.Worksheets(1).Cells(i, j + 1).Value2 = CStr(a(j)): Next j
        i = i + 1
    Next a
    temp.SaveAs Filename:=CStr(file), FileFormat:=62, Local:=False
    temp.Close SaveChanges:=False: Set temp = Nothing
    MsgBox rows.Count & " baris diekspor. Serahkan CSV ke operator lokasi lawan; tidak ada pengiriman jaringan otomatis.", vbInformation
    Exit Sub
Failed:
    message = Err.Description
    On Error Resume Next
    If Not temp Is Nothing Then temp.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox message, vbExclamation, "CSV belum diekspor"
End Sub
Public Sub EksporKirimCSV()
    ExportCSV "SHIPMENT"
End Sub
Public Sub EksporTerimaCSV()
    ExportCSV "RECEIPT"
End Sub
Public Sub ImporKirimCSV()
    Dim file As Variant, rows As Collection, a As Variant, changed As Collection, item As Variant
    Dim dayText As String, received As Date, r As Long, existing As Long, skipped As Long, message As String
    On Error GoTo Failed
    file = Application.GetOpenFilename("CSV (*.csv), *.csv"): If VarType(file) = vbBoolean Then Exit Sub
    Set rows = ReadCSV(CStr(file))
    dayText = Trim$(InputBox("Tanggal terima aktual (yyyy-mm-dd):", "Terima Barang", Format$(Date, "yyyy-mm-dd")))
    If Len(dayText) = 0 Then Exit Sub
    received = ISODate(dayText)
    For Each a In rows
        Validate a, "SHIPMENT": Need a(6) = Loc(), "CSV bukan tujuan lokasi ini."
        Need received >= ISODate(CStr(a(3))), "Tanggal terima sebelum dikirim."
    Next a
    Set changed = New Collection
    For Each a In rows
        existing = RowBy("MUTASI", "H", CStr(a(2)), 4006)
        If existing > 0 Then
            With W("MUTASI")
                Need .Cells(existing, 3).Value = "TERIMA" And .Cells(existing, 13).Value = "AKTIF", "ID sudah dipakai untuk jenis/status lain."
                Need .Cells(existing, 9).Value = a(5) And .Cells(existing, 10).Value = a(6) And .Cells(existing, 4).Value = a(7), "ID sama memiliki data berbeda."
                Need Format$(.Cells(existing, 20).Value, "yyyy-mm-dd") = a(3) And Abs(CDbl(.Cells(existing, 21).Value) - Q(CStr(a(9)))) <= 0.000001 And CStr(.Cells(existing, 11).Value) = a(11), "Metadata kirim berubah."
            End With
            skipped = skipped + 1
        Else
            r = FreeRow("MUTASI", 4006, "N")
            Need Application.CountIf(W("MUTASI").Range("A7:A4006"), "RX-" & a(2)) = 0, "ID gerak penerimaan bentrok."
            changed.Add r
            Put "MUTASI", r, Array("RX-" & a(2), received, "TERIMA", a(7), Q(CStr(a(9))), Empty, Empty, a(2), a(5), a(6), a(11), Empty, "AKTIF", a(12))
            W("MUTASI").Cells(r, 20).Value = ISODate(CStr(a(3))): W("MUTASI").Cells(r, 21).Value = Q(CStr(a(9)))
        End If
    Next a
    Modul_PizzaFuzta.ValidasiPizza
    MsgBox changed.Count & " penerimaan baru; " & skipped & " duplikat identik dilewati. Periksa qty fisik sebelum ekspor konfirmasi; sumber Qty Kirim tidak diubah. Tekan Ctrl+S.", vbInformation
    Exit Sub
Failed:
    message = Err.Description
    If Not changed Is Nothing Then
        For Each item In changed
            W("MUTASI").Range("A" & item & ":N" & item).ClearContents
            W("MUTASI").Range("T" & item & ":V" & item).ClearContents
        Next item
    End If
    Application.CalculateFull
    MsgBox message, vbExclamation, "Batch penerimaan dibatalkan"
End Sub
Public Sub ImporKonfirmasiCSV()
    Dim file As Variant, rows As Collection, a As Variant, changed As Collection, item As Variant
    Dim sent As Long, existing As Long, r As Long, skipped As Long, message As String
    On Error GoTo Failed
    file = Application.GetOpenFilename("CSV (*.csv), *.csv"): If VarType(file) = vbBoolean Then Exit Sub
    Set rows = ReadCSV(CStr(file))
    For Each a In rows
        Validate a, "RECEIPT": Need a(5) = Loc(), "Konfirmasi bukan pengiriman dari lokasi ini."
        sent = RowBy("MUTASI", "H", CStr(a(2)), 4006): Need sent > 0, "Pengiriman ID tersebut tidak ditemukan."
        With W("MUTASI")
            Need .Cells(sent, 3).Value = "KIRIM" And .Cells(sent, 13).Value = "AKTIF", "Pengiriman tidak aktif."
            Need .Cells(sent, 9).Value = a(5) And .Cells(sent, 10).Value = a(6) And .Cells(sent, 4).Value = a(7), "Data pengiriman berbeda."
            Need Format$(.Cells(sent, 2).Value, "yyyy-mm-dd") = a(3) And Abs(CDbl(.Cells(sent, 5).Value) - Q(CStr(a(9)))) <= 0.000001 And CStr(.Cells(sent, 11).Value) = a(11), "Tanggal/qty/permintaan kirim berbeda."
        End With
    Next a
    Set changed = New Collection
    For Each a In rows
        existing = RowBy("KONFIRMASI", "A", CStr(a(2)), 2006)
        If existing > 0 Then
            With W("KONFIRMASI")
                Need Format$(.Cells(existing, 2).Value, "yyyy-mm-dd") = a(4) And Format$(.Cells(existing, 9).Value, "yyyy-mm-dd") = a(3), "Tanggal konfirmasi ID sama berbeda."
                Need .Cells(existing, 3).Value = a(5) And .Cells(existing, 4).Value = a(6) And .Cells(existing, 5).Value = a(7) And .Cells(existing, 8).Value = a(8), "Metadata konfirmasi ID sama berbeda."
                Need Abs(CDbl(.Cells(existing, 6).Value) - Q(CStr(a(9)))) <= 0.000001 And Abs(CDbl(.Cells(existing, 7).Value) - Q(CStr(a(10)))) <= 0.000001, "Qty konfirmasi ID sama berbeda."
            End With
            skipped = skipped + 1
        Else
            r = FreeRow("KONFIRMASI", 2006, "J"): changed.Add r
            Put "KONFIRMASI", r, Array(a(2), ISODate(CStr(a(4))), a(5), a(6), a(7), Q(CStr(a(9))), Q(CStr(a(10))), a(8), ISODate(CStr(a(3))), a(12))
        End If
    Next a
    Application.CalculateFull
    MsgBox changed.Count & " konfirmasi baru; " & skipped & " duplikat dilewati. Konfirmasi tidak menjadi mutasi stok. Periksa SELISIH dan tekan Ctrl+S.", vbInformation
    Exit Sub
Failed:
    message = Err.Description
    If Not changed Is Nothing Then
        For Each item In changed: W("KONFIRMASI").Range("A" & item & ":J" & item).ClearContents: Next item
    End If
    Application.CalculateFull: MsgBox message, vbExclamation, "Batch konfirmasi dibatalkan"
End Sub
