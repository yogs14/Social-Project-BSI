Attribute VB_Name = "Modul_MuritaCrepes"
Option Explicit
'ASCII source. Import into a saved .xlsm copy in Excel desktop.
Private Const LAST_NOTA As Long = 1006
Private Const LAST_ITEM As Long = 3006
Private Const LAST_GERAK As Long = 6006

Private Function S(ByVal title As String) As Range
    Select Case title
        Case "DASHBOARD": Set S = ThisWorkbook.Worksheets("RINGKASAN").Range("A1")
        Case "MULAI": Set S = ThisWorkbook.Worksheets("RINGKASAN").Range("A81")
        Case "KASIR": Set S = ThisWorkbook.Worksheets("KASIR").Range("A1")
        Case "STRUK": Set S = ThisWorkbook.Worksheets("KASIR").Range("A47")
        Case "PRODUKSI": Set S = ThisWorkbook.Worksheets("HARIAN").Range("A1")
        Case "FORM": Set S = ThisWorkbook.Worksheets("HARIAN").Range("A39")
        Case "TUTUP KASIR": Set S = ThisWorkbook.Worksheets("TUTUP KASIR").Range("A1")
        Case "BAHAN": Set S = ThisWorkbook.Worksheets("KATALOG").Range("A225")
        Case "MENU": Set S = ThisWorkbook.Worksheets("KATALOG").Range("A1")
        Case "KALIBRASI": Set S = ThisWorkbook.Worksheets("KATALOG").Range("A63")
        Case "PREMIKS": Set S = ThisWorkbook.Worksheets("KATALOG").Range("A101")
        Case "ADONAN": Set S = ThisWorkbook.Worksheets("KATALOG").Range("A135")
        Case "REFERENSI": Set S = ThisWorkbook.Worksheets("KATALOG").Range("A167")
        Case "PELANGGAN": Set S = ThisWorkbook.Worksheets("PELANGGAN").Range("A1")
        Case "NOTA": Set S = ThisWorkbook.Worksheets("JURNAL").Range("A1")
        Case "ITEM": Set S = ThisWorkbook.Worksheets("JURNAL").Range("A1015")
        Case "GERAK BAHAN": Set S = ThisWorkbook.Worksheets("JURNAL").Range("A4029")
        Case "KEUANGAN": Set S = ThisWorkbook.Worksheets("JURNAL").Range("A10043")
        Case "POIN": Set S = ThisWorkbook.Worksheets("JURNAL").Range("A11057")
        Case "ARSIP TUTUP": Set S = ThisWorkbook.Worksheets("JURNAL").Range("A12071")
        Case "RESEP": Set S = ThisWorkbook.Worksheets("_RESEP").Range("A1")
        Case Else: Err.Raise vbObjectError + 604, , "Blok tidak dikenal: " & title
    End Select
End Function

Private Sub Must(ByVal valid As Boolean, ByVal message As String)
    If Not valid Then Err.Raise vbObjectError + 601, "Mufrita Crepes", message
End Sub

Private Function Num(ByVal value As Variant, ByVal label As String) As Double
    Must Not IsError(value), label & " mengandung error."
    Must IsNumeric(value) And Len(CStr(value)) > 0, "Isi angka untuk " & label & "."
    Num = CDbl(value)
End Function

Private Function DayValue(ByVal value As Variant) As Date
    Must Not IsError(value), "Tanggal mengandung error."
    Must IsDate(value) And Len(CStr(value)) > 0, "Isi tanggal valid."
    DayValue = DateValue(CDate(value))
    Must DayValue > 0 And DayValue <= Date, "Tanggal harus sampai hari ini."
End Function

Private Function Find(ByVal title As String, ByVal id As String, ByVal last As Long) As Long
    Dim r As Long
    Must Len(id) > 0, "Pilih kode pada " & title & "."
    Must Application.CountIf(S(title).Range("A7:A" & last), id) = 1, "Kode hilang atau ganda: " & id
    For r = 7 To last
        If StrComp(CStr(S(title).Cells(r, 1).Value), id, vbTextCompare) = 0 Then Find = r: Exit Function
    Next r
End Function

Private Function FreeRows(ByVal title As String, ByVal last As Long, ByVal inputLast As String, ByVal count As Long) As Long
    Dim r As Long, j As Long
    For r = 7 To last
        If Len(CStr(S(title).Cells(r, 1).Value)) = 0 Then
            Must r + count - 1 <= last, "Kapasitas " & title & " tidak cukup. Arsipkan file."
            For j = r To r + count - 1
                Must Application.CountA(S(title).Range("A" & j & ":" & inputLast & j)) = 0, "Ada baris parsial/terisi pada " & title & " baris " & j & "."
                If title = "NOTA" Then Must Len(CStr(S(title).Cells(j, 14).Value)) = 0, "Baris poin NOTA belum bersih."
            Next j
            FreeRows = r: Exit Function
        End If
    Next r
    Err.Raise vbObjectError + 602, , "Kapasitas " & title & " penuh."
End Function

Private Function ID(ByVal prefix As String, ByVal title As String, ByVal last As Long, Optional ByVal column As String = "A") As String
    Dim i As Long, candidate As String
    For i = 1 To 999
        candidate = "MC-" & prefix & "-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & Format$(i, "00")
        If Application.CountIf(S(title).Range(column & "7:" & column & last), candidate) = 0 Then ID = candidate: Exit Function
    Next i
    Err.Raise vbObjectError + 603, , "Tidak dapat membuat ID unik."
End Function

Private Sub CheckAll()
    Dim pair As Variant
    Application.CalculateFull
    For Each pair In Array(Array("NOTA", "O", 1006), Array("ITEM", "M", 3006), Array("GERAK BAHAN", "N", 6006), Array("KEUANGAN", "H", 1006), Array("POIN", "H", 1006), Array("ARSIP TUTUP", "R", 372))
        Must Application.CountIfs(S(CStr(pair(0))).Range("A7:A" & pair(2)), "<>", S(CStr(pair(0))).Range(pair(1) & "7:" & pair(1) & pair(2)), "<>OK") = 0, "Periksa kolom Cek pada " & pair(0) & "."
    Next pair
    Must Application.CountIf(S("GERAK BAHAN").Range("O7:O6006"), "<-0.000001") = 0, "Stok historis negatif. Isi AWAL/BELI/produksi yang benar; periksa satuan dan tanggal."
    Must Application.CountIf(S("POIN").Range("J7:J1006"), "<0") = 0, "Poin historis negatif; batalkan penukaran terkait terlebih dahulu."
End Sub

Private Sub AddMove(ByVal row As Long, ByVal idValue As String, ByVal day As Date, ByVal doc As String, ByVal kind As String, ByVal material As String, ByVal qty As Double, Optional ByVal noteID As String = "", Optional ByVal note As String = "")
    With S("GERAK BAHAN")
        .Cells(row, 1).Value = idValue: .Cells(row, 2).Value = day: .Cells(row, 3).Value = doc
        .Cells(row, 4).Value = kind: .Cells(row, 5).Value = material: .Cells(row, 6).Value = qty
        .Cells(row, 9).Value = noteID: .Cells(row, 10).Value = "AKTIF": .Cells(row, 11).Value = note
    End With
End Sub

Public Sub SimpanNota()
    Dim k As Range, day As Date, hour As Double, customer As String, method As String
    Dim r As Long, menuRow As Long, customerRow As Long, i As Long, j As Long, n As Long
    Dim qty As Double, price As Double, gross As Double, discount As Double, net As Double, received As Double, points As Long
    Dim needed(1 To 60) As Double, carts As Collection, cart As Variant, countMoves As Long
    Dim nr As Long, ir As Long, gr As Long, noteID As String, written As Boolean, committed As Boolean, message As String
    On Error GoTo Failed
    Set k = S("KASIR"): Set carts = New Collection
    Application.CalculateFull
    day = DayValue(k.Range("B6").Value)
    Must IsDate(k.Range("E6").Value), "Isi jam transaksi."
    hour = CDbl(TimeValue(CDate(k.Range("E6").Value)))
    method = Trim$(CStr(k.Range("H6").Value))
    Must Application.CountIf(S("REFERENSI").Range("O2:O5"), method) = 1, "Pilih metode bayar."
    customer = Trim$(CStr(k.Range("B8").Value))
    If Len(customer) > 0 Then
        customerRow = Find("PELANGGAN", customer, 306)
        Must S("PELANGGAN").Cells(customerRow, 4).Value = "YA", "Pelanggan belum memberi izin CRM; gunakan transaksi anonim."
    End If
    Must S("PREMIKS").Range("B21").Value = "OK" And S("ADONAN").Range("B21").Value = "OK", "Perbaiki parameter resep premiks/adonan."
    For r = 13 To 24
        If Len(CStr(k.Cells(r, 2).Value)) > 0 Or Len(CStr(k.Cells(r, 4).Value)) > 0 Then
            menuRow = Find("MENU", Trim$(CStr(k.Cells(r, 2).Value)), 56)
            Must S("MENU").Cells(menuRow, 7).Value = "AKTIF" And S("MENU").Cells(menuRow, 11).Value = "OK", "Menu tidak aktif atau master belum valid."
            qty = Num(k.Cells(r, 4).Value, "Qty"): price = Num(k.Cells(r, 5).Value, "Harga jual")
            Must qty > 0 And qty = Fix(qty) And price > 0, "Qty harus bulat positif; harga harus positif."
            carts.Add Array(CStr(k.Cells(r, 2).Value), qty, price, S("MENU").Cells(menuRow, 8).Value, S("MENU").Cells(menuRow, 9).Value)
            gross = gross + qty * price
            For j = 1 To 60: needed(j) = needed(j) + qty * CDbl(S("RESEP").Cells(menuRow, j + 2).Value): Next j
        End If
    Next r
    Must carts.Count > 0, "Keranjang kosong."
    discount = Num(k.Range("F27").Value, "Diskon (isi 0 bila tidak ada)")
    Must discount >= 0 And discount <= gross, "Diskon harus antara nol dan bruto."
    net = gross - discount
    If method = "Tunai" Then
        received = Num(k.Range("F29").Value, "Uang diterima")
        Must received >= net, "Uang tunai kurang."
    Else
        received = net
    End If
    If S("REFERENSI").Range("B16").Value = "YA" And Len(customer) > 0 Then
        price = Num(S("REFERENSI").Range("B17").Value, "Rupiah per poin")
        Must price > 0, "Rupiah per poin harus positif."
        points = Fix(net / price)
    End If
    For j = 1 To 60: If needed(j) > 0 Then countMoves = countMoves + 1
    Next j
    Must countMoves > 0, "Resep pemakaian bahan kosong."
    nr = FreeRows("NOTA", LAST_NOTA, "I", 1): ir = FreeRows("ITEM", LAST_ITEM, "H", carts.Count)
    gr = FreeRows("GERAK BAHAN", LAST_GERAK, "K", countMoves): noteID = ID("N", "NOTA", LAST_NOTA)
    written = True
    With S("NOTA")
        .Cells(nr, 1).Value = noteID: .Cells(nr, 2).Value = day: .Cells(nr, 3).Value = hour
        .Cells(nr, 4).Value = customer: .Cells(nr, 5).Value = method: .Cells(nr, 6).Value = discount
        .Cells(nr, 7).Value = received: .Cells(nr, 8).Value = "AKTIF"
        .Cells(nr, 9).Value = CStr(k.Range("B33").Value) & " | Kasir: " & CStr(k.Range("H8").Value)
        .Cells(nr, 14).Value = points
    End With
    i = 0
    For Each cart In carts
        With S("ITEM")
            .Cells(ir + i, 1).Value = noteID & "-I" & Format$(i + 1, "00")
            .Cells(ir + i, 2).Value = noteID: .Cells(ir + i, 3).Value = cart(0)
            .Cells(ir + i, 4).Value = cart(1): .Cells(ir + i, 5).Value = cart(2)
            If CStr(cart(4)) = "LENGKAP" Then .Cells(ir + i, 6).Value = cart(3)
            .Cells(ir + i, 7).Value = cart(4)
        End With
        i = i + 1
    Next cart
    i = 0
    For j = 1 To 60
        If needed(j) > 0 Then
            AddMove gr + i, noteID & "-G" & Format$(i + 1, "00"), day, noteID, "JUAL PAKAI", CStr(S("BAHAN").Cells(j + 6, 1).Value), needed(j), noteID, "Snapshot pemakaian resep saat nota"
            i = i + 1
        End If
    Next j
    CheckAll
    written = False: committed = True
    S("STRUK").Range("B5").Value = noteID
    NotaBaru
    MsgBox "Nota " & noteID & " tersimpan di jurnal. Tekan Ctrl+S. HPP yang belum lengkap tetap ditandai pada ITEM/NOTA.", vbInformation, "Mufrita Crepes"
    Exit Sub
Failed:
    message = Err.Description
    If committed Then
        MsgBox "Nota " & noteID & " SUDAH tercatat. Jangan simpan ulang. Gagal membersihkan form: " & message & vbCrLf & "Tekan Ctrl+S.", vbExclamation, "Nota tercatat"
        Exit Sub
    End If
    If written Then
        S("NOTA").Range("A" & nr & ":I" & nr).ClearContents: S("NOTA").Cells(nr, 14).ClearContents
        S("ITEM").Range("A" & ir & ":H" & ir + carts.Count - 1).ClearContents
        S("GERAK BAHAN").Range("A" & gr & ":K" & gr + countMoves - 1).ClearContents
    End If
    Application.CalculateFull
    MsgBox message, vbExclamation, "Nota belum disimpan"
End Sub

Public Sub NotaBaru()
    Dim r As Long
    With S("KASIR")
        .Range("B13:B24,D13:D24,F29").ClearContents
        .Range("F27").Value = 0: .Range("B33").MergeArea.ClearContents
        .Range("B8").ClearContents
        For r = 13 To 24
            .Cells(r, 5).Formula = "=IF(B" & r & "="""","""",IFERROR(INDEX(KATALOG!$C$7:$C$56,MATCH(B" & r & ",KATALOG!$A$7:$A$56,0)),0))"
        Next r
    End With
End Sub

Public Sub SimpanProduksi()
    Dim p As Range, day As Date, kind As String, doc As String, gr As Long, i As Long, r As Long
    Dim count As Long, output As Double, written As Boolean, message As String
    On Error GoTo Failed
    Set p = S("PRODUKSI"): Application.CalculateFull
    day = DayValue(p.Range("B8").Value): kind = CStr(p.Range("B6").Value)
    Must kind = "PREMIKS" Or kind = "ADONAN BASAH", "Pilih tahap produksi."
    Must Num(p.Range("B10").Value, "Gram premiks") > 0, "Qty produksi harus positif."
    Must S("PREMIKS").Range("B21").Value = "OK" And S("ADONAN").Range("B21").Value = "OK", "Perbaiki resep produksi."
    For r = 14 To 20: If Len(CStr(p.Cells(r, 1).Value)) > 0 And p.Cells(r, 3).Value > 0 Then count = count + 1
    Next r
    output = Num(p.Range("E24").Value, "Hasil produksi"): Must output > 0, "Hasil harus positif."
    gr = FreeRows("GERAK BAHAN", LAST_GERAK, "K", count + 1)
    doc = ID("P", "GERAK BAHAN", LAST_GERAK, "C"): written = True
    For r = 14 To 20
        If Len(CStr(p.Cells(r, 1).Value)) > 0 And p.Cells(r, 3).Value > 0 Then
            AddMove gr + i, doc & "-" & Format$(i + 1, "00"), day, doc, "PRODUKSI PAKAI", CStr(p.Cells(r, 1).Value), CDbl(p.Cells(r, 3).Value), , kind
            i = i + 1
        End If
    Next r
    AddMove gr + i, doc & "-" & Format$(i + 1, "00"), day, doc, "PRODUKSI MASUK", CStr(p.Range("B24").Value), output, , kind
    CheckAll: written = False
    MsgBox "Produksi tercatat: " & doc & ". Tekan Ctrl+S.", vbInformation, "Mufrita Crepes"
    Exit Sub
Failed:
    message = Err.Description
    If written Then S("GERAK BAHAN").Range("A" & gr & ":K" & gr + count).ClearContents
    Application.CalculateFull
    MsgBox message, vbExclamation, "Produksi belum disimpan"
End Sub

Public Sub SimpanGerak()
    Dim f As Range, r As Long, day As Date, material As String, mr As Long, kind As String
    Dim qty As Double, amount As Double, method As String, written As Boolean, message As String
    On Error GoTo Failed
    Set f = S("FORM"): day = DayValue(f.Range("B6").Value)
    kind = CStr(f.Range("E6").Value)
    Must InStr(1, "|AWAL|BELI|RUSAK|KOREKSI MASUK|KOREKSI KELUAR|", "|" & kind & "|", vbBinaryCompare) > 0 And Len(kind) > 0, "Gunakan produksi/kasir untuk gerakan otomatis."
    material = Trim$(CStr(f.Range("B8").Value)): mr = Find("BAHAN", material, 66)
    qty = Num(f.Range("E8").Value, "Qty satuan dasar"): Must qty > 0, "Qty harus positif."
    If kind = "BELI" Then
        amount = Num(f.Range("B10").Value, "Nominal pembelian")
        Must amount >= 0, "Nominal tidak boleh negatif."
        method = CStr(f.Range("E10").Value)
        Must Application.CountIf(S("REFERENSI").Range("O2:O5"), method) = 1, "Pilih metode pembelian."
    End If
    r = FreeRows("GERAK BAHAN", LAST_GERAK, "K", 1)
    written = True
    AddMove r, ID("G", "GERAK BAHAN", LAST_GERAK), day, CStr(f.Range("B14").Value), kind, material, qty, , CStr(f.Range("B12").Value)
    If kind = "BELI" Then S("GERAK BAHAN").Cells(r, 7).Value = amount: S("GERAK BAHAN").Cells(r, 8).Value = method
    CheckAll: written = False
    f.Range("E8,B10,B12,B14").ClearContents
    MsgBox "Gerakan bahan tercatat. Tekan Ctrl+S.", vbInformation, "Mufrita Crepes"
    Exit Sub
Failed:
    message = Err.Description
    If written Then S("GERAK BAHAN").Range("A" & r & ":K" & r).ClearContents
    Application.CalculateFull
    MsgBox message, vbExclamation, "Gerakan belum disimpan"
End Sub

Public Sub SimpanKeuangan()
    Dim f As Range, r As Long, day As Date, amount As Double, written As Boolean, message As String
    On Error GoTo Failed
    Set f = S("FORM"): day = DayValue(f.Range("B20").Value): amount = Num(f.Range("B22").Value, "Nominal kas")
    Must amount > 0, "Nominal harus positif."
    r = FreeRows("KEUANGAN", 1006, "G", 1): written = True
    With S("KEUANGAN")
        .Cells(r, 1).Value = ID("K", "KEUANGAN", 1006): .Cells(r, 2).Value = day
        .Cells(r, 3).Value = f.Range("E20").Value: .Cells(r, 4).Value = amount
        .Cells(r, 5).Value = f.Range("E22").Value: .Cells(r, 6).Value = f.Range("B24").Value: .Cells(r, 7).Value = "AKTIF"
    End With
    CheckAll: written = False: f.Range("B22,B24").ClearContents
    MsgBox "Keuangan tercatat. Tekan Ctrl+S.", vbInformation, "Mufrita Crepes"
    Exit Sub
Failed:
    message = Err.Description
    If written Then S("KEUANGAN").Range("A" & r & ":G" & r).ClearContents
    Application.CalculateFull
    MsgBox message, vbExclamation, "Keuangan belum disimpan"
End Sub

Public Sub SimpanPoin()
    Dim f As Range, r As Long, cr As Long, day As Date, amount As Double, written As Boolean, message As String
    On Error GoTo Failed
    Set f = S("FORM"): day = DayValue(f.Range("B30").Value)
    cr = Find("PELANGGAN", CStr(f.Range("E30").Value), 306)
    amount = Num(f.Range("E32").Value, "Jumlah poin")
    Must amount > 0 And amount = Fix(amount), "Poin harus bulat positif."
    If f.Range("B32").Value = "TUKAR" Then Must S("REFERENSI").Range("B16").Value = "YA", "Program poin belum aktif."
    r = FreeRows("POIN", 1006, "G", 1): written = True
    With S("POIN")
        .Cells(r, 1).Value = ID("L", "POIN", 1006): .Cells(r, 2).Value = day
        .Cells(r, 3).Value = f.Range("E30").Value: .Cells(r, 4).Value = f.Range("B32").Value
        .Cells(r, 5).Value = amount: .Cells(r, 6).Value = f.Range("B34").Value: .Cells(r, 7).Value = "AKTIF"
    End With
    CheckAll: written = False: f.Range("E32,B34").ClearContents
    MsgBox "Gerak poin tercatat. Hadiah/diskon tidak diberikan otomatis; tekan Ctrl+S.", vbInformation, "Mufrita Crepes"
    Exit Sub
Failed:
    message = Err.Description
    If written Then S("POIN").Range("A" & r & ":G" & r).ClearContents
    Application.CalculateFull
    MsgBox message, vbExclamation, "Poin belum disimpan"
End Sub

Public Sub VoidNota()
    Dim t As Range, r As Long, reason As String, changed As Boolean, message As String, oldNote As String
    On Error GoTo Failed
    Set t = S("NOTA"): Must (ActiveSheet Is t.Worksheet), "Pilih baris nota di JURNAL bagian NOTA."
    r = ActiveCell.Row - t.Row + 1: Must r >= 7 And r <= LAST_NOTA, "Pilih baris nota."
    Must Len(CStr(t.Cells(r, 1).Value)) > 0 And t.Cells(r, 8).Value = "AKTIF", "Pilih nota AKTIF."
    reason = Trim$(InputBox("Alasan pembatalan nota " & t.Cells(r, 1).Value & ":", "Void Nota"))
    If Len(reason) = 0 Then Exit Sub
    oldNote = CStr(t.Cells(r, 9).Value): Must Len(oldNote) + Len(reason) + 10 < 32767, "Catatan terlalu panjang."
    t.Cells(r, 8).Value = "VOID": changed = True
    CheckAll
    t.Cells(r, 9).Value = oldNote & " | VOID: " & reason
    MsgBox "Nota VOID. Efek penjualan, stok dan poin dibatalkan; riwayat tetap ada. Tekan Ctrl+S.", vbInformation, "Mufrita Crepes"
    Exit Sub
Failed:
    message = Err.Description
    If changed Then t.Cells(r, 8).Value = "AKTIF": Application.CalculateFull
    MsgBox message, vbExclamation, "Pembatalan gagal"
End Sub

Public Sub EksporStrukPDF()
    Dim t As Range, file As Variant, idValue As String, i As Long, message As String
    On Error GoTo Failed
    Set t = S("STRUK"): Application.CalculateFull
    Must t.Range("A10").Value = "SIAP CETAK", "Pilih nota AKTIF dan valid."
    idValue = CStr(t.Range("B5").Value)
    For i = 1 To Len(idValue)
        If InStr(1, "\/:*?""<>|", Mid$(idValue, i, 1), vbBinaryCompare) > 0 Then Mid$(idValue, i, 1) = "_"
    Next i
    file = Application.GetSaveAsFilename(InitialFileName:=idValue & ".pdf", FileFilter:="PDF (*.pdf), *.pdf")
    If VarType(file) = vbBoolean Then Exit Sub
    If LCase$(Right$(CStr(file), 4)) <> ".pdf" Then file = CStr(file) & ".pdf"
    t.Range("A7:F35").ExportAsFixedFormat Type:=xlTypePDF, Filename:=CStr(file), IgnorePrintAreas:=False, OpenAfterPublish:=False
    MsgBox "Struk PDF tersimpan.", vbInformation, "Mufrita Crepes"
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "PDF belum dibuat"
End Sub

Private Sub AddButton(ByVal title As String, ByVal anchor As String, ByVal caption As String, ByVal action As String)
    Dim b As Button
    Set b = S(title).Worksheet.Buttons.Add(S(title).Range(anchor).Left, S(title).Range(anchor).Top, 180, 27)
    b.Name = "mc_" & action: b.Caption = caption
    b.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & action
    b.Font.Name = "Calibri": b.Font.Size = 11
End Sub

Public Sub PasangTombol()
    Dim t As Worksheet, i As Long
    On Error GoTo Failed
    For Each t In ThisWorkbook.Worksheets
        For i = t.Shapes.Count To 1 Step -1
            If Left$(t.Shapes(i).Name, 3) = "mc_" Then t.Shapes(i).Delete
        Next i
    Next t
    AddButton "KASIR", "J13", "SIMPAN NOTA", "SimpanNota"
    AddButton "KASIR", "J16", "NOTA BARU", "NotaBaru"
    AddButton "PRODUKSI", "H8", "SIMPAN PRODUKSI", "SimpanProduksi"
    AddButton "FORM", "H7", "SIMPAN BARANG", "SimpanGerak"
    AddButton "FORM", "H21", "SIMPAN KEUANGAN", "SimpanKeuangan"
    AddButton "FORM", "H31", "SIMPAN POIN", "SimpanPoin"
    AddButton "STRUK", "D5", "EKSPOR PDF", "EksporStrukPDF"
    AddButton "TUTUP KASIR", "H5", "SIMPAN TUTUP KASIR", "SimpanTutupKasir"
    MsgBox "Delapan tombol terpasang. Simpan .xlsm. Untuk pembatalan: pilih baris nota pada JURNAL lalu Alt+F8 > VoidNota.", vbInformation, "Mufrita Crepes"
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Tombol belum terpasang"
End Sub

'Closing snapshot: append revisions; no period locking or deletion of history.
Public Sub SimpanTutupKasir()
    Dim t As Range, a As Range, day As Date, r As Long, rev As Long
    Dim opening As Double, physical As Double, diff As Double, note As String
    Dim values(1 To 1, 1 To 17) As Variant
    Dim written As Boolean, message As String
    On Error GoTo Failed
    Set t = S("TUTUP KASIR"): Set a = S("ARSIP TUTUP")
    Application.CalculateFull
    day = DayValue(t.Range("C5").Value)
    opening = Num(t.Range("I8").Value, "Modal awal laci (isi 0 bila tidak ada)")
    physical = Num(t.Range("I11").Value, "Uang fisik di laci (isi 0 bila kosong)")
    Must opening >= 0 And physical >= 0, "Modal dan uang fisik tidak boleh negatif."
    Must Len(Trim$(CStr(t.Range("C26").Value))) > 0, "Isi nama kasir."
    CheckAll
    Must Num(t.Range("C18").Value, "Nota perlu diperiksa") = 0, "Perbaiki nota bermasalah sebelum menutup kasir."
    diff = Num(t.Range("I12").Value, "Selisih kas")
    note = Trim$(CStr(t.Range("B20").Value))
    Must Abs(diff) < 0.01 Or Len(note) > 0, "Isi catatan/alasan bila ada selisih kas."
    r = FreeRows("ARSIP TUTUP", 372, "Q", 1)
    rev = Application.CountIf(a.Range("B7:B372"), day) + 1
    values(1, 1) = ID("T", "ARSIP TUTUP", 372)
    values(1, 2) = day: values(1, 3) = rev
    values(1, 4) = CStr(t.Range("C26").Value)
    values(1, 5) = opening
    values(1, 6) = t.Range("F8").Value
    values(1, 7) = t.Range("F12").Value
    values(1, 8) = CDbl(t.Range("F15").Value) - CDbl(t.Range("F13").Value) - CDbl(t.Range("F14").Value)
    values(1, 9) = t.Range("I10").Value
    values(1, 10) = physical: values(1, 11) = diff
    values(1, 12) = t.Range("C13").Value
    values(1, 13) = t.Range("C8").Value
    values(1, 14) = t.Range("C10").Value
    If IsNumeric(t.Range("C15").Value) Then
        values(1, 15) = "LENGKAP"
    Else
        values(1, 15) = "BELUM LENGKAP"
    End If
    values(1, 16) = note: values(1, 17) = Now
    written = True
    a.Cells(r, 1).Resize(1, 17).Value = values
    Application.CalculateFull
    Must a.Cells(r, 18).Value = "OK", "Rekap belum valid; periksa angka dan identitas."
    written = False
    MsgBox "Tutup kasir " & Format$(day, "dd/mm/yyyy") & " revisi " & rev & " tersimpan. Tekan Ctrl+S. Riwayat lama tetap ada; periode tidak dikunci.", vbInformation, "Mufrita Crepes"
    Exit Sub
Failed:
    message = Err.Description
    If written Then a.Cells(r, 1).Resize(1, 17).ClearContents
    Application.CalculateFull
    MsgBox message, vbExclamation, "Tutup kasir belum disimpan"
End Sub
