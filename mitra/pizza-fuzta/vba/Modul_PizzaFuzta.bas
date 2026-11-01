Attribute VB_Name = "Modul_PizzaFuzta"
Option Explicit
'Shared module for MEDOKAN, KARAH and GUDANG. ASCII; no restaurant VBA inherited.
Private Function S(ByVal name As String) As Worksheet
    Set S = ThisWorkbook.Worksheets(name)
End Function
Private Sub Must(ByVal ok As Boolean, ByVal message As String)
    If Not ok Then Err.Raise vbObjectError + 701, "Pizza Fuzta", message
End Sub
Private Function Role() As String
    Role = CStr(S("CONFIG").Range("B6").Value)
End Function
Private Function Num(ByVal v As Variant, ByVal title As String) As Double
    Must Not IsError(v), title & " error."
    Must IsNumeric(v) And Len(CStr(v)) > 0, "Isi angka untuk " & title & "."
    Num = CDbl(v)
End Function
Private Function DayValue(ByVal v As Variant) As Date
    Must Not IsError(v), "Tanggal error."
    Must IsDate(v) And Len(CStr(v)) > 0, "Isi tanggal yang valid."
    DayValue = DateValue(CDate(v))
    Must DayValue > 0 And DayValue <= Date, "Tanggal harus sampai hari ini."
End Function
Private Function Find(ByVal tab As String, ByVal id As String, ByVal last As Long, Optional ByVal column As String = "A") As Long
    Dim r As Long
    Must Len(id) > 0, "Pilih kode pada " & tab & "."
    Must Application.CountIf(S(tab).Range(column & "7:" & column & last), id) = 1, "Kode hilang/ganda: " & id
    For r = 7 To last
        If StrComp(CStr(S(tab).Range(column & r).Value), id, vbTextCompare) = 0 Then Find = r: Exit Function
    Next r
End Function
Private Function Free(ByVal tab As String, ByVal last As Long, ByVal lastInput As String, ByVal count As Long) As Long
    Dim r As Long, j As Long
    For r = 7 To last
        If Len(CStr(S(tab).Cells(r, 1).Value)) = 0 Then
            Must r + count - 1 <= last, "Kapasitas " & tab & " tidak cukup."
            For j = r To r + count - 1
                Must Application.CountA(S(tab).Range("A" & j & ":" & lastInput & j)) = 0, "Ada baris parsial pada " & tab & " baris " & j
                If tab = "MUTASI" Then Must Application.CountA(S(tab).Range("T" & j & ":V" & j)) = 0, "Metadata transfer belum kosong."
                If tab = "NOTA" Then Must Len(CStr(S(tab).Cells(j, 17).Value)) = 0, "Status stok nota belum kosong."
            Next j
            Free = r: Exit Function
        End If
    Next r
    Err.Raise vbObjectError + 702, , "Tabel " & tab & " penuh; arsipkan file."
End Function
Private Function ID(ByVal prefix As String, ByVal tab As String, ByVal last As Long, Optional ByVal column As String = "A") As String
    Dim i As Long, candidate As String
    For i = 1 To 999
        candidate = "PF-" & Role() & "-" & prefix & "-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & Format$(i, "00")
        If Application.CountIf(S(tab).Range(column & "7:" & column & last), candidate) = 0 Then ID = candidate: Exit Function
    Next i
    Err.Raise vbObjectError + 703, , "Tidak dapat membuat ID unik."
End Function
Private Sub Put(ByVal tab As String, ByVal row As Long, ByVal values As Variant)
    Dim i As Long
    For i = LBound(values) To UBound(values)
        If Not IsEmpty(values(i)) Then
            If VarType(values(i)) = vbString Then S(tab).Cells(row, i + 1).NumberFormat = "@"
            S(tab).Cells(row, i + 1).Value = values(i)
        End If
    Next i
End Sub
Private Sub ClearMoves(ByVal first As Long, ByVal count As Long)
    If count <= 0 Then Exit Sub
    S("MUTASI").Range("A" & first & ":N" & first + count - 1).ClearContents
    S("MUTASI").Range("T" & first & ":V" & first + count - 1).ClearContents
End Sub
Private Sub CheckAll()
    Dim spec As Variant
    Application.CalculateFull
    For Each spec In Array(Array("MUTASI", "Q", 4006), Array("KEUANGAN", "H", 1006))
        Must Application.CountIfs(S(CStr(spec(0))).Range("A7:A" & spec(2)), "<>", S(CStr(spec(0))).Range(spec(1) & "7:" & spec(1) & spec(2)), "<>OK") = 0, "Periksa kolom Cek pada " & spec(0)
    Next spec
    If Role() = "MEDOKAN" Then
        Must Application.CountIfs(S("NOTA").Range("A7:A1006"), "<>", S("NOTA").Range("O7:O1006"), "<>OK") = 0, "Periksa NOTA."
        Must Application.CountIfs(S("ITEM").Range("A7:A3006"), "<>", S("ITEM").Range("M7:M3006"), "<>OK") = 0, "Periksa ITEM."
    End If
    Must Application.CountIf(S("MUTASI").Range("R7:R4006"), "<-0.000001") = 0, "Saldo historis negatif. Periksa stok awal, satuan, pembelian atau penerimaan."
End Sub
Public Sub ValidasiPizza()
    CheckAll
End Sub

Public Sub SimpanNota()
    Dim k As Worksheet, day As Date, hour As Double, method As String, cashier As String
    Dim carts As Collection, item As Variant, r As Long, mr As Long, j As Long, i As Long
    Dim qty As Double, price As Double, gross As Double, discount As Double, received As Double
    Dim needed(1 To 80) As Double, allRecipe As Boolean, countMoves As Long, nr As Long, ir As Long, gr As Long
    Dim noteID As String, written As Boolean, committed As Boolean, message As String
    On Error GoTo Failed
    Must Role() = "MEDOKAN", "Kasir ini hanya untuk Medokan; Karah memakai rekap Majoo."
    Set k = S("KASIR"): Set carts = New Collection: allRecipe = True
    Application.CalculateFull
    day = DayValue(k.Range("B6").Value): Must IsDate(k.Range("E6").Value), "Isi jam transaksi."
    hour = CDbl(TimeValue(CDate(k.Range("E6").Value)))
    cashier = Trim$(CStr(k.Range("B8").Value)): Must Len(cashier) > 0, "Isi kasir."
    method = CStr(k.Range("H6").Value): Must Application.CountIf(S("CONFIG").Range("O2:O8"), method) = 1, "Pilih metode."
    For r = 13 To 24
        If Len(CStr(k.Cells(r, 2).Value)) > 0 Or Len(CStr(k.Cells(r, 4).Value)) > 0 Then
            mr = Find("MENU", CStr(k.Cells(r, 2).Value), 86)
            Must S("MENU").Cells(mr, 6).Value = "AKTIF" And S("MENU").Cells(mr, 11).Value = "OK", "Menu belum valid/aktif."
            qty = Num(k.Cells(r, 4).Value, "Qty"): price = Num(k.Cells(r, 5).Value, "Harga")
            Must qty > 0 And qty = Fix(qty) And price > 0, "Qty harus bulat positif; harga positif."
            carts.Add Array(CStr(k.Cells(r, 2).Value), qty, price, S("MENU").Cells(mr, 8).Value, S("MENU").Cells(mr, 9).Value)
            gross = gross + qty * price
            If S("MENU").Cells(mr, 12).Value <> "LENGKAP" Then allRecipe = False
            For j = 1 To 80: needed(j) = needed(j) + qty * CDbl(S("RESEP").Cells(mr, j + 2).Value): Next j
        End If
    Next r
    Must carts.Count > 0, "Keranjang kosong."
    discount = Num(k.Range("F27").Value, "Diskon (isi 0 jika tidak ada)")
    Must discount >= 0 And discount <= gross, "Diskon tidak valid."
    If method = "Tunai" Then
        received = Num(k.Range("F29").Value, "Diterima"): Must received >= gross - discount, "Uang tunai kurang."
    Else
        received = gross - discount
    End If
    If allRecipe Then
        For j = 1 To 80: If needed(j) > 0 Then countMoves = countMoves + 1
        Next j
        Must countMoves > 0, "Resep kosong."
        gr = Free("MUTASI", 4006, "N", countMoves)
    End If
    nr = Free("NOTA", 1006, "I", 1): ir = Free("ITEM", 3006, "H", carts.Count)
    noteID = ID("N", "NOTA", 1006): written = True
    Put "NOTA", nr, Array(noteID, day, hour, cashier, method, discount, received, "AKTIF", "Snapshot kasir")
    S("NOTA").Cells(nr, 17).Value = IIf(allRecipe, "LENGKAP", "BELUM LENGKAP")
    For Each item In carts
        Put "ITEM", ir + i, Array(noteID & "-I" & Format$(i + 1, "00"), noteID, item(0), item(1), item(2), Empty, item(4), "Snapshot saat nota")
        If CStr(item(4)) = "LENGKAP" Then S("ITEM").Cells(ir + i, 6).Value = item(3)
        i = i + 1
    Next item
    If allRecipe Then
        i = 0
        For j = 1 To 80
            If needed(j) > 0 Then
                Put "MUTASI", gr + i, Array(noteID & "-G" & Format$(i + 1, "00"), day, "PAKAI", S("BAHAN").Cells(j + 6, 1).Value, needed(j), Empty, Empty, Empty, Empty, Empty, Empty, noteID, "AKTIF", "Snapshot resep")
                i = i + 1
            End If
        Next j
    End If
    CheckAll: written = False: committed = True
    S("STRUK").Range("B5").Value = noteID
    NotaBaru
    MsgBox "Nota " & noteID & " tercatat. Stok: " & IIf(allRecipe, "LENGKAP", "BELUM LENGKAP - lengkapi pemakaian satu kali") & ". Tekan Ctrl+S.", vbInformation, "Pizza Fuzta"
    Exit Sub
Failed:
    message = Err.Description
    If committed Then MsgBox "Nota SUDAH tercatat, jangan simpan ulang. " & message, vbExclamation: Exit Sub
    If written Then
        S("NOTA").Range("A" & nr & ":I" & nr).ClearContents: S("NOTA").Cells(nr, 17).ClearContents
        S("ITEM").Range("A" & ir & ":H" & ir + carts.Count - 1).ClearContents
        ClearMoves gr, countMoves
    End If
    Application.CalculateFull
    MsgBox message, vbExclamation, "Nota belum disimpan"
End Sub
Public Sub NotaBaru()
    Dim r As Long
    Must Role() = "MEDOKAN", "Hanya untuk Medokan."
    With S("KASIR")
        .Range("B13:B24,D13:D24,F29").ClearContents: .Range("F27").Value = 0
        For r = 13 To 24
            .Cells(r, 5).Formula = "=IF(B" & r & "="""","""",IFERROR(INDEX(MENU!$E$7:$E$86,MATCH(B" & r & ",MENU!$A$7:$A$86,0)),0))"
        Next r
    End With
End Sub
Public Sub SimpanBarang()
    Dim f As Worksheet, day As Date, kind As String, code As String, qty As Double, amount As Double, method As String
    Dim target As String, transfer As String, r As Long, mr As Long, written As Boolean, committed As Boolean, message As String
    On Error GoTo Failed
    Set f = S("FORM"): day = DayValue(f.Range("B6").Value): kind = CStr(f.Range("E6").Value)
    Must InStr(1, "|AWAL|BELI|KIRIM|PAKAI|RUSAK|KOREKSI MASUK|KOREKSI KELUAR|", "|" & kind & "|", vbBinaryCompare) > 0 And Len(kind) > 0, "Jenis tidak sesuai form."
    code = CStr(f.Range("B8").Value): mr = Find("BAHAN", code, 86)
    qty = Num(f.Range("E8").Value, "Qty"): Must qty > 0, "Qty positif."
    If kind = "BELI" Then amount = Num(f.Range("B10").Value, "Nominal beli"): Must amount >= 0, "Nominal negatif."
    method = CStr(f.Range("E10").Value)
    If kind = "KIRIM" Then
        target = CStr(f.Range("B12").Value)
        Must target <> Role() And Application.CountIf(S("CONFIG").Range("Q2:Q4"), target) = 1, "Tujuan tidak valid."
        transfer = CStr(f.Range("E12").Value)
        If Len(transfer) = 0 Then transfer = ID("T", "MUTASI", 4006, "H")
        Must Application.CountIf(S("MUTASI").Range("H7:H4006"), transfer) = 0, "ID Transfer sudah dipakai."
    End If
    r = Free("MUTASI", 4006, "N", 1): written = True
    Put "MUTASI", r, Array(ID("G", "MUTASI", 4006), day, kind, code, qty, Empty, Empty, Empty, Empty, Empty, f.Range("B14").Value, Empty, "AKTIF", f.Range("E14").Value)
    If kind = "BELI" Then S("MUTASI").Cells(r, 6).Value = amount: S("MUTASI").Cells(r, 7).Value = method
    If kind = "KIRIM" Then Put "MUTASI", r, Array(Empty, Empty, Empty, Empty, Empty, Empty, Empty, transfer, Role(), target)
    CheckAll: written = False: committed = True
    f.Range("E8,B10,E12,B14,E14").ClearContents
    MsgBox "Gerak barang tercatat. Tekan Ctrl+S. Untuk transfer, ekspor CSV berikutnya.", vbInformation
    Exit Sub
Failed:
    message = Err.Description
    If committed Then MsgBox "Barang SUDAH tercatat. Jangan simpan ulang. " & message, vbExclamation: Exit Sub
    If written Then ClearMoves r, 1
    Application.CalculateFull: MsgBox message, vbExclamation, "Barang belum disimpan"
End Sub
Public Sub SimpanKeuangan()
    Dim f As Worksheet, r As Long, day As Date, amount As Double, written As Boolean, committed As Boolean, message As String
    On Error GoTo Failed
    Set f = S("FORM"): day = DayValue(f.Range("B21").Value): amount = Num(f.Range("B23").Value, "Nominal")
    Must amount > 0, "Nominal positif."
    r = Free("KEUANGAN", 1006, "G", 1): written = True
    Put "KEUANGAN", r, Array(ID("K", "KEUANGAN", 1006), day, f.Range("E21").Value, amount, f.Range("E23").Value, f.Range("B25").Value, "AKTIF")
    CheckAll: written = False: committed = True
    f.Range("B23,B25").ClearContents
    MsgBox "Keuangan tercatat; tekan Ctrl+S.", vbInformation: Exit Sub
Failed:
    message = Err.Description
    If committed Then MsgBox "Keuangan SUDAH tercatat. Jangan simpan ulang. " & message, vbExclamation: Exit Sub
    If written Then S("KEUANGAN").Range("A" & r & ":G" & r).ClearContents
    Application.CalculateFull: MsgBox message, vbExclamation, "Keuangan belum disimpan"
End Sub
Public Sub SimpanProduksi()
    Dim t As Worksheet, day As Date, doc As String, r As Long, n As Long, inputs As Long, outputs As Long, mr As Long
    Dim first As Long, i As Long, rows As Collection, a As Variant, written As Boolean, committed As Boolean, message As String
    On Error GoTo Failed
    Must Role() = "GUDANG", "Produksi house dicatat di file GUDANG."
    Set t = S("PRODUKSI"): Set rows = New Collection: day = DayValue(t.Range("B6").Value)
    For r = 11 To 22
        If Len(CStr(t.Cells(r, 1).Value)) > 0 Or Len(CStr(t.Cells(r, 2).Value)) > 0 Then
            mr = Find("BAHAN", CStr(t.Cells(r, 1).Value), 86): Must Num(t.Cells(r, 2).Value, "Input") > 0, "Input harus positif."
            rows.Add Array("PRODUKSI PAKAI", t.Cells(r, 1).Value, CDbl(t.Cells(r, 2).Value)): inputs = inputs + 1
        End If
        If Len(CStr(t.Cells(r, 5).Value)) > 0 Or Len(CStr(t.Cells(r, 6).Value)) > 0 Then
            mr = Find("BAHAN", CStr(t.Cells(r, 5).Value), 86): Must Num(t.Cells(r, 6).Value, "Output") > 0, "Output positif."
            rows.Add Array("PRODUKSI MASUK", t.Cells(r, 5).Value, CDbl(t.Cells(r, 6).Value)): outputs = outputs + 1
        End If
    Next r
    Must inputs > 0 And outputs > 0, "Isi minimal satu input dan output aktual."
    doc = CStr(t.Range("E6").Value): If Len(doc) = 0 Then doc = ID("P", "MUTASI", 4006, "V")
    Must Application.CountIf(S("MUTASI").Range("V7:V4006"), doc) = 0, "ID produksi sudah dipakai."
    first = Free("MUTASI", 4006, "N", rows.Count): written = True
    For Each a In rows
        Put "MUTASI", first + i, Array(doc & "-" & Format$(i + 1, "00"), day, a(0), a(1), a(2), Empty, Empty, Empty, Empty, Empty, Empty, Empty, "AKTIF", "Batch aktual production house")
        S("MUTASI").Cells(first + i, 22).Value = doc: i = i + 1
    Next a
    CheckAll: written = False: committed = True
    t.Range("A11:B22,E11:F22,E6").ClearContents
    MsgBox "Produksi " & doc & " tercatat. Tekan Ctrl+S.", vbInformation: Exit Sub
Failed:
    message = Err.Description
    If committed Then MsgBox "Produksi SUDAH tercatat. Jangan simpan ulang. " & message, vbExclamation: Exit Sub
    If written Then ClearMoves first, rows.Count
    Application.CalculateFull: MsgBox message, vbExclamation, "Produksi belum disimpan"
End Sub
Public Sub VoidNota()
    Dim r As Long, reason As String, changed As Boolean, message As String
    On Error GoTo Failed
    Must Role() = "MEDOKAN", "Hanya untuk Medokan.": Must (ActiveSheet Is S("NOTA")), "Pilih baris NOTA."
    r = ActiveCell.Row: Must r >= 7 And r <= 1006, "Pilih baris nota."
    Must Len(CStr(S("NOTA").Cells(r, 1).Value)) > 0 And S("NOTA").Cells(r, 8).Value = "AKTIF", "Pilih nota AKTIF."
    reason = Trim$(InputBox("Alasan VOID nota:", "Pizza Fuzta")): If Len(reason) = 0 Then Exit Sub
    Must Len(CStr(S("NOTA").Cells(r, 9).Value)) + Len(reason) + 10 < 32767, "Catatan terlalu panjang."
    S("NOTA").Cells(r, 8).Value = "VOID": changed = True: CheckAll
    S("NOTA").Cells(r, 9).Value = CStr(S("NOTA").Cells(r, 9).Value) & " | VOID: " & reason
    MsgBox "Nota VOID; penjualan dan pemakaian terkait dibatalkan. Tekan Ctrl+S.", vbInformation: Exit Sub
Failed:
    message = Err.Description: If changed Then S("NOTA").Cells(r, 8).Value = "AKTIF"
    Application.CalculateFull: MsgBox message, vbExclamation
End Sub
Public Sub EksporStrukPDF()
    Dim file As Variant, text As String, i As Long
    On Error GoTo Failed
    Must Role() = "MEDOKAN", "Hanya Medokan.": Application.CalculateFull
    Must S("STRUK").Range("A10").Value = "SIAP CETAK", "Pilih nota AKTIF dan valid."
    text = CStr(S("STRUK").Range("B5").Value)
    For i = 1 To Len(text)
        If InStr(1, "\/:*?""<>|", Mid$(text, i, 1), vbBinaryCompare) > 0 Then Mid$(text, i, 1) = "_"
    Next i
    file = Application.GetSaveAsFilename(InitialFileName:=text & ".pdf", FileFilter:="PDF (*.pdf), *.pdf")
    If VarType(file) = vbBoolean Then Exit Sub
    If LCase$(Right$(CStr(file), 4)) <> ".pdf" Then file = CStr(file) & ".pdf"
    S("STRUK").ExportAsFixedFormat Type:=xlTypePDF, Filename:=CStr(file), IgnorePrintAreas:=False, OpenAfterPublish:=False
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "PDF belum dibuat"
End Sub
Private Sub AddButton(ByVal tab As String, ByVal anchor As String, ByVal title As String, ByVal action As String)
    Dim btn As Button
    Set btn = S(tab).Buttons.Add(S(tab).Range(anchor).Left, S(tab).Range(anchor).Top, 180, 27)
    btn.Name = "pf_" & action: btn.Caption = title: btn.Font.Name = "Calibri"
    btn.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & action
End Sub
Public Sub PasangTombol()
    Dim t As Worksheet, i As Long
    On Error GoTo Failed
    For Each t In ThisWorkbook.Worksheets
        For i = t.Shapes.Count To 1 Step -1
            If Left$(t.Shapes(i).Name, 3) = "pf_" Then t.Shapes(i).Delete
        Next i
    Next t
    AddButton "FORM", "H6", "SIMPAN BARANG", "SimpanBarang": AddButton "FORM", "H22", "SIMPAN KEUANGAN", "SimpanKeuangan"
    AddButton "TRANSFER", "H5", "EKSPOR KIRIM CSV", "EksporKirimCSV": AddButton "TRANSFER", "H8", "IMPOR KIRIM CSV", "ImporKirimCSV"
    AddButton "TRANSFER", "H11", "EKSPOR TERIMA CSV", "EksporTerimaCSV": AddButton "TRANSFER", "H14", "IMPOR KONFIRMASI", "ImporKonfirmasiCSV"
    If Role() = "MEDOKAN" Then
        AddButton "KASIR", "J13", "SIMPAN NOTA", "SimpanNota": AddButton "KASIR", "J16", "NOTA BARU", "NotaBaru": AddButton "STRUK", "D5", "EKSPOR PDF", "EksporStrukPDF"
    ElseIf Role() = "GUDANG" Then
        AddButton "PRODUKSI", "I6", "SIMPAN PRODUKSI", "SimpanProduksi"
    End If
    MsgBox "Tombol dipasang. Impor juga Modul_TransferPizza.bas untuk CSV. Simpan .xlsm. VBA perlu diuji di Excel desktop.", vbInformation
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation
End Sub
