Attribute VB_Name = "Modul_GianFataya"
Option Explicit

'Import into POS_Gian_Fataya.xlsx and save as .xlsm in Excel desktop.
'No restaurant macros, external libraries, network calls or credentials.
Private Const FIRST_ROW As Long = 7
Private Const LAST_MOVE As Long = 2006
Private Const LAST_INV As Long = 406
Private Const LAST_PAY As Long = 1006

Private Function WS(ByVal title As String) As Worksheet
    Set WS = ThisWorkbook.Worksheets(title)
End Function

Private Sub Require(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 501, "Gian Fataya", message
End Sub

Private Function NumberOf(ByVal value As Variant, ByVal label As String) As Double
    Require Not IsError(value), label & " mengandung error."
    Require Len(Trim$(CStr(value))) > 0, "Isi " & label & "."
    Require IsNumeric(value), label & " harus angka."
    NumberOf = CDbl(value)
End Function

Private Function DayOf(ByVal value As Variant, ByVal label As String) As Date
    Require Not IsError(value), label & " mengandung error."
    Require Len(Trim$(CStr(value))) > 0, "Isi " & label & "."
    Require IsDate(value), label & " harus tanggal."
    DayOf = DateValue(CDate(value))
    Require DayOf > 0, label & " tidak valid."
End Function

Private Function FindID(ByVal sheet As Worksheet, ByVal id As String, ByVal last As Long) As Long
    Dim r As Long
    Require Len(id) > 0, "Pilih kode/ID terlebih dahulu."
    Require Application.CountIf(sheet.Range("A7:A" & last), id) = 1, "Kode tidak ditemukan atau ganda: " & id
    For r = FIRST_ROW To last
        If StrComp(CStr(sheet.Cells(r, 1).Value), id, vbTextCompare) = 0 Then FindID = r: Exit Function
    Next r
End Function

Private Function EmptyRow(ByVal sheet As Worksheet, ByVal last As Long, ByVal lastInput As String) As Long
    Dim r As Long, occupied As Boolean
    For r = FIRST_ROW To last
        If Len(CStr(sheet.Cells(r, 1).Value)) = 0 Then
            occupied = Application.CountA(sheet.Range("A" & r & ":" & lastInput & r)) > 0
            If sheet.Name = "MUTASI" Then occupied = occupied Or Application.CountA(sheet.Range("U" & r & ":W" & r)) > 0
            If sheet.Name = "INVOICE" Then occupied = occupied Or Len(CStr(sheet.Cells(r, 18).Value)) > 0
            Require Not occupied, "Lengkapi atau bersihkan baris parsial " & r & " pada " & sheet.Name & "."
            EmptyRow = r: Exit Function
        End If
    Next r
    Err.Raise vbObjectError + 502, , "Kapasitas tabel " & sheet.Name & " penuh. Arsipkan atau perluas versi workbook."
End Function

Private Function NewID(ByVal prefix As String, ByVal sheet As Worksheet, ByVal last As Long) As String
    Dim i As Long, candidate As String
    For i = 1 To 999
        candidate = "GF-" & prefix & "-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & Format$(i, "00")
        If Application.CountIf(sheet.Range("A7:A" & last), candidate) = 0 Then NewID = candidate: Exit Function
    Next i
    Err.Raise vbObjectError + 503, , "Tidak dapat membuat ID unik."
End Function

Private Sub CheckMoves()
    Application.CalculateFull
    Require Application.CountIfs(WS("MUTASI").Range("A7:A2006"), "<>", WS("MUTASI").Range("R7:R2006"), "<>OK") = 0, "Masih ada mutasi tidak valid. Periksa kolom Cek."
    Require Application.CountIf(WS("MUTASI").Range("T7:T2006"), "<0") = 0, "Ada saldo historis negatif. Periksa stok, urutan tanggal, dan lot."
End Sub

Public Sub SimpanMutasi()
    Dim f As Worksheet, t As Worksheet, r As Long, locRow As Long, lotRow As Long
    Dim kind As String, location As String, lot As String, note As String, po As String
    Dim qty As Double, price As Double, freight As Double, day As Date, written As Boolean
    On Error GoTo Failed
    Set f = WS("FORM"): Set t = WS("MUTASI")
    day = DayOf(f.Range("B6").Value, "Tanggal mutasi")
    Require day <= Date, "Tanggal mutasi tidak boleh di masa depan."
    kind = Trim$(CStr(f.Range("E6").Value))
    Require InStr(1, "|AWAL|PRODUKSI|KIRIM|JUAL|RETUR|RUSAK|TAMBAH|KURANG|", "|" & kind & "|", vbBinaryCompare) > 0 And Len(kind) > 0, "Pilih jenis mutasi yang valid."
    location = Trim$(CStr(f.Range("B8").Value)): lot = Trim$(CStr(f.Range("E8").Value))
    locRow = FindID(WS("OUTLET"), location, 67): lotRow = FindID(WS("LOT"), lot, 86)
    If kind = "KIRIM" Or kind = "JUAL" Then Require WS("OUTLET").Cells(locRow, 7).Value = "AKTIF", "Outlet NONAKTIF tidak menerima pengiriman/penjualan baru."
    qty = NumberOf(f.Range("B10").Value, "Qty")
    Require qty > 0 And qty = Fix(qty), "Qty harus bilangan bulat positif."
    note = Trim$(CStr(f.Range("B12").Value)): po = Trim$(CStr(f.Range("H8").Value))
    If kind = "KIRIM" Then Require Len(po) > 0, "Isi nomor PO atau surat jalan."
    If kind = "RUSAK" Or kind = "TAMBAH" Or kind = "KURANG" Then Require Len(note) > 0, "Isi alasan mutasi/koreksi."
    If kind = "JUAL" Then
        price = NumberOf(f.Range("E10").Value, "Harga jual saat transaksi")
        Require price > 0, "Harga jual harus positif."
    End If
    If Len(CStr(f.Range("H12").Value)) > 0 Then freight = NumberOf(f.Range("H12").Value, "Ongkir aktual")
    Require freight >= 0, "Ongkir tidak boleh negatif."
    r = EmptyRow(t, LAST_MOVE, "J")
    t.Cells(r, 1).Value = NewID("M", t, LAST_MOVE): written = True
    t.Cells(r, 2).Value = day: t.Cells(r, 3).Value = kind
    t.Cells(r, 4).Value = location: t.Cells(r, 5).Value = lot: t.Cells(r, 6).Value = qty
    If kind = "JUAL" Then t.Cells(r, 7).Value = price
    t.Cells(r, 9).Value = "AKTIF": t.Cells(r, 10).Value = note
    t.Cells(r, 21).Value = po: t.Cells(r, 22).Value = f.Range("H10").Value
    t.Cells(r, 23).Value = freight
    Application.CalculateFull
    Require CStr(t.Cells(r, 18).Value) = "OK", "Mutasi ditolak: " & CStr(t.Cells(r, 18).Value)
    CheckMoves
    written = False
    f.Range("B10,E10,B12,H8,H10,H12").ClearContents
    MsgBox "Mutasi tercatat: " & t.Cells(r, 1).Value & vbCrLf & "Tekan Ctrl+S untuk menyimpan file.", vbInformation, "Gian Fataya"
    Exit Sub
Failed:
    Dim message As String
    message = Err.Description
    If written Then t.Range("A" & r & ":J" & r & ",U" & r & ":W" & r).ClearContents
    Application.CalculateFull
    MsgBox message, vbExclamation, "Mutasi belum disimpan"
End Sub

Public Sub TerbitkanInvoice()
    Dim f As Worksheet, t As Worksheet, m As Worksheet, r As Long, i As Long, areaRow As Long
    Dim area As String, id As String, day As Date, start As Date, finish As Date, due As Date
    Dim discount As Double, freight As Double, gross As Double, chosen As Collection, item As Variant
    Dim written As Boolean, reason As String
    On Error GoTo Failed
    Set f = WS("FORM"): Set t = WS("INVOICE"): Set m = WS("MUTASI")
    Set chosen = New Collection
    day = DayOf(f.Range("B20").Value, "Tanggal terbit")
    start = DayOf(f.Range("B22").Value, "Periode awal")
    finish = DayOf(f.Range("E22").Value, "Periode akhir")
    due = DayOf(f.Range("B24").Value, "Jatuh tempo")
    Require day <= Date And start <= finish And finish <= day And due >= day, "Periksa urutan tanggal dan periode."
    area = Trim$(CStr(f.Range("E20").Value)): areaRow = FindID(WS("WILAYAH"), area, 26)
    Require CStr(WS("WILAYAH").Cells(areaRow, 7).Value) = "OK", "Perbaiki master wilayah."
    discount = NumberOf(f.Range("E24").Value, "Diskon")
    freight = NumberOf(f.Range("B26").Value, "Ongkir ditagih (isi 0 bila tidak ada)")
    Require discount >= 0 And discount <= 1 And freight >= 0, "Diskon harus 0-100%; ongkir tidak negatif."
    CheckMoves
    For i = FIRST_ROW To LAST_MOVE
        If CStr(m.Cells(i, 3).Value) = "JUAL" And CStr(m.Cells(i, 9).Value) = "AKTIF" And CStr(m.Cells(i, 8).Value) = "" And CStr(m.Cells(i, 13).Value) = area Then
            If m.Cells(i, 2).Value >= start And m.Cells(i, 2).Value <= finish Then
                chosen.Add i: gross = gross + CDbl(m.Cells(i, 17).Value)
            End If
        End If
    Next i
    Require chosen.Count > 0 And gross > 0, "Tidak ada penjualan belum ditagih pada wilayah/periode tersebut."
    Require MsgBox("Terbitkan invoice untuk " & chosen.Count & " baris penjualan?" & vbCrLf & "Bruto: Rp" & Format$(gross, "#,##0") & vbCrLf & "Diskon: " & Format$(discount, "0.0%") & vbCrLf & "Ongkir: Rp" & Format$(freight, "#,##0"), vbYesNo + vbQuestion, "Periksa invoice") = vbYes, "Penerbitan dibatalkan."
    r = EmptyRow(t, LAST_INV, "J"): id = NewID("INV", t, LAST_INV)
    t.Cells(r, 1).Value = id: written = True
    t.Cells(r, 2).Value = day: t.Cells(r, 3).Value = area
    t.Cells(r, 4).Value = start: t.Cells(r, 5).Value = finish: t.Cells(r, 6).Value = due
    t.Cells(r, 7).Value = discount: t.Cells(r, 8).Value = freight: t.Cells(r, 9).Value = "DRAF"
    t.Cells(r, 10).Value = gross: t.Cells(r, 18).Value = f.Range("B28").Value
    For Each item In chosen: m.Cells(CLng(item), 8).Value = id: Next item
    Application.CalculateFull
    Require CStr(t.Cells(r, 16).Value) = "OK", "Rekonsiliasi invoice gagal: " & CStr(t.Cells(r, 16).Value)
    t.Cells(r, 9).Value = "TERBIT"
    Application.CalculateFull
    Require CStr(t.Cells(r, 16).Value) = "OK", "Invoice tidak valid setelah diterbitkan."
    written = False
    WS("CETAK INVOICE").Range("B5").Value = id
    WS("CETAK INVOICE").Activate
    MsgBox "Invoice diterbitkan: " & id & vbCrLf & "Tekan Ctrl+S untuk menyimpan file.", vbInformation, "Gian Fataya"
    Exit Sub
Failed:
    reason = Err.Description
    If written Then
        For Each item In chosen
            If CStr(m.Cells(CLng(item), 8).Value) = id Then m.Cells(CLng(item), 8).ClearContents
        Next item
        t.Range("A" & r & ":J" & r).ClearContents: t.Cells(r, 18).ClearContents
    End If
    Application.CalculateFull
    MsgBox reason, vbExclamation, "Invoice belum diterbitkan"
End Sub

Public Sub SimpanPembayaran()
    Dim f As Worksheet, t As Worksheet, inv As Worksheet, ir As Long, r As Long
    Dim id As String, method As String, day As Date, amount As Double, written As Boolean, reason As String
    On Error GoTo Failed
    Set f = WS("FORM"): Set t = WS("PEMBAYARAN"): Set inv = WS("INVOICE")
    Application.CalculateFull
    id = Trim$(CStr(f.Range("E37").Value)): ir = FindID(inv, id, LAST_INV)
    Require inv.Cells(ir, 9).Value = "TERBIT" And inv.Cells(ir, 16).Value = "OK", "Invoice harus TERBIT dan valid."
    day = DayOf(f.Range("B37").Value, "Tanggal pembayaran")
    Require day <= Date And day >= inv.Cells(ir, 2).Value, "Tanggal pembayaran harus sejak tanggal invoice sampai hari ini."
    amount = NumberOf(f.Range("E39").Value, "Nominal pembayaran")
    Require amount > 0 And amount <= inv.Cells(ir, 14).Value, "Nominal harus positif dan tidak melebihi sisa piutang."
    method = Trim$(CStr(f.Range("B39").Value))
    Require Application.CountIf(WS("REFERENSI").Range("O2:O5"), method) = 1, "Pilih metode pembayaran."
    r = EmptyRow(t, LAST_PAY, "G")
    t.Cells(r, 1).Value = NewID("PAY", t, LAST_PAY): written = True
    t.Cells(r, 2).Value = day: t.Cells(r, 3).Value = id: t.Cells(r, 4).Value = method
    t.Cells(r, 5).Value = amount: t.Cells(r, 6).Value = "AKTIF": t.Cells(r, 7).Value = f.Range("B41").Value
    Application.CalculateFull
    Require t.Cells(r, 9).Value = "OK", "Pembayaran ditolak: " & CStr(t.Cells(r, 9).Value)
    written = False
    f.Range("E39,B41").ClearContents
    MsgBox "Pembayaran tercatat. Sisa piutang Rp" & Format$(inv.Cells(ir, 14).Value, "#,##0") & vbCrLf & "Tekan Ctrl+S.", vbInformation, "Gian Fataya"
    Exit Sub
Failed:
    reason = Err.Description
    If written Then t.Range("A" & r & ":G" & r).ClearContents
    Application.CalculateFull
    MsgBox reason, vbExclamation, "Pembayaran belum disimpan"
End Sub

Public Sub EksporInvoicePDF()
    Dim s As Worksheet, t As Worksheet, r As Long, i As Long, file As Variant
    Dim id As String, safeID As String, reason As String, hidden(17 To 66) As Boolean, hiddenSaved As Boolean
    On Error GoTo Failed
    Set s = WS("CETAK INVOICE"): Set t = WS("INVOICE")
    Application.CalculateFull
    id = Trim$(CStr(s.Range("B5").Value)): r = FindID(t, id, LAST_INV)
    CheckMoves
    Require s.Range("A14").Value = "SIAP CETAK", "Belum siap cetak: " & CStr(s.Range("A14").Value)
    safeID = id
    For i = 1 To Len(safeID)
        If InStr(1, "\/:*?""<>|", Mid$(safeID, i, 1), vbBinaryCompare) > 0 Then Mid$(safeID, i, 1) = "_"
    Next i
    file = Application.GetSaveAsFilename(InitialFileName:=safeID & ".pdf", FileFilter:="PDF (*.pdf), *.pdf", Title:="Simpan invoice Gian Fataya")
    If VarType(file) = vbBoolean Then Exit Sub
    If LCase$(Right$(CStr(file), 4)) <> ".pdf" Then file = CStr(file) & ".pdf"
    For i = 17 To 66
        hidden(i) = s.Rows(i).Hidden
    Next i
    hiddenSaved = True
    For i = 17 To 66
        s.Rows(i).Hidden = (Val(s.Cells(i, 4).Value) = 0)
    Next i
    s.ExportAsFixedFormat Type:=xlTypePDF, Filename:=CStr(file), Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    For i = 17 To 66: s.Rows(i).Hidden = hidden(i): Next i
    MsgBox "PDF tersimpan: " & CStr(file), vbInformation, "Gian Fataya"
    Exit Sub
Failed:
    reason = Err.Description
    If hiddenSaved Then
        For i = 17 To 66: s.Rows(i).Hidden = hidden(i): Next i
    End If
    MsgBox reason, vbExclamation, "PDF belum dibuat"
End Sub

Public Sub VoidMutasi()
    Dim t As Worksheet, r As Long, reason As String, oldNote As String, changed As Boolean, message As String
    On Error GoTo Failed
    Set t = WS("MUTASI")
    Require (ActiveSheet Is t), "Pilih baris pada sheet MUTASI."
    r = ActiveCell.Row
    Require r >= FIRST_ROW And r <= LAST_MOVE, "Pilih baris transaksi."
    Require Len(CStr(t.Cells(r, 1).Value)) > 0 And t.Cells(r, 9).Value = "AKTIF", "Pilih transaksi AKTIF."
    Require Len(CStr(t.Cells(r, 8).Value)) = 0, "Batalkan invoice terkait lebih dahulu."
    reason = Trim$(InputBox("Alasan pembatalan mutasi " & t.Cells(r, 1).Value & ":", "Void Mutasi"))
    If Len(reason) = 0 Then Exit Sub
    oldNote = CStr(t.Cells(r, 10).Value)
    Require Len(oldNote) + Len(reason) + 10 < 32767, "Catatan mutasi terlalu panjang."
    t.Cells(r, 9).Value = "VOID": changed = True
    CheckMoves
    t.Cells(r, 10).Value = oldNote & " | VOID: " & reason
    MsgBox "Mutasi dibatalkan. Riwayat tetap disimpan; tekan Ctrl+S.", vbInformation, "Gian Fataya"
    Exit Sub
Failed:
    message = Err.Description
    If changed Then t.Cells(r, 9).Value = "AKTIF": Application.CalculateFull
    MsgBox message, vbExclamation, "Pembatalan gagal"
End Sub

Public Sub VoidInvoice()
    Dim t As Worksheet, m As Worksheet, p As Worksheet, r As Long, i As Long, id As String, reason As String
    Dim changes As Collection, item As Variant, oldState As String, oldNote As String, editing As Boolean, message As String
    On Error GoTo Failed
    Set t = WS("INVOICE"): Set m = WS("MUTASI"): Set p = WS("PEMBAYARAN")
    Require (ActiveSheet Is t), "Pilih baris pada sheet INVOICE."
    r = ActiveCell.Row
    Require r >= FIRST_ROW And r <= LAST_INV, "Pilih baris invoice."
    id = CStr(t.Cells(r, 1).Value)
    Require Len(id) > 0 And t.Cells(r, 9).Value <> "VOID", "Pilih invoice yang belum VOID."
    Require Application.CountIfs(p.Range("C7:C1006"), id, p.Range("F7:F1006"), "AKTIF") = 0, "Batalkan semua pembayaran terkait terlebih dahulu."
    reason = Trim$(InputBox("Alasan pembatalan invoice " & id & ":", "Void Invoice"))
    If Len(reason) = 0 Then Exit Sub
    Set changes = New Collection
    oldState = CStr(t.Cells(r, 9).Value): oldNote = CStr(t.Cells(r, 18).Value)
    Require Len(oldNote) + Len(reason) + 10 < 32767, "Catatan invoice terlalu panjang."
    For i = FIRST_ROW To LAST_MOVE
        If CStr(m.Cells(i, 8).Value) = id Then
            Require Len(CStr(m.Cells(i, 10).Value)) + Len(id) + Len(reason) + 25 < 32767, "Catatan mutasi terlalu panjang."
            changes.Add Array(i, CStr(m.Cells(i, 10).Value))
        End If
    Next i
    editing = True
    For i = FIRST_ROW To LAST_MOVE
        If CStr(m.Cells(i, 8).Value) = id Then
            m.Cells(i, 8).ClearContents
            m.Cells(i, 10).Value = CStr(m.Cells(i, 10).Value) & " | Invoice " & id & " VOID: " & reason
        End If
    Next i
    t.Cells(r, 9).Value = "VOID": t.Cells(r, 18).Value = CStr(t.Cells(r, 18).Value) & " | VOID: " & reason
    Application.CalculateFull
    editing = False
    MsgBox "Invoice VOID. Penjualan dilepas untuk penagihan ulang; tekan Ctrl+S.", vbInformation, "Gian Fataya"
    Exit Sub
Failed:
    message = Err.Description
    If editing Then
        On Error Resume Next
        For Each item In changes
            m.Cells(CLng(item(0)), 8).Value = id
            m.Cells(CLng(item(0)), 10).Value = CStr(item(1))
        Next item
        t.Cells(r, 9).Value = oldState: t.Cells(r, 18).Value = oldNote
        Application.CalculateFull
        On Error GoTo 0
    End If
    MsgBox message, vbExclamation, "Pembatalan gagal"
End Sub

Public Sub VoidPembayaran()
    Dim t As Worksheet, r As Long, reason As String
    On Error GoTo Failed
    Set t = WS("PEMBAYARAN")
    Require (ActiveSheet Is t), "Pilih baris pada sheet PEMBAYARAN."
    r = ActiveCell.Row
    Require r >= FIRST_ROW And r <= LAST_PAY, "Pilih baris pembayaran."
    Require Len(CStr(t.Cells(r, 1).Value)) > 0 And t.Cells(r, 6).Value = "AKTIF", "Pilih pembayaran AKTIF."
    reason = Trim$(InputBox("Alasan pembatalan pembayaran:", "Void Pembayaran"))
    If Len(reason) = 0 Then Exit Sub
    Require Len(CStr(t.Cells(r, 7).Value)) + Len(reason) + 10 < 32767, "Catatan pembayaran terlalu panjang."
    t.Cells(r, 6).Value = "VOID": t.Cells(r, 7).Value = CStr(t.Cells(r, 7).Value) & " | VOID: " & reason
    Application.CalculateFull
    MsgBox "Pembayaran VOID. Piutang dihitung ulang; tekan Ctrl+S.", vbInformation, "Gian Fataya"
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Pembatalan gagal"
End Sub

Private Sub ButtonAt(ByVal sheet As Worksheet, ByVal anchor As String, ByVal title As String, ByVal action As String)
    Dim b As Button
    Set b = sheet.Buttons.Add(sheet.Range(anchor).Left, sheet.Range(anchor).Top, 200, 27)
    b.Name = "gf_" & action: b.Caption = title
    b.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & action
    b.Font.Name = "Calibri": b.Font.Size = 11
End Sub

Public Sub PasangTombol()
    Dim sheet As Worksheet, i As Long
    On Error GoTo Failed
    For Each sheet In ThisWorkbook.Worksheets
        For i = sheet.Shapes.Count To 1 Step -1
            If Left$(sheet.Shapes(i).Name, 3) = "gf_" Then sheet.Shapes(i).Delete
        Next i
    Next sheet
    ButtonAt WS("FORM"), "H6", "SIMPAN MUTASI", "SimpanMutasi"
    ButtonAt WS("FORM"), "H21", "TERBITKAN INVOICE", "TerbitkanInvoice"
    ButtonAt WS("FORM"), "H38", "SIMPAN PEMBAYARAN", "SimpanPembayaran"
    ButtonAt WS("CETAK INVOICE"), "D5", "EKSPOR INVOICE PDF", "EksporInvoicePDF"
    'Void actions use Alt+F8 from the selected ledger row; never placed over data.
    MsgBox "Empat tombol terpasang. Simpan workbook sebagai .xlsm. Untuk pembatalan, pilih baris jurnal lalu Alt+F8 > VoidMutasi / VoidInvoice / VoidPembayaran.", vbInformation, "Gian Fataya"
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Pemasangan tombol gagal"
End Sub
