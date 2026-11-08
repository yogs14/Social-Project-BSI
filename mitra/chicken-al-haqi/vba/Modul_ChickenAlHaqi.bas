Attribute VB_Name = "Modul_ChickenAlHaqi"
Option Explicit
'Simple single-outlet POS. ASCII source; no drug inventory, CRM or multi-outlet logic.
Private Function S(ByVal title As String) As Worksheet
    Set S = ThisWorkbook.Worksheets(title)
End Function
Private Sub Need(ByVal valid As Boolean, ByVal message As String)
    If Not valid Then Err.Raise vbObjectError + 901, "Chicken Al-Haqi", message
End Sub
Private Function Num(ByVal value As Variant, ByVal title As String) As Double
    Need Not IsError(value), title & " mengandung error."
    Need IsNumeric(value) And Len(CStr(value)) > 0, "Isi angka untuk " & title & "."
    Num = CDbl(value)
End Function
Private Function DayValue(ByVal value As Variant) As Date
    Need Not IsError(value), "Tanggal error."
    Need IsDate(value) And Len(CStr(value)) > 0, "Isi tanggal valid."
    DayValue = DateValue(CDate(value))
    Need DayValue > 0 And DayValue <= Date, "Tanggal harus sampai hari ini."
End Function
Private Function Find(ByVal tab As String, ByVal key As String, ByVal first As Long, ByVal last As Long) As Long
    Dim r As Long
    Need Len(key) > 0, "Pilih kode pada " & tab & "."
    Need Application.CountIf(S(tab).Range("A" & first & ":A" & last), key) = 1, "Kode hilang/ganda: " & key
    For r = first To last
        If StrComp(CStr(S(tab).Cells(r, 1).Value), key, vbTextCompare) = 0 Then Find = r: Exit Function
    Next r
End Function
Private Function FreeRows(ByVal tab As String, ByVal last As Long, ByVal lastInput As String, ByVal count As Long) As Long
    Dim r As Long, j As Long
    For r = 7 To last
        If Len(CStr(S(tab).Cells(r, 1).Value)) = 0 Then
            Need r + count - 1 <= last, "Kapasitas " & tab & " tidak cukup; arsipkan file."
            For j = r To r + count - 1
                Need Application.CountA(S(tab).Range("A" & j & ":" & lastInput & j)) = 0, "Ada baris parsial pada " & tab & " baris " & j
            Next j
            FreeRows = r: Exit Function
        End If
    Next r
    Err.Raise vbObjectError + 902, , "Kapasitas tabel penuh."
End Function
Private Function NewID(ByVal prefix As String, ByVal tab As String, ByVal last As Long) As String
    Dim i As Long, id As String
    For i = 1 To 999
        id = "CAH-" & prefix & "-" & Format$(Now, "yyyymmdd-hhnnss") & "-" & Format$(i, "00")
        If Application.CountIf(S(tab).Range("A7:A" & last), id) = 0 Then NewID = id: Exit Function
    Next i
    Err.Raise vbObjectError + 903, , "Tidak dapat membuat ID unik."
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
Private Sub ValidateAll()
    Dim pair As Variant, r As Long, state As String
    Application.CalculateFull
    For Each pair In Array(Array("NOTA", "O", 1506), Array("ITEM", "M", 6006), Array("BELANJA", "L", 1506), Array("KAS USAHA", "H", 1506), Array("SETUP", "J", 26))
        Need Application.CountIfs(S(CStr(pair(0))).Range("A7:A" & pair(2)), "<>", S(CStr(pair(0))).Range(pair(1) & "7:" & pair(1) & pair(2)), "<>OK") = 0, "Periksa kolom Cek pada " & pair(0)
    Next pair
    Need Application.CountIf(S("SETUP").Range("E31:E34"), "<>OK") = 0, "Periksa data rasa."
    Need Application.CountIfs(S("STOK").Range("A36:A235"), "<>", S("STOK").Range("H36:H235"), "<>OK") = 0, "Periksa koreksi stok."
    For r = 7 To 26
        If Len(CStr(S("STOK").Cells(r, 1).Value)) > 0 Then
            state = CStr(S("STOK").Cells(r, 10).Value)
            Need state = "OK" Or state = "BELUM AWAL", "Periksa stok awal: " & state
        End If
    Next r
    Need Application.CountIf(S("ITEM").Range("O7:O6006"), "<0") = 0, "Ada saldo historis negatif. Periksa stok awal, belanja dan tanggal."
    Need Application.CountIf(S("STOK").Range("J36:J235"), "<0") = 0 And Application.CountIf(S("STOK").Range("H7:H26"), "<0") = 0, "Saldo stok/koreksi negatif."
End Sub
Public Sub PeriksaData()
    On Error GoTo Failed
    ValidateAll
    MsgBox "Data jurnal valid. Jika ada BELUM AWAL, jumlah stok belum dapat diverifikasi. Tekan Ctrl+S untuk menyimpan file.", vbInformation
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Periksa data"
End Sub

Public Sub SimpanTransaksi()
    SimpanNota
End Sub
Public Sub SimpanNota()
    Dim k As Worksheet, rows As Collection, a As Variant, day As Date, hour As Double
    Dim r As Long, mr As Long, fr As Long, selector As Long, qty As Double, price As Double, hpp As Variant, state As String
    Dim gross As Double, discount As Double, received As Double, method As String, cashier As String
    Dim channel As String, customers As Double, code As String, flavor As String
    Dim nr As Long, ir As Long, i As Long, noteID As String, written As Boolean, committed As Boolean, message As String
    On Error GoTo Failed
    Set k = S("KASIR"): Set rows = New Collection: Application.CalculateFull
    day = DayValue(k.Range("C7").Value): Need IsDate(k.Range("C8").Value), "Isi jam transaksi."
    hour = CDbl(TimeValue(CDate(k.Range("C8").Value)))
    cashier = Trim$(CStr(k.Range("C9").Value)): Need Len(cashier) > 0, "Isi nama kasir."
    method = CStr(k.Range("C10").Value)
    Need Application.CountIf(S("SETUP").Range("N2:N5"), method) = 1, "Pilih metode bayar."
    channel = CStr(k.Range("C11").Value)
    Need Application.CountIf(S("SETUP").Range("Q2:Q4"), channel) = 1, "Pilih channel penjualan."
    customers = Num(k.Range("C12").Value, "Jumlah pelanggan")
    Need customers > 0 And customers = Fix(customers), "Jumlah pelanggan harus bulat positif."
    For r = 16 To 27
        If Len(CStr(k.Cells(r, 2).Value)) > 0 Or Len(CStr(k.Cells(r, 3).Value)) > 0 Then
            selector = Find("SETUP", CStr(k.Cells(r, 2).Value), 69, 88)
            code = CStr(S("SETUP").Cells(selector, 2).Value)
            flavor = CStr(S("SETUP").Cells(selector, 3).Value)
            mr = Find("SETUP", code, 7, 26): fr = Find("SETUP", flavor, 31, 34)
            Need code <> "NASI" Or flavor = "Original", "Nasi tanpa tambahan saus."
            Need S("SETUP").Cells(mr, 9).Value = "AKTIF" And S("SETUP").Cells(mr, 10).Value = "OK", "Menu belum valid/aktif."
            qty = Num(k.Cells(r, 3).Value, "Jumlah"): price = Num(k.Cells(r, 4).Value, "Harga jual")
            Need qty > 0 And qty = Fix(qty) And price > 0, "Jumlah harus bulat positif; harga positif."
            hpp = Empty: state = "BELUM DIKONFIRMASI"
            If S("SETUP").Cells(mr, 7).Value = "YA" And S("SETUP").Cells(fr, 4).Value = "YA" Then
                hpp = Num(S("SETUP").Cells(mr, 6).Value, "HPP dasar") + Num(S("SETUP").Cells(fr, 3).Value, "HPP rasa")
                Need hpp >= 0, "HPP tidak boleh negatif.": state = "DIKONFIRMASI"
            End If
            rows.Add Array(code, flavor, qty, price, hpp, state)
            gross = gross + qty * price
        End If
    Next r
    Need rows.Count > 0, "Keranjang kosong."
    discount = Num(k.Range("C30").Value, "Diskon (isi 0 bila tidak ada)")
    Need discount >= 0 And discount <= gross, "Diskon harus antara nol dan bruto."
    If method = "Tunai" Then
        received = Num(k.Range("C32").Value, "Uang diterima"): Need received >= gross - discount, "Uang tunai kurang."
    Else
        received = gross - discount
    End If
    nr = FreeRows("NOTA", 1506, "I", 1): ir = FreeRows("ITEM", 6006, "H", rows.Count)
    Need Application.CountA(S("NOTA").Range("Q" & nr & ":R" & nr)) = 0, "Ada metadata nota parsial."
    noteID = NewID("N", "NOTA", 1506): written = True
    Put "NOTA", nr, Array(noteID, day, hour, cashier, method, discount, received, "AKTIF", "Kasir menu + jumlah")
    S("NOTA").Cells(nr, 17).NumberFormat = "@": S("NOTA").Cells(nr, 17).Value = channel
    S("NOTA").Cells(nr, 18).Value = customers
    For Each a In rows
        Put "ITEM", ir + i, Array(noteID & "-I" & Format$(i + 1, "00"), noteID, a(0), a(1), a(2), a(3), a(4), a(5))
        i = i + 1
    Next a
    ValidateAll: written = False: committed = True
    NotaBaru
    MsgBox "Nota " & noteID & " tercatat. Tekan Ctrl+S. Stok tanpa saldo awal tetap BELUM AWAL; HPP belum dikonfirmasi tidak ditampilkan sebagai laba pasti.", vbInformation, "Chicken Al-Haqi"
    Exit Sub
Failed:
    message = Err.Description
    If committed Then MsgBox "Nota SUDAH tercatat. Jangan simpan ulang. Gagal membersihkan form: " & message, vbExclamation: Exit Sub
    If written Then
        S("NOTA").Range("A" & nr & ":I" & nr).ClearContents
        S("NOTA").Range("Q" & nr & ":R" & nr).ClearContents
        S("ITEM").Range("A" & ir & ":H" & ir + rows.Count - 1).ClearContents
    End If
    Application.CalculateFull: MsgBox message, vbExclamation, "Nota belum disimpan"
End Sub
Public Sub NotaBaru()
    Dim r As Long
    With S("KASIR")
        .Range("B16:C27,C32").ClearContents
        .Range("C30").Value = 0
        For r = 16 To 27
            .Cells(r, 4).Formula = "=IF(B" & r & "="""","""",IFERROR(INDEX(SETUP!$D$69:$D$88,MATCH(B" & r & ",SETUP!$A$69:$A$88,0)),0))"
        Next r
    End With
End Sub
Public Sub TutupKasir()
    S("TUTUP KASIR").Activate
    Application.CalculateFull
    S("TUTUP KASIR").Range("K8").Select
End Sub
Public Sub VoidNotaTerakhir()
    BatalNota
End Sub
Public Sub SimpanBelanja()
    Dim k As Worksheet, day As Date, code As String, mr As Long, kg As Variant, qty As Double, amount As Double
    Dim r As Long, written As Boolean, committed As Boolean, message As String
    On Error GoTo Failed
    Set k = S("KASIR"): day = DayValue(k.Range("I53").Value)
    code = CStr(k.Range("I54").Value): mr = Find("SETUP", code, 7, 26)
    If Len(CStr(k.Range("I55").Value)) > 0 Then kg = Num(k.Range("I55").Value, "Kg")
    qty = Num(k.Range("I56").Value, "Pcs/porsi aktual"): amount = Num(k.Range("I57").Value, "Nominal belanja")
    Need qty > 0 And qty = Fix(qty) And amount >= 0, "Qty bulat positif; nominal tidak negatif."
    r = FreeRows("BELANJA", 1506, "I", 1): written = True
    Put "BELANJA", r, Array(NewID("B", "BELANJA", 1506), day, code, kg, qty, amount, k.Range("I58").Value, "AKTIF", "Qty aktual; kg hanya pembanding")
    ValidateAll: written = False: committed = True
    k.Range("I55:I57").ClearContents
    MsgBox "Belanja tercatat memakai pcs aktual saja. Tekan Ctrl+S.", vbInformation
    Exit Sub
Failed:
    message = Err.Description
    If committed Then MsgBox "Belanja SUDAH tercatat. Jangan simpan ulang. " & message, vbExclamation: Exit Sub
    If written Then S("BELANJA").Range("A" & r & ":I" & r).ClearContents
    Application.CalculateFull: MsgBox message, vbExclamation, "Belanja belum disimpan"
End Sub
Public Sub SimpanKas()
    Dim k As Worksheet, day As Date, amount As Double, r As Long, written As Boolean, committed As Boolean, message As String
    On Error GoTo Failed
    Set k = S("KASIR"): day = DayValue(k.Range("I65").Value): amount = Num(k.Range("I67").Value, "Nominal kas")
    Need amount > 0, "Nominal harus positif."
    r = FreeRows("KAS USAHA", 1506, "G", 1): written = True
    Put "KAS USAHA", r, Array(NewID("K", "KAS USAHA", 1506), day, k.Range("I66").Value, amount, k.Range("I68").Value, k.Range("I69").Value, "AKTIF")
    ValidateAll: written = False: committed = True
    k.Range("I67,I69").ClearContents
    MsgBox "Arus uang ayam tercatat. Tekan Ctrl+S.", vbInformation: Exit Sub
Failed:
    message = Err.Description
    If committed Then MsgBox "Kas SUDAH tercatat. Jangan simpan ulang. " & message, vbExclamation: Exit Sub
    If written Then S("KAS USAHA").Range("A" & r & ":G" & r).ClearContents
    Application.CalculateFull: MsgBox message, vbExclamation, "Kas belum disimpan"
End Sub
Public Sub BatalNota()
    Dim key As String, r As Long, reason As String, changed As Boolean, message As String, latest As String, lastRow As Long
    On Error GoTo Failed
    lastRow = S("NOTA").Cells(S("NOTA").Rows.Count, 1).End(xlUp).Row
    If lastRow >= 7 Then latest = CStr(S("NOTA").Cells(lastRow, 1).Value)
    key = Trim$(InputBox("Nomor nota yang dibatalkan (default: nota terakhir):", "Batal Nota", latest)): If Len(key) = 0 Then Exit Sub
    r = Find("NOTA", key, 7, 1506)
    Need S("NOTA").Cells(r, 8).Value = "AKTIF", "Nota sudah VOID."
    reason = Trim$(InputBox("Alasan pembatalan " & key & ":", "Batal Nota")): If Len(reason) = 0 Then Exit Sub
    Need Len(CStr(S("NOTA").Cells(r, 9).Value)) + Len(reason) + 10 < 32767, "Catatan terlalu panjang."
    S("NOTA").Cells(r, 8).Value = "VOID": changed = True: ValidateAll
    S("NOTA").Cells(r, 9).Value = CStr(S("NOTA").Cells(r, 9).Value) & " | VOID: " & reason
    MsgBox "Nota VOID; stok/penjualan dihitung ulang. Laksanakan refund aktual bila diperlukan dan dokumentasikan. Tekan Ctrl+S.", vbInformation
    Exit Sub
Failed:
    message = Err.Description
    If changed Then S("NOTA").Cells(r, 8).Value = "AKTIF"
    Application.CalculateFull: MsgBox message, vbExclamation, "Pembatalan gagal"
End Sub
Private Sub AddButton(ByVal anchor As String, ByVal title As String, ByVal action As String)
    Dim b As Button
    With S("KASIR")
        Set b = .Buttons.Add(.Range(anchor).Left, .Range(anchor).Top, 170, 27)
        b.Name = "cah_" & action: b.Caption = title: b.Font.Name = "Calibri"
        b.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & action
    End With
End Sub
Public Sub PasangTombol()
    Dim i As Long
    On Error GoTo Failed
    With S("KASIR")
        For i = .Shapes.Count To 1 Step -1
            If Left$(.Shapes(i).Name, 4) = "cah_" Then .Shapes(i).Delete
        Next i
    End With
    AddButton "B43", "SIMPAN TRANSAKSI", "SimpanTransaksi"
    AddButton "E43", "NOTA BARU", "NotaBaru"
    AddButton "B45", "VOID NOTA", "VoidNotaTerakhir"
    AddButton "E45", "TUTUP KASIR", "TutupKasir"
    AddButton "I62", "SIMPAN BELANJA", "SimpanBelanja"
    AddButton "I71", "SIMPAN KAS", "SimpanKas"
    MsgBox "Enam tombol terpasang. Simpan .xlsm. Untuk memeriksa jurnal gunakan Alt+F8 > PeriksaData.", vbInformation
    Exit Sub
Failed:
    MsgBox Err.Description, vbExclamation, "Pemasangan tombol gagal"
End Sub
