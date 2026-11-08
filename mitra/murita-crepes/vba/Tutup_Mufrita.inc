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
