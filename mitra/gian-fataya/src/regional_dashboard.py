"""Prominent, live regional charts inside the existing BERANDA sheet."""
from copy import copy

from openpyxl.chart import AreaChart, BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, Font as DrawingFont
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

TEAL, AMBER, GREY = "087F8C", "C18A20", "71808C"
PALE, INK, MUTED, BLUE = "F0F5F7", "203B4A", "526674", "EDF5FC"


def chart_text(size=11):
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(
        sz=size * 100, latin=DrawingFont(typeface="Calibri"))), endParaRPr=CharacterProperties(lang="id-ID"))])


def configure(book, areas, band, link):
    ws = book["BERANDA"]
    # Move the detail table and alerts as a unit; accounting input and VBA form
    # coordinates stay stable. Only local references in the table need moving.
    snapshots = {}
    for old, new in [(r, r + 12) for r in range(24, 45)] + [(19, 60), (21, 62)]:
        for c in range(1, 13):
            cell = ws.cell(old, c)
            value = cell.value
            if isinstance(value, str) and value.startswith("=") and 24 <= old <= 44:
                value = value.replace(f"'BERANDA'!A{old}", f"'BERANDA'!A{new}")
            snapshots[(new, c)] = (value, copy(cell._style))
    for merged in list(ws.merged_cells.ranges):
        if 17 <= merged.min_row and merged.max_row <= 65:
            ws.unmerge_cells(str(merged))
    for r in range(17, 66):
        ws.row_dimensions[r].height = 22
        for c in range(1, 13):
            ws._cells.pop((r, c), None)
    for (r, c), (value, style) in snapshots.items():
        ws.cell(r, c, value)._style = style
    for key in list(ws.conditional_formatting._cf_rules):
        if all(17 <= cell_range.min_row <= cell_range.max_row <= 44 for cell_range in key.sqref.ranges):
            del ws.conditional_formatting._cf_rules[key]

    ws["G5"] = "Tanggal biru = periode perbandingan wilayah"
    ws["G5"].font = Font(name="Calibri", size=10, color=MUTED)
    link(ws, "A16", "1. Catat barang", "TRANSAKSI", "A6", "C16")
    link(ws, "E16", "2. Buat tagihan", "TAGIHAN", "A6", "G16")
    link(ws, "I16", "3. Catat pembayaran", "TAGIHAN", "A436", "L16")
    band(ws, "A17:F17", "01  PERBANDINGAN PENJUALAN WILAYAH")
    band(ws, "G17:L17", "02  TREN PENJUALAN WILAYAH")
    band(ws, "A18:F18", '="Periode: "&TEXT(TglAwal,"dd mmm yyyy")&" – "&TEXT(TglAkhir,"dd mmm yyyy")', PALE, INK, 10, False)
    band(ws, "G18:H18", "Pilih wilayah →", PALE, INK, 10)
    band(ws, "I18:L18", book["DATA USAHA"]["B7"].value, BLUE, "1668A8", 11)
    book.defined_names.add(DefinedName("NamaWilayahGrafik", attr_text=f"'DATA USAHA'!$B$7:$B${6 + areas}"))
    selector = DataValidation(type="list", formula1="NamaWilayahGrafik", allow_blank=False)
    selector.showErrorMessage, selector.errorStyle = True, "stop"
    selector.errorTitle, selector.error = "Pilih wilayah", "Gunakan nama wilayah dari daftar."
    selector.showInputMessage, selector.promptTitle = True, "Grafik tren bulanan"
    selector.prompt = "Pilih salah satu wilayah. Grafik mengikuti tanggal akhir di E5."
    ws.add_data_validation(selector)
    selector.add("I18")
    band(ws, "A19:F19", '=COUNTIF(H37:H56,"TINGGI")&" tinggi  •  "&COUNTIF(H37:H56,"SEDANG")&" sedang  •  "&COUNTIF(H37:H56,"DATAR")&" datar"', "FFFFFF", INK, 10, False)
    band(ws, "G19:L19", "12 bulan hingga tanggal akhir; bulan terakhir dapat belum penuh.", "FFFFFF", MUTED, 10, False)

    ws["H36"] = "Performa"
    ws["H36"]._style = copy(ws["G36"]._style)
    ws.row_dimensions[36].height = 34
    for r in range(37, 57):
        ws[f"H{r}"] = (f'=IF(A{r}="","",IF(D{r}=0,"DATAR",'
                       f'IF(D{r}>=MAX($D$37:$D$56)*\'DATA USAHA\'!$B$276,"TINGGI","SEDANG")))')
        ws[f"H{r}"]._style = copy(ws[f"B{r}"]._style)
        ws[f"H{r}"].alignment = Alignment(indent=1, vertical="center")
        ws[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[r].height = 28
    for label, bg, fg in [("TINGGI", "E0F2F0", TEAL), ("SEDANG", "FFF2D5", "80590D"), ("DATAR", "EDF0F2", GREY)]:
        ws.conditional_formatting.add("H37:H56", CellIsRule(operator="equal", formula=[f'"{label}"'],
            fill=PatternFill("solid", fgColor=bg), font=Font(name="Calibri", bold=True, color=fg)))
    # Collapse spare capacity, while keeping all official areas visible.
    if areas < 20:
        ws.row_dimensions.group(37 + areas, 56, outline_level=1, hidden=True)
        ws.row_dimensions[57].collapsed = True

    # Concrete A1 chart sources survive Excel/LibreOffice recalculation. Hidden
    # helper columns are explicitly included by plotVisOnly=False on both charts.
    ws.column_dimensions.group("N", "Q", outline_level=1, hidden=True)
    ws["N2"], ws["O2"] = "Kode wilayah tren", '=IFERROR(INDEX(\'DATA USAHA\'!$A$7:$A$26,MATCH($I$18,\'DATA USAHA\'!$B$7:$B$26,0)),"")'
    ws["N5"], ws["O5"], ws["P5"] = "Awal bulan", '="Penjualan bruto • "&$I$18', "Bulan"
    for r in range(6, 18):
        ws[f"N{r}"] = f'=DATE(YEAR(TglAkhir),MONTH(TglAkhir)-{17-r},1)'
        ws[f"N{r}"].number_format = "mmm yyyy"
        ws[f"P{r}"] = f'=TEXT(N{r},"mmm yy")'
        ws[f"O{r}"] = (f'=IF($O$2="",0,SUMIFS(TRANSAKSI!$W$7:$W$2006,'
            f'TRANSAKSI!$S$7:$S$2006,$O$2,TRANSAKSI!$B$7:$B$2006,">="&N{r},'
            f'TRANSAKSI!$B$7:$B$2006,"<"&DATE(YEAR(N{r}),MONTH(N{r})+1,1),'
            'TRANSAKSI!$B$7:$B$2006,"<="&TglAkhir))')
        ws[f"O{r}"].number_format = '#,##0'
    for c, text in [("N", "Wilayah / kategori"), ("O", "TINGGI"), ("P", "SEDANG"), ("Q", "DATAR")]:
        ws[f"{c}20"] = text
    for r in range(21, 21 + areas):
        table_row = r + 16
        ws[f"N{r}"] = f'=B{table_row}&" · "&H{table_row}&IF(D{table_row}=0," (0)","")'
        for c, category in [("O", "TINGGI"), ("P", "SEDANG"), ("Q", "DATAR")]:
            ws[f"{c}{r}"] = f'=IF(H{table_row}="{category}",D{table_row},0)'
            ws[f"{c}{r}"].number_format = '#,##0;;;'

    chart = BarChart()
    chart.type, chart.grouping, chart.overlap = "bar", "stacked", 100
    chart.title = "Penjualan bruto per wilayah"
    chart.x_axis.title, chart.y_axis.title = "Wilayah / kategori", "Penjualan (Rp)"
    chart.add_data(Reference(ws, min_col=15, max_col=17, min_row=20, max_row=20 + areas), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=14, min_row=21, max_row=20 + areas))
    chart.x_axis.scaling.orientation, chart.x_axis.tickLblSkip = "maxMin", 1
    chart.y_axis.crosses, chart.y_axis.scaling.min = "max", 0
    chart.y_axis.numFmt = '#,##0'
    chart.legend.position = "b"
    chart.dLbls = DataLabelList(showVal=True, showCatName=False, showSerName=False,
                               showPercent=False, showLegendKey=False, numFmt='#,##0;;;', dLblPos="ctr")
    chart.dLbls.txPr = chart_text(10)
    for series, color in zip(chart.series, [TEAL, AMBER, GREY]):
        series.graphicalProperties.solidFill = color
        series.graphicalProperties.line.noFill = True

    trend = AreaChart()
    trend.title = "Tren penjualan bruto bulanan"
    trend.x_axis.title, trend.y_axis.title = "Bulan", "Penjualan (Rp)"
    trend.y_axis.scaling.min, trend.y_axis.numFmt = 0, '#,##0'
    trend.add_data(Reference(ws, min_col=15, min_row=5, max_row=17), titles_from_data=True)
    trend.set_categories(Reference(ws, min_col=16, min_row=6, max_row=17))
    trend.series[0].graphicalProperties.solidFill = "CFE8EC"
    trend.series[0].graphicalProperties.line.solidFill = TEAL
    trend.series[0].graphicalProperties.line.width = 24000
    trend.legend.position = "b"
    trend.x_axis.tickLblSkip = 2
    for plot, anchor in [(chart, "A20"), (trend, "G20")]:
        plot.width, plot.height = 20.2, 9.6
        plot.visible_cells_only = False
        plot.graphical_properties = GraphicalProperties(solidFill="FFFFFF")
        plot.graphical_properties.line.noFill = True
        plot.txPr = chart_text(11)
        plot.x_axis.txPr = chart_text(10)
        plot.y_axis.txPr = chart_text(10)
        plot.legend.txPr = chart_text(10)
        ws.add_chart(plot, anchor)

    band(ws, "A33:F34", '=IF(SUM(D37:D56)=0,"Belum ada penjualan valid pada periode ini. Semua wilayah DATAR; grafik akan terisi setelah JUAL dicatat.","Panjang batang = bruto penjualan. Nama wilayah memuat kategori; angka menunjukkan rupiah.")', PALE, INK, 10, False)
    band(ws, "G33:L34", '=IF(SUM(O6:O17)=0,"Belum ada penjualan valid untuk wilayah ini dalam 12 bulan terakhir.","Tren "&I18&": "&TEXT(N6,"mmm yyyy")&" – "&TEXT(TglAkhir,"dd mmm yyyy")&". Bulan tanpa JUAL bernilai 0.")', PALE, INK, 10, False)
    band(ws, "A35:H35", "RINCIAN PER WILAYAH  /  sesuai periode perbandingan")
    band(ws, "I35:L35", "CARA MEMBACA GRAFIK")
    band(ws, "I37:L39", '="TINGGI ≥ "&TEXT(\'DATA USAHA\'!B276,"0%")&" dari area teratas."&CHAR(10)&"SEDANG: positif di bawah ambang."&CHAR(10)&"DATAR: belum ada penjualan (0)."', PALE, INK, 11, False)
    band(ws, "I40:L42", "Kategori membandingkan bruto antarwilayah, bukan pertumbuhan. KIRIM belum menjadi penjualan. Toko independen ikut wilayahnya.", PALE, INK, 11, False)
    link(ws, "I43", "Ubah ambang kategori", "DATA USAHA", "A274", "L43")
    link(ws, "I44", "Lihat stok per outlet", "STOK", "A6", "L44")
    band(ws, "A57:H57", "PO minimum 10 pcs per nomor PO + outlet; periksa sebelum penagihan.", PALE, MUTED, 10, False)
    ws["I57"] = '=COUNTIF(TRANSAKSI!K7:K2006,"PO < 10 PCS")'
    band(ws, "J57:L57", "baris PO perlu dilengkapi", "FFFFFF", MUTED, 10, False)
    band(ws, "A59:L59", "PERIKSA SEBELUM MENGIRIM TAGIHAN")
    for r in [60, 62]:
        for start, end, value in [("A", "B", "C"), ("E", "F", "G"), ("I", "J", "K")]:
            ws.merge_cells(f"{start}{r}:{end}{r}")
            ws.conditional_formatting.add(f"{value}{r}", CellIsRule(operator="greaterThan", formula=["0"],
                fill=PatternFill("solid", fgColor="FFF2D5"), font=Font(color="80590D")))
    link(ws, "A64", "Lengkapi data usaha", "DATA USAHA", "A6", "D64")
    link(ws, "E64", "Input cepat (opsional)", "BERANDA", "A69", "H64")
    link(ws, "I64", "Panduan grafik: sel biru", "BERANDA", "I18", "L64")
    ws.sheet_view.topLeftCell = "A1"
    ws.sheet_view.selection[0].activeCell = "B5"
    ws.sheet_view.selection[0].sqref = "B5"
