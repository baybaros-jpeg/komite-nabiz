# -*- coding: utf-8 -*-
"""TUPRS (Tupras) kapsamli yatirim analizi calisma kitabi.
Veri kaynagi: borsa MCP konnektoru (yfinance / borsapy / isyatirim / TCMB / TradingView)
Veri tarihi: 2026-09-02
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

FONT = "Arial"
TITLE = Font(name=FONT, size=14, bold=True, color="FFFFFF")
HDR   = Font(name=FONT, size=10, bold=True, color="FFFFFF")
SUB   = Font(name=FONT, size=10, bold=True)
BODY  = Font(name=FONT, size=10)
BLUE  = Font(name=FONT, size=10, color="0000FF")          # girdi / varsayim
GREEN = Font(name=FONT, size=10, color="008000")          # baska sayfadan link
NOTE  = Font(name=FONT, size=9, italic=True, color="595959")

FILL_T = PatternFill("solid", fgColor="1F3864")
FILL_H = PatternFill("solid", fgColor="2F5597")
FILL_S = PatternFill("solid", fgColor="D9E2F3")
FILL_Y = PatternFill("solid", fgColor="FFFF00")
FILL_G = PatternFill("solid", fgColor="C6EFCE")
FILL_R = PatternFill("solid", fgColor="FFC7CE")
FILL_O = PatternFill("solid", fgColor="FFE699")

THIN = Side(style="thin", color="BFBFBF")
BOX  = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MN  = '#,##0;(#,##0);-'      # milyon TL
TL2 = '#,##0.00;(#,##0.00);-'
PCT = '0.0%;(0.0%);-'
MUL = '0.0"x"'

wb = Workbook()

def title_row(ws, text, ncols, row=1):
    ws.cell(row=row, column=1, value=text)
    ws.cell(row=row, column=1).font = TITLE
    for c in range(1, ncols + 1):
        ws.cell(row=row, column=c).fill = FILL_T
    ws.row_dimensions[row].height = 22

def header_row(ws, row, values, start=1):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=start + i, value=v)
        c.font = HDR; c.fill = FILL_H; c.border = BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 30

def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w

def put(ws, row, col, value, font=BODY, fmt=None, fill=None, align=None, border=True):
    c = ws.cell(row=row, column=col, value=value)
    c.font = font
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if align: c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=(align == "left"))
    if border: c.border = BOX
    return c

# ============================================================ 1) OZET
ws = wb.active
ws.title = "1-Ozet"
widths(ws, {"A": 42, "B": 20, "C": 20, "D": 20, "E": 52})
title_row(ws, "TÜPRAŞ (TUPRS) — YATIRIM ANALİZİ ÖZETİ  |  Veri tarihi: 02.09.2026", 5)
ws["A2"] = "Kaynak: borsa MCP konnektörü (yfinance / borsapy / İş Yatırım / TCMB / TradingView) + KAP & basın taraması"
ws["A2"].font = NOTE
ws.merge_cells("A2:E2")

r = 4
header_row(ws, r, ["Gösterge", "Değer", "Karşılaştırma", "Birim / Dönem", "Yorum"])
summary = [
    ("Kapanış fiyatı",                392.25,  402.75, "TL (02.09 / 01.09)", "Gün içi -%2,6; 52 hafta zirvesi 417,25 TL'nin %6 altında"),
    ("Piyasa değeri",                 755786,  None,   "mn TL",              "1.926.795.598 adet pay"),
    ("52 hafta aralığı (düşük)",      117.30,  None,   "TL",                 "Dipten +%234"),
    ("52 hafta aralığı (yüksek)",     417.25,  None,   "TL",                 "01.09.2026'da görüldü"),
    ("F/K (TTM)",                     10.40,   14.15,  "x — vs XUSIN medyan","Sektöre göre %26 iskontolu"),
    ("PD/DD",                         1.70,    1.70,   "x — vs XUSIN medyan","Sektör medyanı ile aynı"),
    ("FD/FAVÖK (TTM)",                5.60,    None,   "x",                  "Tarihsel ortalamanın altında, ancak zirve FAVÖK üzerinden"),
    ("FD/Satışlar",                   0.91,    None,   "x",                  "Rafinericilikte düşük marj yapısı nedeniyle normal"),
    ("Temettü verimi (TTM)",          0.0626,  None,   "%",                  "2026 toplam 33 mlr TL; brüt 17,13 TL/pay ≈ %4,4 + Mart taksiti"),
    ("Hisse başı kazanç (TTM)",       36.04,   None,   "TL",                 "2Ç26'da 19,49 TL beklentiye karşı 36,04 TL gerçekleşti"),
    ("Analist konsensüsü",            "AL",    None,   "20 AL / 7 TUT / 0 SAT","yfinance konsensüs"),
    ("Ortalama hedef fiyat",          396.98,  392.25, "TL — vs spot",       "Yükseliş potansiyeli ≈ %1,2 (yurt içi 16 kurum ort. 405,44 TL)"),
    ("DCF içsel değer (Buffett mdl.)",1022.31, 392.25, "TL — vs spot",       "Güvenlik marjı %61,6 — ancak zirve nakit akışını ekstrapole eder"),
    ("RSI (14, günlük)",              67.29,   70.00,  "— vs aşırı alım",    "Aşırı alım sınırına yakın"),
    ("Brent",                         94.33,   None,   "USD/varil",          "Yıl başından bu yana +%55 (USD)"),
    ("USD/TRY",                       48.2913, None,   "TL",                 "YBB +%12,3"),
    ("TÜFE (yıllık, Tem-26)",         0.3175,  None,   "%",                  "Reel getiri eşiği yüksek"),
]
r += 1
for name, val, cmp_, unit, note in summary:
    put(ws, r, 1, name, SUB, align="left")
    fmt = None
    if isinstance(val, float) and val < 1: fmt = PCT
    elif isinstance(val, (int, float)) and abs(val) > 10000: fmt = MN
    elif isinstance(val, float): fmt = TL2
    put(ws, r, 2, val, BODY, fmt, align="center")
    put(ws, r, 3, cmp_, BODY, fmt if isinstance(cmp_, float) and cmp_ < 1 else (MN if isinstance(cmp_, (int, float)) and abs(cmp_) > 10000 else TL2), align="center")
    put(ws, r, 4, unit, BODY, align="center")
    put(ws, r, 5, note, BODY, align="left")
    r += 1

r += 1
put(ws, r, 1, "SONUÇ / KANAAT", TITLE, fill=FILL_T); 
for c in range(2, 6): ws.cell(row=r, column=c).fill = FILL_T
r += 1
verdicts = [
    ("Kısa vade (0-3 ay)", "NÖTR / BEKLE", FILL_O,
     "2 ayda +%73 yükseliş, RSI 67, konsensüs hedefi (397 TL) fiyatın üzerinde değil. Koç Holding blok satışı (384,50 TL, ~6,16 mlr TL, takas 07.09) kısa vadeli arz baskısı ve fiyat çıpası yaratıyor. 30.09 temettü kesintisi (brüt 6,75 TL) mekanik düşüş getirecek. Yeni alım için 365-385 TL bandı beklenmeli."),
    ("Orta vade (6-18 ay)", "KADEMELİ AL", FILL_G,
     "Net rafineri marjı beklentisi 6-7 $/varilden 13-15 $/varile yükseltildi; 2026 FAVÖK tahmini 2,4 mlr $'dan 3,1 mlr $'a çıktı. Net nakit pozisyonu, %7,7 temettü verimi ve 10,4 F/K (sektör 14,15) orta vadede destekleyici. Ancak marjlar savaş/arz kesintisi kaynaklı — normalleşme senaryosu fiyatlanmalı."),
    ("Ana risk", "MARJ NORMALLEŞMESİ", FILL_R,
     "Mevcut fiyat, orta-çevrim (~8-10 $/varil) yerine zirve marjı fiyatlıyor. Senaryo sayfasına göre 2,0 mlr $ normalize FAVÖK'te adil değer ~342 TL (spot'un %13 altında)."),
]
for lbl, verdict, fill, txt in verdicts:
    put(ws, r, 1, lbl, SUB, align="left")
    put(ws, r, 2, verdict, Font(name=FONT, size=10, bold=True), fill=fill, align="center")
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
    put(ws, r, 3, txt, BODY, align="left")
    ws.row_dimensions[r].height = 58
    r += 1

r += 1
put(ws, r, 1, "UYARI: Bu çalışma yatırım danışmanlığı değildir. Kamuya açık verilerden derlenmiş analiz içerir; yatırım kararları kişisel risk profilinize göre alınmalıdır.", NOTE, fill=FILL_Y, align="left", border=False)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)

# ============================================================ 2) BILANCO
ws = wb.create_sheet("2-Bilanco")
widths(ws, {"A": 40, "B": 17, "C": 17, "D": 17, "E": 17, "F": 15, "G": 15})
title_row(ws, "TÜPRAŞ — ÖZET BİLANÇO (milyon TL, çeyrek sonu, TMS-29 enflasyon muhasebesi uygulanmış)", 7)
ws["A2"] = "Kaynak: borsa MCP / borsapy — KAP finansal tabloları. Kalemler cari alım gücüne endekslendiğinden dönemler arası karşılaştırma yaklaşıktır."
ws["A2"].font = NOTE; ws.merge_cells("A2:G2")

r = 4
header_row(ws, r, ["Kalem", "2026/Q2", "2026/Q1", "2025/Q4", "2025/Q2", "Ç/Ç Değişim", "Y/Y Değişim"])
bs = [
    ("Nakit ve Nakit Benzerleri",        195840.940, 129900.408, 126280.900,  90351.730),
    ("Diğer Dönen Varlıklar",             64043.307,  45612.602,  16687.779,  27112.428),
    ("DÖNEN VARLIKLAR",                  439480.874, 315176.951, 279006.483, 229193.632),
    ("Maddi Duran Varlıklar",            350569.987, 325672.138, 348859.596, 265376.996),
    ("Maddi Olmayan Duran Varlıklar",     11226.804,  10789.887,  11924.025,   8833.094),
    ("Özkaynak Yöntemiyle Değ. Yatırımlar",19800.794,  18000.193,  20573.159,  14815.375),
    ("DURAN VARLIKLAR",                  420506.642, 389568.508, 418092.864, 316436.625),
    ("TOPLAM VARLIKLAR",                 859987.516, 704745.459, 697099.347, 545630.257),
    ("KISA VADELİ YÜKÜMLÜLÜKLER",        325369.321, 286503.719, 199892.886, 187463.553),
    ("  Diğer Kısa Vadeli Yükümlülükler", 60957.559,  99595.641,  45133.347,  48613.586),
    ("UZUN VADELİ YÜKÜMLÜLÜKLER",         71311.414,  58163.199,  61702.669,  46678.151),
    ("  Uzun Vadeli Finansal Borçlar",    36443.100,  35817.792,  33477.291,  31975.783),
    ("  Ertelenmiş Vergi Yükümlülüğü",    30853.182,  18621.056,  24098.265,  12079.385),
    ("ÖZKAYNAKLAR",                      463306.781, 360078.541, 435503.792, 311488.553),
    ("  Ana Ortaklığa Ait Özkaynaklar",  456139.125, 353670.740, 428757.720, 306637.509),
    ("  Geçmiş Yıllar Kar/Zararları",    276784.892, 258644.457, 277100.624, 209751.941),
    ("  Dönem Net Karı (kümülatif)",      49847.684,   3709.759,  34766.049,   8986.698),
    ("  Ödenmiş Sermaye",                  1926.796,   1926.796,   1926.796,   1926.796),
    ("Net Yabancı Para Pozisyonu",      -186333.514,-102097.017, -97459.871, -99317.532),
]
r += 1
first = r
for name, q2, q1, y4, py in bs:
    bold = name.isupper() or name.startswith("TOPLAM")
    f = SUB if bold else BODY
    put(ws, r, 1, name, f, align="left")
    for i, v in enumerate([q2, q1, y4, py]):
        put(ws, r, 2 + i, v, f, MN, align="right")
    put(ws, r, 6, f"=IF(C{r}=0,\"\",B{r}/C{r}-1)", f, PCT, align="right")
    put(ws, r, 7, f"=IF(E{r}=0,\"\",B{r}/E{r}-1)", f, PCT, align="right")
    if bold:
        for c in range(1, 8): ws.cell(row=r, column=c).fill = FILL_S
    r += 1

r += 1
put(ws, r, 1, "TÜRETİLMİŞ GÖSTERGELER", HDR, fill=FILL_H, align="left")
for c in range(2, 8): ws.cell(row=r, column=c).fill = FILL_H
r += 1
derived = [
    ("Cari Oran (Dönen V. / KV Yük.)",  f"=B{first+2}/B{first+8}",  f"=C{first+2}/C{first+8}",  f"=D{first+2}/D{first+8}",  f"=E{first+2}/E{first+8}",  MUL),
    ("Nakit / Toplam Varlık",           f"=B{first}/B{first+7}",    f"=C{first}/C{first+7}",    f"=D{first}/D{first+7}",    f"=E{first}/E{first+7}",    PCT),
    ("Özkaynak / Toplam Varlık",        f"=B{first+13}/B{first+7}", f"=C{first+13}/C{first+7}", f"=D{first+13}/D{first+7}", f"=E{first+13}/E{first+7}", PCT),
    ("Toplam Yükümlülük / Özkaynak",    f"=(B{first+8}+B{first+10})/B{first+13}", f"=(C{first+8}+C{first+10})/C{first+13}", f"=(D{first+8}+D{first+10})/D{first+13}", f"=(E{first+8}+E{first+10})/E{first+13}", MUL),
    ("Defter değeri / pay (TL)",        f"=B{first+14}/B{first+17}", f"=C{first+14}/C{first+17}", f"=D{first+14}/D{first+17}", f"=E{first+14}/E{first+17}", TL2),
]
for name, *cells in derived:
    fmt = cells[-1]
    put(ws, r, 1, name, SUB, align="left")
    for i, fx in enumerate(cells[:-1]):
        put(ws, r, 2 + i, fx, BODY, fmt, align="right")
    r += 1

r += 1
put(ws, r, 1, "Not: Tabloda 'Finansal Borçlar' yalnızca uzun vadeli dilimi gösterir; kısa vadeli finansal borç veri setinde ayrıştırılmamıştır. Şirketin açıkladığı net nakit pozisyonu 1Ç26'da 74,7 mlr TL idi; 2Ç26 nakit hareketinden türetilen tahmin ≈ 127,7 mlr TL (bkz. 7-Senaryo).", NOTE, align="left", border=False)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
ws.row_dimensions[r].height = 30

# ============================================================ 3) GELIR TABLOSU
ws = wb.create_sheet("3-Gelir-Tablosu")
widths(ws, {"A": 38, "B": 17, "C": 17, "D": 16, "E": 17, "F": 17, "G": 17, "H": 15})
title_row(ws, "TÜPRAŞ — GELİR TABLOSU (milyon TL, kümülatif dönem)", 8)
ws["A2"] = "Kaynak: borsa MCP / borsapy. 2026/6A ve 2025/6A altı aylık kümülatif; diğerleri yıl sonu. 2Ç26 tek çeyrek = 2026/6A eksi 1Ç26."
ws["A2"].font = NOTE; ws.merge_cells("A2:H2")

r = 4
header_row(ws, r, ["Kalem", "2026/6A", "2025/6A", "Y/Y", "2025", "2024", "2023", "2Ç26 (tek çeyrek)"])
inc = [
    ("Satış Gelirleri",          662788.010, 464119.398, 977814.685, 1060732.568, 991202.993, 404534.160),
    ("Satışların Maliyeti (-)", -578662.326,-422131.404,-882156.796, -971686.119,-832772.469,-341983.909),
    ("BRÜT KAR",                  84125.684,  41987.994,  95657.889,   89046.449, 158430.524,  62550.251),
    ("Pazarlama Giderleri (-)",   -7520.531,  -6011.496, -12853.234,  -13045.252,  -9263.098,  -3985.329),
    ("Genel Yönetim Gid. (-)",   -13882.060, -12239.818, -26346.314,  -21800.405, -18567.271,  -7666.509),
    ("Ar-Ge Giderleri (-)",        -349.166,   -285.700,   -585.688,    -419.836,   -404.209,   -180.683),
    ("Net Faaliyet Karı",         62373.927,  23450.980,  55872.653,   53780.956, 130195.946,  50717.730),
    ("FAALİYET KARI",             59025.244,  18160.918,  49042.749,   46741.606, 106303.044,  47632.760),
    ("Amortisman & İtfa",          8872.312,   8554.334,  17223.191,   12586.750,   8754.469,   4732.955),
    ("FAVÖK (Net Faal.K + Amort.)",71246.239,  32005.314,  73095.844,   66367.706, 138950.415,  55450.685),
    ("Net Finansal Gelir/Gider",   1343.142,   -848.646,    981.198,   -7056.318, -24459.428,   1416.373),
    ("VERGİ ÖNCESİ KAR",          61029.781,  18104.917,  51541.378,   41552.429,  83480.454,  50646.934),
    ("Vergi Gideri (-)",         -10757.308,  -6123.357, -16363.776,  -16638.855,  -5700.367,  -4189.561),
    ("DÖNEM NET KARI",            50272.473,  11981.560,  35177.602,   24913.575,  77780.087,  46457.373),
    ("  Ana Ortaklık Payı",       49847.684,  11872.237,  34766.049,   23973.136,  77354.421,  46137.925),
    ("Hisse Başına Kazanç (TL)",     25.870,      6.161,     18.043,      12.442,     40.147,     23.945),
]
r += 1
first = r
for name, a, b, y25, y24, y23, q2 in inc:
    bold = name.isupper() or "FAVÖK" in name
    f = SUB if bold else BODY
    fmt = TL2 if "Hisse Başına" in name else MN
    put(ws, r, 1, name, f, align="left")
    put(ws, r, 2, a, f, fmt, align="right")
    put(ws, r, 3, b, f, fmt, align="right")
    put(ws, r, 4, f"=IF(C{r}=0,\"\",IF(SIGN(C{r})<>SIGN(B{r}),\"n.a.\",B{r}/C{r}-1))", f, PCT, align="right")
    for i, v in enumerate([y25, y24, y23, q2]):
        put(ws, r, 5 + i, v, f, fmt, align="right")
    if bold:
        for c in range(1, 9): ws.cell(row=r, column=c).fill = FILL_S
    r += 1

r += 1
put(ws, r, 1, "MARJLAR", HDR, fill=FILL_H, align="left")
for c in range(2, 9): ws.cell(row=r, column=c).fill = FILL_H
r += 1
margins = [("Brüt Kar Marjı", first+2), ("Faaliyet Kar Marjı", first+7), ("FAVÖK Marjı", first+9), ("Net Kar Marjı", first+14)]
for name, src in margins:
    put(ws, r, 1, name, SUB, align="left")
    for col in ["B", "C", "E", "F", "G", "H"]:
        cc = {"B":2,"C":3,"E":5,"F":6,"G":7,"H":8}[col]
        put(ws, r, cc, f"={col}{src}/{col}{first}", BODY, PCT, align="right")
    put(ws, r, 4, f"=B{r}-C{r}", BODY, '0.0"pp";(0.0"pp")', align="right")
    r += 1

r += 1
put(ws, r, 1, "Not: 2Ç26 tek çeyrek satış geliri kümülatiflerden türetilmiştir (404,5 mlr TL). Şirketin sunumunda açıklanan 2Ç26 satış geliri 386,4 mlr TL, FAVÖK 54,3 mlr TL, FAVÖK marjı %14,14'tür — fark TMS-29 yeniden değerlemesinden kaynaklanır.", NOTE, align="left", border=False)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
ws.row_dimensions[r].height = 30

# ============================================================ 4) NAKIT AKIM
ws = wb.create_sheet("4-Nakit-Akim")
widths(ws, {"A": 42, "B": 18, "C": 18, "D": 15, "E": 18, "F": 18})
title_row(ws, "TÜPRAŞ — NAKİT AKIM TABLOSU (milyon TL, kümülatif)", 6)
ws["A2"] = "Kaynak: borsa MCP / borsapy. 2026 ve 2025 sütunları altı aylık kümülatiftir."
ws["A2"].font = NOTE; ws.merge_cells("A2:F2")

r = 4
header_row(ws, r, ["Kalem", "2026/6A", "2025/6A", "Y/Y", "2025 (yıl)", "2024 (yıl)"])
cf = [
    ("Dönem Karı",                          50272.473,  11981.560,  35177.602,  24913.575),
    ("Amortisman & İtfa",                    8872.312,   8554.334,  17223.191,  12586.750),
    ("İşletme Sermayesindeki Değişim",      34035.937, -49016.033, -21123.427,  -9726.430),
    ("İŞLETME FAALİYETLERİNDEN NET NAKİT",  132366.363,  15012.559,  58248.915,  46167.074),
    ("Sabit Sermaye Yatırımları (Capex)",  -11856.999, -10889.286, -22464.984, -18050.188),
    ("Yatırım Faaliyetlerinden Nakit",     -13564.562, -11371.542, -22864.431, -15563.691),
    ("SERBEST NAKİT AKIM",                 118801.801,   3641.017,  35384.483,  30603.383),
    ("Temettü Ödemeleri",                  -21402.951, -20720.229, -41032.800, -63827.710),
    ("Finansal Borçlardaki Değişim",        16377.204,  29220.354,  36940.277, -32168.215),
    ("Finansman Faaliyetlerinden Nakit",   -45160.214,  11744.926,    169.600, -75571.511),
    ("Nakit ve Benzerlerindeki Net Değişim",77926.167,  17814.631,  41512.199, -42074.706),
    ("DÖNEM SONU NAKİT",                   164585.006,  96754.791, 106727.848,  77507.034),
    ("Yurt İçi Satışlar",                  529793.167, 365938.212, 819013.868, 880709.835),
    ("Yurt Dışı Satışlar (ihracat)",       141303.619, 103650.280, 171547.735, 192073.559),
]
r += 1
first = r
for name, a, b, y25, y24 in cf:
    bold = name.isupper()
    f = SUB if bold else BODY
    put(ws, r, 1, name, f, align="left")
    put(ws, r, 2, a, f, MN, align="right")
    put(ws, r, 3, b, f, MN, align="right")
    put(ws, r, 4, f"=IF(C{r}=0,\"\",IF(SIGN(C{r})<>SIGN(B{r}),\"n.a.\",B{r}/C{r}-1))", f, PCT, align="right")
    put(ws, r, 5, y25, f, MN, align="right")
    put(ws, r, 6, y24, f, MN, align="right")
    if bold:
        for c in range(1, 7): ws.cell(row=r, column=c).fill = FILL_S
    r += 1

r += 1
put(ws, r, 1, "TÜRETİLMİŞ", HDR, fill=FILL_H, align="left")
for c in range(2, 7): ws.cell(row=r, column=c).fill = FILL_H
r += 1
put(ws, r, 1, "Serbest Nakit Akım Verimi (PD 755.786 mn TL)", SUB, align="left")
put(ws, r, 2, f"=B{first+6}*2/755786", BODY, PCT, align="right")
put(ws, r, 3, f"=C{first+6}*2/755786", BODY, PCT, align="right")
put(ws, r, 5, f"=E{first+6}/755786", BODY, PCT, align="right")
put(ws, r, 6, f"=F{first+6}/755786", BODY, PCT, align="right")
ws.cell(row=r, column=2).comment = Comment("2026 ve 2025 sütunlarında 6 aylık SNA yıllıklandırılmıştır (x2). Piyasa değeri 02.09.2026 kapanışı ile sabittir.", "Analiz")
r += 1
put(ws, r, 1, "Temettü / Serbest Nakit Akım (ödeme oranı)", SUB, align="left")
put(ws, r, 2, f"=-B{first+7}/B{first+6}", BODY, PCT, align="right")
put(ws, r, 3, f"=-C{first+7}/C{first+6}", BODY, PCT, align="right")
put(ws, r, 5, f"=-E{first+7}/E{first+6}", BODY, PCT, align="right")
put(ws, r, 6, f"=-F{first+7}/F{first+6}", BODY, PCT, align="right")
r += 1
put(ws, r, 1, "İhracat / Toplam Satış", SUB, align="left")
for col, cc in [("B",2),("C",3),("E",5),("F",6)]:
    put(ws, r, cc, f"={col}{first+13}/({col}{first+12}+{col}{first+13})", BODY, PCT, align="right")

# ============================================================ 5) RASYOLAR & SEKTOR
ws = wb.create_sheet("5-Rasyo-Sektor")
widths(ws, {"A": 32, "B": 16, "C": 16, "D": 14, "E": 14, "F": 16, "G": 34})
title_row(ws, "TÜPRAŞ — DEĞERLEME RASYOLARI VE SEKTÖR KARŞILAŞTIRMASI (XUSIN)", 7)
ws["A2"] = "Kaynak: borsa MCP / İş Yatırım + borsapy. Sektör medyanı F/K 14,15x — PD/DD 1,70x."
ws["A2"].font = NOTE; ws.merge_cells("A2:G2")

r = 4
header_row(ws, r, ["Rasyo", "TUPRS", "Sektör Medyanı", "Fark", "", "", "Yorum"])
ratios = [
    ("F/K (TTM)",            10.40, 14.15, "Sektöre göre %26 iskonto"),
    ("PD/DD",                 1.70,  1.70, "Sektör ile aynı seviyede"),
    ("FD/FAVÖK",              5.60,  None, "Zirve FAVÖK üzerinden hesaplandığı için yanıltıcı olabilir"),
    ("FD/Satışlar",           0.91,  None, "Rafineri iş modeli için normal"),
    ("Temettü Verimi (TTM)", 6.26,   None, "2026 toplamı üzerinden ≈ %7,7"),
    ("Owner Earnings (mn TL)",42902.42, None, "Buffett modeli gerçek nakit üretimi"),
    ("OE Verimi",            22.70,  10.00, "Buffett eşiği %10'un çok üzerinde"),
    ("DCF İçsel Değer (TL)",1022.31, 392.25, "Güvenlik marjı %61,6 — zirve nakit akışı varsayımına duyarlı"),
]
r += 1
for name, v, cmp_, note in ratios:
    put(ws, r, 1, name, SUB, align="left")
    fmt = MN if (isinstance(v, float) and v > 1000) else (MUL if "F/K" in name or "PD/DD" in name or "FD/" in name else TL2)
    put(ws, r, 2, v, BODY, fmt, align="center")
    put(ws, r, 3, cmp_, BODY, fmt, align="center")
    if cmp_ is not None:
        put(ws, r, 4, f"=B{r}/C{r}-1", BODY, PCT, align="center")
    else:
        put(ws, r, 4, "", BODY, align="center")
    put(ws, r, 5, "", BODY); put(ws, r, 6, "", BODY)
    put(ws, r, 7, note, BODY, align="left")
    r += 1

r += 2
put(ws, r, 1, "XUSIN SEKTÖR EMSALLERİ", HDR, fill=FILL_H, align="left")
for c in range(2, 8): ws.cell(row=r, column=c).fill = FILL_H
r += 1
header_row(ws, r, ["Şirket", "Kod", "Piyasa Değeri (mn TL)", "F/K", "PD/DD", "F/K vs TUPRS", "Not"])
peers = [
    ("Tüpraş",                  "TUPRS", 776016.93, 10.67, 1.70, "İnceleme konusu"),
    ("Astor Enerji",            "ASTOR", 316116.50, 28.22, 7.22, ""),
    ("Ford Otosan",             "FROTO", 271077.98,  8.16, 1.51, ""),
    ("Ereğli Demir Çelik",      "EREGL", 264460.00, 33.89, 0.83, ""),
    ("Coca-Cola İçecek",        "CCOLA", 214752.53,  9.76, 2.35, ""),
    ("Gübre Fabrikaları",       "GUBRF", 157314.00, 19.51, 4.45, ""),
    ("Anadolu Efes",            "AEFES", 107052.63, 10.67, 0.85, ""),
    ("Oyak Çimento",            "OYAKC", 105497.93, 13.64, 1.39, ""),
    ("Borusan Boru",            "BRSAN", 100232.51, 37.17, 2.31, ""),
    ("Aksa",                    "AKSA",   41569.50,  6.40, 1.13, ""),
    ("Çimsa",                   "CIMSA",  41416.89,  9.54, 1.11, ""),
    ("Kardemir (D)",            "KRDMD",  34642.03, 11.68, 0.47, ""),
]
r += 1
pfirst = r
for name, code, mc, pe, pb, note in peers:
    f = SUB if code == "TUPRS" else BODY
    fill = FILL_G if code == "TUPRS" else None
    put(ws, r, 1, name, f, fill=fill, align="left")
    put(ws, r, 2, code, f, fill=fill, align="center")
    put(ws, r, 3, mc, f, MN, fill=fill, align="right")
    put(ws, r, 4, pe, f, MUL, fill=fill, align="center")
    put(ws, r, 5, pb, f, MUL, fill=fill, align="center")
    put(ws, r, 6, f"=D{r}/$D${pfirst}-1", f, PCT, fill=fill, align="center")
    put(ws, r, 7, note, f, fill=fill, align="left")
    r += 1
put(ws, r, 1, "Medyan (emsaller)", SUB, align="left")
put(ws, r, 3, f"=MEDIAN(C{pfirst}:C{r-1})", SUB, MN, align="right")
put(ws, r, 4, f"=MEDIAN(D{pfirst}:D{r-1})", SUB, MUL, align="center")
put(ws, r, 5, f"=MEDIAN(E{pfirst}:E{r-1})", SUB, MUL, align="center")

# ============================================================ 6) TEKNIK ANALIZ
ws = wb.create_sheet("6-Teknik")
widths(ws, {"A": 30, "B": 16, "C": 16, "D": 18, "E": 56})
title_row(ws, "TÜPRAŞ — TEKNİK ANALİZ (günlük, 02.09.2026 kapanışı 392,25 TL)", 5)
ws["A2"] = "Kaynak: borsa MCP / yfinance. BIST hisselerinde teknik göstergeler günlük veriyle hesaplanır."
ws["A2"].font = NOTE; ws.merge_cells("A2:E2")

r = 4
header_row(ws, r, ["Gösterge", "Değer", "Fiyata Uzaklık", "Sinyal", "Yorum"])
tech = [
    ("Kapanış",       392.25, None,  "—",   "01.09'da 417,25 TL zirvesi görüldükten sonra geri çekilme"),
    ("EMA-20",        385.26, "=B{r}/$B${p}-1", "AL",  "Fiyat EMA-20 üzerinde — kısa vadeli trend pozitif"),
    ("SMA-20",        372.35, "=B{r}/$B${p}-1", "AL",  "Ana destek bölgesi"),
    ("EMA-50",        365.01, "=B{r}/$B${p}-1", "AL",  "Orta vadeli trend desteği; kaybı trend bozar"),
    ("RSI-14",         67.29, None,  "AL / dikkat", "70 aşırı alım sınırına yakın — yeni alım için ideal değil"),
    ("MACD",           20.25, None,  "AL",  "Sinyal çizgisinin (19,69) üzerinde"),
    ("MACD Histogram",  0.554, None, "ZAYIFLIYOR", "Pozitif ama daralıyor — momentum yavaşlıyor"),
    ("Trend (algoritmik)", None, None, "YATAY", "Zirve sonrası konsolidasyon"),
]
r += 1
price_row = r
for name, v, dist, sig, note in tech:
    put(ws, r, 1, name, SUB, align="left")
    put(ws, r, 2, v, BODY, TL2 if (v and v > 100) else '0.00', align="center")
    if dist:
        put(ws, r, 3, dist.format(r=r, p=price_row), BODY, PCT, align="center")
    else:
        put(ws, r, 3, "", BODY, align="center")
    fill = FILL_G if sig == "AL" else (FILL_O if sig in ("YATAY", "AL / dikkat", "ZAYIFLIYOR") else None)
    put(ws, r, 4, sig, SUB, fill=fill, align="center")
    put(ws, r, 5, note, BODY, align="left")
    r += 1

r += 1
put(ws, r, 1, "PIVOT SEVİYELERİ (klasik)", HDR, fill=FILL_H, align="left")
for c in range(2, 6): ws.cell(row=r, column=c).fill = FILL_H
r += 1
header_row(ws, r, ["Seviye", "Fiyat (TL)", "Uzaklık", "Tip", "Açıklama"])
r += 1
pv_first = r
pivots = [
    ("R3", 430.58, "Direnç", "Güçlü marj haberi olmadan zorlu"),
    ("R2", 423.92, "Direnç", "52 hafta zirvesi üstü hedef"),
    ("52h Zirve", 417.25, "Direnç", "01.09.2026'da test edildi"),
    ("R1", 413.33, "Direnç", "İlk ciddi satış bölgesi"),
    ("Pivot", 406.67, "Direnç", "Gün içi kaybedildi — geri alınması gerek"),
    ("S1", 396.08, "Destek", "Bugün kırıldı"),
    ("S2", 389.42, "Destek", "Fiyat hemen üzerinde — kritik"),
    ("Blok satış fiyatı", 384.50, "Destek", "Koç Holding satış çıpası — psikolojik zemin"),
    ("S3", 378.83, "Destek", "Kademeli alım için ilk bölge"),
    ("SMA-20", 372.35, "Destek", "Kademeli alım için ikinci bölge"),
    ("EMA-50", 365.01, "Destek", "Orta vade trend çizgisi — kaybı stop seviyesi"),
]
for lvl, px, tip, note in pivots:
    fill = FILL_R if tip == "Direnç" else FILL_G
    put(ws, r, 1, lvl, SUB, align="left")
    put(ws, r, 2, px, BODY, TL2, align="center")
    put(ws, r, 3, f"=B{r}/$B${price_row}-1", BODY, PCT, align="center")
    put(ws, r, 4, tip, BODY, fill=fill, align="center")
    put(ws, r, 5, note, BODY, align="left")
    r += 1

r += 1
put(ws, r, 1, "FİYAT PERFORMANSI", HDR, fill=FILL_H, align="left")
for c in range(2, 6): ws.cell(row=r, column=c).fill = FILL_H
r += 1
header_row(ws, r, ["Dönem / Varlık", "Getiri (TRY)", "Getiri (USD)", "Başlangıç", "Not"])
r += 1
perf = [
    ("TUPRS — YBB (02.01→02.09.2026)", 1.0976, 0.8682, 187.00, "BIST'in en güçlü büyük ölçekli hisselerinden"),
    ("Brent — YBB",                    0.7435, 0.5528, 60.75,  "94,33 USD/varil"),
    ("Gram Altın — YBB",               0.1381, 0.0136, 5985.70,""),
    ("USD/TRY — YBB",                  0.1228, 0.0000, 43.03,  "48,29 TL"),
    ("TUPRS — 2 aylık (30.06→02.09)",  0.7242, None,   227.50, "Bilanço sonrası sert ralli"),
    ("TUPRS — 52 hafta",               1.2493, None,   174.00, "Eyl-2025 seviyesinden"),
]
for name, tr, usd, start, note in perf:
    put(ws, r, 1, name, SUB, align="left")
    put(ws, r, 2, tr, BODY, PCT, align="center")
    put(ws, r, 3, usd, BODY, PCT, align="center")
    put(ws, r, 4, start, BODY, TL2, align="center")
    put(ws, r, 5, note, BODY, align="left")
    r += 1

# ============================================================ 7) SENARYO
ws = wb.create_sheet("7-Senaryo")
widths(ws, {"A": 40, "B": 18, "C": 18, "D": 18, "E": 18, "F": 18, "G": 30})
title_row(ws, "TÜPRAŞ — DEĞERLEME SENARYO ANALİZİ (FD/FAVÖK yöntemi)", 7)
ws["A2"] = "Mavi hücreler girdi/varsayımdır — değiştirilebilir. Siyah hücreler formüldür."
ws["A2"].font = NOTE; ws.merge_cells("A2:G2")

r = 4
put(ws, r, 1, "VARSAYIMLAR", HDR, fill=FILL_H, align="left")
for c in range(2, 8): ws.cell(row=r, column=c).fill = FILL_H
r += 1
assum = [
    ("Pay adedi (mn adet)",              1926.796, "KAP — ödenmiş sermaye 1.926.795.598 TL, 1 TL nominal"),
    ("Güncel fiyat (TL)",                  392.25, "BIST kapanış 02.09.2026"),
    ("USD/TRY (değerleme kuru)",           48.2913,"Spot 02.09.2026"),
    ("Net nakit pozisyonu (mn TL)",     127700.00, "Türetilmiş: şirket açıklamalı 1Ç26 net nakit 74,7 mlr TL + 2Ç26 nakit artışı 65,9 mlr TL - borç artışı 12,9 mlr TL"),
]
afirst = r
for name, v, note in assum:
    put(ws, r, 1, name, SUB, align="left")
    put(ws, r, 2, v, BLUE, MN if v > 1000 else TL2, fill=FILL_Y, align="center")
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
    put(ws, r, 3, note, NOTE, align="left")
    r += 1
SH, PX, FX_, NC = afirst, afirst+1, afirst+2, afirst+3

r += 1
put(ws, r, 1, "PİYASA DEĞERİ (kontrol)", SUB, align="left")
put(ws, r, 2, f"=B{SH}*B{PX}", BODY, MN, align="center")
put(ws, r, 3, "mn TL", BODY, align="left")
mc_row = r
r += 1
put(ws, r, 1, "FİRMA DEĞERİ (PD - Net Nakit)", SUB, align="left")
put(ws, r, 2, f"=B{mc_row}-B{NC}", BODY, MN, align="center")
put(ws, r, 3, "mn TL", BODY, align="left")
r += 2

put(ws, r, 1, "SENARYOLAR — 2026/normalize FAVÖK", HDR, fill=FILL_H, align="left")
for c in range(2, 8): ws.cell(row=r, column=c).fill = FILL_H
r += 1
header_row(ws, r, ["Senaryo", "Net Raf. Marjı ($/varil)", "FAVÖK (mn USD)", "FAVÖK (mn TL)", "Dayanak", "", ""])
r += 1
scen_first = r
scenarios = [
    ("Boğa — marj kalıcı yüksek",   14.0, 3400, "Şirket üst bant 15 $/varil; Akdeniz arz sıkışıklığı sürer"),
    ("Baz — şirket rehberliği",     13.0, 3100, "2Ç26 sonrası konsensüs FAVÖK tahmini 3,1 mlr USD"),
    ("Ayı — hızlı normalleşme",      9.0, 2200, "Marjların 2H26'da 9 $/varile gerilemesi"),
    ("Normalize — orta çevrim",      7.0, 2000, "Tarihsel orta-çevrim marjı 6-8 $/varil (şirketin yıl başı rehberliği)"),
    ("Sert daralma",                 5.0, 1500, "Talep şoku + arz normalleşmesi birlikte"),
]
for name, nrm, ebitda, note in scenarios:
    put(ws, r, 1, name, SUB, align="left")
    put(ws, r, 2, nrm, BLUE, '0.0', fill=FILL_Y, align="center")
    put(ws, r, 3, ebitda, BLUE, MN, fill=FILL_Y, align="center")
    put(ws, r, 4, f"=C{r}*$B${FX_}", BODY, MN, align="center")
    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7)
    put(ws, r, 5, note, NOTE, align="left")
    r += 1
scen_last = r - 1

r += 1
put(ws, r, 1, "HEDEF FİYAT MATRİSİ (TL/pay)", HDR, fill=FILL_H, align="left")
for c in range(2, 8): ws.cell(row=r, column=c).fill = FILL_H
r += 1
header_row(ws, r, ["Senaryo", "4,0x FD/FAVÖK", "5,0x FD/FAVÖK", "5,6x (mevcut)", "6,5x FD/FAVÖK", "5,6x'te potansiyel", "Değerlendirme"])
mult_row = r
r += 1
mults = {2: 4.0, 3: 5.0, 4: 5.6, 5: 6.5}
verdict_txt = ["Güçlü yükseliş", "Makul yükseliş", "Yatay / hafif düşüş", "Aşağı yönlü risk", "Ciddi düşüş riski"]
for i, (name, nrm, ebitda, note) in enumerate(scenarios):
    src = scen_first + i
    put(ws, r, 1, name, SUB, align="left")
    for col, m in mults.items():
        put(ws, r, col, f"=(D{src}*{m}+$B${NC})/$B${SH}", BODY, TL2, align="center")
    put(ws, r, 6, f"=D{r}/$B${PX}-1", SUB, PCT, align="center")
    fill = FILL_G if i < 2 else (FILL_O if i == 2 else FILL_R)
    ws.cell(row=r, column=6).fill = fill
    put(ws, r, 7, verdict_txt[i], BODY, fill=fill, align="center")
    r += 1
# etiketleri carpan basliklarina yaz
for col, m in mults.items():
    ws.cell(row=mult_row, column=col).value = f"{m:.1f}x FD/FAVÖK" + (" (mevcut)" if m == 5.6 else "")

r += 1
put(ws, r, 1, "Olasılık ağırlıklı adil değer (Boğa %20 / Baz %30 / Ayı %25 / Normalize %20 / Sert %5, 5,0x çarpan)", SUB, align="left")
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
wr = r
put(ws, r, 6, f"=0.2*C{mult_row+1}+0.3*C{mult_row+2}+0.25*C{mult_row+3}+0.2*C{mult_row+4}+0.05*C{mult_row+5}", SUB, TL2, fill=FILL_Y, align="center")
put(ws, r, 7, "TL/pay", BODY, align="center")
r += 1
put(ws, r, 1, "Spot fiyata göre potansiyel", SUB, align="left")
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
put(ws, r, 6, f"=F{wr}/B{PX}-1", SUB, PCT, fill=FILL_Y, align="center")

r += 2
put(ws, r, 1, "Not: FD/FAVÖK çarpanı, rafineri sektöründe çevrimin tepesinde daraltılır (yatırımcı zirve kârı ekstrapole etmez). Mevcut 5,6x, son 12 ayın (zirve) FAVÖK'ü üzerinden hesaplanmıştır; normalize FAVÖK üzerinden mevcut fiyat 9-10x'e denk gelir.", NOTE, align="left", border=False)
ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
ws.row_dimensions[r].height = 30

# ============================================================ 8) HABERLER
ws = wb.create_sheet("8-Haberler")
widths(ws, {"A": 13, "B": 40, "C": 74, "D": 14, "E": 16})
title_row(ws, "TÜPRAŞ — YAKIN DÖNEM HABER VE KAP AKIŞI", 5)
ws["A2"] = "Kaynak: KAP bildirimleri, şirket sunumları ve basın taraması (Haziran–Eylül 2026). borsa MCP haber servisi TUPRS için boş döndüğünden web taramasıyla derlenmiştir."
ws["A2"].font = NOTE; ws.merge_cells("A2:E2")

r = 4
header_row(ws, r, ["Tarih", "Başlık", "İçerik / Rakamlar", "Etki", "Vade"])
news = [
    ("02.09.2026", "Koç Holding + Temel Ticaret blok pay satışı",
     "Toplam ~%0,83 pay (11 mn TL + 5,02 mn TL nominal), 1 TL nominal başına 384,50 TL fiyatla ~6,16 mlr TL büyüklüğünde toptan satış. BofA aracılığıyla 02.09'da BİST'e başvuru, işlem 03.09, takas 07.09. Sonrası: Koç doğrudan payı %3,7'ye, Koç+Enerji Yatırımları toplam %50,1'e, halka açıklık %49,7'ye çıkıyor.",
     "NEGATİF", "Kısa"),
    ("01.09.2026", "Entek Elektrik sermaye artırımına katılım",
     "Bağlı ortaklık Entek Elektrik Üretim A.Ş. sermayesi 2,065 mlr TL'den 2,500 mlr TL artışla 4,565 mlr TL'ye çıkarılıyor. Amaç: yatırım finansmanı, işletme sermayesi ve mali yapının güçlendirilmesi.",
     "NÖTR", "Orta"),
    ("Eylül 2026", "Eylül yatırımcı bilgilendirme sunumu",
     "Petrol fiyatları, ürün marjları ve Türkiye akaryakıt piyasası güncellemesi yayımlandı.",
     "NÖTR", "—"),
    ("04.08.2026", "2Ç26 bilançosu — rekor kâr, rekor nakit",
     "2Ç26 ana ortaklık net kârı 45,9 mlr TL (yıllık +%291). Satış geliri 386,4 mlr TL. FAVÖK 54,3 mlr TL — beklentiyi %41,4 aştı; net kâr beklentiyi %51,7 aştı. FAVÖK marjı %6,12'den %14,14'e (+802 bp). Çeyreklik FAVÖK, önceki dört çeyreğin toplamına yaklaştı. 6 aylık net kâr 49,85 mlr TL (2025/6A: 11,87 mlr TL).",
     "ÇOK POZİTİF", "Orta"),
    ("04.08.2026", "2026 net rafineri marjı beklentisi yükseltildi",
     "Yıl başı rehberliği 6-7 USD/varil iken 13-15 USD/varil aralığına revize edildi. Diğer hedefler korundu: %95-100 kapasite kullanımı, ~29 mn ton üretim, ~30 mn ton satış, ~700 mn USD yatırım harcaması. Piyasa 2026 FAVÖK tahminini 2,4 mlr USD'den 3,1 mlr USD'ye yükseltti.",
     "ÇOK POZİTİF", "Orta"),
    ("Ağustos 2026", "16 aracı kurum hedef fiyat güncellemesi",
     "Bilanço sonrası ortalama hedef fiyat 405,44 TL; en düşük 346,00 TL, en yüksek 470,90 TL. Vakıf Yatırım 510 TL hedefle 'AL' tavsiyesiyle takibe başladı.",
     "POZİTİF", "Orta"),
    ("Haziran 2026", "Akdeniz'de rafineri marjı şoku",
     "İsrail'in Hayfa'daki 197 bin varil/gün kapasiteli rafinerisine saldırı sonrası Akdeniz rafine ürün piyasası sıkıştı. Dizel crack spread ~45 USD/varile çıktı (önceki çeyrek ort. 27,7 USD/varil). Avrupa'da yapısal dizel açığı devam ediyor.",
     "POZİTİF (geçici)", "Kısa/Orta"),
    ("30.09.2026", "2026 ikinci temettü taksiti (yaklaşan)",
     "13 mlr TL dağıtım; brüt 6,7469533 TL/pay, net 5,7349103 TL/pay. Hak kullanım (ex-date) 30.09.2026, ödeme 02.10.2026. 2026 toplam temettü 33 mlr TL — brüt 17,13 TL/pay, net 14,56 TL/pay.",
     "POZİTİF (fiyat düzeltmeli)", "Kısa"),
    ("16.03.2026", "2026 birinci temettü taksiti (ödendi)",
     "20 mlr TL; brüt 10,3799282 TL/pay, net 8,8229389 TL/pay. Ödeme 18.03.2026.",
     "POZİTİF", "Geçmiş"),
    ("29.10.2026", "3Ç26 bilanço tarihi (beklenen)",
     "TradingView'e göre bir sonraki finansal açıklama tarihi. 2Ç26'da HBK beklentisi 19,49 TL iken gerçekleşme 36,04 TL oldu.",
     "İZLENECEK", "Kısa"),
]
r += 1
for date, title, body, impact, horizon in news:
    put(ws, r, 1, date, SUB, align="center")
    put(ws, r, 2, title, SUB, align="left")
    put(ws, r, 3, body, BODY, align="left")
    fill = FILL_G if "POZİTİF" in impact else (FILL_R if "NEGATİF" in impact else FILL_O)
    put(ws, r, 4, impact, SUB, fill=fill, align="center")
    put(ws, r, 5, horizon, BODY, align="center")
    ws.row_dimensions[r].height = 74
    r += 1

# ============================================================ 9) ANALIST
ws = wb.create_sheet("9-Analist")
widths(ws, {"A": 34, "B": 20, "C": 20, "D": 20, "E": 46})
title_row(ws, "TÜPRAŞ — ANALİST BEKLENTİLERİ VE HEDEF FİYATLAR", 5)
ws["A2"] = "Kaynak: borsa MCP / yfinance konsensüsü + basında derlenen yurt içi aracı kurum hedefleri (Ağustos 2026)."
ws["A2"].font = NOTE; ws.merge_cells("A2:E2")

r = 4
header_row(ws, r, ["Gösterge", "Değer", "Spot (392,25 TL)", "Potansiyel", "Not"])
r += 1
an_price = r
analyst = [
    ("Güçlü AL",                     0, None, "yfinance konsensüs dağılımı"),
    ("AL",                          20, None, "Toplam 27 kurumun 20'si"),
    ("TUT",                          7, None, ""),
    ("SAT / Güçlü SAT",              0, None, "Hiç satış tavsiyesi yok"),
    ("Konsensüs ort. hedef (TL)",  396.98, 392.25, "Uluslararası konsensüs"),
    ("En düşük hedef (TL)",        278.00, 392.25, "Marj normalleşmesi senaryosu"),
    ("En yüksek hedef (TL)",       510.00, 392.25, "Vakıf Yatırım — AL ile takibe başlangıç"),
    ("Yurt içi 16 kurum ort. (TL)",405.44, 392.25, "2Ç26 bilanço sonrası güncellenen hedefler"),
    ("Yurt içi en düşük (TL)",     346.00, 392.25, ""),
    ("Yurt içi en yüksek (TL)",    470.90, 392.25, ""),
]
for name, v, spot, note in analyst:
    put(ws, r, 1, name, SUB, align="left")
    put(ws, r, 2, v, BODY, TL2 if v > 100 else '0', align="center")
    put(ws, r, 3, spot, BODY, TL2, align="center")
    if spot:
        put(ws, r, 4, f"=B{r}/C{r}-1", BODY, PCT, align="center")
    else:
        put(ws, r, 4, "", BODY, align="center")
    put(ws, r, 5, note, BODY, align="left")
    r += 1

r += 1
put(ws, r, 1, "KAZANÇ BEKLENTİLERİ", HDR, fill=FILL_H, align="left")
for c in range(2, 6): ws.cell(row=r, column=c).fill = FILL_H
r += 1
header_row(ws, r, ["Dönem", "Beklenti (HBK, TL)", "Gerçekleşen (HBK, TL)", "Sürpriz", "Not"])
r += 1
earn = [
    ("2Ç26 (04.08.2026)", 19.4928, 36.0441, "Beklentinin %85 üzerinde"),
    ("HBK — TTM",         None,    36.0441, "Son 12 ay"),
    ("3Ç26 (29.10.2026)", 19.4928, None,    "Bir sonraki açıklama tarihi"),
]
for per, est, act, note in earn:
    put(ws, r, 1, per, SUB, align="left")
    put(ws, r, 2, est, BODY, TL2, align="center")
    put(ws, r, 3, act, BODY, TL2, align="center")
    if est and act:
        put(ws, r, 4, f"=C{r}/B{r}-1", BODY, PCT, fill=FILL_G, align="center")
    else:
        put(ws, r, 4, "", BODY, align="center")
    put(ws, r, 5, note, BODY, align="left")
    r += 1

r += 1
put(ws, r, 1, "ŞİRKET REHBERLİĞİ (2026)", HDR, fill=FILL_H, align="left")
for c in range(2, 6): ws.cell(row=r, column=c).fill = FILL_H
r += 1
header_row(ws, r, ["Kalem", "Yıl Başı Hedefi", "2Ç26 Sonrası Revize", "Değişim", "Not"])
r += 1
guide = [
    ("Net rafineri marjı (USD/varil)", "6-7",     "13-15",   "Yukarı revize", "Tek revize edilen kalem — marj şoku kaynaklı"),
    ("Kapasite kullanım oranı",        "%95-100", "%95-100", "Değişmedi",     ""),
    ("Üretim (mn ton)",                "~29",     "~29",     "Değişmedi",     ""),
    ("Satış (mn ton)",                 "~30",     "~30",     "Değişmedi",     ""),
    ("Yatırım harcaması (mn USD)",     "~700",    "~700",    "Değişmedi",     "Serbest nakit akımı koruyucu"),
]
for name, a, b, ch, note in guide:
    put(ws, r, 1, name, SUB, align="left")
    put(ws, r, 2, a, BODY, align="center")
    put(ws, r, 3, b, BODY, align="center")
    put(ws, r, 4, ch, SUB, fill=(FILL_G if "Yukarı" in ch else None), align="center")
    put(ws, r, 5, note, BODY, align="left")
    r += 1

# ============================================================ 10) FIYAT VERISI
ws = wb.create_sheet("10-Fiyat-Verisi")
widths(ws, {"A": 14, "B": 12, "C": 12, "D": 12, "E": 12, "F": 16, "G": 14, "H": 14})
title_row(ws, "TÜPRAŞ — FİYAT VERİSİ (günlük, son 1 ay ve aylık, son 1 yıl)", 8)
ws["A2"] = "Kaynak: borsa MCP / borsapy. Bölünme düzeltilmiş kapanışlar."
ws["A2"].font = NOTE; ws.merge_cells("A2:H2")

r = 4
put(ws, r, 1, "GÜNLÜK (03.08 – 02.09.2026)", HDR, fill=FILL_H, align="left")
for c in range(2, 9): ws.cell(row=r, column=c).fill = FILL_H
r += 1
header_row(ws, r, ["Tarih", "Açılış", "Yüksek", "Düşük", "Kapanış", "Hacim (adet)", "Günlük %", "Gün İçi Aralık %"])
daily = [
 ("2026-08-03",293,295.75,286.5,290.5,21759393),("2026-08-04",293.25,295.5,283,290.25,30714332),
 ("2026-08-05",299,310.5,295.5,303.75,54128289),("2026-08-06",304.25,323.5,303.5,321.5,42412282),
 ("2026-08-07",322.75,332,318.25,323.5,36486193),("2026-08-10",325,342.5,323.75,340,38603793),
 ("2026-08-11",342.75,344.75,329.75,336.5,39267266),("2026-08-12",337.5,353.5,337.5,351.25,35147541),
 ("2026-08-13",351.5,357,343.25,346,22784976),("2026-08-14",350,364.5,345.75,361.75,34621634),
 ("2026-08-17",362.5,373.5,361,368.75,25997605),("2026-08-18",372,374.5,364.25,374,26573759),
 ("2026-08-19",374.25,398.25,373,395.5,40012823),("2026-08-20",398,402.75,390.75,395.25,29319183),
 ("2026-08-21",396,406.75,391,406.75,25718838),("2026-08-24",407,407.5,399.25,400.5,22238012),
 ("2026-08-25",398.5,402,373.75,373.75,44737737),("2026-08-26",369.75,382.75,366.5,380,34726936),
 ("2026-08-27",380,386.75,378.75,380.75,21093852),("2026-08-28",382,396,380.75,396,28201225),
 ("2026-08-31",406,406.25,397.25,400.25,35328555),("2026-09-01",404,417.25,400,402.75,29832344),
 ("2026-09-02",402.75,403,390.5,392.25,34117361),
]
r += 1
dfirst = r
for d, o, h, l, c_, v in daily:
    put(ws, r, 1, d, BODY, align="center")
    for i, x in enumerate([o, h, l, c_]):
        put(ws, r, 2 + i, x, BODY, TL2, align="right")
    put(ws, r, 6, v, BODY, '#,##0', align="right")
    if r > dfirst:
        put(ws, r, 7, f"=E{r}/E{r-1}-1", BODY, PCT, align="right")
    else:
        put(ws, r, 7, "", BODY, align="right")
    put(ws, r, 8, f"=C{r}/D{r}-1", BODY, PCT, align="right")
    r += 1
dlast = r - 1
put(ws, r, 1, "Dönem", SUB, align="left")
put(ws, r, 3, f"=MAX(C{dfirst}:C{dlast})", SUB, TL2, align="right")
put(ws, r, 4, f"=MIN(D{dfirst}:D{dlast})", SUB, TL2, align="right")
put(ws, r, 5, f"=E{dlast}/E{dfirst}-1", SUB, PCT, align="right")
put(ws, r, 6, f"=AVERAGE(F{dfirst}:F{dlast})", SUB, '#,##0', align="right")
put(ws, r, 2, "Zirve/Dip/Getiri/Ort.hacim", NOTE, align="left")

r += 2
put(ws, r, 1, "AYLIK (Eyl-2025 – Eyl-2026)", HDR, fill=FILL_H, align="left")
for c in range(2, 9): ws.cell(row=r, column=c).fill = FILL_H
r += 1
header_row(ws, r, ["Ay", "Açılış", "Yüksek", "Düşük", "Kapanış", "Hacim (adet)", "Aylık %", ""])
monthly = [
 ("2025-09",174,201.7,163.3,186.5,467982477),("2025-10",186.2,199.3,170.8,197.6,393189757),
 ("2025-11",197.6,209.3,186.9,194.8,421025768),("2025-12",195,201.7,182.4,184.4,385588325),
 ("2026-01",184.5,248.4,184.4,244.8,484863425),("2026-02",239.9,239.9,212.5,218.5,591290887),
 ("2026-03",218.5,277.25,218.5,258.25,963583060),("2026-04",253.25,279.5,240.2,271,757445037),
 ("2026-05",270.75,274.75,232.2,236.2,353575981),("2026-06",236.2,252.75,215.1,227.5,487861759),
 ("2026-07",227.5,320.25,226,295.25,576627652),("2026-08",293,407.5,283,400.25,689874224),
 ("2026-09",404,417.25,390.5,392.25,63949705),
]
r += 1
mfirst = r
for d, o, h, l, c_, v in monthly:
    put(ws, r, 1, d, BODY, align="center")
    for i, x in enumerate([o, h, l, c_]):
        put(ws, r, 2 + i, x, BODY, TL2, align="right")
    put(ws, r, 6, v, BODY, '#,##0', align="right")
    if r > mfirst:
        put(ws, r, 7, f"=E{r}/E{r-1}-1", BODY, PCT, align="right")
    else:
        put(ws, r, 7, "", BODY, align="right")
    r += 1
mlast = r - 1
put(ws, r, 1, "12 aylık getiri", SUB, align="left")
put(ws, r, 5, f"=E{mlast}/E{mfirst}-1", SUB, PCT, fill=FILL_G, align="right")

# ---- freeze panes
for name in wb.sheetnames:
    s = wb[name]
    s.freeze_panes = "A5"
    s.sheet_view.showGridLines = False

wb.save("/home/user/komite-nabiz/analiz/TUPRAS_Analiz_2026-09-02.xlsx")
print("OK")
