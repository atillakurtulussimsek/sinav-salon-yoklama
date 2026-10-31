#!/usr/bin/env python3
"""Excel sınav yerleşim listesini salon bazlı yoklama listesi PDF'ine çevirir.

Kurulum : pip install openpyxl reportlab
Kullanım: python main.py liste.xlsx [cikti.pdf] [--ayar ayarlar.json]

Tüm görünüm ve içerik ayarları ayarlar.json dosyasından okunur.
"""
import json
import os
import re
import sys
from collections import OrderedDict
from datetime import datetime

from openpyxl import load_workbook
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (Flowable, KeepTogether, PageBreak, Paragraph,
                                SimpleDocTemplate, Spacer, Table, TableStyle)

VARSAYILAN_AYAR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ayarlar.json")

# Sütun tanımları: anahtar -> (göreli genişlik, hizalama)
SUTUN_TANIM = OrderedDict([
    ("koltuk", (10, TA_CENTER)),
    ("ogrenciNo", (12, TA_CENTER)),
    ("adSoyad", (30, TA_LEFT)),
    ("sinif", (20, TA_CENTER)),
    ("kitapcik", (10, TA_CENTER)),
    ("ogrenciTelefonu", (18, TA_CENTER)),
    ("veliAdi", (26, TA_LEFT)),
    ("veliTelefonu", (18, TA_CENTER)),
    ("imza", (22, TA_CENTER)),
])

FONT_ADAYLARI = [  # (normal, kalın) — Türkçe karakter için TTF gerekli
    ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
    ("/System/Library/Fonts/Supplemental/Arial.ttf", "/System/Library/Fonts/Supplemental/Arial Bold.ttf"),
    ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
]


# ───────────── yardımcılar ─────────────
def ayar_yukle(yol):
    with open(yol, encoding="utf-8") as f:
        return json.load(f)


def font_yukle():
    for normal, kalin in FONT_ADAYLARI:
        if os.path.exists(normal) and os.path.exists(kalin):
            pdfmetrics.registerFont(TTFont("F", normal))
            pdfmetrics.registerFont(TTFont("FB", kalin))
            pdfmetrics.registerFontFamily("F", normal="F", bold="FB", italic="F", boldItalic="FB")
            return "F", "FB"
    print("UYARI: TTF font bulunamadı, Türkçe karakterler bozuk çıkabilir.")
    return "Helvetica", "Helvetica-Bold"


def md(t):
    """Basit markdown -> ReportLab işaretlemesi (**kalın**, *italik*, `kod`)."""
    t = t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    t = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", t)
    t = re.sub(r"__(.+?)__", r"<b>\1</b>", t)
    t = re.sub(r"(?<!\*)\*(?!\s)(.+?)(?<!\s)\*(?!\*)", r"<i>\1</i>", t)
    t = re.sub(r"(?<!_)_(?!\s)(.+?)(?<!\s)_(?!_)", r"<i>\1</i>", t)
    t = re.sub(r"`(.+?)`", r"<font face='Courier'>\1</font>", t)
    return t


def metin(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def sayi(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return 10 ** 9


def sutun_bul(baslik, ad, zorunlu=True):
    for i, b in enumerate(baslik):
        if b.strip().lower() == ad.strip().lower():
            return i
    if zorunlu:
        sys.exit(f"Excel'de '{ad}' sütunu bulunamadı. Mevcut sütunlar: {baslik}")
    return None


# ───────────── veri ─────────────
def salonlari_oku(yol, ayar):
    """{salon: [ {alan: değer, ...}, ... ]} döndürür."""
    ex = ayar["excelSutunlari"]
    zorunlu = {"adSoyad", "ogrenciNo", "sinif", "salon", "sira", "sutun"}
    salonlar = OrderedDict()
    for ws in load_workbook(yol, data_only=True).worksheets:
        satirlar = [[metin(c) for c in r] for r in ws.iter_rows(values_only=True)]
        satirlar = [r for r in satirlar if any(r)]
        if len(satirlar) < 2:
            continue
        baslik, veri = satirlar[0], satirlar[1:]
        idx = {alan: sutun_bul(baslik, ad, alan in zorunlu) for alan, ad in ex.items()}
        for r in veri:
            kayit = {alan: (r[i] if i is not None and i < len(r) else "")
                     for alan, i in idx.items()}
            kayit["sira"], kayit["sutun"] = sayi(kayit["sira"]), sayi(kayit["sutun"])
            kayit["koltuk"] = f"{kayit['sira']}-{kayit['sutun']}"
            salon = kayit["salon"] or ws.title
            salonlar.setdefault(salon, []).append(kayit)

    siralama = ayar.get("siralama", "koltuk")
    for kayitlar in salonlar.values():
        if siralama in ("adSoyad", "ad"):
            kayitlar.sort(key=lambda k: k["adSoyad"])
        else:
            kayitlar.sort(key=lambda k: (k["sira"], k["sutun"]))
    return OrderedDict(sorted(salonlar.items()))


# ───────────── pdf ─────────────
class AltaYasla(Flowable):
    """Kalan boşluğu doldurur, içeriği sayfanın altına yaslar."""

    def __init__(self, icerik, bosluk=6 * mm):
        super().__init__()
        self.icerik, self.bosluk = icerik, bosluk

    def wrap(self, aw, ah):
        self.aw = aw
        self.w, self.h = self.icerik.wrap(aw, ah)
        gerekli = self.h + self.bosluk
        self.avail_h = (ah - 1) if ah >= gerekli else gerekli
        return aw, self.avail_h

    def draw(self):
        self.icerik.drawOn(self.canv, 0, 0, _sW=self.aw - self.w)


def sayfa_cizici(ayar, renk):
    sayfa = ayar["sayfa"]

    def ciz(canv, doc):
        canv.saveState()
        w, h = doc.pagesize
        if sayfa.get("cerceve", True):
            canv.setStrokeColor(renk["koyu"])
            canv.setLineWidth(1.8)
            canv.rect(8 * mm, 8 * mm, w - 16 * mm, h - 16 * mm)
        canv.setFont("F", 8)
        canv.setFillColor(colors.HexColor("#555555"))
        canv.setStrokeColor(renk["cizgi"])
        canv.setLineWidth(0.5)
        canv.line(doc.leftMargin, 13 * mm, w - doc.rightMargin, 13 * mm)
        if sayfa.get("altBilgi"):
            canv.drawString(doc.leftMargin, 9.5 * mm, sayfa["altBilgi"])
        if sayfa.get("sayfaNumarasi", True):
            canv.drawCentredString(w / 2, 9.5 * mm, f"Sayfa {doc.page}")
        if sayfa.get("olusturmaTarihi", True):
            canv.drawRightString(w - doc.rightMargin, 9.5 * mm,
                                 datetime.now().strftime("%d.%m.%Y %H:%M"))
        canv.restoreState()
    return ciz


def pdf_olustur(salonlar, cikti, ayar):
    f, fb = font_yukle()
    renk = {k: colors.HexColor(v) for k, v in ayar["renkler"].items()}
    boyut = ayar["sayfa"].get("yaziBoyutu", 9.5)
    kurum, sinav = ayar["kurum"], ayar["sinav"]
    bilgi_ayar = ayar.get("bilgiKutusu", {})
    uyari_ayar = ayar.get("uyarilar", {})
    imza_ayar = ayar.get("gozetmenImza", {})
    basliklar = ayar.get("sutunBasliklari", {})

    aktif = [k for k in SUTUN_TANIM if ayar["sutunlar"].get(k)]
    if not aktif:
        sys.exit("ayarlar.json: en az bir sütun açık olmalı.")

    st_kurum = ParagraphStyle("k", fontName=fb, fontSize=18, alignment=TA_CENTER,
                              textColor=renk["koyu"], spaceAfter=3, leading=22)
    st_sinav = ParagraphStyle("s", fontName=fb, fontSize=13, alignment=TA_CENTER,
                              textColor=colors.HexColor("#333333"), spaceAfter=8, leading=16)
    st_etiket = ParagraphStyle("e", fontName=fb, fontSize=10, leading=13)
    st_deger = ParagraphStyle("d", fontName=f, fontSize=10, leading=13)
    st_bas = ParagraphStyle("b", fontName=fb, fontSize=10, leading=12,
                            alignment=TA_CENTER, textColor=colors.white)
    st_hucre = {hiza: ParagraphStyle(f"h{hiza}", fontName=f, fontSize=boyut,
                                     leading=boyut + 2.5, alignment=hiza)
                for hiza in (TA_CENTER, TA_LEFT)}
    st_ub = ParagraphStyle("ub", fontName=fb, fontSize=11, leading=14,
                           textColor=renk["koyu"], spaceAfter=3)
    st_u = ParagraphStyle("u", fontName=f, fontSize=9.5, leading=13.5, alignment=TA_JUSTIFY,
                          leftIndent=12, bulletIndent=2, spaceAfter=2)
    st_imza = ParagraphStyle("i", fontName=fb, fontSize=10, leading=13)

    doc = SimpleDocTemplate(cikti, pagesize=A4,
                            title=f"{sinav['ad']} — {sinav.get('listeAdi', '')}",
                            leftMargin=15 * mm, rightMargin=15 * mm,
                            topMargin=15 * mm, bottomMargin=18 * mm)
    W = doc.width
    toplam = sum(SUTUN_TANIM[k][0] for k in aktif)
    genislik = [W * SUTUN_TANIM[k][0] / toplam for k in aktif]

    akis = []
    for n, (salon, ogr) in enumerate(salonlar.items()):
        if n:
            akis.append(PageBreak())
        alt_baslik = " — ".join(x for x in [sinav["ad"], sinav.get("listeAdi")] if x)
        akis += [Paragraph(kurum["ad"], st_kurum), Paragraph(alt_baslik, st_sinav)]

        # Bilgi kutusu
        if bilgi_ayar.get("goster", True):
            hucreler = []
            if bilgi_ayar.get("salon", True):
                hucreler.append(("Salon", salon))
            if bilgi_ayar.get("ogrenciSayisi", True):
                hucreler.append(("Öğrenci Sayısı", str(len(ogr))))
            if bilgi_ayar.get("tarihSaat", True):
                hucreler.append(("Tarih / Saat", f"{sinav.get('tarih', '')}  {sinav.get('saat', '')}".strip()))
            if bilgi_ayar.get("gozetmen", True):
                hucreler.append(("Gözetmen", ""))
            if hucreler:
                if len(hucreler) % 2:
                    hucreler.append(("", ""))
                satirlar = [[Paragraph(hucreler[i][0], st_etiket), Paragraph(hucreler[i][1], st_deger),
                             Paragraph(hucreler[i + 1][0], st_etiket), Paragraph(hucreler[i + 1][1], st_deger)]
                            for i in range(0, len(hucreler), 2)]
                bilgi = Table(satirlar, colWidths=[W * 0.17, W * 0.43, W * 0.20, W * 0.20])
                bilgi.setStyle(TableStyle([
                    ("BOX", (0, 0), (-1, -1), 1, renk["koyu"]),
                    ("INNERGRID", (0, 0), (-1, -1), 0.5, renk["cizgi"]),
                    ("BACKGROUND", (0, 0), (0, -1), renk["acik"]),
                    ("BACKGROUND", (2, 0), (2, -1), renk["acik"]),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ]))
                akis += [bilgi, Spacer(1, 6 * mm)]

        # Öğrenci tablosu
        veri = [[Paragraph(basliklar.get(k, k), st_bas) for k in aktif]]
        for kayit in ogr:
            veri.append([Paragraph(kayit.get(k, ""), st_hucre[SUTUN_TANIM[k][1]])
                         if k != "imza" else "" for k in aktif])
        tablo = Table(veri, colWidths=genislik, repeatRows=1)
        stil = [
            ("BOX", (0, 0), (-1, -1), 1, renk["koyu"]),
            ("INNERGRID", (0, 0), (-1, -1), 0.5, renk["cizgi"]),
            ("BACKGROUND", (0, 0), (-1, 0), renk["koyu"]),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ]
        if ayar["sayfa"].get("zebra", True):
            stil += [("BACKGROUND", (0, i), (-1, i), renk["zebra"]) for i in range(2, len(veri), 2)]
        tablo.setStyle(TableStyle(stil))
        akis += [tablo, Spacer(1, 6 * mm)]

        # Uyarı kutusu
        if uyari_ayar.get("goster", True) and uyari_ayar.get("maddeler"):
            uyari_ic = []
            if uyari_ayar.get("baslik"):
                uyari_ic.append(Paragraph(md(uyari_ayar["baslik"]), st_ub))
            uyari_ic += [Paragraph(md(u), st_u, bulletText="•") for u in uyari_ayar["maddeler"]]
            uyari = Table([[uyari_ic]], colWidths=[W])
            uyari.setStyle(TableStyle([
                ("BOX", (0, 0), (-1, -1), 0.8, renk["koyu"]),
                ("LINEBEFORE", (0, 0), (0, -1), 4, renk["koyu"]),
                ("BACKGROUND", (0, 0), (-1, -1), renk["acik"]),
                ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ("LEFTPADDING", (0, 0), (-1, -1), 12), ("RIGHTPADDING", (0, 0), (-1, -1), 12),
            ]))
            akis.append(KeepTogether(uyari))

        # Gözetmen imza bloğu (sağ alt)
        if imza_ayar.get("goster", True):
            satirlar = [[Paragraph(imza_ayar.get("baslik", "Salon Gözetmeni"), st_imza)]]
            for s in imza_ayar.get("satirlar", ["Adı Soyadı :", "İmza :"]):
                satirlar += [[Paragraph(s, st_deger)], [Spacer(1, 3 * mm)]]
            satirlar.append([Spacer(1, 4 * mm)])
            imza = Table(satirlar, colWidths=[W * 0.38], hAlign="RIGHT")
            imza.setStyle(TableStyle([
                ("BOX", (0, 0), (-1, -1), 0.8, renk["koyu"]),
                ("BACKGROUND", (0, 0), (-1, 0), renk["acik"]),
                ("LINEBELOW", (0, 0), (-1, 0), 0.5, renk["cizgi"]),
                ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ]))
            akis.append(AltaYasla(imza, bosluk=3 * mm))

    ciz = sayfa_cizici(ayar, renk)
    doc.build(akis, onFirstPage=ciz, onLaterPages=ciz)


# ───────────── giriş ─────────────
def argumanlar(argv):
    ayar_yolu, kalan = VARSAYILAN_AYAR, []
    i = 0
    while i < len(argv):
        if argv[i] == "--ayar":
            if i + 1 >= len(argv):
                sys.exit("--ayar için dosya yolu verilmedi.")
            ayar_yolu, i = argv[i + 1], i + 2
        else:
            kalan.append(argv[i])
            i += 1
    if not kalan:
        sys.exit(__doc__)
    girdi = kalan[0]
    cikti = kalan[1] if len(kalan) > 1 else os.path.splitext(girdi)[0] + ".pdf"
    return girdi, cikti, ayar_yolu


if __name__ == "__main__":
    girdi, cikti, ayar_yolu = argumanlar(sys.argv[1:])
    if not os.path.exists(ayar_yolu):
        sys.exit(f"Ayar dosyası bulunamadı: {ayar_yolu}")
    ayar = ayar_yukle(ayar_yolu)
    salonlar = salonlari_oku(girdi, ayar)
    if not salonlar:
        sys.exit("Excel dosyasında liste bulunamadı.")
    pdf_olustur(salonlar, cikti, ayar)
    print(f"{len(salonlar)} salon → {cikti}")
